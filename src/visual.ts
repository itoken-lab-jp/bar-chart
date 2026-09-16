/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import "./../style/visual.less";
import { App } from "./App";
import { VisualFormattingSettingsModel } from "./settings";
import { transform, tooltipStackOf, lineTooltipItems, ribbonTooltipItems, ViewModel, DataPoint } from "./viewModel";
import { tooltipItemsOf, toRootCoordinates } from "./tooltip";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ITooltipService = powerbi.extensibility.ITooltipService;
import ISelectionId = powerbi.visuals.ISelectionId;

export class Visual implements IVisual {
    private root: Root;
    private element: HTMLElement;
    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    /** このビジュアルで選ばれている棒。選択を変えたら描き直す */
    private selectedIds: ISelectionId[] = [];
    /** 最後の update の内容で描き直す（選択が変わったときに使う） */
    private renderLatest: (() => void) | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.element = options.element;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);

        // ブックマークの適用などで、選択が外から変わったとき
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            this.selectedIds = ids;
            this.renderLatest?.();
        });
    }

    public update(options: VisualUpdateOptions): void {
        // レンダリングイベントは認定要件。必ず started / finished(failed) を対で呼ぶ。
        this.events.renderingStarted(options);

        try {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                options.dataViews?.[0]
            );

            // グラフの種類による既定（100% 積み上げのデータラベルは値がオフで詳細がオン）は、保存していない項目だけに効かせる
            this.formattingSettings.applyChartTypeDefaults(options.dataViews?.[0]);
            const viewModel: ViewModel = transform(options.dataViews?.[0], this.host, this.formattingSettings);
            this.formattingSettings.applyTargets(viewModel.columnTargets, {
                seriesMode: viewModel.seriesMode,
                labelTargets: viewModel.labelTargets,
                lineTargets: viewModel.lineTargets,
            });
            this.formattingSettings.applyCardVisibility(viewModel.lines.length > 0);
            this.selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];

            const render = () =>
                this.root.render(
                    React.createElement(App, {
                        viewModel,
                        viewport: options.viewport,
                        settings: this.formattingSettings,
                        selectedIds: this.selectedIds,
                        onSelect: (id, multiSelect) => {
                            this.selectionManager.select(id, multiSelect).then((ids) => {
                                this.selectedIds = ids as ISelectionId[];
                                render();
                            });
                        },
                        onClearSelection: () => {
                            if (this.selectedIds.length === 0) return;
                            this.selectionManager.clear().then(() => {
                                this.selectedIds = [];
                                render();
                            });
                        },
                        onContextMenu: (id, x, y) => {
                            this.selectionManager.showContextMenu(id, { x, y });
                        },
                        onTooltipShow: (d, x, y) => this.showTooltip(viewModel, d, x, y, false),
                        onTooltipMove: (d, x, y) => this.showTooltip(viewModel, d, x, y, true),
                        onTooltipHide: () => this.tooltipService.hide({ isTouchEvent: false, immediately: false }),
                        onLineTooltipShow: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, false),
                        onLineTooltipMove: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, true),
                        onRibbonTooltipShow: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, false),
                        onRibbonTooltipMove: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, true),
                    })
                );
            this.renderLatest = render;
            render();

            this.events.renderingFinished(options);
        } catch (error) {
            console.error("update failed", error);
            this.events.renderingFailed(options, String(error));
        }
    }

    /**
     * 棒のツールヒント。標準と同じく カテゴリ・凡例（あれば）・値・追加フィールド を出し、棒の selectionId を渡す
     * （ドリルスルーやレポート ページのツールヒントが対象の行を知るため）
     */
    private showTooltip(viewModel: ViewModel, d: DataPoint, clientX: number, clientY: number, move: boolean): void {
        if (!viewModel.tooltip || !this.tooltipService.enabled()) return;
        const options = {
            coordinates: toRootCoordinates(clientX, clientY, this.element),
            isTouchEvent: false,
            dataItems: tooltipItemsOf(viewModel.tooltip, d.rowIndex, d.category, d.seriesIndex, tooltipStackOf(viewModel, d)),
            identities: [d.selectionId],
        };
        if (move) {
            this.tooltipService.move(options);
        } else {
            this.tooltipService.show(options);
        }
    }

    /** 折れ線の点のツールヒント。カテゴリ と 折れ線の値。点の selectionId（カテゴリ＋メジャー）を渡す */
    private showLineTooltip(viewModel: ViewModel, lineIndex: number, pointIndex: number, clientX: number, clientY: number, move: boolean): void {
        const point = viewModel.lines[lineIndex]?.points[pointIndex];
        if (!point || !this.tooltipService.enabled()) return;
        const options = {
            coordinates: toRootCoordinates(clientX, clientY, this.element),
            isTouchEvent: false,
            dataItems: lineTooltipItems(viewModel, lineIndex, pointIndex),
            identities: [point.selectionId],
        };
        if (move) {
            this.tooltipService.move(options);
        } else {
            this.tooltipService.show(options);
        }
    }

    /** リボンの帯のツールヒント。前後の値・変化・順位。帯の両端の棒の selectionId を渡す */
    private showRibbonTooltip(viewModel: ViewModel, seriesIndex: number, fromIndex: number, clientX: number, clientY: number, move: boolean): void {
        const a = viewModel.categoryGroups[fromIndex]?.points[seriesIndex];
        const b = viewModel.categoryGroups[fromIndex + 1]?.points[seriesIndex];
        if (!a || !b || !this.tooltipService.enabled()) return;
        const options = {
            coordinates: toRootCoordinates(clientX, clientY, this.element),
            isTouchEvent: false,
            dataItems: ribbonTooltipItems(viewModel, seriesIndex, fromIndex),
            identities: [a.selectionId, b.selectionId],
        };
        if (move) {
            this.tooltipService.move(options);
        } else {
            this.tooltipService.show(options);
        }
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        this.root?.unmount();
    }
}
