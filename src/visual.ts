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
import { VisualFormattingSettingsModel, CALCULATION_MODES } from "./settings";
import {
    VisualState,
    EMPTY_VISUAL_STATE,
    VISUAL_STATE_OBJECT,
    readVisualState,
    serializeVisualState,
    toPersistedProperties,
    effectiveCumulative,
    withoutStale,
} from "./visualState";

/** 保存の応答（update）を待つ上限 (ms)。過ぎたら、保存に失敗したものとして読み戻しを受け付ける */
const PENDING_TIMEOUT_MS = 5000;
import { transform, savedCumulativeReset, tooltipStackOf, lineTooltipItems, ribbonTooltipItems, ViewModel, DataPoint, LOADING_NOTICE, TRUNCATED_TITLE, TRUNCATED_NOTICE } from "./viewModel";
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
    /** 最後の update（閲覧者が累計を切り替えたときに、同じ内容で作り直す） */
    private lastOptions: VisualUpdateOptions | null = null;
    /**
     * 続きの読み込み。カテゴリが 30,000 件を超えると、Power BI は続きを残して届ける（metadata.segment）。
     * update() でだけ決める（閲覧者の操作で描き直すときに、もう一度読みに行かない）
     */
    private loadState: "complete" | "loading" | "truncated" = "complete";
    /** 閲覧者が触った累計の状態。ページを移ってもレポートから読み直す */
    private visualState: VisualState = EMPTY_VISUAL_STATE;
    /** persistProperties 直後の値。保存が返る前の古い dataView で操作を巻き戻さないための印 */
    private pendingVisualState: string | null = null;
    /** pendingVisualState を立てた時刻 */
    private pendingAt = 0;
    /** 同じ状態を何度も当て直さないための印 */
    private lastRestoredVisualState: string | null = null;
    /** 操作を受け付けるか（ダッシュボードのタイルに固定すると false）。false なら切り替えボタンを出さない */
    private allowInteractions = true;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.element = options.element;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);
        this.allowInteractions = options.host.hostCapabilities?.allowInteractions !== false;

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
            this.lastOptions = options;
            // 続きがあれば全部読む（「その他」や積み上げの合計を、読み込めた分だけで計算しないため）。
            // 読めない（100 MB の上限など）ときは、読めた分で描いて警告を出す
            this.loadState = !options.dataViews?.[0]?.metadata?.segment
                ? "complete"
                : this.host.fetchMoreData(true)
                  ? "loading"
                  : "truncated";
            this.restoreVisualState(options.dataViews?.[0]);
            this.build(options);
            this.events.renderingFinished(options);
        } catch (error) {
            console.error("update failed", error);
            this.events.renderingFailed(options, String(error));
        }
    }

    /** 書式と閲覧者の操作から viewModel を作り、描く。累計は transform の中で効くので、切り替えたら作り直す */
    private build(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            options.dataViews?.[0]
        );

        // 軸のタイトル・凡例のタイトルと位置は、保存が無ければテーマ（基本テーマの Fluent 2 など）の値に従う
        this.formattingSettings.applyThemeDefaults(options.dataViews?.[0]?.metadata?.objects);
        // グラフの種類による既定（100% 積み上げのデータラベルは値がオフで詳細がオン）は、保存していない項目だけに効かせる
        this.formattingSettings.applyChartTypeDefaults(options.dataViews?.[0]);
        const calc = this.formattingSettings.calculation;
        const baseCumulative = String(calc.mode.value?.value ?? CALCULATION_MODES.none) === CALCULATION_MODES.cumulative;
        const authorReset = savedCumulativeReset(options.dataViews?.[0], calc);
        // 作成者が書式を変えて古くなった閲覧者の操作は捨て、捨てた状態を保存し直す（保存が返す update では、もう捨てるものが無いので輪にならない）
        const current = withoutStale(this.visualState, baseCumulative, authorReset);
        if (current !== this.visualState) this.persistVisualState(current);
        const viewModel: ViewModel = transform(options.dataViews?.[0], this.host, this.formattingSettings, {
            cumulative: effectiveCumulative(current, baseCumulative),
            cumulativeReset: current.cumulativeReset,
        });
        if (this.loadState === "loading") viewModel.notice = LOADING_NOTICE;
        // 警告は描き直すたびに消えるので、そのたびに出し直す
        if (this.loadState === "truncated") this.host.displayWarningIcon(TRUNCATED_TITLE, TRUNCATED_NOTICE);
        // 書式ペインの区切りは書式の値のまま出す（閲覧者の選んだ区切りは書式を書き換えない）
        calc.applyCumulativeLevels(viewModel.cumulative.levels, authorReset);
        this.formattingSettings.lines.includeCumulative.visible = viewModel.cumulative.available;
        this.formattingSettings.applyTargets(viewModel.columnTargets, {
            seriesMode: viewModel.seriesMode,
            labelTargets: viewModel.labelTargets,
            lineTargets: viewModel.lineTargets,
        });
        this.formattingSettings.applySingleSeriesFill(viewModel.seriesMode, viewModel.columns.fill, options.dataViews?.[0]?.metadata?.objects);
        this.formattingSettings.applyCardVisibility(viewModel.lines.length > 0);
        this.selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];

        const render = () =>
            this.root.render(
                React.createElement(App, {
                    viewModel,
                    viewport: options.viewport,
                    settings: this.formattingSettings,
                    selectedIds: this.selectedIds,
                    // 操作できない場所（ダッシュボードのタイルなど）では、選択・右クリックのメニューを送らない
                    onSelect: (id, multiSelect) => {
                        if (!this.allowInteractions) return;
                        this.selectionManager.select(id, multiSelect).then((ids) => {
                            this.selectedIds = ids as ISelectionId[];
                            render();
                        });
                    },
                    onSelectMany: (ids, multiSelect) => {
                        if (!this.allowInteractions || ids.length === 0) return;
                        // 同じランクをもう一度押したら外す（旧パレート図と同じ）
                        const current = this.selectedIds;
                        const same = !multiSelect && current.length === ids.length && ids.every((id) => current.some((c) => c.equals(id)));
                        const action = same ? this.selectionManager.clear().then(() => []) : this.selectionManager.select(ids, multiSelect);
                        action.then((selected) => {
                            this.selectedIds = selected as ISelectionId[];
                            render();
                        });
                    },
                    onClearSelection: () => {
                        if (!this.allowInteractions || this.selectedIds.length === 0) return;
                        this.selectionManager.clear().then(() => {
                            this.selectedIds = [];
                            render();
                        });
                    },
                    onContextMenu: (id, x, y) => {
                        if (!this.allowInteractions) return;
                        this.selectionManager.showContextMenu(id, { x, y });
                    },
                    onTooltipShow: (d, x, y) => this.showTooltip(viewModel, d, x, y, false),
                    onTooltipMove: (d, x, y) => this.showTooltip(viewModel, d, x, y, true),
                    onTooltipHide: () => this.tooltipService.hide({ isTouchEvent: false, immediately: false }),
                    onLineTooltipShow: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, false),
                    onLineTooltipMove: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, true),
                    onRibbonTooltipShow: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, false),
                    onRibbonTooltipMove: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, true),
                    interactive: this.allowInteractions,
                    onToggleCumulative: () =>
                        this.changeVisualState({
                            ...this.visualState,
                            cumulative: !viewModel.cumulative.enabled,
                            cumulativeBase: baseCumulative,
                        }),
                    onChangeCumulativeReset: (reset) =>
                        this.changeVisualState({ ...this.visualState, cumulativeReset: reset, cumulativeResetBase: authorReset }),
                })
            );
        this.renderLatest = render;
        render();
    }

    /** 閲覧者の操作を保存して作り直す。update() からは呼ばない（保存が update を呼び、また保存する輪になる） */
    private changeVisualState(next: VisualState): void {
        if (!this.allowInteractions || !this.lastOptions) return;
        this.persistVisualState(next);
        this.build(this.lastOptions);
    }

    /** 閲覧者の操作をこのセッションに持ち、レポートに保存する */
    private persistVisualState(next: VisualState): void {
        this.visualState = next;
        this.pendingVisualState = serializeVisualState(next);
        this.pendingAt = Date.now();
        try {
            this.host.persistProperties({
                merge: [{ objectName: VISUAL_STATE_OBJECT, properties: toPersistedProperties(next), selector: null }],
            });
        } catch (error) {
            // 保存に失敗しても、このセッションの見た目は保つ
            this.pendingVisualState = null;
            console.error("累計の状態を保存できませんでした", error);
        }
    }

    /** 保存済みの閲覧者の操作を読み戻す */
    private restoreVisualState(dataView: powerbi.DataView | undefined): void {
        // 保存の応答が来ないまま時間が過ぎたら、保存に失敗したものとして待つのをやめる（ブックマークなどの読み戻しを拒み続けない）
        if (this.pendingVisualState !== null && Date.now() - this.pendingAt > PENDING_TIMEOUT_MS) this.pendingVisualState = null;
        const persisted = readVisualState(dataView);
        if (!persisted) {
            // 読み戻した保存が消えた（保存の無いブックマークに切り替えたなど）なら、閲覧者の操作は無い状態に戻す。
            // 一度も読み戻していない（保存に失敗して、このセッションにだけある操作）なら残す
            if (this.pendingVisualState === null && this.lastRestoredVisualState !== null) {
                this.visualState = EMPTY_VISUAL_STATE;
                this.lastRestoredVisualState = null;
            }
            return;
        }
        // persistProperties の直後に古い dataView が来ても、押したばかりの操作を戻さない
        if (this.pendingVisualState !== null && persisted.raw !== this.pendingVisualState) return;
        if (this.pendingVisualState === persisted.raw) this.pendingVisualState = null;
        if (this.lastRestoredVisualState === persisted.raw) return;
        this.visualState = persisted.state;
        this.lastRestoredVisualState = persisted.raw;
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
            // 「その他」は、まとめたカテゴリの ID を全部渡す（ドリルスルーやレポート ページのツールヒントの対象）
            identities: d.selectionIds ?? [d.selectionId],
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
