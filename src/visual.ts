/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { displayWarnings } from "./shared/warnings";
import { sentimentTextColors, showThemeToneColors } from "./shared/numberFormat";
import { syncToneColors } from "./shared/formatCards";

import "./../style/visual.less";
import { App } from "./App";
import { hasBrowserMenu } from "./shared/copyImage";
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
import { savedCumulativeReset, lineTooltipItems, ribbonTooltipItems, ViewModel, DataPoint, TRUNCATED_TITLE, TRUNCATED_NOTICE, SERIES_TRUNCATED_TITLE, SERIES_TRUNCATED_NOTICE, LINE_RATIO_WARNING_TITLE, VALUE_LINE_WARNING_TITLE } from "./viewModel";
import { showTooltipAt } from "./shared/tooltip";
import { GuardedRenderer } from "./shared/renderGuard";
import { selectionLookup } from "./shared/selectionLookup";
import { contextMenuIdOf } from "./shared/contextMenu";
import { transformWithLayers, nextVisibleLayers, labelRoomChanges, barTooltipItems, COMPARE_LIMIT_TITLE, COMPARE_LINE_ONLY_TITLE } from "./compareLayers";
import { receivedCounts, ROW_LIMIT, SERIES_LIMIT } from "./matrixDataView";
import { pickedLayersOf } from "./layout";
import { MAX_AXIS_REBUILDS, initialValueAxisLength, needsAxisRebuild } from "./axisLength";

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
    /** 描画の守り。update では描き終えてからレンダリング イベントを送り、失敗したら前の絵を消して知らせを出す */
    private renderer: GuardedRenderer;
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
     * この update で、値の軸の長さで組み直した回数。長さが 2 つの値を行き来しても描き直しが止まるよう、
     * MAX_AXIS_REBUILDS までにする（ふつうは 1 回で落ち着く）。update と、閲覧者の操作で作り直すとき（凡例のレイヤー・累計）に数え直す
     */
    private axisRebuilds = 0;
    /**
     * 上限に達して切られたか。行（カテゴリ）は top 30000・列（凡例）は top 2000 で受け、届いた数が上限に達したら切られたとみなす。
     * 行を window にして続きを読むと、凡例が 60 で切られ、列の小計（全体の値）も届かなくなったため。update() でだけ決める
     */
    private truncated = { rows: false, series: false };
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
    /**
     * 凡例で選んだ比較レイヤー（compareLayers の添字）。空ならすべて描く。選んだレイヤーだけを並べ直して描く
     * （通常の選択とは別に持ち、ほかのビジュアルを絞り込まない）。保存はしない
     */
    private visibleLayers: number[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.element = options.element;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);
        this.renderer = new GuardedRenderer(this.root);
        this.allowInteractions = options.host.hostCapabilities?.allowInteractions !== false;

        // ブックマークの適用などで、選択が外から変わったとき
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            this.selectedIds = ids;
            this.renderLatest?.();
        });
    }

    public update(options: VisualUpdateOptions): void {
        // レンダリングイベントは認定要件。started と、描き終えてから finished（失敗なら failed）を対で 1 回ずつ送る。
        // 値の軸の長さが決まってからの描き直し（onValueAxisLength）も、同じ update の中で描き終えてから送る
        const rendered = this.renderer.update(this.events, options, () => {
            this.lastOptions = options;
            this.axisRebuilds = 0;
            const counts = receivedCounts(options.dataViews?.[0]);
            this.truncated = { rows: counts.rows >= ROW_LIMIT, series: counts.series >= SERIES_LIMIT };
            this.restoreVisualState(options.dataViews?.[0]);
            this.build(options);
        });
        // 描けなかったら、選択が外から変わっても前の内容で描き直さない（消した前の絵が戻る）
        if (!rendered) this.renderLatest = null;
    }

    /**
     * 書式と閲覧者の操作から viewModel を作り、描く。累計は transform の中で効くので、切り替えたら作り直す。
     * 値の軸の長さ axisLength は、描いて知らせた長さで組み直すときだけ渡す。update と閲覧者の操作で作り直すときは、入力（ビューの大きさと書式）から
     * 見積もった長さで組む（initialValueAxisLength。前の update で描いた長さは使わないので、同じ入力なら、それまでの大きさの順によらず同じ絵になる）
     */
    private build(options: VisualUpdateOptions, axisLength?: number): void {
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
        const viewModel: ViewModel = transformWithLayers(options.dataViews?.[0], this.host, this.formattingSettings, {
            cumulative: effectiveCumulative(current, baseCumulative),
            cumulativeReset: current.cumulativeReset,
            valueAxisLength: axisLength ?? initialValueAxisLength(options.viewport, this.formattingSettings),
            visibleLayers: this.visibleLayers,
        });
        // 警告は描き直すたびに消えるので、そのたびに出し直す。アイコンは 1 つしか出せないので、2 つあれば本文を並べる
        displayWarnings(this.host, [
            this.truncated.rows && { title: TRUNCATED_TITLE, detail: TRUNCATED_NOTICE },
            this.truncated.series && { title: SERIES_TRUNCATED_TITLE, detail: SERIES_TRUNCATED_NOTICE },
            viewModel.lineWarning && { title: LINE_RATIO_WARNING_TITLE, detail: viewModel.lineWarning },
            viewModel.compareWarning && { title: COMPARE_LIMIT_TITLE, detail: viewModel.compareWarning },
            viewModel.compareLineOnlyWarning && { title: COMPARE_LINE_ONLY_TITLE, detail: viewModel.compareLineOnlyWarning },
            viewModel.valueLineWarning && { title: VALUE_LINE_WARNING_TITLE, detail: viewModel.valueLineWarning },
        ]);
        // 書式ペインの区切りは書式の値のまま出す（閲覧者の選んだ区切りは書式を書き換えない）
        calc.applyCumulativeLevels(viewModel.cumulative.levels, authorReset);
        this.formattingSettings.lines.includeCumulative.visible = viewModel.cumulative.available;
        // X 軸の「種類」は、連続で描けるとき（X 軸が日付か数値の 1 列で縦棒など。viewModel の continuousAvailable）だけ出す
        this.formattingSettings.categoryAxis.applyAxisKind(!!viewModel.continuousAvailable);
        // 縦のグリッド線の「線の位置」に、X 軸のレベルの名前を並べる
        this.formattingSettings.gridlines.applyLevels(viewModel.categoryAxis.levelNames ?? [], viewModel.gridlines.verticalPosition);
        this.formattingSettings.applyTargets(viewModel.columnTargets, {
            seriesMode: viewModel.seriesMode,
            labelTargets: viewModel.labelTargets,
            lineTargets: viewModel.lineTargets,
        });
        this.formattingSettings.applySingleSeriesFill(viewModel.seriesMode, viewModel.columns.fill, options.dataViews?.[0]?.metadata?.objects);
        // 単位ラベルの文字サイズ・カラーは、保存していなければ目盛りと同じものを書式ペインにも出す（保存はしない）
        this.formattingSettings.showUnitDefaults();
        // 符号の色は、書式で決めていなければテーマから作った色を書式ペインにも出す（保存はしない）
        showThemeToneColors(
            this.formattingSettings.dataLabels,
            options.dataViews?.[0]?.metadata?.objects?.dataLabels,
            sentimentTextColors(this.host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette)
        );
        // 出し分けは描画に使ったのと同じ条件で（パレートの累積比の線は「折れ線の値」に数えない、第 2 Y 軸は描くときだけ、など）
        this.formattingSettings.applyCardVisibility(viewModel.pane ?? false, false, viewModel.measureGroups ?? false);
        this.formattingSettings.legend.applySeries(viewModel.seriesMode);
        // 「比較値」「比較の列」が無ければ「比較」のカードは出さない。「手前にする値」は「比較の列」のときだけ
        this.formattingSettings.compare.visible = hasRole(options.dataViews?.[0], "compare") || hasRole(options.dataViews?.[0], "compareBy");
        this.formattingSettings.compare.applyCompareBy(hasRole(options.dataViews?.[0], "compareBy"));
        // 「X 軸の定数線」の欄が空なら、そのカードは出さない
        this.formattingSettings.categoryLine.visible = hasRole(options.dataViews?.[0], "categoryLine");
        // 比較レイヤーが変わったら（枚数が減る・比較を外す）、凡例で選んだレイヤーは外す（組むときも、外の番号のある選択は選択なしとして組んだ）
        this.visibleLayers = pickedLayersOf(viewModel.compareLayers.length, this.visibleLayers);
        this.selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];

        const render = () =>
            this.renderer.paint(
                React.createElement(App, {
                    viewModel,
                    viewport: options.viewport,
                    settings: this.formattingSettings,
                    // 描いた値の軸の長さで、棒の外に出すラベルの余白（長さが 1px 以上変わったとき）と範囲の丸め（長さで決まる本数が変わったとき）を
                    // 決め直して描き直す（needsAxisRebuild）。ふつうは 1 回で落ち着くが、軸の範囲で目盛りの文字数・X 軸の名前の回し方が変わると
                    // 長さが行き来しうるので、組み直しは MAX_AXIS_REBUILDS 回まで
                    onValueAxisLength: (length) => {
                        if (this.lastOptions !== options || !needsAxisRebuild(viewModel, length)) return;
                        if (this.axisRebuilds >= MAX_AXIS_REBUILDS) return;
                        this.axisRebuilds++;
                        this.build(options, length);
                    },
                    selectedIds: this.selectedIds,
                    // 操作できない場所（ダッシュボードのタイルなど）では、選択・右クリックのメニューを送らない
                    onSelect: (id, multiSelect) => {
                        if (!this.allowInteractions) return;
                        this.selectionManager.select(id, multiSelect).then((ids) => {
                            this.selectedIds = ids as ISelectionId[];
                            // 選択が返るまでに次の update が来ていれば、そのときの最新の内容で描く（失敗していれば描かない）
                            this.renderLatest?.();
                        });
                    },
                    onSelectMany: (ids, multiSelect) => {
                        if (!this.allowInteractions || ids.length === 0) return;
                        // 同じランクをもう一度押したら外す（旧パレート図と同じ）
                        const current = this.selectedIds;
                        // 選んだ ID は行の数だけあることがある。1 つずつ比べず、選んだ ID を引いて比べる
                        const inCurrent = selectionLookup(current);
                        const same = !multiSelect && current.length === ids.length && ids.every((id) => inCurrent(id));
                        const action = same ? this.selectionManager.clear().then(() => []) : this.selectionManager.select(ids, multiSelect);
                        action.then((selected) => {
                            this.selectedIds = selected as ISelectionId[];
                            // 選択が返るまでに次の update が来ていれば、そのときの最新の内容で描く（失敗していれば描かない）
                            this.renderLatest?.();
                        });
                    },
                    onClearSelection: () => {
                        if (!this.allowInteractions || this.selectedIds.length === 0) return;
                        this.selectionManager.clear().then(() => {
                            this.selectedIds = [];
                            // 選択が返るまでに次の update が来ていれば、そのときの最新の内容で描く（失敗していれば描かない）
                            this.renderLatest?.();
                        });
                    },
                    onContextMenu: (ids, x, y) => {
                        if (!this.allowInteractions) return;
                        // 1 行の棒はその行、まとめた棒（その他）と背景はビジュアルのメニュー
                        this.selectionManager.showContextMenu(contextMenuIdOf(ids), { x, y });
                    },
                    onTooltipShow: (d, x, y) => this.showTooltip(viewModel, d, x, y, false),
                    onTooltipMove: (d, x, y) => this.showTooltip(viewModel, d, x, y, true),
                    onTooltipHide: () => this.tooltipService.hide({ isTouchEvent: false, immediately: false }),
                    onLineTooltipShow: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, false),
                    onLineTooltipMove: (j, i, x, y) => this.showLineTooltip(viewModel, j, i, x, y, true),
                    onRibbonTooltipShow: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, false),
                    onRibbonTooltipMove: (s, i, x, y) => this.showRibbonTooltip(viewModel, s, i, x, y, true),
                    interactive: this.allowInteractions,
                    editing: isEditing(options),
                    visibleLayers: this.visibleLayers,
                    onSelectLayer: (layer, multiSelect) => {
                        if (!this.allowInteractions) return;
                        const next = nextVisibleLayers(this.visibleLayers, layer, multiSelect);
                        // データ ラベルを描くレイヤーが変わって値の軸の余白が変わるなら、組み直す
                        const rebuild = labelRoomChanges(viewModel, this.visibleLayers, next) && this.lastOptions === options;
                        this.visibleLayers = next;
                        if (rebuild) this.rebuildFromViewer(options);
                        else render();
                    },
                    browserMenu: hasBrowserMenu(this.host.hostEnv),
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
        this.rebuildFromViewer(this.lastOptions);
    }

    /** 閲覧者の操作で作り直す。update と同じく、見積もった値の軸の長さから組み、組み直しの回数も数え直す */
    private rebuildFromViewer(options: VisualUpdateOptions): void {
        this.axisRebuilds = 0;
        this.build(options);
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
        // 比較レイヤーの奥の棒は、そのレイヤーの値で出す。
        // 「その他」は、まとめたカテゴリの ID を全部渡す（系列があればその系列の ID。ドリルスルーやレポート ページのツールヒントの対象）
        showTooltipAt(this.tooltipService, this.element, clientX, clientY, move, barTooltipItems(viewModel, d), d.selectionIds ?? [d.selectionId]);
    }

    /**
     * 折れ線の点のツールヒント。カテゴリ と 折れ線の値。点の selectionId（カテゴリ＋メジャー）を渡す。
     * 「その他」の点は、まとめた行の線の点の ID を全部渡す（ドリルスルーやレポート ページのツールヒントの対象）
     */
    private showLineTooltip(viewModel: ViewModel, lineIndex: number, pointIndex: number, clientX: number, clientY: number, move: boolean): void {
        const point = viewModel.lines[lineIndex]?.points[pointIndex];
        if (!point) return;
        showTooltipAt(this.tooltipService, this.element, clientX, clientY, move, lineTooltipItems(viewModel, lineIndex, pointIndex), point.selectionIds ?? [point.selectionId]);
    }

    /** リボンの帯のツールヒント。前後の値・変化・順位。帯の両端の棒の ID を渡す（「その他」の棒は、まとめたカテゴリの帯の系列の ID を全部） */
    private showRibbonTooltip(viewModel: ViewModel, seriesIndex: number, fromIndex: number, clientX: number, clientY: number, move: boolean): void {
        const a = viewModel.categoryGroups[fromIndex]?.points[seriesIndex];
        const b = viewModel.categoryGroups[fromIndex + 1]?.points[seriesIndex];
        if (!a || !b) return;
        const ids = [...(a.selectionIds ?? [a.selectionId]), ...(b.selectionIds ?? [b.selectionId])];
        showTooltipAt(this.tooltipService, this.element, clientX, clientY, move, ribbonTooltipItems(viewModel, seriesIndex, fromIndex), ids);
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        // 使わない符号の色は出さない（触っても何も起きない）
        syncToneColors(this.formattingSettings.dataLabels);
        return this.formattingSettings.formattingModel(this.formattingSettingsService);
    }

    public destroy(): void {
        this.root?.unmount();
    }
}

/** その欄にフィールドが入っているか（メタデータの列の役割で見る） */
function hasRole(dataView: powerbi.DataView | undefined, role: string): boolean {
    return (dataView?.metadata?.columns ?? []).some((column) => column.roles?.[role]);
}

/** レポートを編集中か。viewMode が View（0）なら閲覧中。渡されなければ編集中として扱う（今までと同じく案内を出す） */
function isEditing(options: VisualUpdateOptions): boolean {
    return options.viewMode === undefined || options.viewMode !== 0;
}
