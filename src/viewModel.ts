"use strict";

import powerbi from "powerbi-visuals-api";
import { scaleLinear } from "d3-scale";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import DataViewObjects = powerbi.DataViewObjects;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import {
    VisualFormattingSettingsModel,
    ColumnTarget,
    LabelTarget,
    LineTarget,
    CHART_TYPES,
    ChartType,
    RIBBON_ORDERS,
    DETAIL_CONTENTS,
    LABEL_LAYOUTS,
    TITLE_CONTENTS,
    LINE_STYLES,
    LINE_SHAPE_DEFAULTS,
    ORIENTATIONS,
    Orientation,
    CALCULATION_MODES,
    CUMULATIVE_RESET_NONE,
    LINE_FORMAT_MODES,
    LINE_PERCENT_FORMAT,
    LinesCardSettings,
    legendPlacementValue,
} from "./settings";
import {
    resolveUnit,
    formatValue,
    formatDynamicValue,
    dynamicShownSign,
    resolveBadgeText,
    composeUnitText,
    UnitDefinition,
    UNIT_DEFINITIONS,
} from "./unitUtils";
import { collapseOthers, lineIndependenceOf } from "./others";
import { categoricalOf, withCategoryOf, withSeriesOf } from "./matrixDataView";
import { TooltipSource, TooltipStack, TooltipColumn, tooltipColumnOf, formatTooltipValue, categoryTooltipRows, BLANK_TEXT } from "./tooltip";
import { ticksUpTo, tickCountOf, boundOf } from "./shared/ticks";
import { PT_TO_PX } from "./shared/text";
import { Halo, haloOf, NO_HALO } from "./shared/halo";
import { formatSigned, shownSignOf, toneOf, NEGATIVE_STYLES, SignStyle, ZERO_STYLES, TONE_MODES, DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR, isPercentFormat, percentDigitsOf, formatPercent } from "./shared/numberFormat";

export { tickCountOf };

export interface DataPoint {
    category: string;
    /** DataView の行番号。並べ替え後もツールヒント等で元の行を引くのに使う */
    rowIndex: number;
    /** 系列の番号（ViewModel.series の添字）。系列が 1 本なら 0 */
    seriesIndex: number;
    /** 値が空白（複数系列で、そのカテゴリにその系列の値が無い）。棒とラベルを描かない */
    blank: boolean;
    value: number;
    /** 棒の始まりの位置（軸の比率）。集合なら 0 の位置、積み上げなら下（負なら上）に積んだ棒の端 */
    startRatio: number;
    /** 棒の終わりの位置（軸の比率） */
    valRatio: number;
    /** 100% 積み上げでの割合（-1〜1、正と負の絶対値の合計に対する比）。それ以外は null */
    share: number | null;
    /** 積み上げの外側の端の棒（角丸を付ける棒）。集合ではすべて true */
    outermost: boolean;
    /** 積み上げで、同じ側（正・負）の下に別の棒がある。描くとき、始まりの側を系列間のスペース（px）だけ縮める */
    insetStart: boolean;
    /** リボンでの順位（カテゴリの中で値の大きい順、1 が最大）。リボン以外は null */
    rank: number | null;
    formattedValue: string;
    dataLabelText: string;
    /** 値の行を符号で塗る色（「符号の色」）。塗らなければ空 */
    labelToneColor: string;
    /** データラベルのタイトルの行（系列名か、ラベルのタイトルのフィールドの値）。無ければ空 */
    titleText: string;
    /** データラベルの詳細の行（全体に対する割合か、ラベルの詳細のフィールドの値）。無ければ空 */
    detailText: string;
    selectionId: ISelectionId;
    /**
     * カテゴリだけの ID。パレートのランクの帯はカテゴリで選ぶので、系列があるときに棒の ID（カテゴリ＋系列）と
     * 一致しない。選ばれているかを見るときはこちらとも比べる。系列 1 本なら selectionId と同じ
     */
    categorySelectionId?: ISelectionId;
    /** 「その他」の棒だけ持つ。押したらまとめたカテゴリを全部選ぶ */
    selectionIds?: ISelectionId[];
    /** ハイライトの値。ハイライトが無い・この棒が該当しないときは null */
    highlight: number | null;
    /** highlight を valRatio と同じ軸で表した比率。highlight が null なら null */
    highlightRatio: number | null;
    color: string;
    transparency: number;
    borderShow: boolean;
    borderColor: string;
    borderTransparency: number;
    borderWidth: number;
    /** データラベルを出すか（系列ごとの「このシリーズに表示」）。カード全体の表示とは別に効く */
    labelShow: boolean;
    /** データラベルの文字色。空 = 自動（棒の色に合わせて白か黒） */
    labelColor: string;
    /** 比較レイヤーの番号（0 が手前の「値」、1 から奥の「比較値」）。比較が無ければ無し */
    layerIndex?: number;
}

/** 系列 1 本ぶん（凡例の 1 項目） */
export interface SeriesInfo {
    name: string;
    color: string;
    /** 凡例のクリックで選ぶ系列の ID。系列が 1 本のときは null */
    selectionId: ISelectionId | null;
}

/** 積み上げのカテゴリの合計（合計ラベルとツールヒント用）。空白は数えない */
export interface StackTotals {
    /** 正と負を足した合計 */
    net: number;
    positive: number;
    negative: number;
    hasPositive: boolean;
    hasNegative: boolean;
    /** 正の棒の上の端・負の棒の下の端（軸の比率） */
    positiveEndRatio: number;
    negativeEndRatio: number;
    netText: string;
    positiveText: string;
    negativeText: string;
}

/** カテゴリ 1 つぶんの棒（系列の順＝凡例の順） */
export interface CategoryGroup {
    /** 名前（ツールヒント・読み上げ・書式の適用先）。階層は上のレベルから全部つなぐ（ドリルした共通の親も入れる） */
    category: string;
    /** 軸に 1 行で出す名前（「ラベルの連結」）。ドリルした共通の親は外す */
    label: string;
    /** 階層のレベルごとの表示（上のレベルから）。1 列なら [category] */
    levels: string[];
    /** レベルごとの元の値を見分けるキー。親の区切りと並べ替えは表示ではなくこれで見る */
    levelKeys: string[];
    rowIndex: number;
    points: DataPoint[];
    /** 積み上げのときの合計。集合・100% 積み上げでは null */
    totals: StackTotals | null;
    /** 比較レイヤーの奥の棒（[0] が 2 枚目＝「比較値」の 1 つ目）。並びは points と同じ系列の順。比較が無ければ無し */
    layerPoints?: DataPoint[][];
}

/** 折れ線の点 1 つ。並びは categoryGroups と同じ（並べ替えのあと） */
export interface LinePoint {
    rowIndex: number;
    /** 「その他」の点だけ持つ。押したらまとめたカテゴリを全部選ぶ */
    selectionIds?: ISelectionId[];
    /** この点の手前で線を切るか（累計の区切りで 0 に戻るところ） */
    breakBefore?: boolean;
    /** 値。空白なら null（線を切る） */
    value: number | null;
    /** 軸の比率。第 2 Y 軸で描くときは第 2 Y 軸の比率 */
    ratio: number | null;
    selectionId: ISelectionId;
}

/** 折れ線 1 本（「折れ線の値」に入れたメジャー 1 つ） */
export interface LineSeriesInfo {
    name: string;
    color: string;
    width: number;
    lineStyle: string;
    /** 線のスタイルがカスタムのときの模様（ダッシュ配列）と端の形（ダッシュ キャップ） */
    dashArray: string;
    dashCap: string;
    /** 破線・点線・カスタムの模様を線の幅に比例させるか */
    scaleWithWidth: boolean;
    /** 線の透過性 (%)。マーカー・網掛け領域には効かない */
    transparency: number;
    /** 結合の種類（"miter" | "round" | "bevel"） */
    lineJoin: string;
    /** 補間の種類（"linear" | "smooth" | "step"） */
    interpolation: string;
    /** スムーズの種類（"monotone" | "cardinal"） */
    smoothing: string;
    /** カーディナルのテンション (%) */
    tension: number;
    /** ステップの位置（"before" | "center" | "after"） */
    stepPosition: string;
    /** ステップの段と段をつなぐ線を出すか */
    stepConnect: boolean;
    /** 段のつなぎを出さないときの線の長さ（"step" カテゴリの間隔 | "bar" 棒の幅） */
    stepWidth: string;
    /** 網掛け領域を出すか（カードの表示と、線ごとの「このシリーズに表示」） */
    areaShow: boolean;
    /** 線（点をつなぐ線）を出すか。線ごとの「このシリーズに表示」、無ければ「すべての系列に表示」。マーカーは別 */
    lineShow: boolean;
    /** 網掛け領域の下の辺（値 0 の高さ。軸の範囲の外なら端）の軸の比率 */
    baselineRatio: number;
    /** 凡例のクリックで選ぶ ID（メジャー） */
    selectionId: ISelectionId;
    /**
     * 凡例・線のクリックで線ごと選べるか。パレートの累積比の線（ビジュアルが計算した線でメジャーが無い）は false。
     * 省略は true
     */
    selectable?: boolean;
    points: LinePoint[];
    /** ツールヒント用。values は DataView の行番号で引く */
    tooltip: TooltipColumn;
}

export interface ValueAxis2Settings {
    /** 第 2 Y 軸を描くか（折れ線を右の軸で描くとき） */
    show: boolean;
    valueShow: boolean;
    ticks: Tick[];
    /** 目盛りを Y 軸と同じ本数の上限で作り直す（viewModel の ticksFor と同じ） */
    ticksFor: (count: number) => Tick[];
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    labelColor: string;
    titleShow: boolean;
    titleText: string;
    titleFontFamily: string;
    titleFontSize: number;
    titleBold: boolean;
    titleItalic: boolean;
    titleUnderline: boolean;
    titleColor: string;
    /** 第 2 Y 軸の上に出す単位ラベル（(万人) など）。空なら出さない */
    badgeText: string;
    unitFontSize: number;
    unitColor: string;
}

/** 凡例の 1 項目。棒の系列のあとに折れ線が並ぶ（標準と同じ） */
export interface LegendItemInfo {
    kind: "bar" | "line" | "layer";
    /** 棒なら series の添字、折れ線なら lines の添字、比較レイヤーなら compareLayers の添字 */
    index: number;
    name: string;
    color: string;
    /** クリックで選ぶ ID。系列 1 本の棒は null（選ばない） */
    selectionId: ISelectionId | null;
    /** 棒の見た目（凡例の四角を棒に合わせる）。折れ線の見た目は lines[index] とマーカーを見る */
    barStyle?: { transparency: number; borderShow: boolean; borderColor: string; borderWidth: number };
    /** 比較レイヤーの印の濃さ（0〜1） */
    layerOpacity?: number;
}

/** 折れ線のマーカー */
export interface MarkerSettings {
    show: boolean;
    /** 型（"circle" | "square" | "diamond" | "triangle" | "cross" | "shortDash" | "longDash" | "plus"） */
    shape: string;
    size: number;
    /** 中心で回す角度（度、時計回り。0〜359） */
    rotation: number;
    /** 空 = 線の色 */
    color: string;
    transparency: number;
    borderShow: boolean;
    borderMatchLine: boolean;
    borderColor: string;
    borderTransparency: number;
    borderWidth: number;
}

/** 網掛け領域の色 */
export interface AreaSettings {
    matchLineColor: boolean;
    fill: string;
    transparency: number;
}

/** リボンの帯の見た目 */
export interface RibbonSettings {
    /** 帯を描くか（グラフの種類がリボンで縦棒のとき） */
    show: boolean;
    matchSeriesColor: boolean;
    fill: string;
    transparency: number;
    borderShow: boolean;
    borderMatchRibbon: boolean;
    borderFill: string;
    borderTransparency: number;
    borderWidth: number;
    /** 列の端と帯の端のすき間（カテゴリの間のすき間に対する %） */
    spacing: number;
}

export interface TotalLabelsSettings {
    /** 描くか（積み上げで、合計ラベルが オン のとき） */
    show: boolean;
    split: boolean;
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
    backgroundShow: boolean;
    backgroundColor: string;
    backgroundTransparency: number;
}

export interface LegendInfo {
    /** 描くか（複数系列で、凡例カードが オン のとき） */
    show: boolean;
    /** 折れ線の印（LEGEND_MARKER_STYLES） */
    markerStyle: string;
    /** 折れ線のマーカーを線の色で塗る */
    matchLineColor: boolean;
    position: string;
    /** 空ならタイトルなし */
    title: string;
    /** 凡例の「タイトル」の設定。オフなら、比較レイヤーの段の見出し「比較」も出さない */
    titleShow: boolean;
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
}

export interface Tick {
    value: number;
    label: string;
    ratio: number; // 0 (min) to 1 (max)
}

export interface UnitInfo {
    unitDef: UnitDefinition;
    badgeText: string;
    unitPosition: string; // "valueAxisTop" | "plotTopRight" | "none"
    precision: string;
    fontSize: number;
    color: string;
}

export interface DataLabelsSettings {
    show: boolean;
    /** 値の行の下線と透過性 (%) */
    underline: boolean;
    transparency: number;
    /** ラベルの表示を最適化：ラベルの行を labelMaxWidth (px) で切る */
    optimizeLabelDisplay: boolean;
    labelMaxWidth: number;
    /** 詳細がカスタムで、フィールドの値が空白のときに出す文字（空なら出さない） */
    detailShowBlankAs: string;
    /** 値の行を出すか */
    valueShow: boolean;
    /** 詳細の行を出すか */
    detailShow: boolean;
    detailFontFamily: string;
    detailFontSize: number;
    detailBold: boolean;
    detailItalic: boolean;
    detailUnderline: boolean;
    /** 空 = 自動（値の行と同じ） */
    detailColor: string;
    detailTransparency: number;
    /** タイトルの行を出すか（系列名か、ラベルのタイトルのフィールド） */
    titleShow: boolean;
    titleFontFamily: string;
    titleFontSize: number;
    titleBold: boolean;
    titleItalic: boolean;
    titleUnderline: boolean;
    /** 空 = 自動（値の行と同じ） */
    titleColor: string;
    titleTransparency: number;
    /** タイトルがカスタムで、フィールドの値が空白のときに出す文字（空なら出さない） */
    titleShowBlankAs: string;
    /** ラベルの値のフィールドが空白のときに出す文字（空なら棒の値） */
    valueShowBlankAs: string;
    /** 行を 1 行に並べるか（単一行） */
    singleLine: boolean;
    /** 複数行のとき、行をそろえる位置 */
    horizontalAlignment: "left" | "center" | "right";
    /** 文字の縁（幅 0 なら付けない） */
    halo: Halo;
    position: string; // "auto" | "outsideEnd" | "insideTop" | "insideCenter" | "insideBottom"
    orientation: string; // "horizontal" | "vertical"
    overflow: boolean;
    fontSize: number;
    fontFamily: string;
    bold: boolean;
    italic: boolean;
    color: string;
    backgroundShow: boolean;
    backgroundColor: string;
    backgroundTransparency: number;
    precision: string;
}

export interface CategoryAxisSettings {
    show: boolean;
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    labelColor: string;
    maxHeight: number;
    titleShow: boolean;
    titleText: string;
    titleFontFamily: string;
    titleFontSize: number;
    titleBold: boolean;
    titleItalic: boolean;
    titleUnderline: boolean;
    titleColor: string;
    minCategoryWidth: number;
    /** はみ出してスクロールするとき、開いたときにどこから見せるか（start・end） */
    scrollStart: string;
    /** 階層のラベルを 1 行につなぐか（標準の「ラベルの連結」）。1 列のときは使わない */
    concatenateLabels: boolean;
    /** 階層のレベルの数。1 なら階層なし */
    levelCount: number;
    /** 段に重ねた上のレベルの見せ方（"lines" 区切り線 | "boxed" 囲み） */
    hierarchyStyle: string;
}

export interface ValueAxisSettings {
    start: string;
    end: string;
    logarithmic: boolean;
    /** 対数が指定されたが 0・正負混在のデータのため線形で描いている */
    logarithmicFallback: boolean;
    invertRange: boolean;
    roundRange: boolean;
    /** 目盛りの本数の目安。0 なら自動（描く範囲の長さで決める） */
    tickCount: number;
    show: boolean;
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    labelColor: string;
    unitNotation: string;
    showUnitOnAxis: boolean;
    switchPosition: boolean;
    titleShow: boolean;
    titleText: string;
    titleStyle: string;
    titleFontFamily: string;
    titleFontSize: number;
    titleBold: boolean;
    titleItalic: boolean;
    titleUnderline: boolean;
    titleColor: string;
}

export interface GridlinesSettings {
    horizontalShow: boolean;
    horizontalColor: string;
    horizontalTransparency: number;
    horizontalStyle: string;
    horizontalWidth: number;
    horizontalScaleWithWidth: boolean;
    /** 線のスタイルがカスタムのときの模様（ダッシュ配列）と端の形 */
    horizontalDashArray: string;
    horizontalDashCap: string;
    verticalShow: boolean;
    verticalColor: string;
    verticalTransparency: number;
    verticalStyle: string;
    verticalWidth: number;
    verticalScaleWithWidth: boolean;
    verticalDashArray: string;
    verticalDashCap: string;
}

export interface ColumnsSettings {
    fill: string;
    transparency: number;
    showBorder: boolean;
    borderMatchColumn: boolean;
    borderFill: string;
    borderTransparency: number;
    borderWidth: number;
    reverseOrder: boolean;
    sortByValue: boolean;
    /** 外側のパディング (%)。null = 自動（カテゴリ間のスペースの半分） */
    outerPadding: number | null;
    categorySpacing: number;
    /** 系列間のスペース (%)。集合で同じカテゴリの棒のすき間 */
    seriesSpacing: number;
    /** 集合で系列の棒を重ねる（系列間のスペースを重ねる割合として読む）。積み上げでは false */
    overlap: boolean;
    /** 重なった棒の前後を入れ替える（既定は凡例の後ろの系列が手前） */
    overlapReverse: boolean;
    /** 積み上げの系列間のスペース。系列の展開がオフなら px、オンなら %（プロットの長さの半分に対する %） */
    stackedSpacing: number;
    /** 系列の展開（積んだ棒を離して並べ、値の軸を消す）。折れ線があるときは false */
    stackedExplode: boolean;
    /** 積み上げで、罫線を積んだ棒の外側だけに引く */
    borderOutlineOnly: boolean;
    /** 0 = 上限なし */
    maxBarWidth: number;
    cornerRadius: number;
}

export interface ViewModel {
    /** 読み込み中などの知らせ（左下に重ねる）。無ければ空か無し */
    notice?: string;
    /** すべての棒。カテゴリの順（並べ替え後）→ 系列の順 */
    dataPoints: DataPoint[];
    /** カテゴリごとにまとめた棒。描画はこちらを使う */
    categoryGroups: CategoryGroup[];
    /** 系列。系列が 1 本でも 1 つ入る */
    series: SeriesInfo[];
    /** 系列が複数（値が 2 つ以上か、凡例にフィールドがある） */
    seriesMode: boolean;
    /** グラフの種類（集合・積み上げ・100% 積み上げ） */
    chartType: ChartType;
    /** 棒の向き（縦棒・横棒）。横棒では折れ線を描かない（標準の複合は縦棒だけ） */
    orientation: Orientation;
    maxValue: number;
    minValue: number;
    niceMin: number;
    niceMax: number;
    zeroRatio: number;
    ticks: Tick[];
    /**
     * 目盛りの本数を決め直す（範囲は変えない）。標準と同じく、描く範囲の長さで本数の上限を決めるために、長さが決まった後で呼ぶ。
     * 対数の軸は本数を変えない
     */
    ticksFor: (count: number) => Tick[];
    unitInfo: UnitInfo;
    /** ドリルダウンした位置（「事業A ＞ 製品A1」）。ドリルしていない・出さないときは空 */
    drillPath: string;
    /** 「画像としてコピー」のボタンを出すか */
    copyButton: boolean;
    /** どのカテゴリでも同じ上の階層の値（軸の段から外したもの。読み上げの名前に使う）。書式でドリルの位置を消しても入る */
    commonLevels: string[];
    columns: ColumnsSettings;
    /** 書式ペインの「設定の適用先」。系列 1 本ならカテゴリ、複数なら系列 */
    columnTargets: ColumnTarget[];
    /** データラベルの系列ごとの設定。系列 1 本なら空 */
    labelTargets: LabelTarget[];
    dataLabels: DataLabelsSettings;
    categoryAxis: CategoryAxisSettings;
    valueAxis: ValueAxisSettings;
    gridlines: GridlinesSettings;
    legend: LegendInfo;
    /** 凡例に並べる項目（棒の系列 → 折れ線） */
    legendEntries: LegendItemInfo[];
    totalLabels: TotalLabelsSettings;
    /** 折れ線（複合）。無ければ空 */
    lines: LineSeriesInfo[];
    lineTargets: LineTarget[];
    /** 折れ線の率の分子と分母の数が合わないときの警告（組にならない分は描かない）。無ければ無し */
    lineWarning?: string;
    /** 「比較の列」の値が多すぎて重ねきれないときの警告（compareLayers.ts）。無ければ無し */
    compareWarning?: string;
    /** 「比較の列」に入れたフィールドの名前（ツールヒントでレイヤーの値の行に使う）。「比較値」で重ねたときは無し */
    compareByName?: string;
    markers: MarkerSettings;
    areas: AreaSettings;
    valueAxis2: ValueAxis2Settings;
    ribbons: RibbonSettings;
    hasHighlights: boolean;
    /** ツールヒントの元データ。データが無いときは null */
    tooltip: TooltipSource | null;
    /** 累計の状態 */
    cumulative: CumulativeInfo;
    /** パレートの状態 */
    pareto: ParetoInfo;
    /** Y 軸の定数線。無ければ空 */
    valueLines: ValueLineInfo[];
    valueLine: ValueLineSettings;
    /** Y 軸の定数線のメジャーが、カテゴリごとに違う値を返したときの警告（先頭のカテゴリの値で引く）。無ければ無し */
    valueLineWarning?: string;
    /** X 軸の定数線（「X 軸の定数線」の欄のメジャーごと）。無ければ空 */
    categoryLines: CategoryLineInfo[];
    categoryLine: CategoryLineSettings;
    /** 比較レイヤー（2 枚以上のときだけ。手前が [0]）。比較が無ければ空 */
    compareLayers: CompareLayerInfo[];
    compare: CompareSettings;
    /** 比較レイヤーごとのツールヒントの元データ（[0] は tooltip と同じ）。比較が無ければ無し */
    layerTooltips?: Array<TooltipSource | null>;
    /** 比較レイヤーで手前にそろえるもの（この viewModel を作ったときの値） */
    basis: { extent: ValueExtent; keepRows: number[] | null; displayOrder: number[] };
    isEmpty: boolean;
}

/**
 * X 軸の定数線 1 本（「X 軸の定数線」の欄のメジャー 1 つ）。メジャーが空白でない値を返したカテゴリに線を引く。
 * marks の index は categoryGroups の添字（並べ替えのあと）
 */
export interface CategoryLineInfo {
    name: string;
    /** label はスタイルで選んだ中身（データ値・名前・両方） */
    marks: Array<{ index: number; label: string }>;
}

export interface CategoryLineSettings {
    show: boolean;
    position: "before" | "center" | "after";
    color: string;
    transparency: number;
    lineStyle: string;
    dashArray: string;
    scaleWithWidth: boolean;
    dashCap: string;
    width: number;
    /** 棒の後ろ（back）か前（front）か */
    layer: "back" | "front";
    shadeShow: boolean;
    shadeRegion: "before" | "after";
    shadeColor: string;
    shadeTransparency: number;
    labelShow: boolean;
    labelHorizontal: "left" | "right";
    labelVertical: "top" | "bottom";
    labelColor: string;
    labelFontSize: number;
    labelAvoidOverlap: boolean;
}

/** Y 軸の定数線 1 本（カードの「値」か、「Y 軸の定数線」の欄のメジャー 1 つ）。軸の範囲の外の線は入れない */
export interface ValueLineInfo {
    name: string;
    value: number;
    /** 載せた軸の比率（0〜1） */
    ratio: number;
    secondary: boolean;
    /** スタイルで選んだ中身（データ値・名前・両方） */
    label: string;
}

export type ValueLineSettings = Omit<CategoryLineSettings, "position" | "shadeRegion"> & {
    /** 網掛け領域を、値の小さい側（below）か大きい側（above）に塗る */
    shadeRegion: "below" | "above";
};

/** 比較レイヤー 1 枚 */
export interface CompareLayerInfo {
    /** 凡例・ツールヒントに出す名前（メジャーの名前） */
    name: string;
}

export interface CompareSettings {
    /** 手前の棒を置く側。縦棒は右・左、横棒は下・上 */
    rightFront: boolean;
    /** 重なり（0〜1）。0 で隣り合わせ、1 で完全に重なる */
    overlap: number;
    /** いちばん奥の不透明度（0〜1） */
    backOpacity: number;
    outline: boolean;
    showInLegend: boolean;
}

export const EMPTY_CATEGORY_LINE: CategoryLineSettings = {
    show: true,
    position: "before",
    color: "#605E5C",
    transparency: 0,
    lineStyle: "dashed",
    dashArray: "",
    scaleWithWidth: true,
    dashCap: "none",
    width: 1,
    layer: "front",
    shadeShow: false,
    shadeRegion: "before",
    shadeColor: "#E1DFDD",
    shadeTransparency: 40,
    labelShow: true,
    labelHorizontal: "right",
    labelVertical: "top",
    labelColor: "#605E5C",
    labelFontSize: 9,
    labelAvoidOverlap: true,
};

export const EMPTY_VALUE_LINE: ValueLineSettings = {
    show: true,
    color: "#605E5C",
    transparency: 0,
    lineStyle: "dashed",
    dashArray: "",
    scaleWithWidth: true,
    dashCap: "none",
    width: 1,
    layer: "front",
    shadeShow: false,
    shadeRegion: "below",
    shadeColor: "#E1DFDD",
    shadeTransparency: 40,
    labelShow: true,
    labelHorizontal: "left",
    labelVertical: "top",
    labelColor: "#605E5C",
    labelFontSize: 9,
    labelAvoidOverlap: true,
};

export const VALUE_LINE_WARNING_TITLE = "Y 軸の定数線の値がカテゴリごとに違います";

export const EMPTY_COMPARE: CompareSettings = { rightFront: true, overlap: 0.6, backOpacity: 0.3, outline: false, showInLegend: true };

const EMPTY_COLUMNS_SETTINGS: ColumnsSettings = {
    fill: "#118DFF",
    transparency: 0,
    showBorder: false,
    borderMatchColumn: false,
    borderFill: "#605E5C",
    borderTransparency: 0,
    borderWidth: 1,
    reverseOrder: false,
    sortByValue: false,
    outerPadding: null,
    categorySpacing: 20,
    seriesSpacing: 0,
    overlap: false,
    overlapReverse: false,
    stackedSpacing: 0,
    stackedExplode: false,
    borderOutlineOnly: false,
    maxBarWidth: 0,
    cornerRadius: 0,
};

const clampPercent = (v: number): number => Math.max(0, Math.min(100, v));

/** 水平方向の配置の保存値（left・center・right。ほかは中央） */
const alignmentOf = (v: unknown): "left" | "center" | "right" => (v === "left" || v === "right" ? v : "center");
const clampBorderWidth = (v: number): number => Math.max(1, Math.min(5, v));

/** 対数目盛りの本数の上限。標準の値軸と同程度（5〜8 本）に収める */
export const MAX_LOG_TICKS = 8;

/** ドリルダウンした位置の区切り */
export const DRILL_PATH_SEPARATOR = " ＞ ";
/** ドリルした親が空白のときに出す文字（ツールヒントの空白の表記と同じ） */
export const BLANK_LEVEL_TEXT = "(空白)";

/** Math.log10 の丸め誤差を吸収する（10 の冪ならちょうどの整数を返す） */
/**
 * 0 を含む範囲 [lo, hi] を広げて、0 が下から zero（0〜1）の割合の高さに来るようにする（第 2 Y 軸の「0 を配置する」）。
 * zero が端（0 か 1）で、反対側に値があって合わせられないときは、そのまま返す
 */
export function alignZeroAt(lo: number, hi: number, zero: number): [number, number] {
    if (!(hi > lo) || lo > 0 || hi < 0) return [lo, hi];
    const current = -lo / (hi - lo);
    if (Math.abs(current - zero) < 1e-9) return [lo, hi];
    if (zero <= 0) return lo === 0 ? [0, hi] : [lo, hi];
    if (zero >= 1) return hi === 0 ? [lo, 0] : [lo, hi];
    // 0 の下が足りなければ下へ、上が足りなければ上へ広げる
    if (current < zero) return [(-zero * hi) / (1 - zero), hi];
    return [lo, (-lo * (1 - zero)) / zero];
}

function log10Snap(v: number): number {
    const l = Math.log10(v);
    const r = Math.round(l);
    return Math.abs(l - r) < 1e-9 ? r : l;
}

/**
 * 対数軸の目盛り（正の大きさ、lower < upper）。10 の冪を基本にし、3 本以上取れるときは
 * MAX_LOG_TICKS 以下になるよう桁を間引く。10 の冪が 2 本以下しか入らない狭い範囲では 1・2・5 を足す。
 */
export function logAxisTicks(lower: number, upper: number): number[] {
    const eps = 1e-9;
    const inRange = (v: number) => v >= lower * (1 - eps) && v <= upper * (1 + eps);

    const powers: number[] = [];
    for (let k = Math.ceil(log10Snap(lower)); k <= Math.floor(log10Snap(upper)); k++) {
        powers.push(Math.pow(10, k));
    }
    if (powers.length >= 3) {
        const step = Math.ceil(powers.length / MAX_LOG_TICKS);
        return powers.filter((_, i) => i % step === 0);
    }

    const withSteps: number[] = [];
    for (let k = Math.floor(log10Snap(lower)); k <= Math.ceil(log10Snap(upper)); k++) {
        for (const m of [1, 2, 5]) {
            const v = m * Math.pow(10, k);
            if (inRange(v)) withSteps.push(v);
        }
    }
    if (withSteps.length >= 2 && withSteps.length <= MAX_LOG_TICKS) return withSteps;
    if (powers.length > 0) return powers;
    return [lower, upper];
}

const EMPTY_CATEGORY_AXIS: CategoryAxisSettings = {
    show: true,
    fontFamily: "Segoe UI",
    fontSize: 9,
    bold: false,
    italic: false,
    underline: false,
    labelColor: "#605E5C",
    maxHeight: 25,
    titleShow: false,
    titleText: "",
    titleFontFamily: "DIN",
    titleFontSize: 12,
    titleBold: false,
    titleItalic: false,
    titleUnderline: false,
    titleColor: "#252423",
    minCategoryWidth: 20,
    scrollStart: "start",
    concatenateLabels: false,
    levelCount: 1,
    hierarchyStyle: "lines",
};

const EMPTY_VALUE_AXIS: ValueAxisSettings = {
    start: "",
    end: "",
    logarithmic: false,
    logarithmicFallback: false,
    invertRange: false,
    roundRange: true,
    tickCount: 0,
    show: true,
    fontFamily: "Segoe UI",
    fontSize: 9,
    bold: false,
    italic: false,
    underline: false,
    labelColor: "#605E5C",
    unitNotation: "japanese",
    showUnitOnAxis: false,
    switchPosition: false,
    titleShow: true,
    titleText: "",
    titleStyle: "showTitleOnly",
    titleFontFamily: "DIN",
    titleFontSize: 12,
    titleBold: false,
    titleItalic: false,
    titleUnderline: false,
    titleColor: "#252423",
};

const EMPTY_GRIDLINES: GridlinesSettings = {
    horizontalShow: true,
    horizontalColor: "#E1DFDD",
    horizontalTransparency: 0,
    horizontalStyle: "dotted",
    horizontalWidth: 1,
    horizontalScaleWithWidth: false,
    horizontalDashArray: "",
    horizontalDashCap: "none",
    verticalShow: false,
    verticalColor: "#E1DFDD",
    verticalTransparency: 0,
    verticalStyle: "dotted",
    verticalWidth: 1,
    verticalScaleWithWidth: false,
    verticalDashArray: "",
    verticalDashCap: "none",
};

const EMPTY_DATA_LABELS: DataLabelsSettings = {
    show: false,
    underline: false,
    transparency: 0,
    optimizeLabelDisplay: false,
    labelMaxWidth: 200,
    detailShowBlankAs: "",
    position: "auto",
    orientation: "horizontal",
    overflow: false,
    fontSize: 9,
    fontFamily: "Segoe UI",
    bold: false,
    italic: false,
    color: "",
    backgroundShow: false,
    backgroundColor: "#FFFFFF",
    backgroundTransparency: 0,
    precision: "auto",
    valueShow: true,
    detailShow: false,
    detailFontFamily: "Segoe UI",
    detailFontSize: 9,
    detailBold: false,
    detailItalic: false,
    detailUnderline: false,
    detailColor: "",
    detailTransparency: 0,
    titleShow: false,
    titleFontFamily: "Segoe UI",
    titleFontSize: 9,
    titleBold: false,
    titleItalic: false,
    titleUnderline: false,
    titleColor: "",
    titleTransparency: 0,
    titleShowBlankAs: "",
    valueShowBlankAs: "",
    singleLine: false,
    horizontalAlignment: "center",
    halo: NO_HALO,
};

const EMPTY_LEGEND: LegendInfo = {
    markerStyle: "lineAndMarker",
    matchLineColor: false,
    show: false,
    position: "topLeft",
    title: "",
    titleShow: true,
    fontFamily: "Segoe UI",
    fontSize: 10,
    bold: false,
    italic: false,
    underline: false,
    color: "#605E5C",
};

const EMPTY_TOTAL_LABELS: TotalLabelsSettings = {
    show: false,
    split: false,
    fontFamily: "Segoe UI",
    fontSize: 9,
    bold: false,
    italic: false,
    underline: false,
    color: "#605E5C",
    backgroundShow: false,
    backgroundColor: "#FFFFFF",
    backgroundTransparency: 0,
};

const EMPTY_PARETO: ParetoInfo = {
    enabled: false,
    thresholds: [0.8, 0.95],
    showThresholds: false,
    showRankBand: false,
    ranks: [],
    labels: { A: "A", B: "B", C: "C" },
    colors: { A: "#118DFF", B: "#74B9FF", C: "#C4E1FF" },
    selections: { A: [], B: [], C: [] },
};

const EMPTY: ViewModel = {
    cumulative: { available: false, enabled: false, toggle: false, levels: [], reset: "none" },
    pareto: EMPTY_PARETO,
    dataPoints: [],
    categoryGroups: [],
    series: [],
    seriesMode: false,
    chartType: CHART_TYPES.clustered,
    orientation: ORIENTATIONS.vertical,
    maxValue: 0,
    minValue: 0,
    niceMin: 0,
    niceMax: 1,
    zeroRatio: 0,
    ticks: [],
    ticksFor: () => [],
    drillPath: "",
    copyButton: false,
    commonLevels: [],
    unitInfo: {
        unitDef: UNIT_DEFINITIONS["0"],
        badgeText: "",
        unitPosition: "valueAxisTop",
        precision: "auto",
        fontSize: 9,
        color: "#605E5C",
    },
    columns: EMPTY_COLUMNS_SETTINGS,
    columnTargets: [],
    labelTargets: [],
    dataLabels: EMPTY_DATA_LABELS,
    categoryAxis: EMPTY_CATEGORY_AXIS,
    valueAxis: EMPTY_VALUE_AXIS,
    gridlines: EMPTY_GRIDLINES,
    legend: EMPTY_LEGEND,
    legendEntries: [],
    totalLabels: EMPTY_TOTAL_LABELS,
    lines: [],
    lineTargets: [],
    markers: {
        show: false,
        shape: "circle",
        size: 5,
        rotation: 0,
        color: "",
        transparency: 0,
        borderShow: false,
        borderMatchLine: false,
        borderColor: "#605E5C",
        borderTransparency: 0,
        borderWidth: 1,
    },
    areas: { matchLineColor: true, fill: "#118DFF", transparency: 60 },
    valueAxis2: {
        show: false,
        valueShow: true,
        ticks: [],
        ticksFor: () => [],
        fontFamily: "Segoe UI",
        fontSize: 9,
        bold: false,
        italic: false,
        underline: false,
        labelColor: "#605E5C",
        titleShow: true,
        titleText: "",
        titleFontFamily: "DIN",
        titleFontSize: 12,
        titleBold: false,
        titleItalic: false,
        titleUnderline: false,
        titleColor: "#252423",
        badgeText: "",
        unitFontSize: 9,
        unitColor: "#605E5C",
    },
    ribbons: {
        show: false,
        matchSeriesColor: true,
        fill: "#C8C6C4",
        transparency: 30,
        borderShow: false,
        borderMatchRibbon: false,
        borderFill: "#605E5C",
        borderTransparency: 0,
        borderWidth: 1,
        spacing: 0,
    },
    hasHighlights: false,
    tooltip: null,
    valueLines: [],
    valueLine: EMPTY_VALUE_LINE,
    categoryLines: [],
    categoryLine: EMPTY_CATEGORY_LINE,
    compareLayers: [],
    compare: EMPTY_COMPARE,
    basis: { extent: { min: 0, max: 0, valueMaxAbs: 0, magMin: Infinity, magMax: 0 }, keepRows: null, displayOrder: [] },
    isEmpty: true,
};

/**
 * リボンの帯のツールヒント（標準と同じ並び）：前後のカテゴリの値、値の変化（差と率）、前後の順位、順位の変化。
 * fromIndex は categoryGroups の添字で、帯は fromIndex と fromIndex + 1 のあいだ
 */
export function ribbonTooltipItems(
    viewModel: ViewModel,
    seriesIndex: number,
    fromIndex: number
): powerbi.extensibility.VisualTooltipDataItem[] {
    const a = viewModel.categoryGroups[fromIndex]?.points[seriesIndex];
    const b = viewModel.categoryGroups[fromIndex + 1]?.points[seriesIndex];
    const tooltip = viewModel.tooltip;
    if (!a || !b || !tooltip) return [];
    const measure = tooltip.series?.[seriesIndex]?.measure ?? tooltip.measure;
    const seriesName = viewModel.seriesMode ? viewModel.series[seriesIndex]?.name ?? "" : "";
    const valueLabel = (category: string) => [category, seriesName, measure.displayName].filter(Boolean).join(" ");
    const rankLabel = (category: string) => [category, seriesName, "順位"].filter(Boolean).join(" ");
    const diff = b.value - a.value;
    const rate = a.value !== 0 ? ` (${((diff / Math.abs(a.value)) * 100).toFixed(1)}%)` : "";
    const rankChange = a.rank !== null && b.rank !== null ? a.rank - b.rank : null;
    return [
        { displayName: valueLabel(a.category), value: formatTooltipValue(a.value, measure.format) },
        { displayName: valueLabel(b.category), value: formatTooltipValue(b.value, measure.format) },
        { displayName: `${measure.displayName} 変更`, value: `${formatTooltipValue(diff, measure.format)}${rate}` },
        ...(rankChange !== null
            ? [
                { displayName: rankLabel(a.category), value: String(a.rank) },
                { displayName: rankLabel(b.category), value: String(b.rank) },
                { displayName: "順位 変更", value: rankChange > 0 ? `▲${rankChange}` : rankChange < 0 ? `▼${-rankChange}` : "0" },
            ]
            : []),
    ];
}

/** 折れ線の点のツールヒント。標準と同じく カテゴリ と 折れ線の値 の 2 行 */
export function lineTooltipItems(
    viewModel: ViewModel,
    lineIndex: number,
    pointIndex: number
): powerbi.extensibility.VisualTooltipDataItem[] {
    const line = viewModel.lines[lineIndex];
    const group = viewModel.categoryGroups[pointIndex];
    if (!line || !group || !viewModel.tooltip) return [];
    return [
        ...categoryTooltipRows(viewModel.tooltip, group.rowIndex, group.category),
        { displayName: line.tooltip.displayName, value: formatTooltipValue(line.tooltip.values[group.rowIndex], line.tooltip.format) },
    ];
}

/** 棒のツールヒントに足す、積み上げの合計（積み上げ）か割合（100% 積み上げ） */
export function tooltipStackOf(viewModel: ViewModel, d: DataPoint): TooltipStack {
    // 比較レイヤーの奥の棒は、合計・割合を持たない（積み上げの合計は手前のレイヤーのもの）
    if ((d.layerIndex ?? 0) > 0) return viewModel.chartType === CHART_TYPES.stacked100 ? { share: d.share } : {};
    if (viewModel.chartType === CHART_TYPES.stacked100) return { share: d.share };
    if (viewModel.chartType === CHART_TYPES.stacked && viewModel.seriesMode) {
        const group = viewModel.categoryGroups.find((g) => g.rowIndex === d.rowIndex);
        return { total: group?.totals?.net ?? null };
    }
    return {};
}

/**
 * 系列 1 本ぶんの元データ。凡例があれば凡例の値ごと、無ければ「値」に入れたフィールドごとに 1 本。
 * capabilities の条件で「凡例あり ⇒ 値は 1 つ」に絞ってあるので、両方が複数になることはない
 */
interface SeriesSlot {
    /** テーマの色を割り当てるキー。凡例の未加工の値か、メジャーの queryName（名前を変えても色が変わらない） */
    key: string;
    name: string;
    column: DataViewValueColumn;
    tooltips: DataViewValueColumn[];
    /** 「ラベルの詳細」の列。無ければ undefined */
    detail?: DataViewValueColumn;
    /** 「ラベルのタイトル」「ラベルの値」の列。無ければ undefined */
    title?: DataViewValueColumn;
    value?: DataViewValueColumn;
    /** 「棒の色」の列。無ければ undefined */
    colorColumn?: DataViewValueColumn;
    group: DataViewValueColumnGroup;
    /** 系列ごとの書式の保存先。凡例の系列はグループ、メジャーの系列は列のメタデータに入る */
    objects: Array<DataViewObjects | undefined>;
}

const isMeasure = (column: DataViewValueColumn): boolean => !!column.source?.roles?.measure;
const isDataColor = (column: DataViewValueColumn): boolean => !!column.source?.roles?.dataColor;

/** 色として読める文字（#RGB・#RRGGBB・#RRGGBBAA、rgb()・hsl()、red などの名前） */
const COLOR_TEXT = /^(#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})|(rgb|hsl)a?\([^()]*\)|[a-z]+)$/i;

/**
 * 「棒の色」の i 行目の色。色として読めない値（空白・数など）は null。
 * 標準の系列の色の fx（フィールド値）の代わり。このビジュアルは matrix で受けるので、凡例があると Power BI が fx の色を渡さない
 */
function dataColorAt(column: DataViewValueColumn | undefined, i: number): string | null {
    const raw = column?.values[i];
    if (typeof raw !== "string") return null;
    const text = raw.trim();
    return COLOR_TEXT.test(text) ? text : null;
}

function seriesSlotsOf(
    groups: DataViewValueColumnGroup[],
    valueColumns: DataViewValueColumn[],
    legendSource: powerbi.DataViewMetadataColumn | undefined
): SeriesSlot[] {
    if (legendSource) {
        const legendFormatter = valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(legendSource) });
        return groups.flatMap((group) => {
            const column = group.values.find(isMeasure);
            if (!column) return [];
            const raw = group.name;
            const blankName = raw === null || raw === undefined;
            return [{
                key: blankName ? BLANK_TEXT : String(raw),
                name: blankName ? BLANK_TEXT : legendFormatter.format(raw),
                column,
                tooltips: group.values.filter((c) => c !== column && c.source?.roles?.tooltips),
                // 「値」と同じフィールドを入れると、Power BI は役割を 2 つ持つ 1 つの列にまとめて渡す
                detail: group.values.find((c) => c.source?.roles?.labelDetail),
                title: group.values.find((c) => c.source?.roles?.labelTitle),
                value: group.values.find((c) => c.source?.roles?.labelValue),
                colorColumn: group.values.find(isDataColor),
                group,
                objects: [group.objects, column.source?.objects],
            }];
        });
    }
    // 凡例なし：「値」に入れたフィールドごとに 1 本。「ツールヒント」の列は位置ではなくロールで分ける
    // （値を空にしてツールヒントだけ入れたときに、ツールヒントの列を棒として描かない）
    const columns = groups[0]?.values ?? valueColumns;
    const measures = columns.filter(isMeasure);
    const tooltips = columns.filter((c) => !measures.includes(c) && c.source?.roles?.tooltips);
    const detail = columns.find((c) => c.source?.roles?.labelDetail);
    const title = columns.find((c) => c.source?.roles?.labelTitle);
    const value = columns.find((c) => c.source?.roles?.labelValue);
    // 「棒の色」は系列 1 本のときだけ使う（値が複数で凡例が無いときは、どの系列の色か決められない）
    const colorColumn = measures.length === 1 ? columns.find(isDataColor) : undefined;
    return measures.map((column) => ({
        key: column.source.queryName ?? column.source.displayName,
        name: column.source.displayName,
        column,
        tooltips,
        detail,
        title,
        value,
        colorColumn,
        group: groups[0],
        objects: [column.source.objects],
    }));
}

/** 積み上げの棒 1 本の、始まりと量（値の単位。100% 積み上げは割合）。ハイライトの量は無ければ null */
interface StackSegment {
    start: number;
    amount: number;
    highlightAmount: number | null;
}

/**
 * 系列の展開。積んだ棒と棒のあいだに gap（プロットの長さに対する比）を空け、全体がプロットに収まるよう値の比率を縮めて置き直す。
 * すき間は同じ側（正・負）の棒と棒のあいだだけ（0 をはさむ所には入れない）。値の比率 r = z + v × s の s を、
 * どのカテゴリでも正の端が 1 − margin 以下・負の端が margin 以上になる範囲で最大にする。すき間が多すぎて入らなければ、すき間を縮める
 */
export function explodeStacks(groups: CategoryGroup[], segments: Map<DataPoint, StackSegment>, gap: number, margin = 0): void {
    const stacks = groups.map((g) => {
        const segs = g.points.filter((p) => !p.blank && segments.has(p)).map((p) => ({ p, ...segments.get(p)! }));
        const side = (positive: boolean) =>
            segs.filter((x) => (positive ? x.amount > 0 : x.amount < 0)).sort((a, b) => Math.abs(a.start) - Math.abs(b.start));
        const pos = side(true);
        const neg = side(false);
        return {
            g,
            pos,
            neg,
            zero: segs.filter((x) => x.amount === 0),
            P: pos.reduce((t, x) => t + x.amount, 0),
            N: neg.reduce((t, x) => t + x.amount, 0),
            gp: Math.max(0, pos.length - 1),
            gn: Math.max(0, neg.length - 1),
        };
    });
    const maxGaps = Math.max(0, ...stacks.map((x) => x.gp + x.gn));
    const g = maxGaps > 0 ? Math.min(gap, (0.9 - 2 * margin) / maxGaps) : gap;
    const lower = (s: number) => Math.max(margin, ...stacks.map((x) => margin - x.N * s + x.gn * g));
    const upper = (s: number) => Math.min(1 - margin, ...stacks.map((x) => 1 - margin - x.P * s - x.gp * g));
    const span = Math.max(1e-12, ...stacks.map((x) => x.P - x.N));
    let lo = 0;
    let hi = 1 / span;
    for (let k = 0; k < 60; k++) {
        const mid = (lo + hi) / 2;
        if (lower(mid) <= upper(mid)) lo = mid;
        else hi = mid;
    }
    const s = lo;
    const z = lower(s);
    for (const x of stacks) {
        const place = (list: typeof x.pos, dir: 1 | -1): number => {
            let cur = z;
            list.forEach((seg, k) => {
                if (k > 0) cur += dir * g;
                seg.p.startRatio = cur;
                seg.p.valRatio = cur + seg.amount * s;
                seg.p.highlightRatio = seg.highlightAmount === null ? null : cur + seg.highlightAmount * s;
                cur = seg.p.valRatio;
            });
            return cur;
        };
        const top = place(x.pos, 1);
        const bottom = place(x.neg, -1);
        for (const seg of x.zero) {
            seg.p.startRatio = z;
            seg.p.valRatio = z;
            seg.p.highlightRatio = seg.highlightAmount === null ? null : z;
        }
        for (const p of x.g.points) {
            if (p.blank) {
                p.startRatio = z;
                p.valRatio = z;
            }
            // 展開したあとは、すき間を縮める必要が無い
            p.insetStart = false;
        }
        if (x.g.totals) {
            x.g.totals.positiveEndRatio = top;
            x.g.totals.negativeEndRatio = bottom;
        }
    }
}

/** 書式の保存先を順に見て、最初に見つかった値を返す */
function firstOf<T>(objectsList: Array<DataViewObjects | undefined>, read: (objects: DataViewObjects | undefined) => T | null): T | null {
    for (const objects of objectsList) {
        const v = read(objects);
        if (v !== null) return v;
    }
    return null;
}

/** 個別に指定された色を読む。無ければ null */
const customColor = (
    objects: DataViewObjects | undefined,
    property: "fill" | "borderFill"
): string | null => {
    const fill = objects?.columns?.[property] as powerbi.Fill | undefined;
    return fill?.solid?.color ? String(fill.solid.color) : null;
};

/** 個別に指定された数値を読む。無ければ null */
const customNumber = (
    objects: DataViewObjects | undefined,
    property: "transparency" | "borderTransparency" | "borderWidth"
): number | null => {
    const raw = objects?.columns?.[property];
    return typeof raw === "number" ? raw : null;
};

/** 個別に指定された真偽値を読む。無ければ null */
const customFlag = (
    objects: DataViewObjects | undefined,
    property: "showBorder" | "borderMatchColumn"
): boolean | null => {
    const raw = objects?.columns?.[property];
    return typeof raw === "boolean" ? raw : null;
};

/** 系列ごとのデータラベルの表示（このシリーズに表示）。無ければ null */
const labelFlag = (objects: DataViewObjects | undefined): boolean | null => {
    const raw = objects?.dataLabels?.show;
    return typeof raw === "boolean" ? raw : null;
};

/** 系列ごとのデータラベルの色。無ければ null */
const labelColorOf = (objects: DataViewObjects | undefined): string | null => {
    const fill = objects?.dataLabels?.color as powerbi.Fill | undefined;
    return fill?.solid?.color ? String(fill.solid.color) : null;
};

/**
 * 値の大きい順に並べる。階層（levels が 2 つ以上）では、上のレベルから順に、
 * 親どうしを配下の合計で並べ、その中で子を並べる。同じ値なら元の順を保つ
 */
export function sortWithinParents<T extends { levels: string[]; levelKeys?: string[] }>(items: T[], totalOf: (item: T) => number): T[] {
    const depth = items[0]?.levels.length ?? 1;
    const sumOf = (list: T[]) => list.reduce((sum, item) => sum + totalOf(item), 0);
    const sortAt = (list: T[], level: number): T[] => {
        if (level >= depth - 1) return [...list].sort((a, b) => totalOf(b) - totalOf(a));
        const buckets = new Map<string, T[]>();
        for (const item of list) {
            // 親は元の値で見分ける（空白と文字の "null" のように、表示が同じでも別の親がある）
            const key = (item.levelKeys ?? item.levels)[level];
            const bucket = buckets.get(key);
            if (bucket) bucket.push(item);
            else buckets.set(key, [item]);
        }
        return [...buckets.values()].sort((a, b) => sumOf(b) - sumOf(a)).flatMap((bucket) => sortAt(bucket, level + 1));
    };
    return sortAt(items, 0);
}

/**
 * 書式の区切り。選択肢はデータ次第で、書式の読み込み（populateFormattingSettingsModel）の時点では「区切らない」しか無く、
 * 保存した値が選択肢に無いとして捨てられる。なので保存先（metadata.objects）から直接読む（Desktop で確認）
 */
export function savedCumulativeReset(dataView: DataView | undefined, calc: VisualFormattingSettingsModel["calculation"]): string {
    const raw = dataView?.metadata?.objects?.calculation?.cumulativeReset;
    if (typeof raw === "string" && raw !== "") return raw;
    const value = calc.cumulativeReset.value;
    return String((value && typeof value === "object" ? value.value : value) ?? CUMULATIVE_RESET_NONE);
}

/**
 * 閲覧者の操作（グラフの上の累計の切り替えボタン・区切り）。visual.ts が保存済みの状態から作る。
 * 無ければ書式の値どおり
 */
export interface CalculationRuntime {
    /** 累計を効かせるか。書式が累計のときだけ意味を持つ */
    cumulative?: boolean;
    /** 閲覧者が選んだ区切り（レベルの queryName か CUMULATIVE_RESET_NONE）。null なら書式の値 */
    cumulativeReset?: string | null;
    /**
     * 前に描いた値の軸の高さ (px)。棒の外に出すデータ ラベルが入るよう、値の軸の範囲を広げるのに使う（labelRoomOf）。
     * 値の軸の高さは凡例と X 軸で決まり、範囲では変わらないので、描いた高さを次の transform に渡せば落ち着く
     */
    plotHeight?: number;
}

/** データ ラベルの文字の行の高さ（文字サイズに対する比）と、行のあいだ・背景の余白 (px)。unitUtils の labelBlock と同じ見積もり */
const LABEL_LINE_RATIO = 1.02;
const LABEL_ROOM_GAP = 1;
const LABEL_ROOM_PADDING = 3;

/**
 * 縦棒の外の端に出すデータ ラベルのまとまりの高さ (px)。棒の外に出しうるときだけ（積み上げ・100%・縦向きのラベルは 0）。
 * 標準は、外に出すラベルが入るよう値の軸の範囲を広げる（ラベルが上の枠で切れたり、X 軸の名前に重なったりしない）
 */
function labelRoomOf(settings: VisualFormattingSettingsModel, stacked: boolean): number {
    const dl = settings.dataLabels;
    if (!(dl.show.value ?? false) || stacked) return 0;
    const position = String(dl.position.value?.value ?? "auto");
    if (position !== "auto" && position !== "outsideEnd") return 0;
    if (String(dl.orientation.value?.value ?? "horizontal") === "vertical") return 0;
    const sizes: number[] = [];
    if (dl.titleShow.value) sizes.push(dl.titleFont.fontSize.value ?? 9);
    if (dl.valueShow.value ?? true) sizes.push(dl.fontSize.value ?? 9);
    if (dl.detailShow.value) sizes.push(dl.detailFont.fontSize.value ?? 9);
    if (!sizes.length) return 0;
    const px = sizes.map((pt) => Math.max(8, Math.min(32, pt)) * PT_TO_PX * LABEL_LINE_RATIO);
    const singleLine = String(dl.labelContentLayout.value?.value) === LABEL_LAYOUTS.singleLine;
    const text = singleLine ? Math.max(...px) : px.reduce((sum, h) => sum + h, 0) + LABEL_ROOM_GAP * (px.length - 1);
    return text + LABEL_ROOM_PADDING * 2;
}

/**
 * 比較レイヤーの奥のレイヤーを組み立てるときに、手前のレイヤーにそろえるもの（compareLayers.ts）。
 * 値の軸はすべてのレイヤーの範囲で決め、カテゴリの並びと「その他」にまとめる行は手前のレイヤーで決める
 */
export interface LayerBasis {
    /** すべてのレイヤーの値の範囲。軸の範囲と表示単位に入れる */
    extent?: ValueExtent;
    /** 「その他」にまとめずに残す行（畳む前の行番号） */
    keepRows?: number[];
    /** カテゴリの並び（畳んだ後の行番号） */
    displayOrder?: number[];
}

/** 棒の値の範囲（折れ線を入れる前）。比較レイヤーで軸をそろえるのに使う */
export interface ValueExtent {
    min: number;
    max: number;
    /** 値そのものの絶対値の最大（集合・100% の表示単位） */
    valueMaxAbs: number;
    /** 0 でない値の絶対値の最小・最大（対数の軸）。値が無ければ magMin は Infinity */
    magMin: number;
    magMax: number;
}

/** パレートのランク。累積比で A・B・C に分ける */
export type ParetoRank = "A" | "B" | "C";

/** パレートの状態。棒の並び（categoryGroups）の順で持つ */
export interface ParetoInfo {
    enabled: boolean;
    /** A と B、B と C の境目（累積比、0〜1） */
    thresholds: [number, number];
    showThresholds: boolean;
    showRankBand: boolean;
    /** 棒の並びの順のランク。値が 0 以下（累積比に数えない）なら null */
    ranks: Array<ParetoRank | null>;
    labels: Record<ParetoRank, string>;
    colors: Record<ParetoRank, string>;
    /** ランクごとの棒の ID（帯を押したときにまとめて選ぶ） */
    selections: Record<ParetoRank, ISelectionId[]>;
}

/** 累計の状態。書式ペインの区切りの選択肢と、グラフの上のボタンに使う */
export interface CumulativeInfo {
    /** 書式の「計算」が累計か（閲覧者のボタンを出せる） */
    available: boolean;
    /** いま累計を効かせているか（閲覧者がボタンで切ったら false） */
    enabled: boolean;
    /** グラフの上に切り替えボタンを出すか */
    toggle: boolean;
    /** 区切りの選択肢。X 軸の階層のレベル（いちばん下を除く）。value は queryName */
    levels: powerbi.IEnumMember[];
    /** いまの区切り（CUMULATIVE_RESET_NONE なら区切らない） */
    reset: string;
}

export function transform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    runtime: CalculationRuntime = {},
    basis: LayerBasis = {}
): ViewModel {
    // matrix で受けたなら categorical の形に詰め替える（折れ線の値は、凡例をまとめたカテゴリ全体の値になる）
    dataView = categoricalOf(dataView);
    // 上位 N 件＋「その他」。読む前に DataView を畳む。「その他」の選択は元のカテゴリの列から作り直す
    const originalCategories = dataView?.categorical?.categories;
    const originalDataView = dataView;
    const others = collapseOthers(dataView, settings.columns.otherCount.value ?? 0, settings.columns.otherLabel.value?.trim() || "その他", basis.keepRows);
    dataView = others.dataView;
    const otherRow = others.otherRow;
    const otherSelectionIds = others.mergedRows.map((r) =>
        withCategoryOf(host.createSelectionIdBuilder(), originalDataView!.categorical!, originalCategories![0], r).createSelectionId()
    );

    const categorical: DataViewCategorical | undefined = dataView?.categorical;
    const categories = categorical?.categories?.[0];
    const valueColumns: DataViewValueColumn[] = categorical?.values ?? [];
    // 凡例のフィールド。あれば values は凡例の値ごとのグループに分かれて届く
    const legendSource = categorical?.values?.source;
    const groups: DataViewValueColumnGroup[] =
        categorical?.values?.grouped?.() ?? [{ values: valueColumns } as DataViewValueColumnGroup];
    const slots = seriesSlotsOf(groups, valueColumns, legendSource);

    if (!categories || !slots.length || !categories.values.length) {
        return EMPTY;
    }

    const seriesMode = !!legendSource || slots.length > 1;
    const rowCount = categories.values.length;

    // 階層。X 軸にフィールドが複数あり、展開していると、上のレベルから順に 1 列ずつ届く。
    // 1 列のときは 1.22 までと同じ（levels が 1 つ）
    const levelColumns: DataViewCategoryColumn[] = categorical!.categories!;
    // 日付のカテゴリはモデルの書式で出す（String() だと JS の日付の文字列になる）。それ以外は 1.4 までと同じ
    const levelTextOf = levelColumns.map((column) => {
        const formatter = column.source?.type?.dateTime
            ? valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(column.source) || "yyyy/MM/dd" })
            : null;
        return (i: number): string => {
            const raw = column.values[i];
            return formatter && raw !== null && raw !== undefined ? formatter.format(raw) : String(raw);
        };
    });
    const levelsAt = (i: number): string[] => levelTextOf.map((text) => text(i));
    const levelKeysAt = (i: number): string[] =>
        levelColumns.map((column) => {
            const raw = column.values[i];
            if (raw === null || raw === undefined) return "blank";
            return raw instanceof Date ? `date:${raw.getTime()}` : `${typeof raw}:${String(raw)}`;
        });
    /**
     * 上の階層のうち、どのカテゴリでも同じレベル。ドリルダウンすると親のレベルも届くが
     * （事業A → 製品A1 に入ると、どの行も「事業A / 製品A1 / …」）、全部同じ親を軸の段に繰り返しても読めない。
     * 軸の段・名前からは外し、ドリルの位置に出す。「すべて展開」で親が混ざるレベルから下は今までどおり段に出す。
     * いちばん下のレベルは外さない（1 行だけでも名前は出す）
     */
    const commonDepth = (() => {
        if (levelColumns.length < 2 || rowCount === 0) return 0;
        const first = levelKeysAt(0);
        let depth = 0;
        while (depth < levelColumns.length - 1 && Array.from({ length: rowCount }, (_, i) => i).every((i) => levelKeysAt(i)[depth] === first[depth])) depth++;
        return depth;
    })();
    /**
     * ドリルダウンした位置。どのカテゴリでも同じ上の階層の値を「＞」でつなぐ（事業A に入り、製品A1 に入ると「事業A ＞ 製品A1」）。
     * ドリルと、絞り込みで親が 1 つに決まった「すべて展開」は区別しない（どちらも、表示している項目がその親の下にあることは正しい）
     */
    const commonLevels =
        rowCount > 0 ? levelsAt(0).slice(0, commonDepth).map((text, level) => (levelKeysAt(0)[level] === "blank" || text === "" ? BLANK_LEVEL_TEXT : text)) : [];
    const drillPath = commonLevels.length && (settings.chart.drillPathShow?.value ?? true) ? commonLevels.join(DRILL_PATH_SEPARATOR) : "";
    /** 軸に出すレベル（共通の親を外したもの） */
    const shownLevelsAt = (i: number): string[] => levelsAt(i).slice(commonDepth);
    const shownLevelKeysAt = (i: number): string[] => levelKeysAt(i).slice(commonDepth);
    /** 1 行で読むときの名前。階層は上のレベルから空白でつなぐ（標準の「ラベルの連結」と同じ） */
    const categoryText = (i: number): string => levelsAt(i).join(" ");
    /**
     * ID と書式の保存先は、いちばん下のレベルの列（1 列なら今までと同じ列）。この列の ID は上のレベルを含んだ
     * 複合で（保存される selector は FY・半期・四半期の And）、別の年の Q1 とも区別できる。
     * レベルごとに withCategory を重ねると selector の data が重複し、保存した書式が objects に戻らない（2026-09-23、Desktop）
     */
    const lowestLevel = levelColumns[levelColumns.length - 1];
    const withCategories = (builder: powerbi.visuals.ISelectionIdBuilder, i: number) => withCategoryOf(builder, categorical!, lowestLevel, i);
    const categoryObjectsAt = (i: number) => lowestLevel.objects?.[i];

    const orientation: Orientation =
        String(settings.chart.orientation?.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal
            ? ORIENTATIONS.horizontal
            : ORIENTATIONS.vertical;
    const horizontal = orientation === ORIENTATIONS.horizontal;

    // 折れ線の値。凡例があると、系列ごとに同じメジャーの列が複製されて届くので queryName でまとめる。
    // 標準の複合は縦棒だけだが、どの種類・向きでも描く（横棒は既定でマーカーだけ）
    const lineColumns = new Map<string, DataViewValueColumn[]>();
    for (const group of groups) {
        for (const column of group.values) {
            if (!column.source?.roles?.lineMeasure) continue;
            const key = column.source.queryName ?? column.source.displayName;
            const found = lineColumns.get(key);
            if (found) found.push(column);
            else lineColumns.set(key, [column]);
        }
    }
    /**
     * 折れ線 1 本の値（行ごと）。凡例があると系列ごとの値が届く。
     * matrix で受けると、凡例があるときはカテゴリ全体の値（列の小計）が 1 列だけ届く（matrixDataView.ts）。
     * 系列ごとの値が届いたとき（小計が届かない・categorical で受けた）は、どの行でも同じならその値、違えば足した値を使う
     */
    const lineIndependenceBeforeCollapse = otherRow >= 0 ? lineIndependenceOf(originalDataView) : null;
    /** 系列ごとに複製された列の、行 i の数（空白・数でないものは除く） */
    const numbersOf = (columns: DataViewValueColumn[], i: number) =>
        columns
            .map((c) => c.values[i])
            .filter((v) => v !== null && v !== undefined)
            .map((v) => (typeof v === "number" ? v : Number(v)))
            .filter((v) => Number.isFinite(v));
    const measureLineDefs: LineDef[] = [...lineColumns.entries()].map(([key, columns]) => {
        // 「その他」にまとめたときは、畳む前の DataView で判定する（行が減ると判定が変わり、残したカテゴリの線の値まで変わる）
        const independent =
            lineIndependenceBeforeCollapse?.get(key) ??
            Array.from({ length: rowCount }, (_, i) => numbersOf(columns, i)).every((vals) => vals.every((v) => v === vals[0]));
        const values: Array<number | null> = Array.from({ length: rowCount }, (_, i) => {
            const vals = numbersOf(columns, i);
            if (!vals.length) return null;
            return independent ? vals[0] : vals.reduce((sum, v) => sum + v, 0);
        });
        const source = columns[0].source;
        return {
            key,
            name: source.displayName,
            format: valueFormatter.getFormatStringByColumn(source),
            objects: source.objects,
            values,
        };
    });
    /**
     * 「折れ線の率の分子」「折れ線の率の分母」。入れた順に組にして 1 本の線にする。分母が 1 つだけなら、すべての分子をその分母で割る
     * （粗利率と営業利益率を、どちらも売上で割るとき）。
     * 分子と分母を別に受け、系列をまたいでそれぞれ足してから最後に 1 回だけ割るので、凡例があっても全体の率になる
     * （率のメジャーをそのまま「折れ線の値」に入れても、カテゴリ全体の値が届くので同じ率になる）
     */
    const numeratorColumns = new Map<string, DataViewValueColumn[]>();
    const denominatorColumns = new Map<string, DataViewValueColumn[]>();
    for (const group of groups) {
        for (const column of group.values) {
            const roles = column.source?.roles;
            const target = roles?.lineRatioNumerator ? numeratorColumns : roles?.lineRatioDenominator ? denominatorColumns : null;
            if (!target) continue;
            const key = column.source.queryName ?? column.source.displayName;
            const found = target.get(key);
            if (found) found.push(column);
            else target.set(key, [column]);
        }
    }
    const numeratorList = [...numeratorColumns.entries()];
    const denominatorList = [...denominatorColumns.values()];
    const sharedDenominator = denominatorList.length === 1;
    const lineWarning = numeratorList.length !== denominatorList.length && !(sharedDenominator && numeratorList.length > 0) ? LINE_RATIO_WARNING : undefined;
    /** 系列をまたいだ合計。すべて空白なら空白 */
    const seriesSumOf = (columns: DataViewValueColumn[], i: number): number | null => {
        const vals = numbersOf(columns, i);
        return vals.length ? vals.reduce((sum, v) => sum + v, 0) : null;
    };
    const pairedNumerators = sharedDenominator ? numeratorList : numeratorList.slice(0, denominatorList.length);
    const ratioLineDefs: LineDef[] = pairedNumerators.map(([key, numerators], k) => {
        const denominators = denominatorList[sharedDenominator ? 0 : k];
        const ratio = {
            numerators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(numerators, i)),
            denominators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(denominators, i)),
        };
        const source = numerators[0].source;
        return {
            key,
            // 線の名前は分子の名前（ビジュアルの中でフィールドの名前を変えれば、線の名前も変わる）
            name: source.displayName,
            format: valueFormatter.getFormatStringByColumn(source),
            objects: source.objects,
            values: divideRatio(ratio.numerators, ratio.denominators),
            ratio,
        };
    });
    const lineDefs = [...measureLineDefs, ...ratioLineDefs].map((def) => ({ ...def, format: lineFormatOf(def, settings.lines) }));
    const isBlankAt = (column: DataViewValueColumn, i: number) => column.values[i] === null || column.values[i] === undefined;
    const numberAt = (column: DataViewValueColumn, i: number) => {
        const raw = column.values[i];
        return typeof raw === "number" ? raw : Number(raw) || 0;
    };

    const getDropdownValue = (sliceVal: any, fallback: string): string => {
        if (sliceVal == null) return fallback;
        if (typeof sliceVal === "object" && sliceVal !== null && "value" in sliceVal) {
            return String(sliceVal.value);
        }
        return String(sliceVal);
    };

    const chartTypeValue = getDropdownValue(settings.chart.chartType?.value, CHART_TYPES.clustered);
    // パレート。100% 積み上げのパレートはどの棒も同じ高さで意味が無いので、そのあいだは積み上げとして描く
    const paretoOn = getDropdownValue(settings.calculation.mode.value, CALCULATION_MODES.none) === CALCULATION_MODES.pareto;
    // 1.18 までの「リボン」は、書式の読み込み（applyChartTypeDefaults）で積み上げ＋リボンに読み替える。ここでは積み上げとして扱う
    const chartType: ChartType =
        chartTypeValue === CHART_TYPES.stacked || chartTypeValue === CHART_TYPES.ribbon || (paretoOn && chartTypeValue === CHART_TYPES.stacked100)
            ? CHART_TYPES.stacked
            : chartTypeValue === CHART_TYPES.stacked100
                ? CHART_TYPES.stacked100
                : CHART_TYPES.clustered;
    const stacked = chartType !== CHART_TYPES.clustered;
    const percent = chartType === CHART_TYPES.stacked100;
    // リボン（帯）は積み上げ・100% 積み上げで出せる。積む順が「値の大きい順」なら、カテゴリごとに値の大きい系列を外側に積み替える
    // （標準のリボン グラフと同じ。凡例の順なら標準の積み上げ＋リボンと同じ）
    const ribbonsOn = stacked && (settings.ribbons.show.value ?? false);
    const rankOrder = ribbonsOn && getDropdownValue(settings.ribbons.order.value, RIBBON_ORDERS.legend) === RIBBON_ORDERS.value;
    // 積み上げ順を逆にする（リボンで値の大きい順に積むときは、その順が優先）
    const reverseStack = stacked && !rankOrder && (settings.columns.reverseStackOrder.value ?? false);

    // 複数系列の空白は棒を描かないので数えない（系列 1 本の空白は 1.4 までと同じく 0 として数える）
    const isSkipped = (column: DataViewValueColumn, i: number) => seriesMode && isBlankAt(column, i);

    // --- 並び順 ---------------------------------------------------------------
    // 表示の順（値順・逆順）は元の値で先に決める。累計はこの順に沿って足し、並べ替えは累計した値ではなく元の値で行う。
    // 値順は、複数系列ではカテゴリの合計（空白は 0）の大きい順。階層では親の中で並べ替える（同じ親が離れると上のレベルのラベルが割れる）
    const rawTotalAt = (i: number) => slots.reduce((sum, slot) => sum + (isSkipped(slot.column, i) ? 0 : numberAt(slot.column, i)), 0);
    // パレートは値の大きい順が図の定義なので、値順の設定によらず大きい順に並べ、逆順にもしない
    let displayOrder = Array.from({ length: rowCount }, (_, i) => i);
    if (paretoOn) {
        // 階層を展開していても、親の中ではなく全体を大きい順に並べる（親の中で並べると、いちばん大きいカテゴリが C になりうる）。
        // 上のレベルの段は細かく割れるが、並びの意味を優先する
        displayOrder = [...displayOrder].sort((a, b) => rawTotalAt(b) - rawTotalAt(a));
    } else if (settings.columns.sortByValue.value ?? false) {
        displayOrder = sortWithinParents(
            displayOrder.map((i) => ({ levels: levelsAt(i), levelKeys: levelKeysAt(i), i })),
            (item) => rawTotalAt(item.i)
        ).map((item) => item.i);
    }
    if (!paretoOn && (settings.columns.reverseOrder.value ?? false)) displayOrder.reverse();
    // 「その他」は、並べ替えによらずいつも最後に置く
    if (otherRow >= 0) displayOrder = [...displayOrder.filter((i) => i !== otherRow), otherRow];
    // 比較レイヤーの奥のレイヤーは、手前のレイヤーの並びに合わせる（値順でも、手前の値の順に並べる）
    if (basis.displayOrder && basis.displayOrder.length === rowCount) displayOrder = basis.displayOrder.slice();

    // --- 累計 ------------------------------------------------------------
    // 表示の順に、系列ごとに値を足していく。区切り（階層のレベル）の値が変わったところで 0 に戻す。
    // 棒の値そのものを差し替えるので、積み上げ・軸・データラベルは累計した値で描く
    const calc = settings.calculation;
    const cumulativeAvailable = getDropdownValue(calc.mode.value, CALCULATION_MODES.none) === CALCULATION_MODES.cumulative;
    const cumulativeEnabled = cumulativeAvailable && (runtime.cumulative ?? true);
    const levelItems: powerbi.IEnumMember[] = levelColumns.map((column) => ({
        value: column.source?.queryName ?? column.source?.displayName ?? "",
        displayName: column.source?.displayName ?? "",
    }));
    const cumulativeReset = runtime.cumulativeReset ?? savedCumulativeReset(dataView, calc);
    // 区切りは、いま届いているどのレベルでも効かせる。ドリルアップして区切りのレベルがいちばん下になったら、
    // 棒ごとに 0 に戻る（= 素の値。「FY ごとに 0 に戻す」を FY の棒で見たとき）
    const resetDepth = levelItems.findIndex((level) => level.value === cumulativeReset);
    // 選択肢はいちばん下を除くレベル（いちばん下で区切っても素の値になるだけ）。ただし選んである区切りは、
    // 閲覧者のものも書式のものも残す（書式ペインの区切りが選択肢から消えて「区切らない」に見えないように）
    const authorResetDepth = levelItems.findIndex((level) => level.value === savedCumulativeReset(dataView, calc));
    const cumulativeLevels = levelItems.filter((level, k) => k < levelItems.length - 1 || k === resetDepth || k === authorResetDepth);
    /** 表示の k 番目の手前で 0 に戻すか */
    const resetBefore = displayOrder.map((i, k) => {
        if (k === 0 || resetDepth < 0) return false;
        const groupOf = (row: number) => JSON.stringify(levelKeysAt(row).slice(0, resetDepth + 1));
        return groupOf(i) !== groupOf(displayOrder[k - 1]);
    });
    /**
     * 行ごとの値を表示の順に足した値にする。空白はそこまでの合計にする。
     * keepBlank なら、区切りの中で値がまだ 1 つも無いあいだの空白は空白のまま（凡例の系列が途中から始まるときに、0 の棒や線を描かない）
     */
    const cumulate = (values: powerbi.PrimitiveValue[], keepBlank: boolean): powerbi.PrimitiveValue[] => {
        const out = values.slice();
        let running = 0;
        let started = false;
        displayOrder.forEach((i, k) => {
            if (resetBefore[k]) {
                running = 0;
                started = false;
            }
            const raw = values[i];
            if (raw === null || raw === undefined) {
                out[i] = keepBlank && !started ? raw : running;
                return;
            }
            running += typeof raw === "number" ? raw : Number(raw) || 0;
            started = true;
            out[i] = running;
        });
        return out;
    };
    /** 累計の前の列（ツールヒントに元の値を出すため）。累計でなければ空 */
    const beforeCumulative = new Map<SeriesSlot, DataViewValueColumn>();
    /** 累計にした折れ線の key */
    const cumulativeLineKeys = new Set<string>();
    if (cumulativeEnabled) {
        for (const slot of slots) {
            const column = slot.column;
            beforeCumulative.set(slot, column);
            slot.column = {
                ...column,
                // 空白の行もそこまでの合計の棒にする。複数系列では、その系列が始まる前の空白は描かない
                values: cumulate(column.values, seriesMode),
                // ハイライトは「ここまでの該当分」。該当しない行も、手前までの該当分を残す
                ...(column.highlights ? { highlights: cumulate(column.highlights, false) } : {}),
            };
        }
        const includeAll = settings.lines.includeCumulative.value ?? false;
        for (const def of lineDefs) {
            const own = def.objects?.lines?.includeCumulative;
            if (!(typeof own === "boolean" ? own : includeAll)) continue;
            // 率の線は、率を足さずに分子と分母をそれぞれ累計してから割る（累計の率）
            def.values = def.ratio
                ? divideRatio(
                    cumulate(def.ratio.numerators, true) as Array<number | null>,
                    cumulate(def.ratio.denominators, true) as Array<number | null>
                )
                : (cumulate(def.values, true) as Array<number | null>);
            cumulativeLineKeys.add(def.key);
        }
    }
    // 「その他」の行の折れ線は、足しようが無いので空白（「値」と同じ列を「折れ線の値」にも入れたとき・累計で空白が埋まったときも）。
    // 率の線は、まとめたカテゴリの分子の合計 ÷ 分母の合計なので出せる（累計でなければ。累計の「その他」は区切りをまたぐので空白）
    if (otherRow >= 0) lineDefs.forEach((def) => {
        if (!def.ratio || cumulativeLineKeys.has(def.key)) def.values[otherRow] = null;
    });
    /** ツールヒントの値の行。累計なら「値（累計）」と、累計の前の値の行を出す */
    const measureTooltipOf = (slot: SeriesSlot): { measure: TooltipColumn; before?: TooltipColumn } => {
        const measure = tooltipColumnOf(slot.column);
        const before = beforeCumulative.get(slot);
        return before ? { measure: { ...measure, displayName: `${measure.displayName}（累計）` }, before: tooltipColumnOf(before) } : { measure };
    };
    /** 表示の順での位置（行番号 → 何番目か） */
    const positionOf = new Map(displayOrder.map((i, k) => [i, k]));

    // --- パレート ------------------------------------------
    // 並んだ順に、カテゴリの合計を足した割合（累積比）。分母は正の合計で、0 以下のカテゴリは数えない
    // （全体を分け合う図なので、負が混ざると累積比が 100% を超えたり戻ったりする。標準のビジュアル計算はそうなる）
    const paretoCard = settings.calculation;
    const percentOf = (raw: number | undefined, fallback: number) =>
        Math.max(0, Math.min(100, Number.isFinite(raw) ? (raw as number) : fallback)) / 100;
    const boundaryA = percentOf(paretoCard.boundaryAB.value, 80);
    const boundaryB = percentOf(paretoCard.boundaryBC.value, 95);
    // 境目が逆に入っても壊れないよう、小さいほうを A と B の境目にする
    const paretoThresholds: [number, number] = [Math.min(boundaryA, boundaryB), Math.max(boundaryA, boundaryB)];
    const paretoRatioByRow = new Array<number | null>(rowCount).fill(null);
    const paretoRankByRow = new Array<ParetoRank | null>(rowCount).fill(null);
    if (paretoOn) {
        // 「その他」は、まとめたカテゴリの正の値の合計で数える（正と負を相殺させない）
        const paretoValueAt = (i: number) => (i === otherRow ? others.otherPositiveTotal : rawTotalAt(i));
        const total = displayOrder.reduce((sum, i) => sum + Math.max(0, paretoValueAt(i)), 0);
        let running = 0;
        let first = true;
        const EPS = 1e-9;
        for (const i of displayOrder) {
            const value = paretoValueAt(i);
            if (!(value > 0) || !(total > 0)) continue;
            running += value;
            const ratio = running / total;
            paretoRatioByRow[i] = ratio;
            // 先頭は必ず A（1 件で境目を越えるデータでも A が 0 件にならないように）。境目ちょうどは A に入れる
            paretoRankByRow[i] = first || ratio <= paretoThresholds[0] + EPS ? "A" : ratio <= paretoThresholds[1] + EPS ? "B" : "C";
            first = false;
        }
    }
    const paretoColors: Record<ParetoRank, string> = {
        A: paretoCard.colorA.value?.value || "#118DFF",
        B: paretoCard.colorB.value?.value || "#74B9FF",
        C: paretoCard.colorC.value?.value || "#C4E1FF",
    };
    /** ランクで棒を塗るときの色。凡例（系列）があれば系列の色のまま。0 以下のカテゴリは C の色 */
    const paretoFillAt = (i: number): string | null =>
        paretoOn && !seriesMode && (paretoCard.colorByRank.value ?? true) ? paretoColors[paretoRankByRow[i] ?? "C"] : null;

    // カテゴリごとの正の合計・負の合計・絶対値の合計（積み上げの軸と 100% の割合に使う）
    const positiveSums = new Array<number>(rowCount).fill(0);
    const negativeSums = new Array<number>(rowCount).fill(0);
    const absoluteSums = new Array<number>(rowCount).fill(0);

    // 値そのものの範囲（集合の軸、対数、100% のデータラベルの単位に使う）
    const allValues: number[] = [];
    let valueMaxAbs = 0;
    for (const slot of slots) {
        for (let i = 0; i < rowCount; i++) {
            if (isSkipped(slot.column, i)) continue;
            const val = numberAt(slot.column, i);
            allValues.push(val);
            valueMaxAbs = Math.max(valueMaxAbs, Math.abs(val));
            if (val >= 0) positiveSums[i] += val;
            else negativeSums[i] += val;
            absoluteSums[i] += Math.abs(val);
        }
    }

    // 軸に描く範囲。集合は値、積み上げは正と負それぞれの積み上げの端、100% は割合（-1〜1）
    const plotted: number[] = !stacked
        ? allValues
        : percent
            ? positiveSums.flatMap((p, i) => (absoluteSums[i] > 0 ? [p / absoluteSums[i], negativeSums[i] / absoluteSums[i]] : [0]))
            : positiveSums.flatMap((p, i) => [p, negativeSums[i]]);
    let maxValue = -Infinity;
    let minValue = Infinity;
    for (const v of plotted) {
        if (v > maxValue) maxValue = v;
        if (v < minValue) minValue = v;
    }
    if (!isFinite(maxValue)) maxValue = 0;
    if (!isFinite(minValue)) minValue = 0;
    let magMin = Infinity;
    let magMax = 0;
    for (const v of allValues) {
        const mag = Math.abs(v);
        if (mag > 0 && mag < magMin) magMin = mag;
        if (mag > magMax) magMax = mag;
    }
    /** このレイヤーだけの範囲（比較レイヤーで、すべてのレイヤーの範囲を決めるのに使う） */
    const ownExtent: ValueExtent = { min: minValue, max: maxValue, valueMaxAbs, magMin, magMax };
    // 比較レイヤーでは、すべてのレイヤーの範囲で軸を決める（どのレイヤーも同じ軸で描く）
    if (basis.extent) {
        maxValue = Math.max(maxValue, basis.extent.max);
        minValue = Math.min(minValue, basis.extent.min);
        valueMaxAbs = Math.max(valueMaxAbs, basis.extent.valueMaxAbs);
        if (Number.isFinite(basis.extent.magMin)) allValues.push(basis.extent.magMin * (minValue < 0 && maxValue <= 0 ? -1 : 1));
        if (basis.extent.magMax > 0) allValues.push(basis.extent.magMax * (minValue < 0 && maxValue <= 0 ? -1 : 1));
    }

    // 表示単位は、積み上げなら積み上げの端の大きさ、それ以外（集合・100%）は値の大きさで決める
    let maxAbs = stacked && !percent ? Math.max(Math.abs(maxValue), Math.abs(minValue)) : valueMaxAbs;

    // 折れ線を右の第 2 Y 軸で描くか、左の軸を共有するか。
    // 第 2 Y 軸の表示を保存していなければ自動（標準と同じく、棒と範囲が近ければ共有する）。
    // 100% 積み上げの軸は割合なので、折れ線はいつも第 2 Y 軸
    const lineNumbers = lineDefs.flatMap((d) => d.values.filter((v): v is number => v !== null));
    const lineMin = lineNumbers.length ? Math.min(...lineNumbers) : 0;
    const lineMax = lineNumbers.length ? Math.max(...lineNumbers) : 0;
    const axis2Saved = (dataView?.metadata?.objects?.valueAxis2 as powerbi.DataViewObject | undefined)?.show;
    let onSecondary = false;
    // パレートでは右の軸を累積比（0〜100%）に使うので、「折れ線の値」の線は左の軸を共有する
    if (lineDefs.length && !paretoOn) {
        if (percent) {
            onSecondary = true;
        } else if (typeof axis2Saved === "boolean") {
            onSecondary = axis2Saved;
        } else {
            const barAbs = Math.max(Math.abs(maxValue), Math.abs(minValue));
            const lineAbs = Math.max(Math.abs(lineMax), Math.abs(lineMin));
            const comparable = barAbs > 0 && lineAbs >= barAbs * 0.2 && lineAbs <= barAbs * 1.5;
            const sameSide = !(lineMin < 0 && minValue >= 0) && !(lineMax > 0 && maxValue <= 0);
            onSecondary = !(comparable && sameSide);
        }
    }
    // 左の軸を共有するなら、軸の範囲と単位に折れ線の値も入れる
    if (lineDefs.length && !onSecondary) {
        maxValue = Math.max(maxValue, lineMax);
        minValue = Math.min(minValue, lineMin);
        allValues.push(...lineNumbers);
        maxAbs = Math.max(maxAbs, Math.abs(lineMax), Math.abs(lineMin));
    }

    // --- Y 軸の定数線の値 -------------------------------------------------------
    // カードの「値」（決まった数）と、「Y 軸の定数線」の欄のメジャー。メジャーはカテゴリごとに計算されて届くので、
    // 並べた順の先頭のカテゴリの値で引き、ほかのカテゴリで値が違えば警告する（凡例があると matrix は全体の値で 1 列届く）
    const vl = settings.valueLine;
    const valueLineOn = vl.show.value ?? true;
    const valueLineDefs: Array<{ name: string; value: number }> = [];
    let valueLineWarning: string | undefined;
    if (valueLineOn) {
        const fixed = boundOf(vl.value.value);
        if (fixed !== null) valueLineDefs.push({ name: vl.lineName.value?.trim() || "定数線", value: fixed });
        const valueLineColumns = new Map<string, DataViewValueColumn[]>();
        for (const group of groups) {
            for (const column of group.values) {
                if (!column.source?.roles?.valueLine) continue;
                const key = column.source.queryName ?? column.source.displayName;
                const found = valueLineColumns.get(key);
                if (found) found.push(column);
                else valueLineColumns.set(key, [column]);
            }
        }
        const varying: string[] = [];
        for (const columns of [...valueLineColumns.values()].slice(0, 4)) {
            const values = displayOrder
                .filter((i) => i !== otherRow)
                .map((i) => numbersOf(columns, i)[0])
                .filter((x): x is number => x !== undefined);
            if (!values.length) continue;
            if (values.some((x) => x !== values[0])) varying.push(columns[0].source.displayName);
            valueLineDefs.push({ name: columns[0].source.displayName, value: values[0] });
        }
        if (varying.length) valueLineWarning = `「${varying.join("」「")}」はカテゴリごとに違う値を返しているので、先頭のカテゴリの値で線を引いています。決まった値を返すメジャーにするか、カテゴリごとの値なら「折れ線の値」に入れてください。`;
    }
    const valueLineOnSecondary = onSecondary && getDropdownValue(vl.axis.value, "primary") === "secondary";
    // 軸を線に合わせて広げる（Y 軸のときだけ。標準は広げない）
    if (valueLineDefs.length && !valueLineOnSecondary && (vl.extendAxis.value ?? false) && !percent) {
        for (const def of valueLineDefs) {
            maxValue = Math.max(maxValue, def.value);
            minValue = Math.min(minValue, def.value);
            maxAbs = Math.max(maxAbs, Math.abs(def.value));
            allValues.push(def.value);
        }
    }

    // Y軸設定および単位設定の抽出
    const valAxis = settings.valueAxis;
    const unitTypeKey = getDropdownValue(valAxis.unitType.value, "auto");
    const unitNotation = getDropdownValue(valAxis.unitNotation?.value, "japanese");
    const showUnitOnAxis = valAxis.showUnitOnAxis?.value ?? false;
    const unitText = valAxis.unitText.value ?? "";
    const unitShow = valAxis.unitShow.value ?? true;
    const unitIncludeDisplayUnit = valAxis.unitIncludeDisplayUnit.value ?? true;
    const unitPosition = getDropdownValue(valAxis.unitPosition.value, "valueAxisTop");
    const unitStyle = getDropdownValue(valAxis.unitStyle.value, "parentheses");
    const precision = getDropdownValue(valAxis.precision.value, "auto");

    // 値のメジャーの書式が % なら（構成比・率。フィールド パラメーターで円のメジャーと切り替えるときも）、100% 積み上げや
    // 第 2 Y 軸の率と同じく、表示単位によらず % で出し、単位ラベル（単位の追加文字の「円」など）を付けない
    const measureFormats = slots.map((slot) => valueFormatter.getFormatStringByColumn(slot.column.source) ?? "");
    const percentMeasure = !percent && measureFormats.length > 0 && measureFormats.every(isPercentFormat);
    /** 書式の % の前の小数の桁（0.0% なら 1）。小数点以下の桁数が「自動」のときに使う */
    const measurePercentDigits = percentDigitsOf(measureFormats[0]);
    const measurePercentText = (v: number, digits: string, sign?: Partial<SignStyle>) => formatPercent(v, digits, measurePercentDigits, sign);
    const unitDef = percentMeasure ? resolveUnit("0", 0, unitNotation, precision) : resolveUnit(unitTypeKey, maxAbs, unitNotation, precision);

    const roundRange = valAxis.roundRange.value ?? true;
    const invertRange = valAxis.invertRange.value ?? false;
    const logarithmic = valAxis.logarithmic.value ?? false;

    let userStart: number | undefined = undefined;
    let userEnd: number | undefined = undefined;
    // 桁区切りのカンマ（1,000）も読む。1.30 までは読めずに自動になっていた
    userStart = boundOf(valAxis.start.value) ?? undefined;
    userEnd = boundOf(valAxis.end.value) ?? undefined;

    let niceMin = 0;
    let niceMax = 1;
    let rawTicks: number[] = [];

    // 対数はデータが全て正か全て負のときだけ効かせる。0 を含む・正負が混在するときは線形で描く。
    // 標準の値軸と同じ扱い（Microsoft Learn「軸のカスタマイズ」: 対数は全て正か全て負が必要で 0 は不可、
    // データが変わって条件を外れると自動で線形に切り替わる）。標準は書式ペインに警告アイコンを出すが
    // カスタムビジュアルからは出せないため、状態は valueAxis.logarithmicFallback に持つ。
    const logSign = minValue > 0 ? 1 : maxValue < 0 ? -1 : 0;
    // 積み上げでは対数を使わない（積んだ棒の長さが値の和にならなくなる）
    const isLogScaleActive = logarithmic && logSign !== 0 && !stacked;
    const logarithmicFallback = logarithmic && !isLogScaleActive;
    // 対数軸の範囲（大きさ）。負のデータでは 0 に近い側 (-logLower) が上端になる
    let logLower = 1;
    let logUpper = 10;

    if (isLogScaleActive) {
        let magMin = Infinity;
        let magMax = 0;
        for (const val of allValues) {
            const mag = Math.abs(val);
            if (mag < magMin) magMin = mag;
            if (mag > magMax) magMax = mag;
        }

        // ユーザー指定の最小値・最大値を大きさに直す。符号がデータと合わない指定は無視する
        const userLower = logSign > 0
            ? (userStart !== undefined && userStart > 0 ? userStart : undefined)
            : (userEnd !== undefined && userEnd < 0 ? -userEnd : undefined);
        const userUpper = logSign > 0
            ? (userEnd !== undefined && userEnd > 0 ? userEnd : undefined)
            : (userStart !== undefined && userStart < 0 ? -userStart : undefined);

        // 下限はデータの最小値より小さい最大の 10 の冪（例: 150 → 100、1,000 → 100）。
        // 最小値ちょうどにすると最小の棒が高さ 0 になって値が無いように見えるため 1 桁下げる
        logLower = userLower ?? Math.pow(10, Math.ceil(log10Snap(magMin)) - 1);
        logUpper = userUpper ?? (roundRange ? Math.pow(10, Math.ceil(log10Snap(magMax))) : magMax);
        if (!(logUpper > logLower)) logUpper = logLower * 10;

        niceMin = logSign > 0 ? logLower : -logUpper;
        niceMax = logSign > 0 ? logUpper : -logLower;
        rawTicks = logAxisTicks(logLower, logUpper).map((m) => logSign * m);
    } else if (percent) {
        // 100% 積み上げの軸は割合。負の値があれば -100%〜100%、正だけなら 0%〜100%（標準と同じ。最小値・最大値は使わない）
        const lowerBound = minValue < 0 ? -1 : 0;
        const upperBound = maxValue > 0 || minValue >= 0 ? 1 : 0;
        niceMin = lowerBound;
        niceMax = upperBound > lowerBound ? upperBound : lowerBound + 1;
        rawTicks = scaleLinear().domain([niceMin, niceMax]).ticks(5);
    } else {
        // 合計ラベルを出す積み上げでは、自動で決める端に範囲の 10% の余白を足す
        // （合計ラベルが軸の外で切れたり、カテゴリ名に掛かったりしないように）
        const totalLabelsOn = chartType === CHART_TYPES.stacked && (settings.totalLabels.show.value ?? false);
        const span = Math.max(maxValue, 0) - Math.min(minValue, 0);
        const headroom = totalLabelsOn ? span * 0.1 : 0;
        // 棒の外の端に出すデータ ラベルの余白（縦棒）。描いた値の軸の高さに対するラベルの高さの比で、自動の端を広げる
        const room = !horizontal && (runtime.plotHeight ?? 0) > 0 ? labelRoomOf(settings, stacked) / runtime.plotHeight! : 0;
        const roomTop = userEnd === undefined && maxValue > 0 ? room : 0;
        const roomBottom = userStart === undefined && minValue < 0 ? room : 0;
        const roomShare = Math.min(0.6, roomTop + roomBottom);
        const roomSpan = roomShare > 0 ? span / (1 - roomShare) : span;
        const lowerBound = userStart !== undefined ? userStart : (minValue < 0 ? minValue - headroom - roomBottom * roomSpan : 0);
        const upperBound = userEnd !== undefined ? userEnd : (maxValue > 0 ? maxValue + headroom + roomTop * roomSpan : 1);
        const scale = scaleLinear().domain([lowerBound, upperBound]);
        if (roundRange && userStart === undefined && userEnd === undefined) {
            scale.nice();
        }
        const domain = scale.domain();
        niceMin = domain[0];
        niceMax = domain[1] > niceMin ? domain[1] : niceMin + 1;
        rawTicks = scale.ticks(5);
    }

    // 対数スケール時は全体一律のスケーリング語（億・M等）は使わず、タイトルやバッジにはユーザー単位（円）のみ出す
    const effectiveUnitWord = isLogScaleActive ? "" : unitDef.unitWord;
    // 100% 積み上げの軸は割合なので、軸の単位ラベル（(億円) など）は出さない
    const badgeText = percent || percentMeasure
        ? ""
        : resolveBadgeText({
            unitShow,
            unitPosition,
            unitIncludeDisplayUnit,
            unitStyle,
            unitWord: effectiveUnitWord,
            unitText,
        });

    const calcRatio = (val: number): number => {
        if (isLogScaleActive) {
            const mag = Math.abs(val);
            if (mag <= 0) return logSign > 0 ? 0 : 1;
            const logMin = Math.log10(logLower);
            const logSpan = Math.log10(logUpper) - logMin;
            const r = logSpan > 0 ? Math.max(0, Math.min(1, (Math.log10(mag) - logMin) / logSpan)) : 0;
            return logSign > 0 ? r : 1 - r;
        } else {
            const span = niceMax - niceMin;
            return span > 0 ? Math.max(0, Math.min(1, (val - niceMin) / span)) : 0;
        }
    };

    // 対数では 0 を置けないため、棒の基準線は軸の端（正のデータは下端、負のデータは上端）
    const zeroRatio = isLogScaleActive ? (logSign > 0 ? 0 : 1) : calcRatio(0);

    // 目盛りラベルの生成
    const tickOf = (t: number): Tick => {
        let label: string;
        if (percent) {
            label = `${Math.round(t * 100)}%`;
        } else if (percentMeasure && !isLogScaleActive) {
            // 目盛りの間隔が 1% 未満なら小数 1 桁（第 2 Y 軸の率と同じ）
            const stepPercent = rawTicks.length > 1 ? Math.abs(rawTicks[1] - rawTicks[0]) * 100 : 1;
            label = `${(t * 100).toFixed(precision !== "auto" ? Number(precision) : stepPercent < 1 ? 1 : 0)}%`;
        } else if (isLogScaleActive) {
            // 対数時は目盛値ごとに動的に単位付与（showUnitOnAxisがfalseならカンマ区切り生数値）
            label = formatDynamicValue(t, unitNotation, precision, showUnitOnAxis);
        } else {
            const numLabel = formatValue(t, unitDef.divisor, precision);
            label = showUnitOnAxis && unitDef.unitWord ? `${numLabel}${unitDef.unitWord}` : numLabel;
        }
        const normRatio = calcRatio(t);
        return {
            value: t,
            label,
            ratio: invertRange ? 1 - normRatio : normRatio,
        };
    };
    const ticks: Tick[] = rawTicks.map(tickOf);
    const ticksFor = (count: number): Tick[] =>
        isLogScaleActive ? ticks : ticksUpTo([niceMin, niceMax], count, (values) => values.map((v) => tickOf(v).label)).map(tickOf);

    // X軸設定の抽出
    const catAxis = settings.categoryAxis;
    const concatenateLabels = catAxis.concatenateLabels.value ?? false;
    const levelNames = levelColumns.map((column) => column.source?.displayName ?? "").slice(commonDepth);
    const autoCategoryTitle = concatenateLabels ? levelNames.join(" ") : levelNames[levelNames.length - 1];
    const categoryAxisSettings: CategoryAxisSettings = {
        show: catAxis.show.value ?? true,
        fontFamily: catAxis.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, catAxis.font.fontSize.value ?? 9)),
        bold: catAxis.font.bold?.value ?? false,
        italic: catAxis.font.italic?.value ?? false,
        underline: catAxis.font.underline?.value ?? false,
        labelColor: catAxis.labelColor.value?.value || "#605E5C",
        maxHeight: Math.max(0, Math.min(100, catAxis.maxHeight.value ?? 25)),
        titleShow: catAxis.titleShow.value ?? false,
        // 自動のタイトルは標準と同じ：段に重ねるならいちばん下のレベル、1 行につなぐならレベルの名前を並べる
        titleText: catAxis.titleText.value?.trim() || autoCategoryTitle,
        titleFontFamily: catAxis.titleFont.fontFamily.value ?? "DIN",
        titleFontSize: Math.max(8, Math.min(32, catAxis.titleFont.fontSize.value ?? 12)),
        titleBold: catAxis.titleFont.bold?.value ?? false,
        titleItalic: catAxis.titleFont.italic?.value ?? false,
        titleUnderline: catAxis.titleFont.underline?.value ?? false,
        titleColor: catAxis.titleColor.value?.value || "#252423",
        minCategoryWidth: Math.max(0, Math.min(500, catAxis.minCategoryWidth.value ?? 20)),
        scrollStart: getDropdownValue(catAxis.scrollStart.value, "start"),
        concatenateLabels,
        levelCount: levelColumns.length - commonDepth,
        hierarchyStyle: getDropdownValue(catAxis.hierarchyStyle.value, "lines") === "boxed" ? "boxed" : "lines",
    };

    // ユーザー設定の単位ラベル（例: "億円", "円", "bn"）を反映したタイトル生成。
    // 値が複数のときの自動のタイトルは、標準と同じく値の名前を「および」でつなぐ
    const composedUnit = percent ? "" : composeUnitText(effectiveUnitWord, unitText, unitIncludeDisplayUnit);
    const autoValueTitle = seriesMode && !legendSource
        ? slots.map((slot) => slot.name).join(" および ")
        : (slots[0].column.source?.displayName ?? "");
    const rawValueTitle = valAxis.titleText.value?.trim() || autoValueTitle;
    const valueTitleStyle = getDropdownValue(valAxis.titleStyle.value, "showTitleOnly");
    /** 軸タイトルのスタイル（タイトルのみ・単位のみ・両方）を当てる。単位が無ければタイトルのまま（第 2 Y 軸も同じ） */
    const styledTitle = (title: string, style: string, composed: string): string => {
        if (!composed) return title;
        if (style === "showUnitOnly") return `(${composed})`;
        if (style === "showBoth") return `${title} (${composed})`;
        return title;
    };
    const formattedValueTitle = styledTitle(rawValueTitle, valueTitleStyle, composedUnit);
    /** タイトルに単位を入れる（単位のみ・両方）ときは、単位のラベルを出さない（同じ単位を 2 か所に出さない。1.30 までは両方に出た） */
    const unitInTitle = (titleShow: boolean, style: string, composed: string): boolean =>
        titleShow && composed !== "" && (style === "showUnitOnly" || style === "showBoth");
    const valueUnitInTitle = unitInTitle(valAxis.titleShow.value ?? true, valueTitleStyle, composedUnit);

    const valueAxisSettings: ValueAxisSettings = {
        start: valAxis.start.value ?? "",
        end: valAxis.end.value ?? "",
        logarithmic,
        logarithmicFallback,
        invertRange,
        roundRange,
        tickCount: tickCountOf(valAxis.tickCount.value),
        show: valAxis.show.value ?? true,
        fontFamily: valAxis.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, valAxis.font.fontSize.value ?? 9)),
        bold: valAxis.font.bold?.value ?? false,
        italic: valAxis.font.italic?.value ?? false,
        underline: valAxis.font.underline?.value ?? false,
        labelColor: valAxis.labelColor.value?.value || "#605E5C",
        unitNotation,
        showUnitOnAxis,
        switchPosition: valAxis.switchPosition.value ?? false,
        titleShow: valAxis.titleShow.value ?? true,
        titleText: formattedValueTitle,
        titleStyle: valueTitleStyle,
        titleFontFamily: valAxis.titleFont.fontFamily.value ?? "DIN",
        titleFontSize: Math.max(8, Math.min(32, valAxis.titleFont.fontSize.value ?? 12)),
        titleBold: valAxis.titleFont.bold?.value ?? false,
        titleItalic: valAxis.titleFont.italic?.value ?? false,
        titleUnderline: valAxis.titleFont.underline?.value ?? false,
        titleColor: valAxis.titleColor.value?.value || "#252423",
    };

    // グリッド線設定の抽出
    const gl = settings.gridlines;
    const gridlinesSettings: GridlinesSettings = {
        horizontalShow: gl.horizontalShow.value ?? true,
        horizontalColor: gl.horizontalColor.value?.value || "#E1DFDD",
        horizontalTransparency: Math.max(0, Math.min(100, gl.horizontalTransparency.value ?? 0)),
        horizontalStyle: String(gl.horizontalStyle.value?.value ?? "dotted"),
        horizontalWidth: Math.max(1, Math.min(10, gl.horizontalWidth.value ?? 1)),
        horizontalScaleWithWidth: gl.horizontalScaleWithWidth.value ?? false,
        horizontalDashArray: gl.horizontalDashArray.value ?? "",
        horizontalDashCap: String(gl.horizontalDashCap.value?.value ?? "none"),
        verticalShow: gl.verticalShow.value ?? false,
        verticalColor: gl.verticalColor.value?.value || "#E1DFDD",
        verticalTransparency: Math.max(0, Math.min(100, gl.verticalTransparency.value ?? 0)),
        verticalStyle: String(gl.verticalStyle.value?.value ?? "dotted"),
        verticalWidth: Math.max(1, Math.min(10, gl.verticalWidth.value ?? 1)),
        verticalScaleWithWidth: gl.verticalScaleWithWidth.value ?? false,
        verticalDashArray: gl.verticalDashArray.value ?? "",
        verticalDashCap: String(gl.verticalDashCap.value?.value ?? "none"),
    };

    // データラベル設定の抽出
    const dl = settings.dataLabels;
    const dataLabelsSettings: DataLabelsSettings = {
        show: dl.show.value ?? false,
        position: String(dl.position.value?.value ?? "auto"),
        orientation: String(dl.orientation.value?.value ?? "horizontal"),
        overflow: dl.overflow.value ?? false,
        fontSize: Math.max(8, Math.min(32, dl.fontSize.value ?? 9)),
        fontFamily: dl.fontFamily.value ?? "Segoe UI",
        bold: dl.bold.value ?? false,
        italic: dl.italic.value ?? false,
        underline: dl.underline.value ?? false,
        transparency: clampPercent(dl.transparency.value ?? 0),
        optimizeLabelDisplay: dl.optimizeLabelDisplay.value ?? false,
        labelMaxWidth: Math.max(10, Math.min(2000, Number(dl.labelMaxWidth.value) || 200)),
        detailShowBlankAs: dl.detailShowBlankAs.value ?? "",
        color: dl.color.value?.value ?? "",
        backgroundShow: dl.backgroundShow.value ?? false,
        backgroundColor: dl.backgroundColor.value?.value ?? "#FFFFFF",
        backgroundTransparency: Math.max(0, Math.min(100, dl.backgroundTransparency.value ?? 0)),
        precision: String(dl.precision.value?.value ?? "auto"),
        valueShow: dl.valueShow.value ?? true,
        detailShow: dl.detailShow.value ?? false,
        detailFontFamily: dl.detailFont.fontFamily.value ?? "Segoe UI",
        detailFontSize: Math.max(8, Math.min(32, dl.detailFont.fontSize.value ?? 9)),
        detailBold: dl.detailFont.bold?.value ?? false,
        detailItalic: dl.detailFont.italic?.value ?? false,
        detailUnderline: dl.detailFont.underline?.value ?? false,
        detailColor: dl.detailColor.value?.value ?? "",
        detailTransparency: clampPercent(dl.detailTransparency.value ?? 0),
        titleShow: dl.titleShow.value ?? false,
        titleFontFamily: dl.titleFont.fontFamily.value ?? "Segoe UI",
        titleFontSize: Math.max(8, Math.min(32, dl.titleFont.fontSize.value ?? 9)),
        titleBold: dl.titleFont.bold?.value ?? false,
        titleItalic: dl.titleFont.italic?.value ?? false,
        titleUnderline: dl.titleFont.underline?.value ?? false,
        titleColor: dl.titleColor.value?.value ?? "",
        titleTransparency: clampPercent(dl.titleTransparency.value ?? 0),
        titleShowBlankAs: dl.titleShowBlankAs.value ?? "",
        valueShowBlankAs: dl.valueShowBlankAs.value ?? "",
        singleLine: String(dl.labelContentLayout.value?.value) === LABEL_LAYOUTS.singleLine,
        horizontalAlignment: alignmentOf(dl.horizontalAlignment.value),
        halo: haloOf(dl.haloShow.value, dl.haloColor.value?.value, dl.haloWidth.value),
    };

    // 列（columns）設定の抽出
    const col = settings.columns;
    // 系列の展開は積み上げだけ。折れ線があると値の軸を消せないので使わない（標準の複合にも無い）
    const explodeOn = stacked && (col.stackedExplode.value ?? false) && lineDefs.length === 0;
    const overlapOn = !stacked && (col.overlap.value ?? false);
    // 系列 1 本の棒の色：保存が無ければ、標準と同じくテーマのデータの色の 1 番目（1.28 までは #118DFF 固定で、
    // テーマの 1 番目を変えたレポートでも従わなかった）。getColor は同じ key なら同じ色を返すので、下の折れ線の色の取り方とずれない
    const singleSeriesFill =
        seriesMode || !slots.length ? null : customColor(dataView?.metadata?.objects, "fill") ?? host.colorPalette.getColor(slots[0].key).value;
    const columnsSettings: ColumnsSettings = {
        fill: singleSeriesFill || col.fill.value?.value || "#118DFF",
        transparency: Math.max(0, Math.min(100, col.transparency.value ?? 0)),
        showBorder: col.showBorder.value ?? false,
        borderMatchColumn: col.borderMatchColumn.value ?? false,
        borderFill: col.borderFill.value?.value || "#605E5C",
        borderTransparency: Math.max(0, Math.min(100, col.borderTransparency.value ?? 0)),
        borderWidth: clampBorderWidth(col.borderWidth.value ?? 1),
        reverseOrder: col.reverseOrder.value ?? false,
        sortByValue: col.sortByValue.value ?? false,
        // 空（保存値なし）は「自動」。範囲 0〜100% は描画側でクランプする
        outerPadding:
            typeof col.outerPadding.value === "number" && isFinite(col.outerPadding.value)
                ? clampPercent(col.outerPadding.value)
                : null,
        // 範囲: 間隔 0〜50%、系列間 0〜90%、角丸 0〜30px、罫線 1〜5px。範囲は描画側でクランプする
        categorySpacing: Math.max(0, Math.min(50, col.categorySpacing.value ?? 20)),
        // 重複のときは重ねる割合なので 0〜100%、並べるときはすき間で 0〜90%
        seriesSpacing: Math.max(0, Math.min(overlapOn ? 100 : 90, col.seriesSpacing.value ?? 0)),
        overlap: overlapOn,
        overlapReverse: overlapOn && (col.overlapReverse.value ?? false),
        // 標準と同じく、展開しないときは 0〜5 px、展開するときは 0〜10 %
        stackedSpacing: Math.max(0, Math.min(explodeOn ? 10 : 5, Number(col.stackedSpacing.value) || 0)),
        stackedExplode: explodeOn,
        borderOutlineOnly: stacked && (col.borderOutlineOnly.value ?? false),
        maxBarWidth: Math.max(0, col.maxBarWidth.value ?? 0),
        cornerRadius: Math.max(0, Math.min(30, col.cornerRadius.value ?? 0)),
    };

    // 凡例（複数系列のときだけ描く）。値が複数で凡例のフィールドが無いときは、標準と同じくタイトルを出さない
    const lg = settings.legend;
    const legendInfo: LegendInfo = {
        show: seriesMode && (lg.show.value ?? true),
        markerStyle: getDropdownValue(lg.markerStyle.value, "lineAndMarker"),
        matchLineColor: lg.matchLineColor.value ?? false,
        position: legendPlacementValue(lg.position.value?.value),
        title: (lg.titleShow.value ?? true) ? (lg.titleText.value?.trim() || (legendSource?.displayName ?? "")) : "",
        titleShow: lg.titleShow.value ?? true,
        fontFamily: lg.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, lg.font.fontSize.value ?? 10)),
        bold: lg.font.bold?.value ?? false,
        italic: lg.font.italic?.value ?? false,
        underline: lg.font.underline?.value ?? false,
        color: lg.labelColor.value?.value || "#605E5C",
    };

    // 系列。複数系列の色は、個別の指定が無ければ標準と同じくレポートのテーマの色を順に割り当てる。
    // 系列 1 本のときは 1.4 までと同じく「列」のカラー（カテゴリごとの指定があればそちら）
    /** 書式ペインに出す系列の色（「棒の色」で塗っても、ペインには自分で選んだ色かテーマの色を出す） */
    const ownSeriesColors: string[] = [];
    const series: SeriesInfo[] = slots.map((slot) => {
        if (!seriesMode) return { name: slot.name, color: columnsSettings.fill, selectionId: null };
        // テーマの色は、個別の指定がある系列でも必ず取る。Desktop の colorPalette は呼んだ順に色を割り当てるので、
        // 取らないと後ろの系列の色が 1 つ前にずれる（標準では、ある系列の色を変えても他の系列の色は変わらない）
        const themeColor = host.colorPalette.getColor(slot.key).value;
        ownSeriesColors.push(firstOf(slot.objects, (o) => customColor(o, "fill")) ?? themeColor);
        // 「棒の色」があれば、凡例の色はその系列で最初に色が入っている行の色（標準の fx と同じ）
        let fieldColor: string | null = null;
        for (let i = 0; i < rowCount && fieldColor === null; i++) fieldColor = dataColorAt(slot.colorColumn, i);
        const color = fieldColor ?? ownSeriesColors[ownSeriesColors.length - 1];
        const selectionId = legendSource
            ? withSeriesOf(host.createSelectionIdBuilder(), categorical!, slot.group).createSelectionId()
            : host.createSelectionIdBuilder().withMeasure(slot.column.source.queryName).createSelectionId();
        return { name: slot.name, color, selectionId };
    });

    /** 書式ペインで系列を指す selector。凡例の系列は系列の ID、メジャーの系列はメジャーの queryName */
    const seriesSelector = (s: number): powerbi.data.Selector =>
        legendSource ? series[s].selectionId!.getSelector() : { metadata: slots[s].column.source.queryName };

    // 系列ごとの見た目（複数系列のとき）。個別の指定が無ければ「すべて」の値
    const seriesStyles = slots.map((slot, s) => {
        const borderMatchColumn = firstOf(slot.objects, (o) => customFlag(o, "borderMatchColumn")) ?? columnsSettings.borderMatchColumn;
        const borderColor = firstOf(slot.objects, (o) => customColor(o, "borderFill")) ?? columnsSettings.borderFill;
        return {
            transparency: clampPercent(firstOf(slot.objects, (o) => customNumber(o, "transparency")) ?? columnsSettings.transparency),
            borderShow: firstOf(slot.objects, (o) => customFlag(o, "showBorder")) ?? columnsSettings.showBorder,
            borderMatchColumn,
            borderColor: borderMatchColumn ? series[s].color : borderColor,
            borderTransparency: clampPercent(firstOf(slot.objects, (o) => customNumber(o, "borderTransparency")) ?? columnsSettings.borderTransparency),
            borderWidth: clampBorderWidth(firstOf(slot.objects, (o) => customNumber(o, "borderWidth")) ?? columnsSettings.borderWidth),
            labelShow: seriesMode ? (firstOf(slot.objects, labelFlag) ?? true) : true,
            explicitLabelColor: seriesMode ? firstOf(slot.objects, labelColorOf) : null,
        };
    });

    // 合計ラベル（積み上げのときだけ描く）。表示単位は「自動」なら Y 軸に従い（単位ラベルがあるので語は付けない）、
    // 選んだときはその単位で割って語を付ける
    const tl = settings.totalLabels;
    const totalPrecisionValue = String(tl.precision.value?.value ?? "auto");
    const totalPrecision = totalPrecisionValue !== "auto" ? totalPrecisionValue : precision;
    const totalUnitKey = getDropdownValue(tl.unitType.value, "auto");
    const totalUnitDef = totalUnitKey === "auto" ? null : resolveUnit(totalUnitKey, maxAbs, unitNotation, totalPrecision);
    const totalText = (v: number): string =>
        percentMeasure
            ? measurePercentText(v, totalPrecision)
            : totalUnitDef
            ? `${formatValue(v, totalUnitDef.divisor, totalPrecision)}${totalUnitDef.unitWord}`
            : formatValue(v, unitDef.divisor, totalPrecision);
    const totalLabelsSettings: TotalLabelsSettings = {
        show: chartType === CHART_TYPES.stacked && (tl.show.value ?? false),
        split: tl.splitPositiveNegative.value ?? false,
        fontFamily: tl.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, tl.font.fontSize.value ?? 9)),
        bold: tl.font.bold?.value ?? false,
        italic: tl.font.italic?.value ?? false,
        underline: tl.font.underline?.value ?? false,
        color: tl.color.value?.value || "#605E5C",
        backgroundShow: tl.backgroundShow.value ?? false,
        backgroundColor: tl.backgroundColor.value?.value || "#FFFFFF",
        backgroundTransparency: Math.max(0, Math.min(100, tl.backgroundTransparency.value ?? 0)),
    };

    // リボンの帯（積み上げ・100% 積み上げで、リボンがオンのとき。縦棒・横棒とも）
    const rb = settings.ribbons;
    const ribbonSettings: RibbonSettings = {
        show: ribbonsOn,
        matchSeriesColor: rb.matchSeriesColor.value ?? true,
        fill: rb.fill.value?.value || "#C8C6C4",
        transparency: clampPercent(rb.transparency.value ?? 30),
        borderShow: rb.borderShow.value ?? false,
        borderMatchRibbon: rb.borderMatchRibbon.value ?? false,
        borderFill: rb.borderFill.value?.value || "#605E5C",
        borderTransparency: clampPercent(rb.borderTransparency.value ?? 0),
        borderWidth: Math.max(1, Math.min(10, rb.borderWidth.value ?? 1)),
        spacing: Math.max(0, Math.min(45, rb.spacing.value ?? 0)),
    };

    const labelPrecision = dataLabelsSettings.precision !== "auto" ? dataLabelsSettings.precision : precision;
    // データラベルのマイナスと 0 の書き方（▲・±0 など）
    const labelSign: Partial<SignStyle> = {
        negative: getDropdownValue(dl.negativeStyle.value, NEGATIVE_STYLES.minus),
        zero: getDropdownValue(dl.zeroStyle.value, ZERO_STYLES.zero),
        negativeZero: dl.negativeZero.value ?? true,
    };
    // 値の行を符号で塗る（見える符号で決める。▲0 はマイナス、±0 は塗らない）
    const toneMode = getDropdownValue(dl.toneMode.value, TONE_MODES.none);
    const toneColors = {
        good: dl.positiveColor.value?.value || DEFAULT_GOOD_COLOR,
        bad: dl.negativeColor.value?.value || DEFAULT_BAD_COLOR,
    };
    const toneColorOf = (sign: number): string => {
        const tone = toneOf(sign, 1, toneMode);
        return tone ? toneColors[tone] : "";
    };
    const formatted = (val: number) =>
        // 100% 積み上げの軸は割合なので、データラベルは値ごとに単位を付ける（1,250億 など）
        percentMeasure && !isLogScaleActive
            ? {
                formattedValue: measurePercentText(val, precision),
                dataLabelText: measurePercentText(val, labelPrecision, labelSign),
                labelToneColor: toneColorOf(shownSignOf(val * 100, 1, labelPrecision === "auto" ? String(measurePercentDigits) : labelPrecision, labelSign)),
            }
            : isLogScaleActive || percent
            ? {
                formattedValue: formatDynamicValue(val, unitNotation, precision, true),
                dataLabelText: formatDynamicValue(val, unitNotation, labelPrecision, true, labelSign),
                labelToneColor: toneColorOf(dynamicShownSign(val, unitNotation, labelPrecision, labelSign)),
            }
            : {
                formattedValue: formatValue(val, unitDef.divisor, precision),
                dataLabelText: formatSigned(val, unitDef.divisor, labelPrecision, labelSign),
                labelToneColor: toneColorOf(shownSignOf(val, unitDef.divisor, labelPrecision, labelSign)),
            };

    // データラベルの詳細の行。全体に対する割合は、100% 積み上げと複数系列ではカテゴリの合計（正と負の絶対値）に対する割合、
    // 系列 1 本ではすべてのカテゴリの合計に対する割合。カスタムは「ラベルの詳細」に入れたフィールドの値
    const detailContent = getDropdownValue(dl.detailContent.value, DETAIL_CONTENTS.percentOfTotal);
    const detailPrecision = getDropdownValue(dl.detailPrecision.value, "auto");
    const detailUnitKey = getDropdownValue(dl.detailUnitType.value, "auto");
    const grandAbsolute = absoluteSums.reduce((sum, v) => sum + v, 0);
    // 自動は標準と同じく小数 2 桁（30.00%）
    const percentDigits = detailPrecision === "auto" ? 2 : Number(detailPrecision);
    const percentText = (ratio: number) =>
        `${(ratio * 100).toLocaleString("ja-JP", {
            minimumFractionDigits: percentDigits,
            maximumFractionDigits: percentDigits,
        })}%`;
    const fieldFormatters = new Map<DataViewValueColumn, ReturnType<typeof valueFormatter.create>>();
    const detailTextOf = (slot: SeriesSlot, i: number, val: number, skipped: boolean): string => {
        if (skipped) return "";
        if (detailContent !== DETAIL_CONTENTS.custom) {
            const denominator = seriesMode || percent ? absoluteSums[i] : grandAbsolute;
            return percentText(denominator > 0 ? val / denominator : 0);
        }
        // 棒はあるが詳細のフィールドが空白：標準の「空白の表示方法」の文字（空なら詳細の行を出さない）
        return fieldText(slot.detail, i, detailUnitKey, detailPrecision, dataLabelsSettings.detailShowBlankAs);
    };

    /**
     * ラベルの欄（タイトル・値・詳細）に入れたフィールドの i 行目の文字。空白なら blankAs。
     * 表示単位を選べばその単位で、「自動」ならフィールドの書式のまま（小数点以下の桁数を選んだら、その桁で出す）
     */
    function fieldText(column: DataViewValueColumn | undefined, i: number, unitKey: string, digits: string, blankAs: string): string {
        if (!column) return "";
        const raw = column.values[i];
        if (raw === null || raw === undefined || raw === "") return blankAs;
        if (typeof raw !== "number") return String(raw);
        if (unitKey !== "auto") {
            const unit = resolveUnit(unitKey, Math.abs(raw), unitNotation, digits);
            return `${formatValue(raw, unit.divisor, digits)}${unit.unitWord}`;
        }
        const format = valueFormatter.getFormatStringByColumn(column.source) ?? "";
        if (digits !== "auto") {
            const fraction = Number(digits);
            return /%/.test(format)
                ? `${(raw * 100).toLocaleString("ja-JP", { minimumFractionDigits: fraction, maximumFractionDigits: fraction })}%`
                : formatValue(raw, 1, digits);
        }
        let formatter = fieldFormatters.get(column);
        if (!formatter) {
            formatter = valueFormatter.create({ format });
            fieldFormatters.set(column, formatter);
        }
        return formatter.format(raw);
    }

    // データ ラベルのタイトルの行。系列名か、「ラベルのタイトル」に入れたフィールドの値
    const titleCustom = getDropdownValue(dl.titleContent.value, TITLE_CONTENTS.seriesName) === TITLE_CONTENTS.custom;
    const titleUnitKey = getDropdownValue(dl.titleUnitType.value, "auto");
    const titlePrecision = getDropdownValue(dl.titlePrecision.value, "auto");
    const titleTextOf = (slot: SeriesSlot, i: number, skipped: boolean): string => {
        if (skipped) return "";
        if (!titleCustom) return slot.name;
        // 「その他」の行はフィールドの値をまとめられないので出さない
        if (i === otherRow) return "";
        return fieldText(slot.title, i, titleUnitKey, titlePrecision, dataLabelsSettings.titleShowBlankAs);
    };

    // 「ラベルの値」：棒の値の代わりに出すフィールドの値。値の行の小数点以下の桁数で出す（表示単位はフィールドの書式のまま）。
    // 空白なら「空白の表示方法」の文字、それも空なら棒の値。「その他」の行は棒の値
    const valueFieldTextOf = (slot: SeriesSlot, i: number): string | null => {
        if (!slot.value || i === otherRow) return null;
        const text = fieldText(slot.value, i, "auto", labelPrecision, dataLabelsSettings.valueShowBlankAs);
        return text === "" ? null : text;
    };

    // データポイントとカテゴリ別ターゲットの生成
    const categoryGroups: CategoryGroup[] = [];
    /** 系列の展開のときだけ、積んだ棒の始まりと量（値の単位）を残し、あとで置き直す */
    const stackSegments = new Map<DataPoint, StackSegment>();
    const columnTargets: ColumnTarget[] = [];

    for (let i = 0; i < rowCount; i++) {
        const category = categoryText(i);
        const targetObjects = categoryObjectsAt(i);

        // 系列 1 本のときの、カテゴリごとの見た目（1.4 までと同じ）。「棒の色」があればそれを優先する
        const categoryOwnFill = customColor(targetObjects, "fill") ?? paretoFillAt(i) ?? columnsSettings.fill;
        const categoryFill =
            i === otherRow
                ? settings.columns.otherFill.value?.value || "#A0A0A0"
                : (!seriesMode ? dataColorAt(slots[0]?.colorColumn, i) : null) ?? categoryOwnFill;
        const categoryTransparency = clampPercent(customNumber(targetObjects, "transparency") ?? columnsSettings.transparency);
        const categoryShowBorder = customFlag(targetObjects, "showBorder") ?? columnsSettings.showBorder;
        const categoryBorderMatchColumn = customFlag(targetObjects, "borderMatchColumn") ?? columnsSettings.borderMatchColumn;
        let categoryBorderFill = customColor(targetObjects, "borderFill") ?? columnsSettings.borderFill;
        if (categoryBorderMatchColumn) {
            categoryBorderFill = categoryFill;
        }
        const categoryBorderTransparency = clampPercent(customNumber(targetObjects, "borderTransparency") ?? columnsSettings.borderTransparency);
        const categoryBorderWidth = clampBorderWidth(customNumber(targetObjects, "borderWidth") ?? columnsSettings.borderWidth);

        // 積み上げ: 正の値は 0 から上へ、負の値は 0 から下へ、凡例の順に積む（標準と同じ）。
        // 100% 積み上げは、正と負の絶対値の合計に対する割合で積む
        const absoluteSum = absoluteSums[i];
        let positiveEnd = 0;
        let negativeEnd = 0;
        // 積み上げ順を逆にする：凡例の後ろの系列から 0 の側に積む。始まりを先に決めておく
        const reversedStarts = new Map<number, number>();
        if (reverseStack) {
            let up = 0;
            let down = 0;
            for (let s = slots.length - 1; s >= 0; s--) {
                const slot = slots[s];
                const v = numberAt(slot.column, i);
                const amount = isSkipped(slot.column, i) ? 0 : percent ? (absoluteSums[i] > 0 ? v / absoluteSums[i] : 0) : v;
                reversedStarts.set(s, amount >= 0 ? up : down);
                if (amount >= 0) up += amount;
                else down += amount;
            }
        }

        // リボン：順位は値の大きい順（1 が最大、帯のツールヒント用）。
        // 積む順が値の大きい順なら、カテゴリの中で値の小さい順に 0 から積む（いちばん大きい系列が外側に来る）。
        // 負の値は 0 に近い順に外へ積む。100% 積み上げは割合で積む
        const ribbonStarts = new Map<number, number>();
        const ribbonRanks = new Map<number, number>();
        if (ribbonsOn) {
            const present = slots
                .map((slot, s) => ({ s, v: isSkipped(slot.column, i) ? null : numberAt(slot.column, i) }))
                .filter((e): e is { s: number; v: number } => e.v !== null);
            [...present].sort((a, b) => b.v - a.v).forEach((e, k) => ribbonRanks.set(e.s, k + 1));
            if (rankOrder) {
                const amountOf = (v: number) => (percent ? (absoluteSum > 0 ? v / absoluteSum : 0) : v);
                let up = 0;
                let down = 0;
                present.filter((e) => e.v >= 0).sort((a, b) => a.v - b.v).forEach((e) => {
                    ribbonStarts.set(e.s, up);
                    up += amountOf(e.v);
                });
                present.filter((e) => e.v < 0).sort((a, b) => b.v - a.v).forEach((e) => {
                    ribbonStarts.set(e.s, down);
                    down += amountOf(e.v);
                });
            }
        }

        const points: DataPoint[] = slots.map((slot, s) => {
            const val = numberAt(slot.column, i);
            const skipped = isSkipped(slot.column, i);
            const share = percent ? (absoluteSum > 0 ? val / absoluteSum : 0) : null;
            // 積む量（100% なら割合）と、この棒の始まり・終わり
            const amount = skipped ? 0 : percent ? share! : val;
            let start = 0;
            if (rankOrder) {
                start = ribbonStarts.get(s) ?? 0;
            } else if (reverseStack) {
                start = reversedStarts.get(s) ?? 0;
            } else if (stacked) {
                start = amount >= 0 ? positiveEnd : negativeEnd;
                if (amount >= 0) positiveEnd += amount;
                else negativeEnd += amount;
            }
            const end = stacked ? start + amount : val;
            // 系列 1 本の ID はカテゴリだけ（1.4 までと同じ。カテゴリごとの色の保存先がこの ID の selector）
            const builder = withCategories(host.createSelectionIdBuilder(), i);
            const selectionId = !seriesMode
                ? builder.createSelectionId()
                : legendSource
                    ? withSeriesOf(builder, categorical!, slot.group).createSelectionId()
                    : builder.withMeasure(slot.column.source.queryName).createSelectionId();

            // ハイライトは該当しない行が null。数値で持ち、棒の高さは軸と同じ calcRatio で出す。
            // 軸の範囲・目盛りは全体の値だけで決め、ハイライトでは動かさない（標準と同じ）
            const rawHighlight = slot.column.highlights?.[i];
            const highlightNumber = rawHighlight === null || rawHighlight === undefined ? NaN : Number(rawHighlight);
            const highlight = isFinite(highlightNumber) ? highlightNumber : null;

            // 積み上げのハイライトは、その棒の始まりから該当分だけ伸ばす
            const highlightEnd =
                highlight === null ? null : !stacked ? highlight : start + (percent ? (absoluteSum > 0 ? highlight / absoluteSum : 0) : highlight);

            const style = seriesStyles[s];
            const insetStart = stacked && !skipped && start !== 0;
            const pointColor = seriesMode ? (i === otherRow ? null : dataColorAt(slot.colorColumn, i)) ?? series[s].color : categoryFill;
            const point: DataPoint = {
                category,
                rowIndex: i,
                seriesIndex: s,
                blank: skipped,
                value: val,
                startRatio: stacked ? calcRatio(start) : zeroRatio,
                valRatio: calcRatio(end),
                share,
                outermost: true,
                insetStart,
                rank: ribbonsOn ? ribbonRanks.get(s) ?? null : null,
                ...formatted(val),
                ...(() => {
                    const text = valueFieldTextOf(slot, i);
                    return text === null ? {} : { dataLabelText: text, labelToneColor: "" };
                })(),
                titleText: titleTextOf(slot, i, skipped),
                detailText: detailTextOf(slot, i, val, skipped),
                selectionId,
                // 「その他」の棒を押したら、まとめたカテゴリを全部選ぶ
                ...(i === otherRow ? { selectionIds: otherSelectionIds } : {}),
                // パレートのランクの帯はカテゴリで選ぶ。パレートのときだけ持たせる（ほかの計算の選択の見た目を変えない）
                ...(paretoOn ? { categorySelectionId: seriesMode ? withCategories(host.createSelectionIdBuilder(), i).createSelectionId() : selectionId } : {}),
                highlight,
                highlightRatio: highlightEnd === null ? null : calcRatio(highlightEnd),
                color: pointColor,
                transparency: seriesMode ? style.transparency : categoryTransparency,
                borderShow: seriesMode ? style.borderShow : categoryShowBorder,
                borderColor: seriesMode ? (style.borderMatchColumn ? pointColor : style.borderColor) : categoryBorderFill,
                borderTransparency: seriesMode ? style.borderTransparency : categoryBorderTransparency,
                borderWidth: seriesMode ? style.borderWidth : categoryBorderWidth,
                labelShow: style.labelShow,
                labelColor: style.explicitLabelColor ?? dataLabelsSettings.color,
            };
            if (explodeOn) {
                const highlightAmount = highlight === null ? null : percent ? (absoluteSum > 0 ? highlight / absoluteSum : 0) : highlight;
                stackSegments.set(point, { start, amount, highlightAmount });
            }
            return point;
        });

        // 積み上げでは、角丸は正と負それぞれいちばん外側の棒だけに付ける（値の大きい順に積むときは、最大と最小）
        if (stacked) {
            const pick = (want: (d: DataPoint) => boolean, better: (a: DataPoint, b: DataPoint) => boolean) =>
                points.reduce((best, d, k) => (want(d) && (best < 0 || better(d, points[best])) ? k : best), -1);
            // 外側の端の棒：凡例の順に積むなら最後の棒、逆に積むなら最初の棒
            const outerOf = (flags: boolean[]) => (reverseStack ? flags.indexOf(true) : flags.lastIndexOf(true));
            const lastPositive = rankOrder
                ? pick((d) => !d.blank && d.value > 0, (a, b) => a.value >= b.value)
                : outerOf(points.map((d) => !d.blank && d.value > 0));
            const lastNegative = rankOrder
                ? pick((d) => !d.blank && d.value < 0, (a, b) => a.value <= b.value)
                : outerOf(points.map((d) => !d.blank && d.value < 0));
            points.forEach((d, k) => (d.outermost = k === lastPositive || k === lastNegative));
        }

        const totals: StackTotals | null =
            chartType === CHART_TYPES.stacked
                ? {
                    net: positiveSums[i] + negativeSums[i],
                    positive: positiveSums[i],
                    negative: negativeSums[i],
                    hasPositive: positiveSums[i] > 0,
                    hasNegative: negativeSums[i] < 0,
                    positiveEndRatio: calcRatio(positiveSums[i]),
                    negativeEndRatio: calcRatio(negativeSums[i]),
                    netText: totalText(positiveSums[i] + negativeSums[i]),
                    positiveText: totalText(positiveSums[i]),
                    negativeText: totalText(negativeSums[i]),
                }
                : null;

        categoryGroups.push({ category, label: shownLevelsAt(i).join(" "), levels: shownLevelsAt(i), levelKeys: shownLevelKeysAt(i), rowIndex: i, points, totals });

        // 「その他」は書式の対象にしない（ID はまとめた最初のカテゴリの仮のもので、そこに保存するとそのカテゴリの書式になる）
        if (!seriesMode && i !== otherRow) {
            columnTargets.push({
                name: category,
                selector: points[0].selectionId.getSelector(),
                color: categoryOwnFill,
                transparency: categoryTransparency,
                borderShow: categoryShowBorder,
                borderMatchColumn: categoryBorderMatchColumn,
                borderColor: categoryBorderFill,
                borderTransparency: categoryBorderTransparency,
                borderWidth: categoryBorderWidth,
            });
        }
    }

    // 系列の展開：すき間を入れても全体がプロットに収まるように置き直し、値の軸とその目盛線を消す（標準と同じ）
    if (explodeOn) {
        // 合計ラベルを出すときは、棒の外のラベルが切れないよう、上下に少し空ける
        explodeStacks(categoryGroups, stackSegments, columnsSettings.stackedSpacing / 200, totalLabelsSettings.show ? 0.08 : 0);
        valueAxisSettings.show = false;
        gridlinesSettings.horizontalShow = false;
    }

    // 複数系列では「設定の適用先」に系列を並べる（標準と同じ）
    const labelTargets: LabelTarget[] = [];
    if (seriesMode) {
        slots.forEach((slot, s) => {
            const style = seriesStyles[s];
            const selector = seriesSelector(s);
            columnTargets.push({
                name: slot.name,
                selector,
                color: ownSeriesColors[s],
                transparency: style.transparency,
                borderShow: style.borderShow,
                borderMatchColumn: style.borderMatchColumn,
                borderColor: style.borderColor,
                borderTransparency: style.borderTransparency,
                borderWidth: style.borderWidth,
            });
            labelTargets.push({
                name: slot.name,
                selector,
                show: style.labelShow,
                color: style.explicitLabelColor ?? "",
            });
        });
    }

    // レイアウトの並び替え。順は「並び順」で元の値から決めてある（累計した値では並べ替えない）
    categoryGroups.sort((a, b) => positionOf.get(a.rowIndex)! - positionOf.get(b.rowIndex)!);
    const dataPoints = categoryGroups.flatMap((g) => g.points);

    // --- 折れ線と第 2 Y 軸 ---------------------------------------------------
    const v2 = settings.valueAxis2;
    const parseOptional = (raw: string | undefined | null): number | undefined => boundOf(raw) ?? undefined;
    const v2Precision = getDropdownValue(v2.precision.value, "auto");
    // 表示単位・単位ラベル・タイトルのスタイルは Y 軸と同じ。率のメジャー（書式に % がある）は、
    // 100% 積み上げの Y 軸と同じく表示単位によらず % で出し、単位の語を付けない
    const isPercentLine = /%/.test(lineDefs[0]?.format ?? "");
    const v2Notation = getDropdownValue(v2.unitNotation.value, "japanese");
    const v2ShowUnitOnAxis = v2.showUnitOnAxis.value ?? false;
    const v2UnitText = v2.unitText.value ?? "";
    const v2IncludeDisplayUnit = v2.unitIncludeDisplayUnit.value ?? true;
    const unitDef2 = isPercentLine
        ? resolveUnit("0", 0, v2Notation, v2Precision)
        : resolveUnit(
            getDropdownValue(v2.unitType.value, "auto"),
            Math.max(Math.abs(lineMin), Math.abs(lineMax)),
            v2Notation,
            v2Precision
        );
    // 第 2 Y 軸の範囲：対数目盛り・範囲を丸める・0 を配置する（標準の第 2 Y 軸と同じ項目）。
    // 対数は、折れ線の値がすべて正かすべて負で、0 を配置しないときだけ効かせる（標準も 0 を配置すると対数は押せない）
    const alignZeros2 = v2.alignZeros.value ?? false;
    const logSign2 = lineMin > 0 ? 1 : lineMax < 0 ? -1 : 0;
    const log2Active = onSecondary && (v2.logarithmic.value ?? false) && logSign2 !== 0 && !alignZeros2;
    // 対数では、全体一律の表示単位の語を単位ラベルとタイトルに付けない（Y 軸と同じ）
    const unitWord2 = log2Active ? "" : unitDef2.unitWord;
    let calcRatio2 = calcRatio;
    /** 第 2 Y 軸の範囲（値）。Y 軸の定数線が範囲の内かを見る。第 2 Y 軸が無ければ null */
    let range2: [number, number] | null = null;
    let ticks2: Tick[] = [];
    /** 第 2 Y 軸の目盛りを、Y 軸と同じ本数の上限で作り直す（対数は変えない） */
    let ticks2For: (count: number) => Tick[] = () => ticks2;
    if (onSecondary) {
        const start2 = parseOptional(v2.start.value);
        const end2 = parseOptional(v2.end.value);
        const roundRange2 = v2.roundRange.value ?? true;
        let raw2: number[];
        if (log2Active) {
            const mags = lineNumbers.map((v) => Math.abs(v));
            // 最小値・最大値の指定は大きさに直す。符号が折れ線と合わない指定は使わない（Y 軸と同じ）
            const userLower = logSign2 > 0
                ? (start2 !== undefined && start2 > 0 ? start2 : undefined)
                : (end2 !== undefined && end2 < 0 ? -end2 : undefined);
            const userUpper = logSign2 > 0
                ? (end2 !== undefined && end2 > 0 ? end2 : undefined)
                : (start2 !== undefined && start2 < 0 ? -start2 : undefined);
            const lower = userLower ?? Math.pow(10, Math.ceil(log10Snap(Math.min(...mags))) - 1);
            let upper = userUpper ?? (roundRange2 ? Math.pow(10, Math.ceil(log10Snap(Math.max(...mags)))) : Math.max(...mags));
            if (!(upper > lower)) upper = lower * 10;
            const logMin = Math.log10(lower);
            const logSpan = Math.log10(upper) - logMin;
            calcRatio2 = (v: number) => {
                const mag = Math.abs(v);
                if (mag <= 0) return logSign2 > 0 ? 0 : 1;
                const r = logSpan > 0 ? Math.max(0, Math.min(1, (Math.log10(mag) - logMin) / logSpan)) : 0;
                return logSign2 > 0 ? r : 1 - r;
            };
            raw2 = logAxisTicks(lower, upper).map((m) => logSign2 * m);
            range2 = logSign2 > 0 ? [lower, upper] : [-upper, -lower];
        } else {
            // 標準と同じく、第 2 Y 軸は 0 から始めず折れ線の範囲に合わせる（率が 90〜113% なら 90% あたりから）
            let min2 = start2 ?? lineMin;
            let max2 = end2 ?? lineMax;
            // 0 を配置する：0 を含めたうえで、その高さを Y 軸の 0 にそろえる（Y 軸が対数のときは 0 が無いのでそろえない）
            const aligning = alignZeros2 && !isLogScaleActive;
            if (aligning) {
                min2 = Math.min(min2, 0);
                max2 = Math.max(max2, 0);
            }
            if (!(max2 > min2)) max2 = min2 + (Math.abs(min2) || 1) * 0.1;
            const scale2 = scaleLinear().domain([min2, max2]);
            if (roundRange2 && start2 === undefined && end2 === undefined) scale2.nice();
            let [lo2, hi2] = scale2.domain();
            if (aligning) [lo2, hi2] = alignZeroAt(lo2, hi2, invertRange ? 1 - zeroRatio : zeroRatio);
            const span2 = hi2 - lo2;
            range2 = [lo2, hi2];
            calcRatio2 = (v: number) => (span2 > 0 ? Math.max(0, Math.min(1, (v - lo2) / span2)) : 0);
            raw2 = scaleLinear().domain([lo2, hi2]).ticks(5);
            const domain2: [number, number] = [lo2, hi2];
            ticks2For = (count) => ticks2Of(ticksUpTo(domain2, count, (values) => ticks2Of(values).map((tick) => tick.label)));
        }
        ticks2 = ticks2Of(raw2);
    }
    function ticks2Of(raw: number[]): Tick[] {
        const step2 = raw.length > 1 ? Math.abs(raw[1] - raw[0]) : 1;
        const percentDigits = v2Precision !== "auto" ? Number(v2Precision) : step2 * 100 < 1 ? 1 : 0;
        return raw.map((t) => {
            let label: string;
            if (isPercentLine) {
                label = `${(t * 100).toFixed(percentDigits)}%`;
            } else if (log2Active) {
                // 対数は目盛りごとに桁の語を付ける（Y 軸と同じ）
                label = formatDynamicValue(t, v2Notation, v2Precision, v2ShowUnitOnAxis);
            } else {
                // Y 軸と同じく、目盛りは数字だけ（「軸ラベルに単位を表示」で 1.5万 のように付ける）
                const num = formatValue(t, unitDef2.divisor, v2Precision);
                label = v2ShowUnitOnAxis && unitDef2.unitWord ? `${num}${unitDef2.unitWord}` : num;
            }
            return { value: t, label, ratio: calcRatio2(t) };
        });
    }
    if (!onSecondary && paretoOn) {
        // パレートの累積比の軸。0〜100% に固定する（第 2 Y 軸の範囲の指定は使わない）
        calcRatio2 = (v: number) => Math.max(0, Math.min(1, v));
        const percentTicks = (values: number[]) => values.map((t) => ({ value: t, label: `${Math.round(t * 100)}%`, ratio: t }));
        ticks2 = percentTicks([0, 0.2, 0.4, 0.6, 0.8, 1]);
        ticks2For = (count) => percentTicks(ticksUpTo([0, 1], count, (values) => values.map((v) => `${Math.round(v * 100)}%`)));
    }

    // --- Y 軸の定数線 ---------------------------------------------------------
    // 載せた軸の範囲の内の線だけを入れる（標準と同じく、範囲の外は描かない）。値は載せた軸の表示単位で書く
    const range1: [number, number] = isLogScaleActive ? (logSign > 0 ? [logLower, logUpper] : [-logUpper, -logLower]) : [niceMin, niceMax];
    const vlPrecision = getDropdownValue(vl.labelPrecision.value, "auto");
    const vlLabelText = getDropdownValue(vl.labelText.value, "value");
    const valueLines: ValueLineInfo[] = valueLineDefs.flatMap((def) => {
        const secondary = valueLineOnSecondary && range2 !== null;
        const [lo, hi] = secondary ? range2! : range1;
        const eps = Math.abs(hi - lo) * 1e-9;
        if (!(def.value >= lo - eps && def.value <= hi + eps)) return [];
        const p1 = vlPrecision !== "auto" ? vlPrecision : precision;
        const p2 = vlPrecision !== "auto" ? vlPrecision : v2Precision;
        const valueText = secondary
            ? isPercentLine
                ? `${(def.value * 100).toFixed(p2 === "auto" ? 1 : Number(p2))}%`
                : formatValue(def.value, unitDef2.divisor, p2)
            : percent
                ? `${(def.value * 100).toFixed(p1 === "auto" ? 0 : Number(p1))}%`
                : percentMeasure
                ? measurePercentText(def.value, p1)
                : isLogScaleActive
                    ? formatDynamicValue(def.value, unitNotation, p1, true)
                    : formatValue(def.value, unitDef.divisor, p1);
        const label = vlLabelText === "name" ? def.name : vlLabelText === "both" ? `${def.name} ${valueText}` : valueText;
        return [{ name: def.name, value: def.value, ratio: secondary ? calcRatio2(def.value) : calcRatio(def.value), secondary, label }];
    });
    const valueLine: ValueLineSettings = {
        show: valueLineOn,
        color: vl.color.value?.value || "#605E5C",
        transparency: Math.max(0, Math.min(100, vl.transparency.value ?? 0)),
        lineStyle: getDropdownValue(vl.lineStyle.value, "dashed"),
        dashArray: vl.dashArray.value ?? "",
        scaleWithWidth: vl.scaleWithWidth.value ?? true,
        dashCap: getDropdownValue(vl.dashCap.value, "none"),
        width: Math.max(1, Math.min(10, vl.width.value ?? 1)),
        layer: getDropdownValue(vl.layer.value, "front") === "back" ? "back" : "front",
        shadeShow: vl.shadeShow.value ?? false,
        shadeRegion: getDropdownValue(vl.shadeRegion.value, "below") === "above" ? "above" : "below",
        shadeColor: (vl.shadeMatchLine.value ?? false) ? vl.color.value?.value || "#605E5C" : vl.shadeColor.value?.value || "#E1DFDD",
        shadeTransparency: Math.max(0, Math.min(100, vl.shadeTransparency.value ?? 40)),
        labelShow: vl.labelShow.value ?? true,
        labelHorizontal: getDropdownValue(vl.labelHorizontal.value, "left") === "right" ? "right" : "left",
        labelVertical: getDropdownValue(vl.labelVertical.value, "top") === "bottom" ? "bottom" : "top",
        labelColor: vl.labelColor.value?.value || "#605E5C",
        labelFontSize: Math.max(6, Math.min(32, vl.labelFontSize.value ?? 9)),
        labelAvoidOverlap: vl.labelAvoidOverlap.value ?? true,
    };

    // 折れ線の色は、棒の系列の続きのテーマの色（系列 1 本の棒はテーマの 1 番目を使う扱いにして、線は 2 番目から）
    if (lineDefs.length && !seriesMode) host.colorPalette.getColor(slots[0].key);
    const lineCard = settings.lines;
    const defaultLineStyle = getDropdownValue(lineCard.lineStyle.value, LINE_STYLES.solid);
    const defaultShape = lineCard.shapeValues();
    const areaCard = settings.areas;
    const areasOn = areaCard.show.value ?? false;
    // 網掛け領域の下の辺は値 0 の高さ（軸の範囲の外なら端。対数なら端）
    const baselineRatio = onSecondary ? calcRatio2(0) : calcRatio(0);
    const lines: LineSeriesInfo[] = lineDefs.map((def) => {
        const own = def.objects?.lines;
        const ownFill = (own?.fill as powerbi.Fill | undefined)?.solid?.color;
        const ownWidth = typeof own?.width === "number" ? own.width : null;
        const ownStyle = own?.lineStyle;
        const ownText = (property: string, fallback: string): string => {
            const raw = own?.[property];
            return raw !== undefined && raw !== null ? String(raw) : fallback;
        };
        const ownAreaShow = def.objects?.areas?.show;
        const ownLineShow = own?.show;
        const measureBuilder = () => host.createSelectionIdBuilder();
        // 棒の系列と同じく、テーマの色は色の指定があっても取る（後ろの線の色がずれないように）
        const themeColor = host.colorPalette.getColor(def.key).value;
        return {
            name: def.name,
            color: ownFill ? String(ownFill) : themeColor,
            width: Math.max(1, Math.min(10, ownWidth ?? lineCard.width.value ?? 3)),
            lineStyle: ownStyle !== undefined && ownStyle !== null ? String(ownStyle) : defaultLineStyle,
            dashArray: ownText("dashArray", lineCard.dashArray.value ?? ""),
            dashCap: ownText("dashCap", String(lineCard.dashCap.value?.value ?? "none")),
            scaleWithWidth: typeof own?.scaleWithWidth === "boolean" ? own.scaleWithWidth : lineCard.scaleWithWidth.value ?? true,
            transparency: clampPercent(typeof own?.transparency === "number" ? own.transparency : lineCard.transparency.value ?? 0),
            lineJoin: ownText("lineJoin", defaultShape.lineJoin),
            interpolation: ownText("interpolation", defaultShape.interpolation),
            smoothing: ownText("smoothing", defaultShape.smoothing),
            tension: clampPercent(typeof own?.tension === "number" ? own.tension : defaultShape.tension),
            stepPosition: ownText("stepPosition", defaultShape.stepPosition),
            stepConnect: typeof own?.stepConnect === "boolean" ? own.stepConnect : defaultShape.stepConnect,
            stepWidth: ownText("stepWidth", defaultShape.stepWidth),
            areaShow: areasOn && (typeof ownAreaShow === "boolean" ? ownAreaShow : true),
            lineShow: typeof ownLineShow === "boolean" ? ownLineShow : lineCard.show.value ?? true,
            baselineRatio,
            selectionId: measureBuilder().withMeasure(def.key).createSelectionId(),
            points: categoryGroups.map((g) => {
                const value = def.values[g.rowIndex];
                return {
                    rowIndex: g.rowIndex,
                    ...(cumulativeLineKeys.has(def.key) && resetBefore[positionOf.get(g.rowIndex)!] ? { breakBefore: true } : {}),
                    value,
                    ratio: value === null ? null : onSecondary ? calcRatio2(value) : calcRatio(value),
                    selectionId: withCategories(measureBuilder(), g.rowIndex).withMeasure(def.key).createSelectionId(),
                };
            }),
            tooltip: { displayName: cumulativeLineKeys.has(def.key) ? `${def.name}（累計）` : def.name, format: def.format, values: def.values },
        };
    });
    const lineTargets: LineTarget[] = lines.map((line, j) => ({
        name: line.name,
        selector: { metadata: lineDefs[j].key },
        color: line.color,
        width: line.width,
        lineStyle: line.lineStyle,
        lineJoin: line.lineJoin,
        interpolation: line.interpolation,
        smoothing: line.smoothing,
        tension: line.tension,
        stepPosition: line.stepPosition,
        stepConnect: line.stepConnect,
        stepWidth: line.stepWidth,
        areaShow: typeof lineDefs[j].objects?.areas?.show === "boolean" ? Boolean(lineDefs[j].objects?.areas?.show) : true,
        lineShow: line.lineShow,
        includeCumulative:
            typeof lineDefs[j].objects?.lines?.includeCumulative === "boolean"
                ? Boolean(lineDefs[j].objects?.lines?.includeCumulative)
                : settings.lines.includeCumulative.value ?? false,
        formatMode: ownTextOf(lineDefs[j].objects, "formatMode") ?? String(settings.lines.formatMode.value?.value ?? LINE_FORMAT_MODES.auto),
        customFormat: ownTextOf(lineDefs[j].objects, "customFormat") ?? settings.lines.customFormat.value ?? "",
        dashArray: line.dashArray,
        dashCap: line.dashCap,
        scaleWithWidth: line.scaleWithWidth,
        transparency: line.transparency,
    }));

    if (paretoOn) {
        const ratioColor = paretoCard.ratioColor.value?.value || "#E66C37";
        lines.push({
            name: "累積比",
            color: ratioColor,
            width: 2,
            lineStyle: LINE_STYLES.solid,
            dashArray: "",
            dashCap: "none",
            scaleWithWidth: true,
            transparency: 0,
            lineJoin: "round",
            interpolation: "linear",
            smoothing: LINE_SHAPE_DEFAULTS.smoothing,
            tension: LINE_SHAPE_DEFAULTS.tension,
            stepPosition: LINE_SHAPE_DEFAULTS.stepPosition,
            stepConnect: true,
            stepWidth: LINE_SHAPE_DEFAULTS.stepWidth,
            areaShow: false,
            lineShow: true,
            baselineRatio: 0,
            selectionId: host.createSelectionIdBuilder().createSelectionId(),
            selectable: false,
            points: categoryGroups.map((g) => {
                const value = paretoRatioByRow[g.rowIndex];
                return {
                    rowIndex: g.rowIndex,
                    value,
                    ratio: value === null ? null : calcRatio2(value),
                    // 点を押したら、そのカテゴリの棒を選ぶ（「その他」なら、まとめたカテゴリを全部）
                    selectionId: withCategories(host.createSelectionIdBuilder(), g.rowIndex).createSelectionId(),
                    ...(g.rowIndex === otherRow ? { selectionIds: otherSelectionIds } : {}),
                };
            }),
            tooltip: { displayName: "累積比", format: "0.0%", values: paretoRatioByRow },
        });
    }

    // 累積比の軸は % なので、第 2 Y 軸の単位（円など）を付けない
    const v2TitleStyle = paretoOn ? "showTitleOnly" : getDropdownValue(v2.titleStyle.value, "showTitleOnly");
    const v2ComposedUnit = paretoOn ? "" : composeUnitText(unitWord2, v2UnitText, v2IncludeDisplayUnit);
    const v2UnitInTitle = unitInTitle(v2.titleShow.value ?? true, v2TitleStyle, v2ComposedUnit);
    const valueAxis2Settings: ValueAxis2Settings = {
        show: onSecondary || paretoOn,
        valueShow: v2.valueShow.value ?? true,
        ticks: ticks2,
        ticksFor: ticks2For,
        fontFamily: v2.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, v2.font.fontSize.value ?? 9)),
        bold: v2.font.bold?.value ?? false,
        italic: v2.font.italic?.value ?? false,
        underline: v2.font.underline?.value ?? false,
        labelColor: v2.labelColor.value?.value || "#605E5C",
        titleShow: v2.titleShow.value ?? true,
        // 自動のタイトルは折れ線の名前（標準と同じ。複数なら「および」でつなぐ）
        titleText: styledTitle(
            v2.titleText.value?.trim() || (paretoOn ? "累積比" : lines.map((l) => l.name).join(" および ")),
            v2TitleStyle,
            v2ComposedUnit
        ),
        titleFontFamily: v2.titleFont.fontFamily.value ?? "DIN",
        titleFontSize: Math.max(8, Math.min(32, v2.titleFont.fontSize.value ?? 12)),
        titleBold: v2.titleFont.bold?.value ?? false,
        titleItalic: v2.titleFont.italic?.value ?? false,
        titleUnderline: v2.titleFont.underline?.value ?? false,
        titleColor: v2.titleColor.value?.value || "#252423",
        badgeText: onSecondary && !v2UnitInTitle
            ? resolveBadgeText({
                unitShow: v2.unitShow.value ?? true,
                unitPosition: "valueAxisTop",
                unitIncludeDisplayUnit: v2IncludeDisplayUnit,
                unitStyle: getDropdownValue(v2.unitStyle.value, "parentheses"),
                unitWord: unitWord2,
                unitText: v2UnitText,
            })
            : "",
        unitFontSize: Math.max(6, Math.min(32, v2.unitFontSize.value ?? 9)),
        unitColor: v2.unitColor.value?.value || "#605E5C",
    };

    // 凡例：棒の系列のあとに折れ線。系列 1 本の棒も、折れ線があれば凡例に出す（標準と同じ）
    /** パレートのツールヒントの行（累積比とランク） */
    const paretoLabels: Record<ParetoRank, string> = {
        A: paretoCard.labelA.value?.trim() || "A",
        B: paretoCard.labelB.value?.trim() || "B",
        C: paretoCard.labelC.value?.trim() || "C",
    };
    const paretoTooltipColumns: TooltipColumn[] = paretoOn
        ? [
            { displayName: "累積比", format: "0.0%", values: paretoRatioByRow },
            { displayName: "ランク", format: undefined, values: paretoRankByRow.map((rank) => (rank ? paretoLabels[rank] : null)) },
        ]
        : [];
    const paretoSelectionsOf = (rank: ParetoRank) =>
        categoryGroups
            .filter((g) => paretoRankByRow[g.rowIndex] === rank)
            .flatMap((g) => (g.rowIndex === otherRow ? otherSelectionIds : [withCategories(host.createSelectionIdBuilder(), g.rowIndex).createSelectionId()]));
    const paretoInfo: ParetoInfo = paretoOn
        ? {
            enabled: true,
            thresholds: paretoThresholds,
            showThresholds: paretoCard.showThresholds.value ?? true,
            showRankBand: paretoCard.showRankBand.value ?? true,
            ranks: categoryGroups.map((g) => paretoRankByRow[g.rowIndex]),
            labels: paretoLabels,
            colors: paretoColors,
            selections: { A: paretoSelectionsOf("A"), B: paretoSelectionsOf("B"), C: paretoSelectionsOf("C") },
        }
        : EMPTY_PARETO;

    const legendEntries: LegendItemInfo[] = [
        ...(seriesMode
            ? series.map((s, index) => ({
                kind: "bar" as const,
                index,
                name: s.name,
                color: s.color,
                selectionId: s.selectionId,
                barStyle: {
                    transparency: seriesStyles[index].transparency,
                    borderShow: seriesStyles[index].borderShow,
                    borderColor: seriesStyles[index].borderColor,
                    borderWidth: seriesStyles[index].borderWidth,
                },
            }))
            : lines.length
                ? [{
                    kind: "bar" as const,
                    index: 0,
                    name: slots[0].name,
                    color: columnsSettings.fill,
                    selectionId: null,
                    barStyle: {
                        transparency: columnsSettings.transparency,
                        borderShow: columnsSettings.showBorder,
                        borderColor: columnsSettings.borderMatchColumn ? columnsSettings.fill : columnsSettings.borderFill,
                        borderWidth: columnsSettings.borderWidth,
                    },
                }]
                : []),
        ...lines.map((l, index) => ({ kind: "line" as const, index, name: l.name, color: l.color, selectionId: l.selectable === false ? null : l.selectionId })),
    ];
    // 表示順を反転：凡例の項目だけを逆に並べる（棒の並び・積む順は変えない）
    if (lg.reverseOrder.value ?? false) legendEntries.reverse();
    legendInfo.show = (seriesMode || legendEntries.length > 1) && (lg.show.value ?? true);

    // --- X 軸の定数線 ---------------------------------------------------
    // 「X 軸の定数線」の欄のメジャーは、カテゴリ（行）ごとに Power BI が計算して返す。空白でない値を返したカテゴリに線を引く。
    // 文字ならラベルにし、それ以外（数・真偽）はメジャーの名前をラベルにする。凡例があると matrix は列の小計（全体の値）で 1 列届く。
    // categorical で凡例ごとに届いたときは、どれかの凡例で空白でなければ引く（先頭の凡例が空白でも消さない）。「その他」の行には引かない
    const categoryLineColumns = new Map<string, DataViewValueColumn[]>();
    for (const group of groups) {
        for (const column of group.values) {
            if (!column.source?.roles?.categoryLine) continue;
            const key = column.source.queryName ?? column.source.displayName;
            const found = categoryLineColumns.get(key);
            if (found) found.push(column);
            else categoryLineColumns.set(key, [column]);
        }
    }
    const cl = settings.categoryLine;
    const labelText = getDropdownValue(cl.labelText.value, "value");
    const categoryLines: CategoryLineInfo[] = [...categoryLineColumns.values()].slice(0, 4).map((columns) => {
        const name = columns[0].source.displayName;
        const marks = categoryGroups.flatMap((g, index) => {
            if (g.rowIndex === otherRow) return [];
            const raw = columns.map((c) => c.values[g.rowIndex]).find((v) => v !== null && v !== undefined && v !== "" && v !== false);
            if (raw === undefined) return [];
            // データ値：返した文字（数は書式を当てた値、真偽はメジャーの名前）。名前：メジャーの名前。両方：名前と値
            const value = typeof raw === "string" ? raw : typeof raw === "boolean" ? name : valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(columns[0].source) }).format(raw);
            const label = labelText === "name" ? name : labelText === "both" ? `${name} ${value}` : value;
            return [{ index, label }];
        });
        return { name, marks };
    });
    const categoryLine: CategoryLineSettings = {
        show: cl.show.value ?? true,
        position: (["before", "center", "after"].includes(getDropdownValue(cl.position.value, "before")) ? getDropdownValue(cl.position.value, "before") : "before") as CategoryLineSettings["position"],
        color: cl.color.value?.value || "#605E5C",
        transparency: Math.max(0, Math.min(100, cl.transparency.value ?? 0)),
        lineStyle: getDropdownValue(cl.lineStyle.value, "dashed"),
        dashArray: cl.dashArray.value ?? "",
        scaleWithWidth: cl.scaleWithWidth.value ?? true,
        dashCap: getDropdownValue(cl.dashCap.value, "none"),
        width: Math.max(1, Math.min(10, cl.width.value ?? 1)),
        layer: getDropdownValue(cl.layer.value, "front") === "back" ? "back" : "front",
        shadeShow: cl.shadeShow.value ?? false,
        shadeRegion: getDropdownValue(cl.shadeRegion.value, "before") === "after" ? "after" : "before",
        shadeColor: (cl.shadeMatchLine.value ?? false) ? cl.color.value?.value || "#605E5C" : cl.shadeColor.value?.value || "#E1DFDD",
        shadeTransparency: Math.max(0, Math.min(100, cl.shadeTransparency.value ?? 40)),
        labelShow: cl.labelShow.value ?? true,
        labelHorizontal: getDropdownValue(cl.labelHorizontal.value, "right") === "left" ? "left" : "right",
        labelVertical: getDropdownValue(cl.labelVertical.value, "top") === "bottom" ? "bottom" : "top",
        labelColor: cl.labelColor.value?.value || "#605E5C",
        labelFontSize: Math.max(6, Math.min(32, cl.labelFontSize.value ?? 9)),
        labelAvoidOverlap: cl.labelAvoidOverlap.value ?? true,
    };

    const mk = settings.markers;
    const unitFontSize = Math.max(6, Math.min(32, valAxis.unitFontSize.value ?? 9));
    const unitColor = valAxis.unitColor.value?.value || "#605E5C";

    return {
        dataPoints,
        categoryGroups,
        series,
        seriesMode,
        chartType,
        orientation,
        maxValue,
        minValue,
        niceMin,
        niceMax,
        zeroRatio,
        ticks,
        ticksFor,
        drillPath,
        copyButton: settings.chart.copyButton?.value ?? true,
        commonLevels,
        unitInfo: {
            unitDef: isLogScaleActive ? { ...unitDef, unitWord: "" } : unitDef,
            badgeText: valueUnitInTitle ? "" : badgeText,
            unitPosition,
            precision,
            fontSize: unitFontSize,
            color: unitColor,
        },
        columns: columnsSettings,
        columnTargets,
        labelTargets,
        dataLabels: dataLabelsSettings,
        categoryAxis: categoryAxisSettings,
        valueAxis: valueAxisSettings,
        gridlines: gridlinesSettings,
        legend: legendInfo,
        legendEntries,
        totalLabels: totalLabelsSettings,
        lines,
        lineTargets,
        ...(lineWarning ? { lineWarning } : {}),
        markers: {
            show: mk.show.value ?? false,
            shape: getDropdownValue(mk.shape.value, "circle"),
            size: Math.max(1, Math.min(20, mk.size.value ?? 5)),
            rotation: (((Number(mk.rotation.value) || 0) % 360) + 360) % 360,
            color: mk.color.value?.value ?? "",
            transparency: clampPercent(mk.transparency.value ?? 0),
            borderShow: mk.borderShow.value ?? false,
            borderMatchLine: mk.borderMatchLine.value ?? false,
            borderColor: mk.borderFill.value?.value || "#605E5C",
            borderTransparency: clampPercent(mk.borderTransparency.value ?? 0),
            borderWidth: Math.max(1, Math.min(10, mk.borderWidth.value ?? 1)),
        },
        areas: {
            matchLineColor: areaCard.matchLineColor.value ?? true,
            fill: areaCard.fill.value?.value || "#118DFF",
            transparency: clampPercent(areaCard.transparency.value ?? 60),
        },
        valueAxis2: valueAxis2Settings,
        ribbons: ribbonSettings,
        hasHighlights: slots.some((slot) => !!slot.column.highlights),
        tooltip: {
            categoryName: categories.source?.displayName ?? "",
            ...(levelColumns.length > 1
                ? {
                    categoryLevels: {
                        names: levelColumns.map((column) => column.source?.displayName ?? ""),
                        texts: Array.from({ length: rowCount }, (_, i) => levelsAt(i)),
                    },
                }
                : {}),
            ...measureTooltipOf(slots[0]),
            extras: [...paretoTooltipColumns, ...slots[0].tooltips.map(tooltipColumnOf)],
            ...(seriesMode
                ? {
                    series: slots.map((slot) => ({
                        legendName: legendSource?.displayName ?? null,
                        seriesName: slot.name,
                        ...measureTooltipOf(slot),
                        extras: [...paretoTooltipColumns, ...slot.tooltips.map(tooltipColumnOf)],
                    })),
                }
                : {}),
        },
        pareto: paretoInfo,
        valueLines,
        valueLine,
        ...(valueLineWarning ? { valueLineWarning } : {}),
        categoryLines,
        categoryLine,
        compareLayers: [],
        compare: EMPTY_COMPARE,
        basis: {
            extent: ownExtent,
            keepRows: otherRow >= 0 ? Array.from({ length: originalCategories![0].values.length }, (_, i) => i).filter((i) => !others.mergedRows.includes(i)) : null,
            displayOrder,
        },
        cumulative: {
            available: cumulativeAvailable,
            enabled: cumulativeEnabled,
            toggle: cumulativeAvailable && (calc.cumulativeToggle.value ?? false),
            levels: cumulativeLevels,
            reset: resetDepth < 0 ? CUMULATIVE_RESET_NONE : cumulativeReset,
        },
        isEmpty: dataPoints.length === 0,
    };
}

/** カテゴリが上限（30,000 件）に達したときの警告 */
export const TRUNCATED_TITLE = "すべてのカテゴリを読み込めていません";
export const TRUNCATED_NOTICE =
    "カテゴリが 30,000 件に達したので、並べた順の先頭の 30,000 件で描いています。「その他」と積み上げの合計、パレートの累積比は、読み込めたカテゴリで計算しています。フィルターで絞ってください。";
/** 凡例の値が上限（2,000）に達したときの警告 */
export const SERIES_TRUNCATED_TITLE = "すべての凡例の値を読み込めていません";
export const SERIES_TRUNCATED_NOTICE = "凡例の値が 2,000 に達したので、先頭の 2,000 で描いています。フィルターで絞ってください。";

/** 折れ線の率の分子と分母の数が合わないときの警告 */
export const LINE_RATIO_WARNING_TITLE = "折れ線の率の分子と分母の数が合いません";
export const LINE_RATIO_WARNING =
    "「折れ線の率の分子」と「折れ線の率の分母」は同じ数だけ入れてください。入れた順に組にして割ります。組にならない分は描いていません。";

/** 折れ線 1 本の元。率の線は、系列をまたいで足した分子と分母を持つ（累計のとき、それぞれ累計してから割り直す） */
interface LineDef {
    key: string;
    name: string;
    format: string;
    objects: DataViewObjects | undefined;
    values: Array<number | null>;
    ratio?: { numerators: Array<number | null>; denominators: Array<number | null> };
}

/** 行ごとに分子 ÷ 分母。分子か分母が空白、分母が 0 なら空白（0% ではない） */
function divideRatio(numerators: Array<number | null>, denominators: Array<number | null>): Array<number | null> {
    return numerators.map((numerator, i) => {
        const denominator = denominators[i];
        if (numerator === null || denominator === null || denominator === 0) return null;
        return numerator / denominator;
    });
}

/** 線ごとに保存した「線」の項目の値。保存が無ければ無し */
function ownTextOf(objects: DataViewObjects | undefined, property: string): string | undefined {
    const raw = objects?.lines?.[property];
    return raw !== undefined && raw !== null ? String(raw) : undefined;
}

/** 線の書式文字列。線ごとの保存があればそれ、無ければ「すべて」の値で決める */
function lineFormatOf(def: LineDef, card: LinesCardSettings): string {
    const mode = ownTextOf(def.objects, "formatMode") ?? String(card.formatMode.value?.value ?? LINE_FORMAT_MODES.auto);
    const custom = (ownTextOf(def.objects, "customFormat") ?? card.customFormat.value ?? "").trim();
    if (mode === LINE_FORMAT_MODES.percent) return LINE_PERCENT_FORMAT;
    if (mode === LINE_FORMAT_MODES.custom && custom) return custom;
    return def.ratio ? LINE_PERCENT_FORMAT : def.format;
}
