"use strict";

import { ReferenceLineStyle, referenceLineStyleOf, referenceLabelOf } from "./shared/referenceLine";
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
    LineMarkerValues,
    CHART_TYPES,
    ChartType,
    PaneConditions,
    DEFAULT_LABEL_BACKGROUND,
    DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
    effectiveChartTypeOf,
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
    LINE_LABEL_POSITIONS,
    LINE_PERCENT_FORMAT,
    LinesCardSettings,
    legendPlacementValue,
    LEGEND_BAR_ITEMS,
    LINE_BLANKS,
    AXIS_KINDS,
} from "./settings";
import { legendRowsOf } from "./shared/legend";
import { ContinuousAxis, axisValueOf } from "./continuousAxis";
import {
    resolveUnit,
    formatValue,
    formatDynamicValue,
    dynamicShownSign,
    resolveBadgeText,
    composeUnitText,
    UnitDefinition,
    UNIT_DEFINITIONS,
    outsideLabelSlack,
} from "./unitUtils";
import { labelMayGoOutside, outsideLabelHeadroomOf, outsideLabelLengthOf, totalLabelLengthOf, totalLabelSpotsOf } from "./labelLines";
import { collapseOthers, lineIndependenceOf } from "./others";
import { categoricalOf, isPlaceholderCategory, withCategoryOf, withPlaceholderCategory, withSeriesOf } from "./matrixDataView";
import { TooltipSource, TooltipStack, TooltipColumn, tooltipColumnOf, formatTooltipValue, categoryTooltipRows, BLANK_TEXT } from "./tooltip";
import { ticksUpTo, tickCountOf, boundOf, niceDomain, recommendedTickCount } from "./shared/ticks";
import { localeNumber } from "./shared/localeNumber";
import { maxOf, minOf } from "./shared/minMax";
import { formatOfColumn, isBlankName, isBlankValue, nameTextOf, numberFormatOf, numberOf, numberValueFormatOf, savedChoiceText, valueTextOf } from "./shared/dataValue";
import { PT_TO_PX } from "./shared/text";
import { LABEL_ROOM_SHARE_MAX } from "./shared/plotTop";
import { Halo, haloOf, NO_HALO } from "./shared/halo";
import { enumValue } from "./shared/formatting";
import { clampFontSize, fontFields, titleFontFields } from "./shared/textStyle";
import { DRILL_PATH_SEPARATOR, SERIES_TRUNCATED_TITLE, SERIES_TRUNCATED_NOTICE } from "./shared/warnings";
import type { Grouping } from "./shared/units";
import { HighContrastColors, highContrastOf, highContrastPattern, withHighContrastColors } from "./shared/highContrast";
import { formatSigned, shownSignOf, toneOf, NEGATIVE_STYLES, SignStyle, ZERO_STYLES, TONE_MODES, sentimentTextColors, ownToneColor, DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR, formatPercent, rateFormatOf, sameRateOf, sectionSignOf, sameSectionSignOf, withSectionSign, formatSectionTick, groupingOf, someGroupingOf, fixedNumber } from "./shared/numberFormat";

export { tickCountOf };

export interface DataPoint {
    category: string;
    /** DataView の行番号。並べ替え後もツールヒント等で元の行を引くのに使う */
    rowIndex: number;
    /** 系列の番号（ViewModel.series の添字）。系列が 1 本なら 0 */
    seriesIndex: number;
    /** 値が空白（複数系列で、そのカテゴリにその系列の値が無い）。棒とラベルを描かない */
    blank: boolean;
    /**
     * 系列 1 本の空白（1.4 までと同じく 0 として数え、高さ 0 の棒で描く）。値が 0 の棒と違い、データ ラベルを出さない
     * （標準も空白の棒にはラベルを出さない）
     */
    blankAsZero?: true;
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
    /**
     * 「その他」の棒だけ持つ。押したらまとめたカテゴリを全部選ぶ。ほかの棒と同じ作りで、系列があればまとめた行ごとのカテゴリ＋系列
     * （ほかの系列の「その他」は選ばない）
     */
    selectionIds?: ISelectionId[];
    /**
     * 「その他」の棒のカテゴリだけの ID（まとめたカテゴリごと）。選ばれているかを見るときはこちらとも比べる（categorySelectionId と同じ役目）。
     * パレートの累積比の点・ランクの帯はこの ID で選び、前の版（系列があっても「その他」はカテゴリだけの ID）で保存した選択もこの形で戻る
     */
    categorySelectionIds?: ISelectionId[];
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
    /** 罫線の模様（stroke-dasharray）。ハイコントラストで系列を見分けるときだけ。無ければ実線 */
    borderDash?: string;
    /** データラベルを出すか（系列ごとの「このシリーズに表示」）。カード全体の表示とは別に効く */
    labelShow: boolean;
    /** データラベルの文字色。空 = 自動（棒の色に合わせて白か黒） */
    labelColor: string;
    /** 系列ごとに変えたデータ ラベルの設定（保存の無い項目は「すべて」の値で埋めたもの）。無ければ「すべて」の設定 */
    labelSettings?: DataLabelsSettings;
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
    /** 奥のレイヤーの積み上げの合計（layerPoints と同じ並び）。メジャーごとに棒を並べるときだけ */
    layerTotals?: Array<StackTotals | null>;
}

/** 折れ線の点 1 つ。並びは categoryGroups と同じ（並べ替えのあと） */
export interface LinePoint {
    rowIndex: number;
    /**
     * 「その他」の点だけ持つ。まとめたカテゴリごとの、この線の点の ID（凡例の値＋カテゴリ＋メジャー。累積比の線はカテゴリだけ）。
     * 押したら全部選ぶ
     */
    selectionIds?: ISelectionId[];
    /** この点の手前で線を切るか（累計の区切りで 0 に戻るところ） */
    breakBefore?: boolean;
    /** 値。空白なら null（線を切る） */
    value: number | null;
    /**
     * 軸の比率。第 2 Y 軸で描くときは第 2 Y 軸の比率。0〜1 に丸めない（軸の範囲の外の値は 0 未満か 1 超え。
     * 描くときは描画域で切り、範囲の外の点のマーカー・データ ラベルは出さない）
     */
    ratio: number | null;
    /**
     * ほかのビジュアルで選ばれたときの該当分（Power BI のハイライト）。該当しなければ null。ハイライトが無ければ持たない。
     * highlightRatio は highlight を ratio と同じ軸で表した比率
     */
    highlight?: number | null;
    highlightRatio?: number | null;
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
    /** この線のマーカーの書式。無ければ「すべて」（markers） */
    marker?: MarkerSettings;
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
    /** データ ラベル。線ごとの「このシリーズに表示」がオンのときだけ（既定はオフ） */
    labels?: LineLabelInfo;
}

/** 折れ線のデータ ラベル。texts は points と同じ並び（値が空白なら null） */
export interface LineLabelInfo {
    /** 点の上か下か。自動は上で、プロットの上に入らなければ下 */
    position: "auto" | "above" | "below";
    /** 空 = 「すべて」のカラー（それも空なら線の色） */
    color: string;
    /** 線ごとに変えたデータ ラベルの設定（保存の無い項目は「すべて」の値）。フォント・透過性・背景・文字の縁を使う */
    settings: DataLabelsSettings;
    texts: Array<string | null>;
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
/** 凡例の棒の項目を 1 つにまとめたときの色（系列の色と混ざらない灰色。比較レイヤーの印と同じ） */
export const SINGLE_BAR_LEGEND_GRAY = "#8A8886";

export interface LegendItemInfo {
    kind: "bar" | "line" | "layer";
    /** 棒なら series の添字、折れ線なら lines の添字、比較レイヤーなら compareLayers の添字 */
    index: number;
    name: string;
    color: string;
    /** クリックで選ぶ ID。系列 1 本の棒は null（選ばない） */
    selectionId: ISelectionId | null;
    /** 棒の見た目（凡例の四角を棒に合わせる）。折れ線の見た目は lines[index] とマーカーを見る */
    barStyle?: { transparency: number; borderShow: boolean; borderColor: string; borderWidth: number; borderDash?: string };
    /** 比較レイヤーの印の濃さ（0〜1） */
    layerOpacity?: number;
}

/** 折れ線のマーカー */
export type MarkerSettings = LineMarkerValues;

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
    /** 上下に置いたときに並べる行の数（1 か 2。左右に置いたときは効かない） */
    rows: number;
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
    /** 書式ペインの適用先の項目に出す文字の縁の表示と幅（保存した値のまま。forPane のときだけ。描くのは halo） */
    haloShow?: boolean;
    haloWidth?: number;
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
    /** 階層のレベルの名前（上から。書式ペインの「線の位置」に並べる） */
    levelNames?: string[];
    /** 段に重ねた上のレベルの見せ方（"lines" 区切り線 | "boxed" 帯） */
    hierarchyStyle: string;
    /** 段に重ねた上のレベルの段の高さ（LEVEL_ROW_PADDINGS のキー） */
    levelRowHeight: string;
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
    /** 縦の線を引く位置（"categories" カテゴリごと | "levels" 1 つ上のレベル | "level<n>" 上から n 段目。layout の gridLevelOf） */
    verticalPosition: string;
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
    /**
     * 目盛りを引く範囲（ラベルの余白を足す前の値の軸の端。丸めるなら丸めた棒の端）の、比の大きい側の端の比（0〜1。縦棒では画面の上）。
     * その側にラベルの余白を取らなければ 1。App は Y 軸の上の単位を、この高さの上に置く（描く範囲の上の端より下になりうる。一番上の目盛りは
     * この高さか、それより下）
     */
    frameTopRatio: number;
    /**
     * 目盛りを引く範囲の、比の小さい側の端の比（0〜1。縦棒では画面の下）。その側にラベルの余白を取らなければ 0。
     * App は、目盛りを引く範囲の長さ（描く範囲の長さ ×（frameTopRatio − frameBottomRatio））に数字の行が入る数で目盛りの本数の目安を下げる
     */
    frameBottomRatio: number;
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
    /**
     * 値の軸の長さが分かれば、ラベルの余白を値の軸に取るか（合計ラベルか、棒の外に置きうるデータ ラベルがある）。App はこのときだけ
     * 値の軸の長さを visual.ts へ知らせる（valueAxisLengthToReport。ほかのときは組み直さない）。無ければ取らない
     */
    valueAxisRoom?: boolean;
    /**
     * 値の軸（Y 軸・第 2 Y 軸）の範囲を、描く範囲の長さで決まる目盛りの本数で丸めるか（範囲を丸める・線形・端を決めていない・
     * 「目盛りの本数 (目安)」が空）。App はこのときも値の軸の長さを知らせる（valueAxisLengthToReport）
     */
    roundsByLength?: boolean;
    /**
     * 範囲を丸めた目盛りの本数の目安（roundsByLength のとき。組み立てに使った長さ builtAxisLength で決めた本数）。App は目盛りもこの本数で
     * 引き（valueTickTarget）、visual.ts は描いた長さで決まる本数と比べて組み直すかを決める（needsAxisRebuild）
     */
    roundCount?: number;
    /** 組み立てに使った値の軸の長さ (px。CalculationRuntime の valueAxisLength、無ければ 0)。描いた長さと比べて組み直すかを決める（needsAxisRebuild） */
    builtAxisLength?: number;
    /** 折れ線（複合）。無ければ空 */
    lines: LineSeriesInfo[];
    lineTargets: LineTarget[];
    /** X 軸が日付か数値の 1 列で、種類「連続」を選べる */
    continuousAvailable?: boolean;
    /** X 軸の種類「連続」で描くときの値（categoryGroups と同じ並び）。カテゴリの軸なら無し */
    continuousAxis?: ContinuousAxis;
    /** 折れ線の値が空白のカテゴリの描き方（LINE_BLANKS：線を切る・前後をつなぐ・0 として描く）。無ければ線を切る */
    lineBlanks?: string;
    /** 折れ線の率の分子と分母の数が合わないときの警告（組にならない分は描かない）。無ければ無し */
    lineWarning?: string;
    /** 「比較の列」の値が多すぎて重ねきれないときの警告（compareLayers.ts）。無ければ無し */
    compareWarning?: string;
    /** 折れ線だけ（値が空）で、比較の欄にフィールドを入れたときの知らせ（compareLayers.ts）。無ければ無し */
    compareLineOnlyWarning?: string;
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
    /**
     * 凡例があり「値」にメジャーが 2 つ以上：メジャーごとに 1 本の棒を隣り合わせに並べ、その中を凡例で積む。
     * 比較レイヤーの仕組みで描く（重ねずに並べ、薄くせず、どのレイヤーにもデータ ラベルと合計ラベルを出す）
     */
    measureGroups?: boolean;
    /**
     * 比較レイヤーごとの、値の軸の余白の要るデータ ラベル（compareLayers の添字）。凡例で描くレイヤーを変えたとき、
     * ラベルを描くレイヤーが変わって余白が変わるなら組み直す（visual.ts）。値の軸の長さが分からないときは空
     */
    layerLabelRooms?: LabelRoom[];
    /** 比較レイヤーで手前にそろえるもの（この viewModel を作ったときの値） */
    basis: { extent: ValueExtent; keepRows: number[] | null; displayOrder: number[]; seriesOrder: number[] };
    /** 系列の番号から、集合で並べる位置（左から何番目か）。系列の並びを変えていなければ番号のまま */
    seriesPositions: number[];
    /** 「値」が空で「折れ線の値」だけを入れたか。棒は描かず、線を左の軸で描く */
    lineOnly?: boolean;
    /** 書式ペインの出し分けの条件（描画に使ったのと同じ値。settings の applyCardVisibility が読む）。データが無ければ無し */
    pane?: PaneConditions;
    /** カテゴリを入れずに値だけ。カテゴリのラベルを出さない（押して選ぶ当たりも置かない） */
    noCategory?: boolean;
    /** ビジュアルの背景（テーマの背景色）。棒の外のデータ ラベルの自動の文字色を決めるのに使う */
    background?: string;
    /** ハイコントラストのテーマのときの 3 色。それ以外は無し */
    highContrast?: HighContrastColors | null;
    isEmpty: boolean;
    /** 値のフィールドは入っているが、描く行が無い（絞り込みで 0 件など）。案内を出さずに何も描かない */
    noData: boolean;
}

/** 凡例の線の項目にマーカーを描くか（style は LEGEND_MARKER_STYLES、markerShow はその線のマーカーの表示） */
export function legendMarkerDrawn(style: string, markerShow: boolean): boolean {
    const markerOnly = style === "markerOnly" || style === "markerCircleDefault";
    return markerOnly || (style !== "lineOnly" && markerShow);
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

export interface CategoryLineSettings extends ReferenceLineStyle {
    position: "before" | "center" | "after";
    shadeRegion: "before" | "after";
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
/** 軸の比を 0〜1 に丸める */
const clampRatio = (r: number): number => Math.max(0, Math.min(1, r));

/** 水平方向の配置の保存値（left・center・right。ほかは中央） */
const alignmentOf = (v: unknown): "left" | "center" | "right" => (v === "left" || v === "right" ? v : "center");
const clampBorderWidth = (v: number): number => Math.max(1, Math.min(5, v));

/** 対数目盛りの本数の上限。標準の値軸と同程度（5〜8 本）に収める */
export const MAX_LOG_TICKS = 8;

/** ドリルダウンした位置の区切り */
export { DRILL_PATH_SEPARATOR };
/** ドリルした親が空白のときに出す文字（ツールヒントの空白の表記と同じ） */
export const BLANK_LEVEL_TEXT = "(空白)";

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

/** Math.log10 の丸め誤差を吸収する（10 の冪ならちょうどの整数を返す） */
function log10Snap(v: number): number {
    const l = Math.log10(v);
    const r = Math.round(l);
    return Math.abs(l - r) < 1e-9 ? r : l;
}

/** v 以上でいちばん小さい 1・2・5 × 10 の冪（対数の軸の端を丸める） */
function logStepCeil(v: number): number {
    const power = Math.pow(10, Math.floor(log10Snap(v)));
    return [1, 2, 5, 10].map((m) => m * power).find((step) => step >= v * (1 - 1e-12)) ?? power * 10;
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
    levelRowHeight: "normal",
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
    verticalPosition: "categories",
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
    backgroundColor: DEFAULT_LABEL_BACKGROUND,
    backgroundTransparency: DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
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
    rows: 1,
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
    backgroundColor: DEFAULT_LABEL_BACKGROUND,
    backgroundTransparency: DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
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
    frameTopRatio: 1,
    frameBottomRatio: 0,
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
    basis: { extent: { min: 0, max: 0, valueMaxAbs: 0, magMin: Infinity, magMax: 0 }, keepRows: null, displayOrder: [], seriesOrder: [] },
    seriesPositions: [],
    isEmpty: true,
    noData: false,
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
        // 空の文字は空のまま（空白と別の値）。NaN・±Infinity は空白にせず字で出す
        const legendText = nameTextOf(formatOfColumn(legendSource), BLANK_TEXT, "");
        return groups.flatMap((group) => {
            const column = group.values.find(isMeasure);
            if (!column) return [];
            const raw = group.name;
            const blankName = raw === null || raw === undefined;
            return [{
                key: blankName ? BLANK_TEXT : String(raw),
                name: legendText(raw),
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

/**
 * 折れ線だけのときに置く、値の無い棒の系列（どのカテゴリも空白）。軸・カテゴリの並び・ツールヒントの組み立てを、
 * 棒があるときと同じ道筋で通すために置く。名前は無く、凡例にも出さない
 */
function lineOnlySlotOf(rowCount: number, groups: DataViewValueColumnGroup[], valueColumns: DataViewValueColumn[]): SeriesSlot {
    const columns = groups[0]?.values ?? valueColumns;
    const column = {
        source: { displayName: "", queryName: "__lineOnly", roles: {}, index: -1 },
        values: Array.from({ length: rowCount }, (): null => null),
    } as unknown as DataViewValueColumn;
    return {
        key: "__lineOnly",
        name: "",
        column,
        tooltips: columns.filter((c) => c.source?.roles?.tooltips),
        group: groups[0],
        objects: [],
    };
}

/** 積み上げの棒 1 本の、始まりと量（値の単位。100% 積み上げは割合）。ハイライトの量は無ければ null */
interface StackSegment {
    start: number;
    amount: number;
    highlightAmount: number | null;
}

/** 系列の展開で合計ラベルを出すときに、描く範囲の両端に空ける割合（描く範囲の長さに対する比） */
const EXPLODE_TOTAL_MARGIN = 0.08;

/** 系列の展開で、描く範囲の両端に空ける割合（描く範囲の長さに対する比）。low は値の小さい端（負の合計の側）、high は大きい端 */
export interface ExplodeMargins {
    low: number;
    high: number;
}

/**
 * 系列の展開で合計ラベルのために空ける割合。端ごとに、ふだんは EXPLODE_TOTAL_MARGIN で、その端のいちばん長い合計ラベル（labelRoom の
 * top は大きい端、bottom は小さい端）がそれより長ければその長さの分（合計ラベルの無い端まで広げない）。両端を合わせて
 * LABEL_ROOM_SHARE_MAX（棒が潰れないように）を超えるなら、短い方の端に要る分を先に取り（上限の半分まで）、残りを長い方の端に回す
 * （比を保って縮めると、半分までで入る短い方の端のラベルまで切れる）。値の軸の長さが分からなければ両端とも EXPLODE_TOTAL_MARGIN
 */
function explodeMarginsOf(room: LabelRoom, axisLength: number): ExplodeMargins {
    if (!(axisLength > 0)) return { low: EXPLODE_TOTAL_MARGIN, high: EXPLODE_TOTAL_MARGIN };
    const shareOf = (needs: RoomNeed[]) => Math.max(EXPLODE_TOTAL_MARGIN, Math.max(0, maxOf(needs.map((need) => need.length))) / axisLength);
    const low = shareOf(room.bottom);
    const high = shareOf(room.top);
    if (low + high <= LABEL_ROOM_SHARE_MAX) return { low, high };
    const short = Math.min(low, high, LABEL_ROOM_SHARE_MAX / 2);
    return low <= high ? { low: short, high: LABEL_ROOM_SHARE_MAX - short } : { low: LABEL_ROOM_SHARE_MAX - short, high: short };
}

/**
 * 系列の展開。積んだ棒と棒のあいだに gap（プロットの長さに対する比）を空け、全体がプロットに収まるよう値の比率を縮めて置き直す。
 * すき間は同じ側（正・負）の棒と棒のあいだだけ（0 をはさむ所には入れない）。値の比率 r = z + v × s の s を、
 * どのカテゴリでも正の端が 1 − margins.high 以下・負の端が margins.low 以上になる範囲で最大にする。すき間が多すぎて入らなければ、すき間を縮める
 */
export function explodeStacks(groups: CategoryGroup[], segments: Map<DataPoint, StackSegment>, gap: number, margins: ExplodeMargins = { low: 0, high: 0 }): void {
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
    const maxGaps = Math.max(0, maxOf(stacks.map((x) => x.gp + x.gn)));
    const g = maxGaps > 0 ? Math.min(gap, (0.9 - margins.low - margins.high) / maxGaps) : gap;
    const lower = (s: number) => Math.max(margins.low, maxOf(stacks.map((x) => margins.low - x.N * s + x.gn * g)));
    const upper = (s: number) => Math.min(1 - margins.high, minOf(stacks.map((x) => 1 - margins.high - x.P * s - x.gp * g)));
    const span = Math.max(1e-12, maxOf(stacks.map((x) => x.P - x.N)));
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

/** 系列・線ごとに保存したデータ ラベルの項目の値（[0] を先に見る）。無ければ null */
const ownLabelValue = (objects: Array<DataViewObjects | undefined> | DataViewObjects | undefined, property: string): powerbi.DataViewPropertyValue | null => {
    const list = Array.isArray(objects) ? objects : [objects];
    for (const o of list) {
        const raw = o?.dataLabels?.[property];
        if (raw !== undefined && raw !== null) return raw;
    }
    return null;
};

/** 系列・線ごとに保存したデータ ラベルの項目があるか（「このシリーズに表示」とカラーは別に読むので数えない） */
const hasOwnLabelSettings = (objects: Array<DataViewObjects | undefined> | DataViewObjects | undefined, ignore: string[] = ["show", "color"]): boolean =>
    (Array.isArray(objects) ? objects : [objects]).some((o) => Object.keys(o?.dataLabels ?? {}).some((key) => !ignore.includes(key)));

/**
 * データ ラベルの設定。own（系列・線の objects）に保存した項目はその値を、無ければカード（「すべて」）の値を使う。
 * own が無ければカードの値だけ（1.36.4 までと同じ）。forPane なら、書式ペインの適用先の項目に出すので、文字サイズ・透過性を
 * 丸めずに保存した値のまま返す（「すべて」の欄と同じく、入れた値が欄に戻る）
 */
function dataLabelsSettingsOf(
    dl: VisualFormattingSettingsModel["dataLabels"],
    own?: Array<DataViewObjects | undefined> | DataViewObjects,
    forPane = false
): DataLabelsSettings {
    const raw = (property: string) => (own ? ownLabelValue(own, property) : null);
    const num = (property: string, card: number | undefined) => {
        const v = raw(property);
        return typeof v === "number" ? v : card;
    };
    const flag = (property: string, card: boolean | undefined) => {
        const v = raw(property);
        return typeof v === "boolean" ? v : card;
    };
    const text = (property: string, card: string | undefined) => {
        const v = raw(property);
        return typeof v === "string" ? v : card;
    };
    const fill = (property: string, card: string | undefined) => {
        const v = raw(property) as powerbi.Fill | null;
        return v?.solid?.color ? String(v.solid.color) : card;
    };
    const size = (property: string, card: number | undefined) => (forPane ? (num(property, card) ?? 9) : clampFontSize(num(property, card), 9));
    const percent = (property: string, card: number | undefined) => (forPane ? (num(property, card) ?? 0) : clampPercent(num(property, card) ?? 0));
    /** 選択肢（ItemDropdown の値は IEnumMember） */
    const item = (property: string, card: powerbi.IEnumMember | undefined) => {
        const v = raw(property);
        return v !== null ? String(v) : card?.value !== undefined ? String(card.value) : undefined;
    };
    return {
        show: dl.show.value ?? false,
        position: item("position", dl.position.value) ?? "auto",
        orientation: item("orientation", dl.orientation.value) ?? "horizontal",
        overflow: flag("overflow", dl.overflow.value) ?? false,
        fontSize: size("fontSize", dl.fontSize.value),
        fontFamily: text("fontFamily", dl.fontFamily.value) ?? "Segoe UI",
        bold: flag("bold", dl.bold.value) ?? false,
        italic: flag("italic", dl.italic.value) ?? false,
        underline: flag("underline", dl.underline.value) ?? false,
        transparency: percent("transparency", dl.transparency.value),
        optimizeLabelDisplay: flag("optimizeLabelDisplay", dl.optimizeLabelDisplay.value) ?? false,
        labelMaxWidth: Math.max(10, Math.min(2000, Number(num("labelMaxWidth", dl.labelMaxWidth.value)) || 200)),
        detailShowBlankAs: text("detailShowBlankAs", dl.detailShowBlankAs.value) ?? "",
        color: dl.color.value?.value ?? "",
        backgroundShow: flag("backgroundShow", dl.backgroundShow.value) ?? false,
        backgroundColor: fill("backgroundColor", dl.backgroundColor.value?.value) ?? DEFAULT_LABEL_BACKGROUND,
        backgroundTransparency: percent("backgroundTransparency", dl.backgroundTransparency.value),
        precision: item("precision", dl.precision.value) ?? "auto",
        valueShow: flag("valueShow", dl.valueShow.value) ?? true,
        detailShow: flag("detailShow", dl.detailShow.value) ?? false,
        detailFontFamily: text("detailFontFamily", dl.detailFont.fontFamily.value) ?? "Segoe UI",
        detailFontSize: size("detailFontSize", dl.detailFont.fontSize.value),
        detailBold: flag("detailBold", dl.detailFont.bold?.value) ?? false,
        detailItalic: flag("detailItalic", dl.detailFont.italic?.value) ?? false,
        detailUnderline: flag("detailUnderline", dl.detailFont.underline?.value) ?? false,
        detailColor: fill("detailColor", dl.detailColor.value?.value) ?? "",
        detailTransparency: percent("detailTransparency", dl.detailTransparency.value),
        titleShow: flag("titleShow", dl.titleShow.value) ?? false,
        titleFontFamily: text("titleFontFamily", dl.titleFont.fontFamily.value) ?? "Segoe UI",
        titleFontSize: size("titleFontSize", dl.titleFont.fontSize.value),
        titleBold: flag("titleBold", dl.titleFont.bold?.value) ?? false,
        titleItalic: flag("titleItalic", dl.titleFont.italic?.value) ?? false,
        titleUnderline: flag("titleUnderline", dl.titleFont.underline?.value) ?? false,
        titleColor: fill("titleColor", dl.titleColor.value?.value) ?? "",
        titleTransparency: percent("titleTransparency", dl.titleTransparency.value),
        titleShowBlankAs: text("titleShowBlankAs", dl.titleShowBlankAs.value) ?? "",
        valueShowBlankAs: text("valueShowBlankAs", dl.valueShowBlankAs.value) ?? "",
        singleLine: item("labelContentLayout", dl.labelContentLayout.value) === LABEL_LAYOUTS.singleLine,
        horizontalAlignment: alignmentOf(text("horizontalAlignment", dl.horizontalAlignment.value as string | undefined)),
        halo: haloOf(flag("haloShow", dl.haloShow.value), fill("haloColor", dl.haloColor.value?.value), num("haloWidth", dl.haloWidth.value)),
        ...(forPane ? { haloShow: flag("haloShow", dl.haloShow.value) ?? false, haloWidth: num("haloWidth", dl.haloWidth.value) ?? 3 } : {}),
    };
}

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
 * 「設定の適用先」の名前が重なったら、2 つ目から区別の印を付ける（「1,000」「1,000 (2)」）。
 * 書式ペインは項目の uid を入れ物と名前から作るので、同じ名前が 2 つあると uid が重なる。
 * 名前はカテゴリ・凡例の書式を通した文字なので、元の値が違っても重なる（1000 と 1000.4 を #,0 で出すなど）。選択の ID は変えない
 */
export function withDistinctNames<T extends { name: string }>(targets: T[]): T[] {
    const used = new Set(targets.map((target) => target.name));
    const seen = new Set<string>();
    return targets.map((target) => {
        if (!seen.has(target.name)) {
            seen.add(target.name);
            return target;
        }
        let k = 2;
        while (used.has(`${target.name} (${k})`)) k++;
        const name = `${target.name} (${k})`;
        used.add(name);
        return { ...target, name };
    });
}

/**
 * 書式の区切り。選択肢はデータ次第で、書式の読み込み（populateFormattingSettingsModel）の時点では「区切らない」しか無く、
 * 保存した値が選択肢に無いとして捨てられる。なので保存先（metadata.objects）から直接読む（Desktop で確認）
 */
export function savedCumulativeReset(dataView: DataView | undefined, calc: VisualFormattingSettingsModel["calculation"]): string {
    // 数字の形の値は、Desktop が数で保存して数で届けるので、数も文字に戻して受ける
    const raw = savedChoiceText(dataView?.metadata?.objects?.calculation?.cumulativeReset);
    if (raw) return raw;
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
     * 値の軸の長さ (px)。縦棒は描く範囲の高さ、横棒は幅。visual.ts は、最初の組み立てでは入力から見積もった長さ（initialValueAxisLength）を、
     * 組み直しでは描いて知らせた長さ（App は、値の軸にラベルの余白を取るときと、長さで範囲を丸めるとき（ViewModel の valueAxisRoom・
     * roundsByLength）だけ知らせる）を渡す。
     * 棒の外に出すデータ ラベルと合計ラベルが入るよう、値の軸の範囲を広げるのと（buildSteps で測るラベルの長さ）、範囲を丸める
     * 目盛りの本数を決めるのに使う（長さが無ければ、ラベルの余白を取らず、d3 の既定の 10 本の目安で丸める）。
     * 値の軸の長さは凡例と軸のラベルで決まり、範囲ではほとんど変わらないので、描いた長さで組み直せばふつうは落ち着く
     */
    valueAxisLength?: number;
    /** 凡例で選んだ比較レイヤー（App の visibleLayers）。値の軸の余白は、データ ラベルを描くレイヤーのラベルで取る */
    visibleLayers?: number[];
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
    /** 系列の並び（系列の番号を、集合なら左から、積み上げなら 0 の側から） */
    seriesOrder?: number[];
    /** 値の軸の自動のタイトルに並べる名前（メジャーごとに棒を並べるとき、すべてのメジャーの名前） */
    valueTitleNames?: string[];
}

/** 値の軸の端に要るラベル 1 つ：ラベルを付ける棒の端（値の軸の上の位置）と、軸の長さに対するラベルの長さの割合 */
export interface RoomShare {
    end: number;
    share: number;
    /** share のうち、字の外の余白の割合（字が切れるかを見るときは除く）。無ければ 0 */
    slack?: number;
}

/** 測ったラベルの長さ (px) を、値の軸の長さ (px) に対する割合にする。free が false（ユーザーが決めた端）なら余白を取らない */
function roomSharesOf(needs: RoomNeed[], axisLength: number, free: boolean, endOf: (value: number) => number = (v) => v): RoomShare[] {
    if (!free || !(axisLength > 0)) return [];
    return needs.map((need) => ({ end: endOf(need.value), share: need.length / axisLength, slack: (need.slack ?? 0) / axisLength }));
}

/**
 * ラベルの余白を取った軸の端。lower・upper は余白を取る前の端（自動の端は棒の端、ユーザーが決めた端はその値）。
 * top は上の端の外に、bottom は下の端の外に出すラベル。どのラベルも、棒の端から、取ったあとの軸の長さ（反対側の実際の端まで）×
 * share だけ離れた所まで入るよう、上の端を上げ・下の端を下げる（軸が 0 から始まるとは限らない。いちばん長いラベルがいちばん外の棒に付くとも限らない）。
 * 軸の長さ D について「上げた上の端 − 下げた下の端 ≤ D」を満たす最小の D を二分法で探す（D が増えると左辺の増え方は D より小さい）。
 * 余白が軸の LABEL_ROOM_SHARE_MAX を超えるときは、軸をその上限の長さまで広げる。広げる分は、shortFirst でなければ上下の要る伸びの比で配り、
 * shortFirst なら要る伸びの短い方の端に先に回し（広げる分の半分まで）、残りを長い方の端に回す（どちらがよいかは呼ぶ側が描いた字の切れ方で選ぶ。
 * labelRangeOf）。対数の軸では、大きさの常用対数で渡す。
 * frame は、範囲を丸めたあとの軸の端（lower・upper を含む）。丸めた棒の端から余白を取るときに渡す（labelRangeOf）。丸めた端は、
 * 反対側の端の位置と、余白を取り始める端としてだけ使い、上限の基準は lower・upper（余白を取る前の端）のままにする（丸めた端を
 * 基準にすると、丸めて延びた分だけ上限が延びて棒が縮む）。丸めた端から取って上限に掛かるときも、上限の長さまでは
 * 丸めた端から広げる（丸めた端がもう上限の長さに届いていれば、丸めた端のまま）
 */
export function withLabelRoom(
    lower: number,
    upper: number,
    top: RoomShare[],
    bottom: RoomShare[],
    frame: [number, number] = [lower, upper],
    shortFirst = false
): [number, number] {
    const base = upper - lower;
    const [frameLow, frameHigh] = frame;
    if (!(base > 0) || (!top.length && !bottom.length)) return [frameLow, frameHigh];
    const endsAt = (whole: number): [number, number] => {
        let low = frameLow;
        let high = frameHigh;
        for (const n of top) high = Math.max(high, n.end + n.share * whole);
        for (const n of bottom) low = Math.min(low, n.end - n.share * whole);
        return [low, high];
    };
    const fits = (whole: number) => {
        const [low, high] = endsAt(whole);
        return high - low <= whole;
    };
    const frameLength = frameHigh - frameLow;
    if (fits(frameLength)) return [frameLow, frameHigh];
    const most = base / (1 - LABEL_ROOM_SHARE_MAX);
    // 丸めた端がもう上限の長さに届いている。比で配るなら、丸めた端のまま
    if (!(frameLength < most) && !shortFirst) return [frameLow, frameHigh];
    if (!(frameLength < most)) {
        // 丸めた端がもう上限の長さに届いている。長い方の端（余白を取る前の端から要る伸びの長い方）は丸めた端のまま。
        // 丸めて軸が延びた分だけ短い方の端のラベルが要る長さも延びるので、短い方の端だけ、入るまで広げる（広げる分は、上限の長さまでに広げられる分の半分まで）
        let reachLow = lower;
        let reachHigh = upper;
        for (const n of top) reachHigh = Math.max(reachHigh, n.end + n.share * frameLength);
        for (const n of bottom) reachLow = Math.min(reachLow, n.end - n.share * frameLength);
        const shortLow = lower - reachLow <= reachHigh - upper;
        const [low, high] = endsAt(frameLength);
        if (!((shortLow ? frameLow - low : high - frameHigh) > 0)) return [frameLow, frameHigh];
        if ((shortLow ? lower - reachLow : reachHigh - upper) > (LABEL_ROOM_SHARE_MAX / 2) * frameLength) return [frameLow, frameHigh];
        // 片方の端だけを e 広げると軸の長さも e 延びる。ラベル（端 end・割合 share）が入るのは e ≥ (外へ要る分 at frameLength) / (1 − share)
        let grow = 0;
        for (const n of shortLow ? bottom : top) {
            const over = shortLow ? frameLow - (n.end - n.share * frameLength) : n.end + n.share * frameLength - frameHigh;
            if (over > 0 && n.share < 1) grow = Math.max(grow, over / (1 - n.share));
        }
        grow = Math.min(grow, (most - base) / 2);
        return shortLow ? [frameLow - grow, frameHigh] : [frameLow, frameHigh + grow];
    }
    if (!fits(most)) {
        // 上限に掛かる。軸を上限の長さまで広げる（丸めた棒の端から取るときは丸めた端から。丸めた端で止めると、余白が上限に届かないうちにラベルが切れる）。
        // 軸の長さが上限で決まるので、端ごとに要る伸びはその長さで測れる。比で配るか、短い方の端に要る分を先に（広げる分の半分まで）取り、残りを長い方の端に回す
        const [low, high] = endsAt(most);
        if (!shortFirst) {
            const scale = (most - frameLength) / (high - low - frameLength);
            return [frameLow - (frameLow - low) * scale, frameHigh + (high - frameHigh) * scale];
        }
        const room = most - frameLength;
        const needLow = frameLow - low;
        const needHigh = high - frameHigh;
        const shortLow = needLow <= needHigh;
        // 短い方の端に回すのは、余白を取る前の端から広げられる分（most − base）の半分まで。丸めて短い方の端がもう延びた分は、その中に数える。
        // 丸めた端から測ると、丸めで長い方の端が延びた分だけ短い方の取り分が減る（目盛りの本数の目安が少ない小さい図は、丸めの刻みが粗い）
        const half = (most - base) / 2 - (shortLow ? lower - frameLow : frameHigh - upper);
        const short = Math.max(0, Math.min(shortLow ? needLow : needHigh, half, room));
        const [growLow, growHigh] = shortLow ? [short, room - short] : [room - short, short];
        return [frameLow - growLow, frameHigh + growHigh];
    }
    let short = frameLength;
    let long = most;
    for (let k = 0; k < 60; k++) {
        const mid = (short + long) / 2;
        if (fits(mid)) long = mid;
        else short = mid;
    }
    return endsAt(long);
}

/**
 * 値の軸の範囲（ラベルの余白を取り、round なら棒の端を丸める）。round なら、余白を取る前の端（棒の端）を丸め、その丸めた端から
 * ラベルに要る分だけ延ばし、延ばした端は丸めない（標準の集合縦棒と同じ。ラベルの分を足した端まで丸めると、ラベルに要るより
 * 目盛り 1 段ぶん延びて棒が縮む。延ばした端は目盛りの間になり、目盛りは軸の中の切りのいい値に引く）。ラベルが丸めた端の中に
 * 入れば丸めた端のまま。余白が上限に掛かるときの配り方（withLabelRoom の shortFirst）は、
 * 両方で範囲を決めてから、描いた字が描く範囲から切れるラベルの数、次に切れる量の最大が小さい方を選ぶ（同じなら比で配る方）。
 * 比で配ると、片方の端だけなら入る短いラベルまで長い方につられて切れることがあり、短い方を先にすると、比なら字の外の余白に
 * 吸われて入っていた長い方の字が切れることがある
 */
function labelRangeOf(
    lower: number,
    upper: number,
    top: RoomShare[],
    bottom: RoomShare[],
    round: boolean,
    axisLength: number,
    roundCount: number | undefined
): [number, number] {
    // 丸めるときは、丸めた棒の端から余白を取る。上限の基準は余白を取る前の端のまま（withLabelRoom の frame）
    const frame = barFrameOf(lower, upper, round, roundCount);
    const rangeOf = (shortFirst: boolean): [number, number] => withLabelRoom(lower, upper, top, bottom, frame, shortFirst);
    const byRatio = rangeOf(false);
    if (!top.length && !bottom.length) return byRatio;
    const shortFirst = rangeOf(true);
    if (shortFirst[0] === byRatio[0] && shortFirst[1] === byRatio[1]) return byRatio;
    /** 字が描く範囲から切れるラベルの数と、切れる量の最大 (px) */
    const clipOf = ([low, high]: [number, number]) => {
        const whole = high - low;
        let count = 0;
        let most = 0;
        const add = (over: number) => {
            const px = (over / whole) * axisLength;
            if (!(px > CLIP_EPS_PX)) return;
            count++;
            most = Math.max(most, px);
        };
        for (const n of top) add(n.end + (n.share - (n.slack ?? 0)) * whole - high);
        for (const n of bottom) add(low - (n.end - (n.share - (n.slack ?? 0)) * whole));
        return { count, most };
    };
    const a = clipOf(byRatio);
    const b = clipOf(shortFirst);
    return b.count < a.count || (b.count === a.count && b.most < a.most - CLIP_EPS_PX) ? shortFirst : byRatio;
}

/** 字が切れたとみなす量 (px)。座標の丸めの分は見ない */
const CLIP_EPS_PX = 0.05;

/**
 * ラベルの余白を足す前の値の軸の端（round なら丸めた棒の端）。余白はここから取り（labelRangeOf）、目盛りはこの中に引く。
 * 丸めは、目盛りと同じ本数の目安 count（描く範囲の長さで決まる）で d3 の nice(count) と同じに丸める（標準と同じ）
 */
function barFrameOf(lower: number, upper: number, round: boolean, count: number | undefined): [number, number] {
    return round ? niceDomain(lower, upper, count) : [lower, upper];
}

/** 棒の外に出しうるデータ ラベル：棒の端の値と、値の向きの長さ (px) */
export interface RoomNeed {
    value: number;
    length: number;
    /** length のうち、字の外の余白 (px)。データ ラベルは背景の余白（背景を描かなくても足す）、合計ラベルは 0（背景の箱で見る） */
    slack?: number;
}

/** データ ラベルの長さを測るときに作った文字。棒を組むときに使い回す。undefined は作っていない（ラベルに出さない行） */
interface MadeLabelTexts {
    /** 「ラベルの値」のフィールドの文字（null はフィールドが無い・空白で、棒の値を出す） */
    valueField: string | null | undefined;
    formatted: { formattedValue: string; dataLabelText: string; labelToneColor: string } | undefined;
    title: string | undefined;
    detail: string | undefined;
}

/**
 * 値の軸の端に余白の要るデータ ラベル。上は正の値と 0 の棒（値の大きい側の外）、下は負の値の棒。
 * 同じ長さのラベルは、いちばん外の棒の 1 つだけを持つ
 */
export interface LabelRoom {
    top: RoomNeed[];
    bottom: RoomNeed[];
}

/** 余白の要るラベルを合わせる（比較レイヤーで、ラベルを描くレイヤーのぶんを合わせる）。同じ長さは外の棒だけ残す */
export function unionLabelRoom(rooms: LabelRoom[]): LabelRoom {
    const outermost = (needs: RoomNeed[], outer: (a: number, b: number) => number): RoomNeed[] => {
        const byLength = new Map<string, RoomNeed>();
        for (const n of needs) {
            const key = `${n.length}|${n.slack ?? 0}`;
            const kept = byLength.get(key);
            byLength.set(key, kept ? { ...kept, value: outer(kept.value, n.value) } : n);
        }
        return [...byLength.values()];
    };
    return {
        top: outermost(rooms.flatMap((r) => r.top), Math.max),
        bottom: outermost(rooms.flatMap((r) => r.bottom), Math.min),
    };
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
    const pending = beginTransform(dataView, host, settings, runtime, basis);
    return pending.finish(pending.labelRoom);
}

/** 組み立ての途中で渡す、データ ラベルの長さを測る関数（呼んだときに測る）と、比較レイヤーを重ねうるか（折れ線だけ・パレートでない） */
interface LabelMeasure {
    measure: () => LabelRoom;
    canLayer: boolean;
}

/** 行を読み終えたところで渡す、このレイヤーの範囲と並び（比較レイヤーで手前にそろえるもの）と、比較レイヤーを重ねうるか */
interface RowsRead {
    basis: ViewModel["basis"];
    canLayer: boolean;
}

/** 行を読んだあとに渡す、すべてのレイヤーで決めるもの（値の範囲と、値の軸の自動のタイトルに並べる名前） */
export type LayerShared = Pick<LayerBasis, "extent" | "valueTitleNames">;

/** 組み立てを止めたところで受け取るもの。1 回目は shared、2 回目は room（取り違えたら組み立てを止める） */
type StepInput = { shared: LayerShared } | { room: LabelRoom };

/**
 * 行を読み終え、値の軸の範囲を決める手前で止めたもの。resume にすべてのレイヤーの範囲を渡すと、データ ラベルの長さを測る手前まで
 * 組む（1 回だけ）。比較レイヤーは、レイヤーごとに行を 1 回だけ読み、範囲をそろえてから続きを組む
 */
export interface ReadViewModel {
    /** このレイヤーの範囲と並び。データが無いなど、読み終える前に組み終えたときは、組み終えた viewModel のもの */
    basis: ViewModel["basis"];
    /** 比較レイヤーを重ねうるか（折れ線だけ・パレートでない）。読み終える前に組み終えたときは false */
    canLayer: boolean;
    resume: (shared: LayerShared) => PendingViewModel;
}

/**
 * transform を、行を読み終えたところ（値の軸の範囲を決める手前）で止める。basis の範囲と軸のタイトルは使わず、resume で受ける
 */
export function readTransform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    runtime: CalculationRuntime = {},
    basis: Omit<LayerBasis, keyof LayerShared> = {}
): ReadViewModel {
    const steps = buildSteps(dataView, host, settings, runtime, basis);
    const first = steps.next();
    const read = first.done ? undefined : (first.value as RowsRead);
    let resumed = false;
    return {
        basis: first.done ? (first.value as ViewModel).basis : read!.basis,
        canLayer: read?.canLayer ?? false,
        resume: (shared) => {
            if (resumed) throw new Error("棒グラフの組み立てを 2 回続けた");
            resumed = true;
            const next = first.done ? first : steps.next({ shared });
            return pendingOf(steps, next, host);
        },
    };
}

/** 組み立てを、値の軸の範囲を決める手前で止めたもの。finish に値の軸の余白を渡すと、続きを組んで仕上げる（1 回だけ） */
export interface PendingViewModel {
    /** このレイヤーのラベルで測った、値の軸の端に要る余白 (px)。初めて読んだときに測る（読まなければ測らない） */
    readonly labelRoom: LabelRoom;
    /** 比較レイヤーを重ねうるか（折れ線だけ・パレートでない）。データが無いなど、測る前に組み終えたときは false */
    canLayer: boolean;
    finish: (room: LabelRoom) => ViewModel;
}

/**
 * transform を、値の軸の範囲を決める手前で止める。比較レイヤーは、ラベルを描くレイヤーの余白をそろえてから続きを組む
 * （組み直さない。ラベルの文字は値の軸の範囲では変わらない）
 */
export function beginTransform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    runtime: CalculationRuntime = {},
    basis: LayerBasis = {}
): PendingViewModel {
    const { extent, valueTitleNames, ...order } = basis;
    return readTransform(dataView, host, settings, runtime, order).resume({ extent, valueTitleNames });
}

/** 組み立ての続き（first は、データ ラベルの長さを測る手前で止めたところか、組み終えたところ） */
function pendingOf(
    steps: Generator<RowsRead | LabelMeasure, ViewModel, StepInput>,
    first: IteratorResult<RowsRead | LabelMeasure, ViewModel>,
    host: IVisualHost
): PendingViewModel {
    // データが無いなど、測る前に組み終えたときは余白は無い
    const step = first.done ? undefined : (first.value as LabelMeasure);
    let labelRoom: LabelRoom | undefined;
    return {
        get labelRoom() {
            labelRoom ??= step ? step.measure() : { top: [], bottom: [] };
            return labelRoom;
        },
        canLayer: step?.canLayer ?? false,
        finish: (room) => {
            const last = first.done ? first : steps.next({ room });
            if (!last.done) throw new Error("棒グラフの組み立てが終わらない");
            return finishViewModel(last.value as ViewModel, host);
        },
    };
}

/** 組み立てたあと、テーマの背景色とハイコントラストを当てる */
function finishViewModel(viewModel: ViewModel, host: IVisualHost): ViewModel {
    const palette = host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette | undefined;
    // テーマの背景色。16 進でなければ（届かない・読めない色なら）白として扱う
    const background = palette?.background?.value;
    viewModel.background = background && /^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(background) ? background : "#FFFFFF";
    const hc = highContrastOf(palette);
    return hc ? toHighContrast(viewModel, hc) : viewModel;
}

/** ハイコントラストの線の幅 (px)。棒の囲み・凡例の印 */
const HC_BORDER_WIDTH = 1.5;
/** ハイコントラストの折れ線の種類。色の代わりに線の種類で見分ける */
const HC_LINE_STYLES = ["solid", "dashed", "dotted"];

/**
 * ハイコントラストのテーマで描く形に直す。文字・軸・線は前景色、背景は背景色にし、
 * 棒は系列ごとに「背景色で塗って囲む／前景色で塗る／囲む線の模様」で見分ける（テーマの色・書式で選んだ色は使わない）。
 * データ ラベルの文字色は自動にする（棒の塗りに合わせて前景色か背景色）
 */
export function toHighContrast(viewModel: ViewModel, hc: HighContrastColors): ViewModel {
    const vm = withHighContrastColors(viewModel, hc, ["tooltip", "layerTooltips", "basis"]);
    const fillOf = (seriesIndex: number) => (highContrastPattern(seriesIndex).solid ? hc.foreground : hc.background);
    const seen = new Set<DataPoint>();
    const paint = (d: DataPoint) => {
        if (seen.has(d)) return;
        seen.add(d);
        const pattern = highContrastPattern(d.seriesIndex);
        d.color = fillOf(d.seriesIndex);
        d.transparency = 0;
        d.borderShow = true;
        d.borderColor = hc.foreground;
        d.borderTransparency = 0;
        d.borderWidth = HC_BORDER_WIDTH;
        d.borderDash = pattern.dash;
        d.labelColor = "";
        if (d.labelSettings) d.labelSettings = autoLabelColors(d.labelSettings);
    };
    vm.dataPoints.forEach(paint);
    for (const g of vm.categoryGroups) {
        g.points.forEach(paint);
        g.layerPoints?.forEach((points) => points.forEach(paint));
    }
    vm.dataLabels = autoLabelColors(vm.dataLabels);
    // 合計ラベルの背景も透かさない（データ ラベルと同じ。既定の透過性 90 のままだと、背景色の地がほとんど見えない）
    vm.totalLabels = { ...vm.totalLabels, backgroundTransparency: 0 };
    // 罫線を外側だけにすると、背景色で塗った段の境目が消える。棒の色に合わせた罫線も背景色になるので使わない
    vm.columns = { ...vm.columns, borderOutlineOnly: false, borderMatchColumn: false };
    // リボンの帯を系列の色（背景色のことがある）に合わせると見えないので、前景色を透かして塗る
    vm.ribbons = { ...vm.ribbons, matchSeriesColor: false, transparency: Math.max(vm.ribbons.transparency, 50) };
    vm.series.forEach((s, i) => (s.color = fillOf(i)));
    // 線が 2 本以上なら線の種類で見分ける（書式で選んだ種類より優先）。1 本なら書式のまま
    if (vm.lines.length > 1) vm.lines.forEach((line, i) => (line.lineStyle = HC_LINE_STYLES[i % HC_LINE_STYLES.length]));
    vm.lines.forEach((line) => {
        line.transparency = 0;
        // 線のラベルも、背景を透かすと前景色で塗った棒の上で読めない
        if (line.labels) line.labels = { ...line.labels, settings: { ...line.labels.settings, backgroundTransparency: 0 } };
    });
    // 網掛けは前景色を透かして塗る。透かさないと、前景色で塗った棒が網掛けに溶ける
    vm.valueLine = { ...vm.valueLine, shadeTransparency: Math.max(vm.valueLine.shadeTransparency, 50) };
    vm.categoryLine = { ...vm.categoryLine, shadeTransparency: Math.max(vm.categoryLine.shadeTransparency, 50) };
    for (const entry of vm.legendEntries) {
        if (entry.kind === "bar") {
            entry.color = fillOf(entry.index);
            entry.barStyle = { transparency: 0, borderShow: true, borderColor: hc.foreground, borderWidth: HC_BORDER_WIDTH, borderDash: highContrastPattern(entry.index).dash };
        }
    }
    vm.pareto = { ...vm.pareto, colors: { A: hc.foreground, B: hc.foreground, C: hc.foreground } };
    vm.highContrast = hc;
    return vm;
}

/** データ ラベルの文字色を自動にし（棒の塗りに合わせて前景色か背景色）、背景は透かさない */
function autoLabelColors(dl: DataLabelsSettings): DataLabelsSettings {
    return { ...dl, color: "", titleColor: "", detailColor: "", backgroundTransparency: 0 };
}

/** 数値のカテゴリの目盛りの書き方（モデルの書式。名前の付いた数の書式は同じ意味の書式に直し、日付の書式は当てない。無ければそのまま） */
function numberFormatterOf(column: DataViewCategoryColumn): (value: number) => string {
    const format = numberValueFormatOf(valueFormatter.getFormatStringByColumn(column.source));
    const formatter = format ? valueFormatter.create({ format }) : null;
    return (value) => (formatter ? formatter.format(value) : String(value));
}

/**
 * viewModel を組む。2 回 yield する。1 回目は行を読み終えたところで、このレイヤーの範囲と並びを渡し、すべてのレイヤーの範囲
 * （LayerShared）を受け取る。2 回目は値の軸の範囲を決める手前で、データ ラベルの長さを測る関数を渡し、値の軸に取る余白を受け取って
 * 続きを組む
 */
function* buildSteps(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    runtime: CalculationRuntime = {},
    basis: LayerBasis = {}
): Generator<RowsRead | LabelMeasure, ViewModel, StepInput> {
    // matrix で受けたなら categorical の形に詰め替える（折れ線の値は、凡例をまとめたカテゴリ全体の値になる）
    dataView = withPlaceholderCategory(categoricalOf(dataView));
    // 上位 N 件＋「その他」。読む前に DataView を畳む。「その他」の選択は元のカテゴリの列から作り直す
    const originalCategories = dataView?.categorical?.categories;
    const originalDataView = dataView;
    const others = collapseOthers(dataView, settings.columns.otherCount.value ?? 0, settings.columns.otherLabel.value?.trim() || "その他", basis.keepRows);
    dataView = others.dataView;
    const otherRow = others.otherRow;
    /**
     * builder に、「その他」にまとめた元の行 r のカテゴリを足す。元の（畳む前の）DataView の行で、列はほかの点と同じいちばん下のレベル
     * （畳むのはカテゴリの列が 1 つのときだけなので、その列）
     */
    const withMergedRow = (builder: powerbi.visuals.ISelectionIdBuilder, r: number) =>
        withCategoryOf(builder, originalDataView!.categorical!, originalCategories![originalCategories!.length - 1], r);
    /** 「その他」のカテゴリだけの ID（まとめたカテゴリごと）。パレートの点と帯で選ぶ。棒は系列があれば系列ごとの ID を作る */
    const otherSelectionIds = others.mergedRows.map((r) => withMergedRow(host.createSelectionIdBuilder(), r).createSelectionId());

    const categorical: DataViewCategorical | undefined = dataView?.categorical;
    const categories = categorical?.categories?.[0];
    // カテゴリを入れずに値だけ（仮の 1 行）。カテゴリのラベル・カテゴリごとの書式の対象を出さない
    const noCategory = !!categories && isPlaceholderCategory(categories);
    const valueColumns: DataViewValueColumn[] = categorical?.values ?? [];
    // 凡例のフィールド。あれば values は凡例の値ごとのグループに分かれて届く
    const legendSource = categorical?.values?.source;
    const groups: DataViewValueColumnGroup[] =
        categorical?.values?.grouped?.() ?? [{ values: valueColumns } as DataViewValueColumnGroup];
    const slots = seriesSlotsOf(groups, valueColumns, legendSource);
    // 折れ線だけ：「値」が空で「折れ線の値」（率の分子を含む）があれば、値の無い棒の系列を 1 本置き、棒は描かずに線だけを描く
    const lineOnly =
        !slots.length && !!categories && valueColumns.some((c) => c.source?.roles?.lineMeasure || c.source?.roles?.lineRatioNumerator);
    if (lineOnly) slots.push(lineOnlySlotOf(categories!.values.length, groups, valueColumns));

    if (!categories || !slots.length || !categories.values.length) {
        // 値のフィールドは入っているのに行が無い（絞り込みで 0 件など）なら、作り手向けの案内は出さずに何も描かない
        const fieldsPlaced = (dataView?.metadata?.columns ?? []).some(
            (column) => column.roles?.measure || column.roles?.lineMeasure || column.roles?.lineRatioNumerator
        );
        return fieldsPlaced ? { ...EMPTY, noData: true } : EMPTY;
    }

    const seriesMode = !lineOnly && (!!legendSource || slots.length > 1);
    const rowCount = categories.values.length;

    // 階層。X 軸にフィールドが複数あり、展開していると、上のレベルから順に 1 列ずつ届く。
    // 1 列のときは 1.22 までと同じ（levels が 1 つ）
    const levelColumns: DataViewCategoryColumn[] = categorical!.categories!;
    // カテゴリはモデルの書式で出す（数値の #,0、日付の yyyy/MM/dd など。String() だと「1000」や JS の日付の文字列になる）。
    // 空白は「(空白)」、空の文字は空のまま（名前の無いカテゴリも空の文字）。NaN・±Infinity は空白にせず字で出す。
    // 表示の文字だけで、並べ替え・親の見分け・選択は元の値（levelKeysAt と ID）で行う
    const levelTextOf = levelColumns.map((column) => {
        const text = nameTextOf(formatOfColumn(column.source), BLANK_TEXT, "");
        return (i: number): string => text(column.values[i]);
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
     * いちばん下のレベルは外さない（1 行だけでも名前は出す）。
     * 共通の親のうち、いちばん下の 1 段（いちばん下のレベルのすぐ上）は軸に残す。FY27 ＞ 1H ＞ 2Q にドリルして 7・8・9 月を
     * 描くなら、2Q は軸の段に残し、FY27 ＞ 1H だけをドリルの位置に出す。絞り込み（スライサーなど）で親が 1 つになっただけのときも、
     * 届くデータはドリルと同じで見分けられないが、こうすれば 1 段だけ絞ったときは軸に残る（標準と同じ）
     */
    const sharedDepth = (() => {
        if (levelColumns.length < 2 || rowCount === 0) return 0;
        const first = levelKeysAt(0);
        let depth = 0;
        while (depth < levelColumns.length - 1 && Array.from({ length: rowCount }, (_, i) => i).every((i) => levelKeysAt(i)[depth] === first[depth])) depth++;
        return depth;
    })();
    const commonDepth = Math.max(0, sharedDepth - 1);
    /**
     * ドリルダウンした位置。どのカテゴリでも同じ上の階層の値を「＞」でつなぐ（事業A に入り、製品A1 に入ると「事業A ＞ 製品A1」）。
     * 共通の親のいちばん下の 1 段は軸に残すので、ここには出さない（事業A に入って製品を描くだけなら、ドリルの位置は出ない）
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
    const numbersOf = (columns: DataViewValueColumn[], i: number, highlights = false) =>
        columns
            .map((c) => (highlights ? c.highlights?.[i] : c.values[i]))
            .map(numberOf)
            .filter((v): v is number => v !== null);
    /**
     * 折れ線だけで凡例があるときは、凡例の値ごとに線を分ける（凡例の値ごとの値で引く。全体の値の 1 本にしない）。
     * 線の名前は凡例の値（メジャーが複数なら「凡例の値 メジャー名」）
     */
    const keyOfColumn = (column: DataViewValueColumn) => column.source.queryName ?? column.source.displayName;
    // 線が 2 本以上（「折れ線の値」が 2 つ、率もある）なら分けない。凡例の値ごとの書式の保存先が線ごとに分けられないため
    const ratioNumeratorKeys = new Set(groups.flatMap((g) => g.values).filter((c) => c.source?.roles?.lineRatioNumerator).map(keyOfColumn));
    const splitLinesByLegend = lineOnly && !!legendSource && lineColumns.size + ratioNumeratorKeys.size === 1;
    const legendText = legendSource ? nameTextOf(formatOfColumn(legendSource), BLANK_TEXT, "") : null;
    /**
     * 分けた線の書式。凡例の値に保存したものを先に、無ければ分ける前の 1 本の線に保存したもの（メジャーの保存先）を引き継ぐ。
     * 色だけは引き継がない（どの線も同じ色になる）
     */
    const splitObjectsOf = (own: DataViewObjects | undefined, shared: DataViewObjects | undefined): DataViewObjects | undefined => {
        if (!shared) return own;
        const merged: DataViewObjects = {};
        for (const name of new Set([...Object.keys(shared), ...Object.keys(own ?? {})])) {
            const { fill: _fill, ...inherited } = (shared[name] ?? {}) as Record<string, powerbi.DataViewPropertyValue>;
            merged[name] = { ...(name === "lines" ? inherited : shared[name]), ...(own?.[name] ?? {}) };
        }
        return merged;
    };
    const numbersAt = (column: DataViewValueColumn | undefined, highlights = false): Array<number | null> =>
        Array.from({ length: rowCount }, (_, i) => {
            return numberOf(highlights ? column?.highlights?.[i] : column?.values[i]);
        });
    // 名前とキーは棒の系列と同じ（凡例の書式を通す。空白は「(空白)」）
    const seriesKeyOf = (group: DataViewValueColumnGroup) => (group.name === null || group.name === undefined ? BLANK_TEXT : String(group.name));
    const seriesNameOf = (group: DataViewValueColumnGroup) =>
        legendText ? legendText(group.name) : group.name === null || group.name === undefined ? BLANK_TEXT : String(group.name);
    const seriesLineDefsOf = (key: string, role: string): LineDef[] =>
        groups.flatMap((group) => {
            const column = group.values.find((c) => c.source?.roles?.[role] && keyOfColumn(c) === key);
            if (!column) return [];
            return [
                {
                    key: `${key}\u001f${seriesKeyOf(group)}`,
                    name: seriesNameOf(group),
                    format: valueFormatter.getFormatStringByColumn(column.source),
                    objects: splitObjectsOf(group.objects, column.source.objects),
                    values: numbersAt(column),
                    ...(column.highlights ? { highlights: numbersAt(column, true) } : {}),
                    series: { group, measureKey: key, measureName: column.source.displayName },
                },
            ];
        });
    const measureLineDefs: LineDef[] = splitLinesByLegend
        ? [...lineColumns.keys()].flatMap((key) => seriesLineDefsOf(key, "lineMeasure"))
        : [...lineColumns.entries()].map(([key, columns]) => {
        // 「その他」にまとめたときは、畳む前の DataView で判定する（行が減ると判定が変わり、残したカテゴリの線の値まで変わる）
        const independent =
            lineIndependenceBeforeCollapse?.get(key) ??
            Array.from({ length: rowCount }, (_, i) => numbersOf(columns, i)).every((vals) => vals.every((v) => v === vals[0]));
        const valuesOf = (highlights: boolean): Array<number | null> => Array.from({ length: rowCount }, (_, i) => {
            const vals = numbersOf(columns, i, highlights);
            if (!vals.length) return null;
            return independent ? vals[0] : vals.reduce((sum, v) => sum + v, 0);
        });
        const source = columns[0].source;
        return {
            key,
            name: source.displayName,
            format: valueFormatter.getFormatStringByColumn(source),
            objects: source.objects,
            values: valuesOf(false),
            ...(columns.some((c) => c.highlights) ? { highlights: valuesOf(true) } : {}),
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
    const seriesSumOf = (columns: DataViewValueColumn[], i: number, highlights = false): number | null => {
        const vals = numbersOf(columns, i, highlights);
        return vals.length ? vals.reduce((sum, v) => sum + v, 0) : null;
    };
    const pairedNumerators = sharedDenominator ? numeratorList : numeratorList.slice(0, denominatorList.length);
    /** 凡例ごとに分けるときの率：凡例の値ごとに、その値の分子 ÷ 分母 */
    const seriesRatioDefsOf = (key: string, k: number): LineDef[] => {
        const denominatorKey = keyOfColumn(denominatorList[sharedDenominator ? 0 : k][0]);
        return seriesLineDefsOf(key, "lineRatioNumerator").map((def) => {
            const group = def.series!.group;
            const denominator = group.values.find((c) => c.source?.roles?.lineRatioDenominator && keyOfColumn(c) === denominatorKey);
            const ratio = { numerators: def.values, denominators: numbersAt(denominator) };
            const highlightRatio = def.highlights && denominator?.highlights ? { numerators: def.highlights, denominators: numbersAt(denominator, true) } : undefined;
            return {
                ...def,
                values: divideRatio(ratio.numerators, ratio.denominators),
                ratio,
                ...(highlightRatio ? { highlights: divideRatio(highlightRatio.numerators, highlightRatio.denominators), highlightRatio } : { highlights: undefined }),
            };
        });
    };
    const ratioLineDefs: LineDef[] = splitLinesByLegend
        ? pairedNumerators.flatMap(([key], k) => seriesRatioDefsOf(key, k))
        : pairedNumerators.map(([key, numerators], k) => {
        const denominators = denominatorList[sharedDenominator ? 0 : k];
        const ratio = {
            numerators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(numerators, i)),
            denominators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(denominators, i)),
        };
        const highlightRatio = [...numerators, ...denominators].some((c) => c.highlights)
            ? {
                numerators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(numerators, i, true)),
                denominators: Array.from({ length: rowCount }, (_, i) => seriesSumOf(denominators, i, true)),
            }
            : undefined;
        const source = numerators[0].source;
        return {
            key,
            // 線の名前は分子の名前（ビジュアルの中でフィールドの名前を変えれば、線の名前も変わる）
            name: source.displayName,
            format: valueFormatter.getFormatStringByColumn(source),
            objects: source.objects,
            values: divideRatio(ratio.numerators, ratio.denominators),
            ratio,
            ...(highlightRatio ? { highlights: divideRatio(highlightRatio.numerators, highlightRatio.denominators), highlightRatio } : {}),
        };
    });
    const lineDefs = [...measureLineDefs, ...ratioLineDefs].map((def) => ({
        ...def,
        format: lineFormatOf(def, settings.lines),
    }));
    // 空白・空の文字・数にならない値・NaN・±Infinity は空白（0 で除算したメジャーなど）。空白の数は 0
    const isBlankAt = (column: DataViewValueColumn, i: number) => numberOf(column.values[i]) === null;
    const numberAt = (column: DataViewValueColumn, i: number) => numberOf(column.values[i]) ?? 0;

    const getDropdownValue = enumValue;

    const chartTypeValue = getDropdownValue(settings.chart.chartType?.value, CHART_TYPES.clustered);
    // パレート。100% 積み上げのパレートはどの棒も同じ高さで意味が無いので、そのあいだは積み上げとして描く
    const paretoOn = getDropdownValue(settings.calculation.mode.value, CALCULATION_MODES.none) === CALCULATION_MODES.pareto;
    // 1.18 までの「リボン」は、書式の読み込み（applyChartTypeDefaults）で積み上げ＋リボンに読み替える。ここでは積み上げとして扱う
    const chartType = effectiveChartTypeOf(chartTypeValue, paretoOn);
    const stacked = chartType !== CHART_TYPES.clustered;
    const percent = chartType === CHART_TYPES.stacked100;
    // リボン（帯）は積み上げ・100% 積み上げで出せる。積む順が「値の大きい順」なら、カテゴリごとに値の大きい系列を外側に積み替える
    // （標準のリボン グラフと同じ。凡例の順なら標準の積み上げ＋リボンと同じ）
    const ribbonsOn = stacked && (settings.ribbons.show.value ?? false);
    const rankOrder = ribbonsOn && getDropdownValue(settings.ribbons.order.value, RIBBON_ORDERS.legend) === RIBBON_ORDERS.value;
    // 積み上げ順を逆にする（リボンで値の大きい順に積むときは、その順が優先）
    const reverseStack = stacked && !rankOrder && (settings.columns.reverseStackOrder.value ?? false);

    // 複数系列の空白は棒を描かないので数えない（系列 1 本の空白は 1.4 までと同じく 0 として数える）
    const isSkipped = (column: DataViewValueColumn, i: number) => (seriesMode || lineOnly) && isBlankAt(column, i);

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

    // --- X 軸の種類「連続」 ------------------------------------------------------
    // X 軸が日付か数値の 1 列（階層なら、上のレベルがどの行でも同じ＝いちばん下まで 1 つの親にドリルしたとき）なら選べる。
    // 縦棒だけで、パレート・「その他」・値の空白があるときは選べない（書式ペインにも出さない）。
    // 連続では値の小さい順に並べる（並べ替え・逆順は効かない。標準と同じ）
    const axisType = lowestLevel.source?.type;
    const axisIsDate = !!axisType?.dateTime;
    const axisIsInteger = !!axisType?.integer;
    const singleAxisLevel = levelColumns.length === 1 || sharedDepth === levelColumns.length - 1;
    const axisValues = singleAxisLevel && (axisIsDate || !!axisType?.numeric || axisIsInteger) ? lowestLevel.values.map((raw) => axisValueOf(raw, axisIsDate)) : [];
    const continuousAvailable = axisValues.length > 0 && !horizontal && !paretoOn && otherRow < 0 && axisValues.every((v) => v !== null);
    const continuousOn = continuousAvailable && getDropdownValue(settings.categoryAxis.axisKind.value, AXIS_KINDS.categorical) === AXIS_KINDS.continuous;
    if (continuousOn) displayOrder = Array.from({ length: rowCount }, (_, i) => i).sort((a, b) => axisValues[a]! - axisValues[b]!);
    // 比較レイヤーの奥のレイヤーは、手前のレイヤーの並びに合わせる（値順でも、手前の値の順に並べる）
    if (basis.displayOrder && basis.displayOrder.length === rowCount) displayOrder = basis.displayOrder.slice();

    // 系列の並び（集合で並べる順・積み上げで 0 から積む順）。値順は全カテゴリの合計（空白は 0）の大きい順、
    // 逆にするのは集合だけ（積み上げは「積み上げ順を逆にする」）。比較レイヤーの奥は手前の並びに合わせる
    let seriesOrder = slots.map((_, s) => s);
    if (seriesMode && (settings.columns.seriesSortByValue.value ?? false)) {
        const totalOf = (s: number) => Array.from({ length: rowCount }, (_, i) => (isSkipped(slots[s].column, i) ? 0 : numberAt(slots[s].column, i))).reduce((a, b) => a + b, 0);
        const totals = seriesOrder.map(totalOf);
        seriesOrder = [...seriesOrder].sort((a, b) => totals[b] - totals[a] || a - b);
    }
    if (seriesMode && !stacked && (settings.columns.seriesReverseOrder.value ?? false)) seriesOrder.reverse();
    if (basis.seriesOrder && basis.seriesOrder.length === slots.length) seriesOrder = basis.seriesOrder.slice();
    const seriesPositions = slots.map((_, s) => seriesOrder.indexOf(s));
    /** 0 の側から積む系列の順 */
    const stackSequence = reverseStack ? [...seriesOrder].reverse() : seriesOrder;
    const stackReordered = stackSequence.some((s, k) => s !== k);

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
            const value = numberOf(raw);
            if (value === null) {
                out[i] = keepBlank && !started ? null : running;
                return;
            }
            running += value;
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
            // ハイライトは棒と同じく「ここまでの該当分」
            if (def.highlights) {
                def.highlights = def.highlightRatio
                    ? divideRatio(
                        cumulate(def.highlightRatio.numerators, false) as Array<number | null>,
                        cumulate(def.highlightRatio.denominators, false) as Array<number | null>
                    )
                    : (cumulate(def.highlights, false) as Array<number | null>);
            }
            cumulativeLineKeys.add(def.key);
        }
    }
    // 「その他」の行の折れ線は、足しようが無いので空白（「値」と同じ列を「折れ線の値」にも入れたとき・累計で空白が埋まったときも）。
    // 率の線は、まとめたカテゴリの分子の合計 ÷ 分母の合計なので出せる（累計でなければ。累計の「その他」は区切りをまたぐので空白）
    if (otherRow >= 0) lineDefs.forEach((def) => {
        if (!def.ratio || cumulativeLineKeys.has(def.key)) {
            def.values[otherRow] = null;
            if (def.highlights) def.highlights[otherRow] = null;
        }
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
    const keptRows = otherRow >= 0 ? keptRowsOf(originalCategories![0].values.length, others.mergedRows) : null;
    // 行はここまでで読み終える。すべてのレイヤーの範囲（と軸のタイトルの名前）を受け取ってから続きを組む
    const sharedInput = yield { basis: { extent: ownExtent, keepRows: keptRows, displayOrder, seriesOrder }, canLayer: !lineOnly && !paretoOn };
    if (!("shared" in sharedInput)) throw new Error("棒グラフの組み立てに、行を読んだあとの範囲が渡されていない");
    basis = { ...basis, ...sharedInput.shared };
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
    const lineMin = lineNumbers.length ? minOf(lineNumbers) : 0;
    const lineMax = lineNumbers.length ? maxOf(lineNumbers) : 0;
    const axis2Saved = (dataView?.metadata?.objects?.valueAxis2 as powerbi.DataViewObject | undefined)?.show;
    let onSecondary = false;
    /** 表示を保存していないときに自動で決めた結果（書式ペインのトグルに出す）。自動で決めなければ無し */
    let autoSecondary: boolean | undefined;
    // パレートでは右の軸を累積比（0〜100%）に使うので、「折れ線の値」の線は左の軸を共有する
    if (lineDefs.length && !paretoOn && !lineOnly) {
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
            autoSecondary = onSecondary;
        }
    }
    // 左の軸を共有するなら、軸の範囲と単位に折れ線の値も入れる
    if (lineDefs.length && !onSecondary) {
        maxValue = Math.max(maxValue, lineMax);
        minValue = Math.min(minValue, lineMin);
        // 1 つずつ足す（push(...xs) は点が多いと RangeError を投げる）
        for (const v of lineNumbers) allValues.push(v);
        maxAbs = Math.max(maxAbs, Math.abs(lineMax), Math.abs(lineMin));
    }

    // --- Y 軸の定数線の値 -------------------------------------------------------
    // カードの「値」（決まった数）と、「Y 軸の定数線」の欄のメジャー。メジャーはカテゴリごとに計算されて届くので、
    // 並べた順の先頭のカテゴリの値で引き、ほかのカテゴリで値が違えば警告する（凡例があると matrix は全体の値で 1 列届く）
    const vl = settings.valueLine;
    const valueLineOn = vl.show.value ?? true;
    /** grouping は、欄のメジャーの線の桁区切り（メジャーの書式で決める。ツールヒントと同じ列）。カードの「値」の線は無く、載せた軸の書式で決める */
    const valueLineDefs: Array<{ name: string; value: number; grouping?: Grouping }> = [];
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
            valueLineDefs.push({
                name: columns[0].source.displayName,
                value: values[0],
                grouping: groupingOf(numberFormatOf(valueFormatter.getFormatStringByColumn(columns[0].source))),
            });
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
    // 第 2 Y 軸の率と同じく、表示単位によらず % で出し、単位ラベル（単位の追加文字の「円」など）を付けない。
    // ‰ の書式なら 1000 倍して ‰。字で書いた %（0.0"%"。値がすでに % の数）は 100 倍せずに %（Power BI の書式と同じ）。名前の付いた Percent は 0.00% と読む
    // 折れ線だけのときは、左の軸の書式を線の書式で決める（線が % なら目盛りも %）
    const measureFormats = lineOnly ? lineDefs.map((def) => def.format) : slots.map((slot) => valueFormatter.getFormatStringByColumn(slot.column.source) ?? "");
    /** 値のメジャーがすべて同じ率の書式なら、その倍率・記号・桁（0.0% なら 100・%・1） */
    const measureRate = percent ? null : sameRateOf(measureFormats.map((format) => numberFormatOf(format)));
    const percentMeasure = measureRate !== null;
    /** 書式の % の前の小数の桁（0.0% なら 1）。小数点以下の桁数が「自動」のときに使う */
    const measurePercentDigits = measureRate?.digits ?? 0;
    const measureScale = measureRate?.scale ?? 100;
    // 自前で書く数（目盛り・データ ラベル・合計・定数線）の桁区切り（1,000）は、標準と同じく値のメジャーの書式に従う
    // （#,0 なら付け、0・0.00・書式なしなら付けない）。目盛り・合計・定数線は、載るメジャーのどれかの書式に桁区切りがあれば付ける。
    // データ ラベルは系列のメジャーの書式で決める（slotGroupings）
    const measureGrouping = someGroupingOf(measureFormats.map((format) => numberFormatOf(format)));
    const measurePercentText = (v: number, digits: string, sign?: Partial<SignStyle>, grouping = measureGrouping) =>
        formatPercent(v, digits, measurePercentDigits, sign, measureRate ?? undefined, grouping);
    // 値のメジャーの書式の「正;負;0」の節（#,0;(#,0);-）。目盛りと合計ラベルは、載るメジャーの節の書き方がすべて同じときだけ当てる
    // （標準と同じく、目盛りには書式ペインの符号の書き方を使わない）。100% 積み上げの軸は割合なので当てない
    const measureSections = percent ? null : sameSectionSignOf(measureFormats.map((format) => numberFormatOf(format)));
    const unitDef = percentMeasure ? resolveUnit("0", 0, unitNotation, precision) : resolveUnit(unitTypeKey, maxAbs, unitNotation, precision);

    const roundRange = valAxis.roundRange.value ?? true;
    const invertRange = valAxis.invertRange.value ?? false;
    const logarithmic = valAxis.logarithmic.value ?? false;

    let userStart: number | undefined = undefined;
    let userEnd: number | undefined = undefined;
    // 桁区切りのカンマ（1,000）も読む。1.30 までは読めずに自動になっていた
    userStart = boundOf(valAxis.start.value) ?? undefined;
    userEnd = boundOf(valAxis.end.value) ?? undefined;

    // 対数はデータが全て正か全て負のときだけ効かせる。0 を含む・正負が混在するときは線形で描く。
    // 標準の値軸と同じ扱い（Microsoft Learn「軸のカスタマイズ」: 対数は全て正か全て負が必要で 0 は不可、
    // データが変わって条件を外れると自動で線形に切り替わる）。標準は書式ペインに警告アイコンを出すが
    // カスタムビジュアルからは出せないため、状態は valueAxis.logarithmicFallback に持つ。
    // 折れ線だけのときは棒の値が無く、範囲の下端の 0 は棒の基準なので、線の値で決める（標準の折れ線と同じく、線がすべて正なら対数）
    const logSign = lineOnly ? (lineMin > 0 ? 1 : lineMax < 0 ? -1 : 0) : minValue > 0 ? 1 : maxValue < 0 ? -1 : 0;
    // 積み上げでは対数を使わない（積んだ棒の長さが値の和にならなくなる）
    const isLogScaleActive = logarithmic && logSign !== 0 && !stacked;
    const logarithmicFallback = logarithmic && !isLogScaleActive;
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

    // X軸設定の抽出
    const catAxis = settings.categoryAxis;
    const concatenateLabels = catAxis.concatenateLabels.value ?? false;
    const levelNames = levelColumns.map((column) => column.source?.displayName ?? "").slice(commonDepth);
    const autoCategoryTitle = concatenateLabels ? levelNames.join(" ") : levelNames[levelNames.length - 1];
    const categoryAxisSettings: CategoryAxisSettings = {
        show: catAxis.show.value ?? true,
        ...fontFields(catAxis.font, "Segoe UI", 9),
        labelColor: catAxis.labelColor.value?.value || "#605E5C",
        maxHeight: Math.max(0, Math.min(100, catAxis.maxHeight.value ?? 25)),
        titleShow: catAxis.titleShow.value ?? false,
        // 自動のタイトルは標準と同じ：段に重ねるならいちばん下のレベル、1 行につなぐならレベルの名前を並べる。
        // X 軸に単位は無いので、スタイルが「単位のみ」ならタイトルを出さず、「両方」はタイトルだけ
        titleText:
            getDropdownValue(catAxis.titleStyle.value, "showTitleOnly") === "showUnitOnly" ? "" : catAxis.titleText.value?.trim() || autoCategoryTitle,
        ...titleFontFields(catAxis.titleFont, "DIN", 12),
        titleColor: catAxis.titleColor.value?.value || "#252423",
        minCategoryWidth: Math.max(0, Math.min(500, catAxis.minCategoryWidth.value ?? 20)),
        scrollStart: getDropdownValue(catAxis.scrollStart.value, "start"),
        concatenateLabels,
        levelCount: levelColumns.length - commonDepth,
        levelNames,
        hierarchyStyle: getDropdownValue(catAxis.hierarchyStyle.value, "lines") === "boxed" ? "boxed" : "lines",
        levelRowHeight: getDropdownValue(catAxis.levelRowHeight.value, "normal"),
    };

    // ユーザー設定の単位ラベル（例: "億円", "円", "bn"）を反映したタイトル生成。
    // 値が複数のときの自動のタイトルは、標準と同じく値の名前を「および」でつなぐ
    const composedUnit = percent ? "" : composeUnitText(effectiveUnitWord, unitText, unitIncludeDisplayUnit);
    // メジャーごとに棒を並べるときは、すべてのメジャーの名前。凡例ごとに分けた線は、凡例の値ではなくメジャーの名前（標準と同じ）
    const autoValueTitle = basis.valueTitleNames?.length
        ? basis.valueTitleNames.join(" および ")
        : lineOnly
        ? [...new Set(lineDefs.map((def) => def.series?.measureName ?? def.name))].join(" および ")
        : seriesMode && !legendSource
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
        ...fontFields(valAxis.font, "Segoe UI", 9),
        labelColor: valAxis.labelColor.value?.value || "#605E5C",
        unitNotation,
        showUnitOnAxis,
        switchPosition: valAxis.switchPosition.value ?? false,
        titleShow: valAxis.titleShow.value ?? true,
        titleText: formattedValueTitle,
        titleStyle: valueTitleStyle,
        ...titleFontFields(valAxis.titleFont, "DIN", 12),
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
        verticalPosition: getDropdownValue(gl.verticalPosition.value, "categories"),
    };

    // データラベル設定の抽出
    const dl = settings.dataLabels;
    const dataLabelsSettings: DataLabelsSettings = dataLabelsSettingsOf(dl);

    // 列（columns）設定の抽出
    const col = settings.columns;
    // 系列の展開は積み上げだけ。折れ線があると値の軸を消せないので使わない（標準の複合にも無い）
    const explodeOn = stacked && (col.stackedExplode.value ?? false) && lineDefs.length === 0;
    const overlapOn = !stacked && (col.overlap.value ?? false);
    // 系列 1 本の棒の色：保存が無ければ、標準と同じくテーマのデータの色の 1 番目（1.28 までは #118DFF 固定で、
    // テーマの 1 番目を変えたレポートでも従わなかった）。getColor は同じ key なら同じ色を返すので、下の折れ線の色の取り方とずれない
    const singleSeriesFill =
        seriesMode || lineOnly || !slots.length ? null : customColor(dataView?.metadata?.objects, "fill") ?? host.colorPalette.getColor(slots[0].key).value;
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
        rows: legendRowsOf(lg.rows.value?.value),
        title: (lg.titleShow.value ?? true) ? (lg.titleText.value?.trim() || (legendSource?.displayName ?? "")) : "",
        titleShow: lg.titleShow.value ?? true,
        ...fontFields(lg.font, "Segoe UI", 10),
        color: lg.labelColor.value?.value || "#605E5C",
    };

    // 系列。複数系列の色は、個別の指定が無ければ標準と同じくレポートのテーマの色を順に割り当てる。
    // 系列 1 本のときは 1.4 までと同じく「列」のカラー（カテゴリごとの指定があればそちら）
    /** 書式ペインに出す系列の色（「棒の色」で塗っても、ペインには自分で選んだ色かテーマの色を出す） */
    const ownSeriesColors: string[] = [];
    /** 「棒の色」で塗る系列（書式ペインのカラーが効かない） */
    const fieldColoredSeries: boolean[] = [];
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
        fieldColoredSeries[ownSeriesColors.length - 1] = fieldColor !== null;
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
            // 系列ごとに保存したデータ ラベルの項目（位置・タイトル・値・詳細・背景・レイアウト・文字の縁）。無ければ「すべて」の設定
            labelSettings: seriesMode && hasOwnLabelSettings(slot.objects) ? dataLabelsSettingsOf(dl, slot.objects) : undefined,
        };
    });

    // 合計ラベル（積み上げのときだけ描く）。表示単位は「自動」なら Y 軸に従い（単位ラベルがあるので語は付けない）、
    // 選んだときはその単位で割って語を付ける
    const tl = settings.totalLabels;
    const totalPrecisionValue = String(tl.precision.value?.value ?? "auto");
    const totalPrecision = totalPrecisionValue !== "auto" ? totalPrecisionValue : precision;
    const totalUnitKey = getDropdownValue(tl.unitType.value, "auto");
    const totalUnitDef = totalUnitKey === "auto" ? null : resolveUnit(totalUnitKey, maxAbs, unitNotation, totalPrecision);
    // 合計ラベルは、データ ラベルの符号の書き方（書式ペイン）を使わず、メジャーの書式の節だけに従う（節が無ければ今までどおり）
    const totalText = (v: number): string =>
        percentMeasure
            ? measurePercentText(v, totalPrecision, measureSections ?? undefined)
            : totalUnitDef
            ? formatSigned(v, totalUnitDef.divisor, totalPrecision, measureSections ?? {}, totalUnitDef.unitWord, measureGrouping)
            : formatSigned(v, unitDef.divisor, totalPrecision, measureSections ?? {}, "", measureGrouping);
    const totalLabelsSettings: TotalLabelsSettings = {
        show: chartType === CHART_TYPES.stacked && (tl.show.value ?? false),
        split: tl.splitPositiveNegative.value ?? false,
        ...fontFields(tl.font, "Segoe UI", 9),
        color: tl.color.value?.value || "#605E5C",
        backgroundShow: tl.backgroundShow.value ?? false,
        backgroundColor: tl.backgroundColor.value?.value || DEFAULT_LABEL_BACKGROUND,
        backgroundTransparency: Math.max(0, Math.min(100, tl.backgroundTransparency.value ?? DEFAULT_LABEL_BACKGROUND_TRANSPARENCY)),
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
    const baseLabelPrecision = labelPrecision;
    // データラベルのマイナスと 0 の書き方（▲・±0 など）
    const labelSign: Partial<SignStyle> = {
        negative: getDropdownValue(dl.negativeStyle.value, NEGATIVE_STYLES.minus),
        zero: getDropdownValue(dl.zeroStyle.value, ZERO_STYLES.zero),
        negativeZero: dl.negativeZero.value ?? true,
    };
    const savedLabels = dataView?.metadata?.objects?.dataLabels;
    // 系列（メジャー）ごとのデータ ラベルの符号の書き方。「負の値の書き方」「0 の書き方」をレポートに保存していなければ、
    // メジャーの書式の「正;負;0」の節に従う（標準と同じ）。符号の書き方は系列ごとの項目に無いので、「すべて」の保存だけを見る
    const sectionLabelSignOf = (format: string | undefined) => withSectionSign(labelSign, sectionSignOf(numberFormatOf(format)), savedLabels);
    const slotLabelSigns = slots.map((slot) => sectionLabelSignOf(valueFormatter.getFormatStringByColumn(slot.column.source) ?? ""));
    // 系列（メジャー）ごとのデータ ラベルの桁区切り。系列のメジャーの書式に桁区切りがあれば付ける（標準と同じ）
    const slotGroupings = slots.map((slot) => groupingOf(numberFormatOf(valueFormatter.getFormatStringByColumn(slot.column.source))));
    // 値の行を符号で塗る（見える符号で決める。▲0 はマイナス、±0 は塗らない）
    const toneMode = getDropdownValue(dl.toneMode.value, TONE_MODES.none);
    // 書式ペインで決めていなければ、テーマの良い・悪いの色を字に使える濃さにして使う
    const themeTones = sentimentTextColors(host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette);
    const toneColors = {
        good: ownToneColor(dl.positiveColor.value?.value, savedLabels?.positiveColor, DEFAULT_GOOD_COLOR) || themeTones.good,
        bad: ownToneColor(dl.negativeColor.value?.value, savedLabels?.negativeColor, DEFAULT_BAD_COLOR) || themeTones.bad,
    };
    const toneColorOf = (sign: number): string => {
        const tone = toneOf(sign, 1, toneMode);
        return tone ? toneColors[tone] : "";
    };
    const formattedOnAxis = (val: number, labelPrecision: string, sign: Partial<SignStyle>, grouping: Grouping) =>
        // 100% 積み上げの軸は割合なので、データラベルは値ごとに単位を付ける（1,250億 など）。率のメジャーは対数でも % で書く
        percentMeasure
            ? {
                formattedValue: measurePercentText(val, precision, undefined, grouping),
                dataLabelText: measurePercentText(val, labelPrecision, sign, grouping),
                labelToneColor: toneColorOf(shownSignOf(val * measureScale, 1, labelPrecision === "auto" ? String(measurePercentDigits) : labelPrecision, sign)),
            }
            : isLogScaleActive || percent
            ? {
                formattedValue: formatDynamicValue(val, unitNotation, precision, true, {}, grouping),
                dataLabelText: formatDynamicValue(val, unitNotation, labelPrecision, true, sign, grouping),
                labelToneColor: toneColorOf(dynamicShownSign(val, unitNotation, labelPrecision, sign)),
            }
            : {
                formattedValue: formatValue(val, unitDef.divisor, precision, grouping),
                dataLabelText: formatSigned(val, unitDef.divisor, labelPrecision, sign, "", grouping),
                labelToneColor: toneColorOf(shownSignOf(val, unitDef.divisor, labelPrecision, sign)),
            };
    // 値の表示単位を選んだら、棒のラベルの文字だけその単位で割って語を付ける（ツールヒントは軸の単位のまま。合計ラベルと同じ）。
    // 「Y 軸と同じ」と % の書式のメジャーは今までどおり
    const labelUnitKey = getDropdownValue(dl.unitType.value, "auto");
    const labelUnitDef = labelUnitKey === "auto" || percentMeasure ? null : resolveUnit(labelUnitKey, maxAbs, unitNotation, baseLabelPrecision);
    /** 棒の値のデータ ラベルの文字。sign は系列の符号の書き方（slotLabelSigns）、grouping は系列の桁区切り（slotGroupings） */
    const formatted = (val: number, labelPrecision: string, sign: Partial<SignStyle>, grouping: Grouping) => {
        const onAxis = formattedOnAxis(val, labelPrecision, sign, grouping);
        if (!labelUnitDef) return onAxis;
        return {
            ...onAxis,
            dataLabelText: formatSigned(val, labelUnitDef.divisor, labelPrecision, sign, labelUnitDef.unitWord, grouping),
            labelToneColor: toneColorOf(shownSignOf(val, labelUnitDef.divisor, labelPrecision, sign)),
        };
    };

    // データラベルの詳細の行。全体に対する割合は、100% 積み上げと複数系列ではカテゴリの合計（正と負の絶対値）に対する割合、
    // 系列 1 本ではすべてのカテゴリの合計に対する割合（累計なら最後の累計に対する割合。下）。カスタムは「ラベルの詳細」に入れたフィールドの値
    const detailContent = getDropdownValue(dl.detailContent.value, DETAIL_CONTENTS.percentOfTotal);
    const detailPrecision = getDropdownValue(dl.detailPrecision.value, "auto");
    const detailUnitKey = getDropdownValue(dl.detailUnitType.value, "auto");
    const grandAbsolute = absoluteSums.reduce((sum, v) => sum + v, 0);
    // 累計の系列 1 本では、全体は最後に描く累計の棒（区切りで 0 に戻すなら、その区切りの最後の棒）。最後の棒が 100%
    // （最後の累計が負なら -100%。割合の符号は棒の向きと同じ）。累計した値の総和で割ると、手前の値を何度も数えて割合が小さくなる。
    // 並べ替え・「その他」（いつも最後）・負の値も、描く順の最後の棒の累計（絶対値）で割る。
    // 区切りのレベルがいちばん下なら棒ごとに 0 に戻って棒は素の値なので、累計でないときと同じくすべての棒の合計で割る
    const cumulativeWholeRow = new Array<number>(rowCount).fill(-1);
    // 区切りの中でいちばん大きい累計（絶対値）。最後の累計が 0 かを、小数の足し算の誤差を見込んで決めるのに使う
    const cumulativePeak = new Array<number>(rowCount).fill(0);
    if (cumulativeEnabled && !(resetDepth >= 0 && resetDepth === levelItems.length - 1)) {
        let last = -1;
        for (let k = displayOrder.length - 1; k >= 0; k--) {
            if (last < 0) last = displayOrder[k];
            cumulativeWholeRow[displayOrder[k]] = last;
            cumulativePeak[last] = Math.max(cumulativePeak[last], absoluteSums[displayOrder[k]]);
            if (resetBefore[k]) last = -1;
        }
    }
    // 自動は標準と同じく小数 2 桁（30.00%）
    const percentDigits = detailPrecision === "auto" ? 2 : Number(detailPrecision);
    const percentText = (ratio: number) =>
        `${localeNumber(ratio * 100, percentDigits, percentDigits)}%`;
    const fieldFormatters = new Map<DataViewValueColumn, ReturnType<typeof valueFormatter.create>>();
    // ラベルの欄の文字・日付などを出す関数（列と空白の字ごとに 1 回作る。行ごとに書式を作り直さない）
    const fieldNameTexts = new Map<DataViewValueColumn, Map<string, (raw: unknown) => string>>();
    const fieldNameTextOf = (column: DataViewValueColumn, blankAs: string) => {
        let byBlank = fieldNameTexts.get(column);
        if (!byBlank) fieldNameTexts.set(column, (byBlank = new Map()));
        let text = byBlank.get(blankAs);
        if (!text) byBlank.set(blankAs, (text = nameTextOf(formatOfColumn(column.source), blankAs)));
        return text;
    };
    const detailTextOf = (slot: SeriesSlot, i: number, val: number, skipped: boolean): string => {
        if (skipped) return "";
        if (detailContent !== DETAIL_CONTENTS.custom) {
            const whole = seriesMode || percent ? -1 : cumulativeWholeRow[i];
            const denominator = seriesMode || percent ? absoluteSums[i] : whole >= 0 ? absoluteSums[whole] : grandAbsolute;
            // 累計で最後の棒が 0 なら、全体が 0 で割合が決まらない（0% と書かない）。小数の累計は足し算の誤差で 0 ちょうどに
            // ならない（1.1 + 2.2 - 3.3 は 4.4e-16）ので、区切りの中のいちばん大きい累計の 10 億分の 1 までは 0 とみなす
            if (whole >= 0 && denominator <= cumulativePeak[whole] * 1e-9) return "";
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
        if (isBlankName(raw)) return blankAs;
        // 文字・日付などは列の書式で出す（日付を String() にすると JS の日付の文字列になる）。
        // NaN・±Infinity は名前と同じく空白にせず字で出す
        if (typeof raw !== "number" || !Number.isFinite(raw)) return fieldNameTextOf(column, blankAs)(raw);
        const format = valueFormatter.getFormatStringByColumn(column.source) ?? "";
        // 自前で書くとき（表示単位・桁を選んだとき）は、フィールドの書式の「正;負;0」の節の符号の書き方を当てる（節が無ければ - と 0）
        const sign = sectionSignOf(numberFormatOf(format)) ?? {};
        // 桁区切りもフィールドの書式に従う（書式のままで書くときと同じ）
        const grouping = groupingOf(numberFormatOf(format));
        if (unitKey !== "auto") {
            const unit = resolveUnit(unitKey, Math.abs(raw), unitNotation, digits);
            return formatSigned(raw, unit.divisor, digits, sign, unit.unitWord, grouping);
        }
        if (digits !== "auto") {
            // 率の書式なら倍率を掛けて記号を付ける（% は 100 倍、‰ は 1000 倍、字で書いた % の 0.0"%" は倍率 1 で %。名前の付いた Percent は 0.00% と読む）
            const rate = rateFormatOf(numberFormatOf(format));
            return rate ? formatSigned(raw * rate.scale, 1, digits, sign, rate.mark, grouping) : formatSigned(raw, 1, digits, sign, "", grouping);
        }
        let formatter = fieldFormatters.get(column);
        if (!formatter) {
            // 名前の付いた数の書式（Currency など）は同じ意味の書式か書式なしに直し、日付の書式は当てない（書式の字そのものを出さない）
            formatter = valueFormatter.create({ format: numberValueFormatOf(format) });
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

    /** 棒のデータ ラベルの値の小数点以下の桁数。系列ごとに決めていればその桁 */
    const labelPrecisionOf = (own: DataLabelsSettings | undefined) => (own && own.precision !== "auto" ? own.precision : labelPrecision);

    // 棒の外の端に出しうるデータ ラベルの長さ（値の向き、px。縦向きは行の幅、横書きは行の高さ）。ラベルの文字は値の軸の範囲では
    // 変わらない（表示単位はデータの最大の絶対値で決まる）ので、範囲を決める前に、棒を組むときと同じ文字で測る。系列ごとのラベルの設定も見る。
    // 測る関数を yield で渡し、値の軸に取る余白を受け取ってから続きを組む（比較レイヤーは、レイヤーのあいだで余白をそろえる）。
    // 測るのは、余白を使う組み立てが関数を呼んだときだけ（比較レイヤーの 1 回目のように、並びと範囲だけを使う組み立てでは測らない）。
    // 文字を作るのは重い（3 万行で数秒）ので、測るときに作った文字は棒を組むときに使い回す
    const madeTexts: Array<Array<MadeLabelTexts | undefined>> = slots.map((): Array<MadeLabelTexts | undefined> => []);
    // 系列のデータ ラベルが値の軸に余白を取りうるか（棒の外に置きうるか）。集合の縦棒・横棒の「自動」と「外側の上」は、入るかによらず
    // 外に置く。積み上げのデータ ラベルは外に出さない（外側は上の棒。App で「自動」「外側の上」を内側の中央に読み替える）
    const labelsOf = (s: number) => seriesStyles[s].labelSettings ?? dataLabelsSettings;
    const labelTakesRoom = (s: number) => !stacked && seriesStyles[s].labelShow && labelMayGoOutside(labelsOf(s));
    // 値の軸の長さが分かれば余白を取るか（App は、このときだけ値の軸の長さを知らせる）
    const valueAxisRoom = totalLabelsSettings.show || slots.some((_, s) => labelTakesRoom(s));
    const measureLabels = (): LabelRoom => {
        const measured: LabelRoom = { top: [], bottom: [] };
        if (!((runtime.valueAxisLength ?? 0) > 0)) return measured;
        // 合計ラベル（積み上げ）。正の合計は正の棒の端の外に、負の合計は負の棒の端の外に置く（どれをどちらに置くかは App と同じ
        // totalLabelSpotsOf）。縦棒の長さは文字の大きさだけで決まり、横棒は文字の幅を測る（同じ文字は測った幅を使い回す）
        if (totalLabelsSettings.show) {
            const tl = totalLabelsSettings;
            const font = { family: tl.fontFamily, size: tl.fontSize * PT_TO_PX, bold: tl.bold, italic: tl.italic, underline: tl.underline };
            const widthByText = new Map<string, number>();
            const lengthOf = (text: string, up: boolean): number => {
                // 範囲の反転では、値の大きい方が下（横棒は左）になる
                const placement = { horizontal, screenUp: up !== invertRange };
                if (!horizontal) return totalLabelLengthOf(text, font, placement);
                let length = widthByText.get(text);
                if (length === undefined) widthByText.set(text, (length = totalLabelLengthOf(text, font, placement)));
                return length;
            };
            for (let i = 0; i < rowCount; i++) {
                const positive = positiveSums[i];
                const negative = negativeSums[i];
                for (const { kind, up } of totalLabelSpotsOf(positive, negative, tl.split)) {
                    const text = totalText(kind === "positive" ? positive : kind === "negative" ? negative : positive + negative);
                    (up ? measured.top : measured.bottom).push({ value: up ? positive : negative, length: lengthOf(text, up) });
                }
            }
        }
        for (let s = 0; s < slots.length; s++) {
            if (!labelTakesRoom(s)) continue;
            const slot = slots[s];
            const style = seriesStyles[s];
            const labels = labelsOf(s);
            // 同じ文字のラベルは、測った長さを使い回す（系列ごとに設定が違いうるので、系列ごとに持つ）
            const roomByText = new Map<string, { length: number; slack: number }>();
            for (let i = 0; i < rowCount; i++) {
                // 空白の棒にはラベルを出さない（系列 1 本の空白は 0 の棒で描くが、ラベルは出さない）
                if (isBlankAt(slot.column, i)) continue;
                const val = numberAt(slot.column, i);
                const valueField = labels.valueShow ? valueFieldTextOf(slot, i) : undefined;
                const made: MadeLabelTexts = {
                    valueField,
                    formatted: labels.valueShow && valueField === null ? formatted(val, labelPrecisionOf(style.labelSettings), slotLabelSigns[s], slotGroupings[s]) : undefined,
                    title: labels.titleShow ? titleTextOf(slot, i, false) : undefined,
                    detail: labels.detailShow ? detailTextOf(slot, i, val, false) : undefined,
                };
                madeTexts[s][i] = made;
                const texts = { dataLabelText: made.valueField ?? made.formatted?.dataLabelText ?? "", titleText: made.title ?? "", detailText: made.detail ?? "" };
                // 縦棒の横書きは、値の端が上か下かで長さが少し違う（範囲の反転では、値の大きい方が下）。
                // 値が 0 の棒のラベルは、正の棒と同じく 0 の線の上（横棒は右）に置く（App の placeLabel・placeBarLabel と同じ）
                const positive = val >= 0;
                const up = positive !== invertRange;
                const key = `${texts.titleText}\u0000${texts.dataLabelText}\u0000${texts.detailText}\u0000${up}`;
                let room = roomByText.get(key);
                if (room === undefined) {
                    const shown = { ...texts, labelShow: true, blank: false };
                    const length = outsideLabelLengthOf(labels, shown, horizontal, up);
                    // 縦棒の上に置く横書きは、標準と同じく字の上をもう少し空ける（outsideLabelHeadroomOf）。字の外の余白として数える
                    const headroom = length > 0 ? outsideLabelHeadroomOf(labels, shown, horizontal, up) : 0;
                    room = { length: length + headroom, slack: outsideLabelSlack(horizontal, labels.backgroundShow, up) + headroom };
                    roomByText.set(key, room);
                }
                if (room.length > 0) (positive ? measured.top : measured.bottom).push({ value: val, length: room.length, slack: room.slack });
            }
        }
        return unionLabelRoom([measured]);
    };
    // 折れ線の色は、棒の系列の続きのテーマの色（系列 1 本の棒はテーマの 1 番目を使う扱いにして、線は 2 番目から）
    // 折れ線だけのときは棒が無いので、線はテーマの 1 番目から。colorPalette は呼んだ順に色を割り当てるので、ここで止める前に取る
    // （比較レイヤーは、どのレイヤーもここまで組んでから続きを組む。手前の棒と線の色を、奥のレイヤーの色より先に取る）
    if (lineDefs.length && !seriesMode && !lineOnly) host.colorPalette.getColor(slots[0].key);
    // 棒の系列と同じく、テーマの色は色の指定があっても取る（後ろの線の色がずれないように）。凡例ごとの線は凡例の値で取る
    const lineThemeColors = lineDefs.map((def) => host.colorPalette.getColor(def.series ? seriesKeyOf(def.series.group) : def.key).value);
    // 折れ線だけ・パレートでは比較レイヤーを重ねない（compareLayers は、重ねるときだけ 1 回目で測らない）
    const roomInput = yield { measure: measureLabels, canLayer: !lineOnly && !paretoOn };
    if (!("room" in roomInput)) throw new Error("棒グラフの組み立てに、値の軸の余白が渡されていない");
    const labelRoom = roomInput.room;

    let niceMin = 0;
    let niceMax = 1;
    let rawTicks: number[] = [];
    /** 目盛りを引く範囲（ラベルの余白を足す前の端）。対数と 100% 積み上げは軸の範囲のまま（null） */
    let tickFrame: [number, number] | null = null;
    // 値の軸の長さ（縦棒は高さ、横棒は幅。visual.ts が見積もった長さか、描いて知らせた長さ）。棒の外に出すラベルの余白を、この長さに対する比で取る
    const axisLength = runtime.valueAxisLength ?? 0;
    // 範囲を丸める目盛りの本数の目安。目盛り（App の tickTargetOf）と同じく、「目盛りの本数 (目安)」があればその本数、無ければ
    // 値の軸の長さで決める（標準の集合縦棒は高さ 150px 未満で 3、300px 未満で 5、それ以上で 8 の目安で丸める。2026-10-20 の Desktop で、
    // -520〜900 は 616×336 で -1,000〜1,000、616×672 で -600〜1,000）。長さが無ければ d3 の既定（10）
    const roundCount = valueAxisSettings.tickCount || (axisLength > 0 ? recommendedTickCount(axisLength, horizontal ? "horizontal" : "vertical") : undefined);
    // 描いた長さで丸めるか（App は、このときも値の軸の長さを知らせる）。第 2 Y 軸も同じ本数で丸める（下）
    let roundsByLength = false;

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
        // 最小値ちょうどにすると最小の棒が高さ 0 になって値が無いように見えるため 1 桁下げる。
        // 棒があるときは、標準の集合縦棒と同じく 1 から描く（データが 1 未満なら上の 10 の冪）。折れ線だけのときは標準の折れ線と同じく上の 10 の冪
        const belowMin = Math.pow(10, Math.ceil(log10Snap(magMin)) - 1);
        logLower = userLower ?? (lineOnly ? belowMin : Math.min(1, belowMin));
        // 棒の外の端に出すデータ ラベルの余白。対数の軸では、大きさの常用対数の空間で取る（0 から遠い側の端だけ。正の値は上、負の値は下の端）
        const farRoom = roomSharesOf(logSign > 0 ? labelRoom.top : labelRoom.bottom, axisLength, userUpper === undefined, (v) => Math.log10(Math.abs(v)));
        const roomUpperLog = withLabelRoom(Math.log10(logLower), Math.log10(magMax), farRoom, [])[1];
        const roomUpper = roomUpperLog > Math.log10(magMax) ? Math.pow(10, roomUpperLog) : magMax;
        // 端を丸めるときは 10 の冪。余白が要るなら、余白の端を 1・2・5 の刻みまで上げる（10 の冪まで上げると、棒が 1 桁ぶん縮む）
        logUpper = userUpper ?? (roundRange ? Math.max(Math.pow(10, Math.ceil(log10Snap(magMax))), logStepCeil(roomUpper)) : roomUpper);
        if (!(logUpper > logLower)) logUpper = logLower * 10;

        niceMin = logSign > 0 ? logLower : -logUpper;
        niceMax = logSign > 0 ? logUpper : -logLower;
        rawTicks = logAxisTicks(logLower, logUpper).map((m) => logSign * m);
    } else if (percent) {
        // 100% 積み上げの軸は割合。負の値があれば -100%〜100%、正だけなら 0%〜100%（標準と同じ）。
        // 最小値・最大値を入れたら、割合（0 なら 0%、-0.2 なら -20%）として使う。負の値がわずかで下の -100% が空くときに、
        // 最小値 0 で 0% を下端にできる（入れた範囲の外は端で切る）
        // -1〜1 の外の値（集合の軸のつもりで入れた 1000 など）は割合として読まず自動にする。上下が逆なら両方とも自動にする
        const ratioOrAuto = (v: number | undefined) => (v !== undefined && v >= -1 && v <= 1 ? v : undefined);
        const autoLower = minValue < 0 ? -1 : 0;
        const autoUpper = maxValue > 0 || minValue >= 0 ? 1 : 0;
        let lowerBound = ratioOrAuto(userStart) ?? autoLower;
        let upperBound = ratioOrAuto(userEnd) ?? autoUpper;
        if (!(upperBound > lowerBound)) {
            lowerBound = autoLower;
            upperBound = autoUpper > autoLower ? autoUpper : autoLower + 1;
        }
        niceMin = lowerBound;
        niceMax = upperBound;
        rawTicks = scaleLinear().domain([niceMin, niceMax]).ticks(5);
    } else {
        // 合計ラベルを出す積み上げでは、自動で決める端に範囲の 10% の余白を足す
        // （合計ラベルが軸の外で切れたり、カテゴリ名に掛かったりしないように）。大きい文字で 10% に入らないときは、
        // 下の labelRoom で合計ラベルの長さの分まで広げる（10% で入る大きさでは、範囲は前と変わらない）
        const totalLabelsOn = chartType === CHART_TYPES.stacked && (settings.totalLabels.show.value ?? false);
        const span = Math.max(maxValue, 0) - Math.min(minValue, 0);
        const headroom = totalLabelsOn ? span * 0.1 : 0;
        // 棒の外の端に出すデータ ラベルと合計ラベルの余白。描いた値の軸の長さに対する、測ったラベルの長さ（labelRoom）の比で、
        // 自動の端を広げる。広げる端は、反対側の実際の端（最小値・最大値で決めた端も）から測る（軸が 0 から始まるとは限らない）
        const topRoom = roomSharesOf(labelRoom.top, axisLength, userEnd === undefined);
        const bottomRoom = roomSharesOf(labelRoom.bottom, axisLength, userStart === undefined);
        // 余白を取る前の端（自動なら棒の端、決めた端はその値）。余白の上限（軸の 6 割）は、この範囲を基準にする。
        // 自動の端は 0 を含む。値が負だけなら上の端を 0 にする（正だけのときの下の端と同じ。標準の集合縦棒・横棒も 0）。すべて 0 なら 0〜1
        const barLower = userStart !== undefined ? userStart : (minValue < 0 ? minValue - headroom : 0);
        const barUpper = userEnd !== undefined ? userEnd : (maxValue > 0 ? maxValue + headroom : minValue < 0 ? 0 : 1);
        const round = roundRange && userStart === undefined && userEnd === undefined;
        roundsByLength = round && !valueAxisSettings.tickCount;
        const domain = labelRangeOf(barLower, barUpper, topRoom, bottomRoom, round, axisLength, roundCount);
        niceMin = domain[0];
        niceMax = domain[1] > niceMin ? domain[1] : niceMin + 1;
        // 目盛りは、ラベルの余白を足す前の範囲の中にだけ引く（標準と同じ。2026-10-20 の Desktop で、標準の集合縦棒は 36pt のラベルの
        // 余白で軸が 800 の上へ延びても、一番上の線は 800 のままで、延ばした所に線を引かなかった）
        const frame = barFrameOf(barLower, barUpper, round, roundCount);
        const low = Math.max(niceMin, frame[0]);
        const high = Math.min(niceMax, frame[1]);
        tickFrame = high > low ? [low, high] : null;
        rawTicks = scaleLinear().domain(tickFrame ?? [niceMin, niceMax]).ticks(5);
    }

    /**
     * 値の軸の比（0 が下端、1 が上端）。丸めないので、軸の範囲の外の値は 0 未満・1 超えになる。
     * 折れ線はこの比で描き、範囲の外は描画域で切る（標準と同じく、軸の端に貼り付けない）
     */
    const lineRatioOf = (val: number): number => {
        if (isLogScaleActive) {
            const mag = Math.abs(val);
            if (mag <= 0) return logSign > 0 ? 0 : 1;
            const logMin = Math.log10(logLower);
            const logSpan = Math.log10(logUpper) - logMin;
            const r = logSpan > 0 ? (Math.log10(mag) - logMin) / logSpan : 0;
            return logSign > 0 ? r : 1 - r;
        } else {
            const span = niceMax - niceMin;
            return span > 0 ? (val - niceMin) / span : 0;
        }
    };
    /** 棒などの比。軸の範囲の外は端（0〜1）で止める */
    const calcRatio = (val: number): number => clampRatio(lineRatioOf(val));

    // 対数では 0 を置けないため、棒の基準線は軸の端（正のデータは下端、負のデータは上端）
    const zeroRatio = isLogScaleActive ? (logSign > 0 ? 0 : 1) : calcRatio(0);

    // 目盛りラベルの生成
    const tickOf = (t: number): Tick => {
        let label: string;
        if (percent) {
            // 狭い範囲（最小値・最大値で 99.5%〜100% など）でも目盛りの名前が重ならないよう、間隔に合わせて小数を足す
            const stepPercent = rawTicks.length > 1 ? Math.abs(rawTicks[1] - rawTicks[0]) * 100 : 1;
            label = `${(t * 100).toFixed(stepPercent >= 1 ? 0 : Math.min(4, Math.ceil(-Math.log10(stepPercent))))}%`;
        } else if (percentMeasure) {
            // 目盛りの間隔が 1%（‰ なら 1‰）未満なら小数 1 桁（第 2 Y 軸の率と同じ。対数でも % で書く）
            const stepPercent = rawTicks.length > 1 ? Math.abs(rawTicks[1] - rawTicks[0]) * measureScale : 1;
            const fraction = precision !== "auto" ? Number(precision) : stepPercent < 1 ? 1 : 0;
            label = measureSections
                ? formatSectionTick(t, fixedNumber(Math.abs(t * measureScale), fraction, measureGrouping), measureSections, measureRate!.mark)
                : `${fixedNumber(t * measureScale, fraction, measureGrouping)}${measureRate!.mark}`;
        } else if (isLogScaleActive) {
            // 対数時は目盛値ごとに動的に単位付与（showUnitOnAxisがfalseなら生数値。桁区切りはメジャーの書式に従う）
            label = formatDynamicValue(t, unitNotation, precision, showUnitOnAxis, measureSections ?? {}, measureGrouping);
        } else {
            const word = showUnitOnAxis && unitDef.unitWord ? unitDef.unitWord : "";
            label = measureSections
                ? formatSectionTick(t, formatValue(Math.abs(t), unitDef.divisor, precision, measureGrouping), measureSections, word)
                : `${formatValue(t, unitDef.divisor, precision, measureGrouping)}${word}`;
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
        isLogScaleActive ? ticks : ticksUpTo(tickFrame ?? [niceMin, niceMax], count, (values) => values.map((v) => tickOf(v).label)).map(tickOf);
    // 目盛りを引く範囲の、比の大きい側の端（範囲の反転では値の小さい側）の比
    const frameTopRatio = tickFrame ? (invertRange ? 1 - calcRatio(tickFrame[0]) : calcRatio(tickFrame[1])) : 1;
    const frameBottomRatio = tickFrame ? (invertRange ? 1 - calcRatio(tickFrame[1]) : calcRatio(tickFrame[0])) : 0;

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
        // 積み上げ順を変えたとき（逆にする・系列を値順）：系列の並びの順に 0 の側から積む。始まりを先に決めておく
        const reversedStarts = new Map<number, number>();
        if (stackReordered) {
            let up = 0;
            let down = 0;
            for (const s of stackSequence) {
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
            } else if (stacked && stackReordered) {
                start = reversedStarts.get(s) ?? 0;
            } else if (stacked) {
                start = amount >= 0 ? positiveEnd : negativeEnd;
                if (amount >= 0) positiveEnd += amount;
                else negativeEnd += amount;
            }
            const end = stacked ? start + amount : val;
            // 系列 1 本の ID はカテゴリだけ（1.4 までと同じ。カテゴリごとの色の保存先がこの ID の selector）。系列があれば凡例の値かメジャーを足す
            const withSlot = (builder: powerbi.visuals.ISelectionIdBuilder) =>
                !seriesMode ? builder : legendSource ? withSeriesOf(builder, categorical!, slot.group) : builder.withMeasure(slot.column.source.queryName);
            const selectionId = withSlot(withCategories(host.createSelectionIdBuilder(), i)).createSelectionId();

            // ハイライトは該当しない行が null。数値で持ち、棒の高さは軸と同じ calcRatio で出す。
            // 軸の範囲・目盛りは全体の値だけで決め、ハイライトでは動かさない（標準と同じ）
            const rawHighlight = slot.column.highlights?.[i];
            const highlight = numberOf(rawHighlight);

            // 積み上げのハイライトは、その棒の始まりから該当分だけ伸ばす
            const highlightEnd =
                highlight === null ? null : !stacked ? highlight : start + (percent ? (absoluteSum > 0 ? highlight / absoluteSum : 0) : highlight);

            const style = seriesStyles[s];
            const insetStart = stacked && !skipped && start !== 0;
            const pointColor = seriesMode ? (i === otherRow ? null : dataColorAt(slot.colorColumn, i)) ?? series[s].color : categoryFill;
            // ラベルの長さを測るときに作った文字（測らなかった棒・行は無し）
            const made = madeTexts[s][i];
            const point: DataPoint = {
                category,
                rowIndex: i,
                seriesIndex: s,
                blank: skipped,
                ...(!skipped && isBlankAt(slot.column, i) ? { blankAsZero: true as const } : {}),
                value: val,
                startRatio: stacked ? calcRatio(start) : zeroRatio,
                valRatio: calcRatio(end),
                share,
                outermost: true,
                insetStart,
                rank: ribbonsOn ? ribbonRanks.get(s) ?? null : null,
                ...(made?.formatted ?? formatted(val, labelPrecisionOf(style.labelSettings), slotLabelSigns[s], slotGroupings[s])),
                ...(style.labelSettings ? { labelSettings: style.labelSettings } : {}),
                ...(() => {
                    const text = made?.valueField !== undefined ? made.valueField : valueFieldTextOf(slot, i);
                    return text === null ? {} : { dataLabelText: text, labelToneColor: "" };
                })(),
                titleText: made?.title ?? titleTextOf(slot, i, skipped),
                detailText: made?.detail ?? detailTextOf(slot, i, val, skipped),
                selectionId,
                // 「その他」の棒を押したら、まとめたカテゴリを全部選ぶ。ほかの棒と同じく系列ごと（ツールヒント・リボンもこの ID を渡す）
                ...(i === otherRow ? { selectionIds: others.mergedRows.map((r) => withSlot(withMergedRow(host.createSelectionIdBuilder(), r)).createSelectionId()) } : {}),
                // パレートのランクの帯はカテゴリで選ぶ。パレートのときだけ持たせる（ほかの計算の選択の見た目を変えない）
                ...(paretoOn ? { categorySelectionId: seriesMode ? withCategories(host.createSelectionIdBuilder(), i).createSelectionId() : selectionId } : {}),
                // 「その他」はパレートでなくても持たせる。前の版で保存した選択（カテゴリだけの ID）を戻したとき、どの系列の「その他」も選ばれて見える
                ...(i === otherRow ? { categorySelectionIds: otherSelectionIds } : {}),
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
            // 外側の端の棒：積む順でいちばん最後の棒
            const outerOf = (flags: boolean[]) => [...stackSequence].reverse().find((s) => flags[s]) ?? -1;
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
        if (!seriesMode && i !== otherRow && !noCategory) {
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
                // 「棒の色」で塗るカテゴリは、書式ペインのカラーが効かない
                ...(dataColorAt(slots[0]?.colorColumn, i) !== null ? { fieldColored: true } : {}),
            });
        }
    }

    // 系列の展開：すき間を入れても全体がプロットに収まるように置き直し、値の軸とその目盛線を消す（標準と同じ）
    if (explodeOn) {
        // 合計ラベルを出すときは、棒の外のラベルが切れないよう、上下（横棒は左右）の端に空ける（合計ラベルが長い端はその分）
        explodeStacks(categoryGroups, stackSegments, columnsSettings.stackedSpacing / 200, totalLabelsSettings.show ? explodeMarginsOf(labelRoom, axisLength) : undefined);
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
                // 「棒の色」が色を返す系列は、色の無い行も最初の色で塗り、凡例もその色なので、書式ペインのカラーが効かない
                ...(fieldColoredSeries[s] ? { fieldColored: true } : {}),
            });
            labelTargets.push({
                name: slot.name,
                selector,
                show: style.labelShow,
                color: style.explicitLabelColor ?? "",
                kind: "bar",
                values: dataLabelsSettingsOf(dl, slot.objects, true),
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
    // 表示単位・単位ラベル・タイトルのスタイルは Y 軸と同じ。載る線の書式がすべて率（% か ‰）なら、
    // 100% 積み上げの Y 軸と同じく表示単位によらず % で出し、単位の語を付けない。率と数の線が混ざれば数の軸（Y 軸と同じ判定）
    const lineRate = sameRateOf(lineDefs.map((def) => numberFormatOf(def.format)));
    // 載る線の書式の「正;負;0」の節の書き方がすべて同じなら、目盛りに当てる（Y 軸と同じ）
    const lineSections = sameSectionSignOf(lineDefs.map((def) => numberFormatOf(def.format)));
    // 第 2 Y 軸の目盛りと、そこに載せた定数線の桁区切り。載る線のどれかの書式に桁区切りがあれば付ける（Y 軸と同じ）。
    // 線の書式は、書式ペインで「カスタム書式」「パーセント」を選べばその書式（lineFormatOf）
    const lineGrouping = someGroupingOf(lineDefs.map((def) => numberFormatOf(def.format)));
    const isPercentLine = lineRate !== null;
    const lineScale = lineRate?.scale ?? 100;
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
    let lineRatio2Of = lineRatioOf;
    const calcRatio2 = (v: number): number => clampRatio(lineRatio2Of(v));
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
            const lower = userLower ?? Math.pow(10, Math.ceil(log10Snap(minOf(mags))) - 1);
            let upper = userUpper ?? (roundRange2 ? Math.pow(10, Math.ceil(log10Snap(maxOf(mags)))) : maxOf(mags));
            if (!(upper > lower)) upper = lower * 10;
            const logMin = Math.log10(lower);
            const logSpan = Math.log10(upper) - logMin;
            lineRatio2Of = (v: number) => {
                const mag = Math.abs(v);
                if (mag <= 0) return logSign2 > 0 ? 0 : 1;
                const r = logSpan > 0 ? (Math.log10(mag) - logMin) / logSpan : 0;
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
            // 丸めは Y 軸と同じ本数の目安（描く範囲の高さで決まる。目盛りも同じ本数を上限に引く）
            const round2 = roundRange2 && start2 === undefined && end2 === undefined;
            if (round2 && !valueAxisSettings.tickCount) roundsByLength = true;
            let [lo2, hi2] = round2 ? niceDomain(min2, max2, roundCount) : [min2, max2];
            if (aligning) [lo2, hi2] = alignZeroAt(lo2, hi2, invertRange ? 1 - zeroRatio : zeroRatio);
            const span2 = hi2 - lo2;
            range2 = [lo2, hi2];
            lineRatio2Of = (v: number) => (span2 > 0 ? (v - lo2) / span2 : 0);
            raw2 = scaleLinear().domain([lo2, hi2]).ticks(5);
            const domain2: [number, number] = [lo2, hi2];
            ticks2For = (count) => ticks2Of(ticksUpTo(domain2, count, (values) => ticks2Of(values).map((tick) => tick.label)));
        }
        ticks2 = ticks2Of(raw2);
    }
    function ticks2Of(raw: number[]): Tick[] {
        const step2 = raw.length > 1 ? Math.abs(raw[1] - raw[0]) : 1;
        const percentDigits = v2Precision !== "auto" ? Number(v2Precision) : step2 * lineScale < 1 ? 1 : 0;
        return raw.map((t) => {
            let label: string;
            if (isPercentLine) {
                label = lineSections
                    ? formatSectionTick(t, fixedNumber(Math.abs(t * lineScale), percentDigits, lineGrouping), lineSections, lineRate!.mark)
                    : `${fixedNumber(t * lineScale, percentDigits, lineGrouping)}${lineRate!.mark}`;
            } else if (log2Active) {
                // 対数は目盛りごとに桁の語を付ける（Y 軸と同じ）
                label = formatDynamicValue(t, v2Notation, v2Precision, v2ShowUnitOnAxis, lineSections ?? {}, lineGrouping);
            } else {
                // Y 軸と同じく、目盛りは数字だけ（「軸ラベルに単位を表示」で 1.5万 のように付ける）
                const word = v2ShowUnitOnAxis && unitDef2.unitWord ? unitDef2.unitWord : "";
                label = lineSections
                    ? formatSectionTick(t, formatValue(Math.abs(t), unitDef2.divisor, v2Precision, lineGrouping), lineSections, word)
                    : `${formatValue(t, unitDef2.divisor, v2Precision, lineGrouping)}${word}`;
            }
            return { value: t, label, ratio: calcRatio2(t) };
        });
    }
    if (!onSecondary && paretoOn) {
        // パレートの累積比の軸。0〜100% に固定する（第 2 Y 軸の範囲の指定は使わない）
        lineRatio2Of = (v: number) => v;
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
        const lineFraction = p2 === "auto" ? 1 : Number(p2);
        // 桁区切りは、欄のメジャーの線ならそのメジャーの書式（ツールヒントと同じ列）、カードの「値」の線なら載せた軸の書式
        const grouping = def.grouping ?? (secondary ? lineGrouping : measureGrouping);
        // 負と 0 は、載せた軸の目盛りと同じく書式の「正;負;0」の節で書く（節が無ければ今までどおり）
        const valueText = secondary
            ? isPercentLine
                ? lineSections
                    ? formatSectionTick(def.value, fixedNumber(Math.abs(def.value * lineScale), lineFraction, grouping), lineSections, lineRate!.mark)
                    : `${fixedNumber(def.value * lineScale, lineFraction, grouping)}${lineRate!.mark}`
                : formatSigned(def.value, unitDef2.divisor, p2, lineSections ?? {}, "", grouping)
            : percent
                ? `${(def.value * 100).toFixed(p1 === "auto" ? 0 : Number(p1))}%`
                : percentMeasure
                ? measurePercentText(def.value, p1, measureSections ?? undefined, grouping)
                : isLogScaleActive
                    ? formatDynamicValue(def.value, unitNotation, p1, true, measureSections ?? {}, grouping)
                    : formatSigned(def.value, unitDef.divisor, p1, measureSections ?? {}, "", grouping);
        const label = referenceLabelOf(vlLabelText, def.name, valueText);
        return [{ name: def.name, value: def.value, ratio: secondary ? calcRatio2(def.value) : calcRatio(def.value), secondary, label }];
    });
    const valueLine: ValueLineSettings = {
        ...referenceLineStyleOf(vl.parts, valueLineOn),
        shadeRegion: getDropdownValue(vl.shadeRegion.value, "below") === "above" ? "above" : "below",
    };

    const lineCard = settings.lines;
    const defaultLineStyle = getDropdownValue(lineCard.lineStyle.value, LINE_STYLES.solid);
    const defaultShape = lineCard.shapeValues();
    const areaCard = settings.areas;
    const areasOn = areaCard.show.value ?? false;
    // 網掛け領域の下の辺は値 0 の高さ（軸の範囲の外なら端。対数なら端）
    const baselineRatio = onSecondary ? calcRatio2(0) : calcRatio(0);
    // 折れ線のデータ ラベル。標準と同じく、カードの表示がオンなら線にも出す（線ごとの「このシリーズに表示」をオフにした線は出さない）。
    // 値の書き方は、率（% の書式）なら %、それ以外は線を描く軸（第 2 Y 軸か Y 軸）の表示単位と桁数。桁数は線ごと・「すべて」の順に選んだもの
    const lineLabelOwnShow = (objects: DataViewObjects | undefined) => ownLabelValue(objects, "show") !== false;
    const lineLabelShown = (def: LineDef) => dataLabelsSettings.show && lineLabelOwnShow(def.objects);
    const linePositionOf = (objects: DataViewObjects | undefined): LineLabelInfo["position"] => {
        const saved = ownLabelValue(objects, "linePosition");
        return saved === LINE_LABEL_POSITIONS.above ? "above" : saved === LINE_LABEL_POSITIONS.below ? "below" : "auto";
    };
    const lineLabelOf = (def: LineDef): LineLabelInfo => {
        const own = dataLabelsSettingsOf(dl, def.objects);
        const chosen = own.precision;
        const ownRate = rateFormatOf(numberFormatOf(def.format));
        const digits = chosen !== "auto" ? chosen : ownRate ? "auto" : onSecondary ? v2Precision : precision;
        const divisor = onSecondary ? unitDef2.divisor : unitDef.divisor;
        // 符号の書き方は、保存の無いペインの項目なら線のメジャーの書式の節に従う（棒のラベルと同じ）
        const sign = sectionLabelSignOf(def.format);
        // 桁区切りは線の書式（書式ペインのカスタム書式を含む）に従う
        const grouping = groupingOf(numberFormatOf(def.format));
        const textOf = (v: number) =>
            ownRate ? formatPercent(v, digits, ownRate.digits, sign, ownRate, grouping) : formatSigned(v, divisor, digits, sign, "", grouping);
        return {
            position: linePositionOf(def.objects),
            color: labelColorOf(def.objects) ?? dataLabelsSettings.color,
            settings: own,
            texts: categoryGroups.map((g) => {
                const v = def.values[g.rowIndex];
                return v === null ? null : textOf(v);
            }),
        };
    };
    const lines: LineSeriesInfo[] = lineDefs.map((def, lineIndex) => {
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
        // 凡例ごとに分けた線は、凡例の値（系列）とメジャーで選ぶ
        const measureKey = def.series?.measureKey ?? def.key;
        const measureBuilder = () =>
            def.series ? withSeriesOf(host.createSelectionIdBuilder(), categorical!, def.series.group) : host.createSelectionIdBuilder();
        const themeColor = lineThemeColors[lineIndex];
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
            marker: settings.markers.values(def.objects?.markers),
            selectionId: measureBuilder().withMeasure(measureKey).createSelectionId(),
            points: categoryGroups.map((g) => {
                const value = def.values[g.rowIndex];
                const highlight = def.highlights ? def.highlights[g.rowIndex] ?? null : undefined;
                // 線の比は丸めない（範囲の外の点は描画域で切り、マーカー・データ ラベルを出さない）
                const ratioOf = (v: number) => (onSecondary ? lineRatio2Of(v) : lineRatioOf(v));
                return {
                    rowIndex: g.rowIndex,
                    ...(cumulativeLineKeys.has(def.key) && resetBefore[positionOf.get(g.rowIndex)!] ? { breakBefore: true } : {}),
                    value,
                    ratio: value === null ? null : ratioOf(value),
                    ...(highlight !== undefined ? { highlight, highlightRatio: highlight === null ? null : ratioOf(highlight) } : {}),
                    selectionId: withCategories(measureBuilder(), g.rowIndex).withMeasure(measureKey).createSelectionId(),
                    // 「その他」の点は、まとめたカテゴリごとに、ほかの点と同じ作り（凡例の値＋カテゴリ＋メジャー）の ID を全部。上の ID はまとめた先頭のカテゴリのもの。
                    // 棒の「その他」の ID（カテゴリだけ）にすると、棒の「その他」を選んだだけで線が選ばれた扱いになり、凡例ごとの線では凡例の値が落ちる
                    ...(g.rowIndex === otherRow
                        ? { selectionIds: others.mergedRows.map((r) => withMergedRow(measureBuilder(), r).withMeasure(measureKey).createSelectionId()) }
                        : {}),
                };
            }),
            tooltip: { displayName: cumulativeLineKeys.has(def.key) ? `${def.name}（累計）` : def.name, format: def.format, values: def.values },
            ...(lineLabelShown(def) ? { labels: lineLabelOf(def) } : {}),
        };
    });
    const lineTargets: LineTarget[] = lines.map((line, j) => ({
        name: line.name,
        // 凡例ごとに分けた線の書式は、凡例の値（棒の系列と同じ保存先）に置く
        selector: lineDefs[j].series ? withSeriesOf(host.createSelectionIdBuilder(), categorical!, lineDefs[j].series!.group).createSelectionId().getSelector() : { metadata: lineDefs[j].key },
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
        marker: line.marker ?? settings.markers.values(),
    }));
    // データ ラベルの「設定の適用先」に線も並べる（棒の系列のあと。標準と同じ）。線のラベルも、標準と同じく既定でオン。
    // 「このシリーズに表示」は線に保存した値をそのまま出す（カードの表示がオフでも。カードの表示を混ぜると、
    // カードがオフのときにオンにしたトグルがすぐオフに戻る）。描くかどうかはカードの表示とあわせて決める（lineLabelShown）
    lineTargets.forEach((target, j) => {
        labelTargets.push({
            name: target.name,
            selector: target.selector,
            show: lineLabelOwnShow(lineDefs[j].objects),
            color: labelColorOf(lineDefs[j].objects) ?? "",
            kind: "line",
            values: dataLabelsSettingsOf(dl, lineDefs[j].objects, true),
            linePosition: linePositionOf(lineDefs[j].objects),
        });
    });

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
    const v2TickFont = fontFields(v2.font, "Segoe UI", 9);
    const v2TickColor = v2.labelColor.value?.value || "#605E5C";
    const valueAxis2Settings: ValueAxis2Settings = {
        show: onSecondary || paretoOn,
        valueShow: v2.valueShow.value ?? true,
        ticks: ticks2,
        ticksFor: ticks2For,
        ...v2TickFont,
        labelColor: v2TickColor,
        titleShow: v2.titleShow.value ?? true,
        // 自動のタイトルは折れ線の名前（標準と同じ。複数なら「および」でつなぐ）
        titleText: styledTitle(
            v2.titleText.value?.trim() || (paretoOn ? "累積比" : lines.map((l) => l.name).join(" および ")),
            v2TitleStyle,
            v2ComposedUnit
        ),
        ...titleFontFields(v2.titleFont, "DIN", 12),
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
        // 単位の文字サイズ・カラーは、保存していなければ第 2 Y 軸の値（目盛り）と同じ（1.36.11.44 までは 9pt・#605E5C）
        unitFontSize: v2.unitFontSize.value == null ? v2TickFont.fontSize : clampFontSize(v2.unitFontSize.value, v2TickFont.fontSize),
        unitColor: v2.unitColor.value?.value || v2TickColor,
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

    // 棒の項目を 1 つにまとめる：値の名前で、系列の色と混ざらない灰色の 1 項目（選ぶものは無い）
    const singleBarItem = seriesMode && getDropdownValue(lg.barItems.value, LEGEND_BAR_ITEMS.series) === LEGEND_BAR_ITEMS.single;
    const allLegendEntries: LegendItemInfo[] = [
        ...(singleBarItem
            ? [{
                kind: "bar" as const,
                index: 0,
                name: legendSource ? (slots[0].column.source?.displayName ?? slots[0].name) : slots.map((slot) => slot.name).join(" および "),
                color: SINGLE_BAR_LEGEND_GRAY,
                selectionId: null,
                barStyle: { transparency: 0, borderShow: false, borderColor: SINGLE_BAR_LEGEND_GRAY, borderWidth: 0 },
            }]
            : seriesMode
            ? // 凡例も系列の並べ替え（値順・逆順）に合わせる（集合の左からの並びと同じ。積み上げは「積み上げ順を逆にする」の前の順）
              seriesOrder.map((index) => ({ s: series[index], index })).map(({ s, index }) => ({
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
            : lines.length && !lineOnly
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
    // 「棒の項目」「折れ線の項目」を切ったら、その種類の項目を凡例から外す（凡例を出すかは、外す前の項目の数で決める）。
    // 「棒の項目」は棒と線の両方があるときだけ書式ペインに出すので、線が無い・折れ線だけのときは保存値を見ない
    const barItemsToggle = lines.length > 0 && !lineOnly;
    const legendEntries = allLegendEntries.filter((e) =>
        e.kind === "bar" ? !barItemsToggle || (lg.showBars.value ?? true) : e.kind === "line" ? (lg.showLines.value ?? true) : true
    );
    // 表示順を反転：凡例の項目だけを逆に並べる（棒の並び・積む順は変えない）
    if (lg.reverseOrder.value ?? false) legendEntries.reverse();
    legendInfo.show = (seriesMode || allLegendEntries.length > 1) && legendEntries.length > 0 && (lg.show.value ?? true);
    /**
     * j 本目の線のマーカーを、グラフか凡例に描くか（凡例のスタイルがマーカーなら、表示がオフでも凡例には描く）。
     * 累積比の線は線ごとの書式を持たず、「すべて」の書式で描く
     */
    const markerDrawnFor = (j: number): boolean => {
        const show = (lines[j]?.marker ?? settings.markers.values()).show;
        const inLegend = legendInfo.show && legendEntries.some((e) => e.kind === "line" && e.index === j) && legendMarkerDrawn(legendInfo.markerStyle, show);
        return show || inLegend;
    };
    // 棒の項目を外したら、凡例の欄の名前（製品など）の見出しは棒の項目の見出しなので出さない（文字で決めた見出しは残す）
    // 1 つにまとめたときも、系列の見出しは要らない
    if (seriesMode && (singleBarItem || !legendEntries.some((e) => e.kind === "bar")) && !lg.titleText.value?.trim()) legendInfo.title = "";

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
            const raw = columns.map((c) => c.values[g.rowIndex]).find((v) => !isBlankValue(v) && v !== false);
            if (raw === undefined) return [];
            // データ値：返した文字（数・日付は書式を当てた値、真偽はメジャーの名前）。名前：メジャーの名前。両方：名前と値
            const value = typeof raw === "string" ? raw : typeof raw === "boolean" ? name : valueTextOf(formatOfColumn(columns[0].source))(raw);
            const label = referenceLabelOf(labelText, name, value);
            return [{ index, label }];
        });
        return { name, marks };
    });
    const categoryLine: CategoryLineSettings = {
        ...referenceLineStyleOf(cl.parts, cl.show.value ?? true),
        position: (["before", "center", "after"].includes(getDropdownValue(cl.position.value, "before")) ? getDropdownValue(cl.position.value, "before") : "before") as CategoryLineSettings["position"],
        shadeRegion: getDropdownValue(cl.shadeRegion.value, "before") === "after" ? "after" : "before",
    };

    const mk = settings.markers;
    // 単位の文字サイズ・カラーは、保存していなければ Y 軸の値（目盛り）と同じ（1.36.11.44 までは 9pt・#605E5C）
    const unitFontSize = valAxis.unitFontSize.value == null ? valueAxisSettings.fontSize : clampFontSize(valAxis.unitFontSize.value, valueAxisSettings.fontSize);
    const unitColor = valAxis.unitColor.value?.value || valueAxisSettings.labelColor;

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
        frameTopRatio,
        frameBottomRatio,
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
        columnTargets: withDistinctNames(columnTargets),
        labelTargets: withDistinctNames(labelTargets),
        dataLabels: dataLabelsSettings,
        categoryAxis: categoryAxisSettings,
        valueAxis: valueAxisSettings,
        gridlines: gridlinesSettings,
        legend: legendInfo,
        legendEntries,
        totalLabels: totalLabelsSettings,
        valueAxisRoom,
        roundsByLength,
        ...(roundsByLength && roundCount !== undefined ? { roundCount } : {}),
        builtAxisLength: axisLength,
        continuousAvailable,
        ...(continuousOn
            ? {
                  continuousAxis: {
                      values: displayOrder.map((i) => axisValues[i]!),
                      isDate: axisIsDate,
                      integer: axisIsInteger || displayOrder.every((i) => Number.isInteger(axisValues[i])),
                      min: minOf(displayOrder.map((i) => axisValues[i]!)),
                      max: maxOf(displayOrder.map((i) => axisValues[i]!)),
                      formatNumber: numberFormatterOf(lowestLevel),
                  },
              }
            : {}),
        lines,
        lineTargets,
        lineBlanks: String(settings.lines.blanks.value?.value ?? LINE_BLANKS.break),
        ...(lineWarning ? { lineWarning } : {}),
        markers: mk.values(),
        areas: {
            matchLineColor: areaCard.matchLineColor.value ?? true,
            fill: areaCard.fill.value?.value || "#118DFF",
            transparency: clampPercent(areaCard.transparency.value ?? 60),
        },
        valueAxis2: valueAxis2Settings,
        ribbons: ribbonSettings,
        // 折れ線だけのときも、線のハイライトで選ばれたと分かるように
        hasHighlights: slots.some((slot) => !!slot.column.highlights) || lineDefs.some((def) => !!def.highlights),
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
            keepRows: keptRows,
            displayOrder,
            seriesOrder,
        },
        seriesPositions,
        cumulative: {
            available: cumulativeAvailable,
            enabled: cumulativeEnabled,
            toggle: cumulativeAvailable && (calc.cumulativeToggle.value ?? false),
            levels: cumulativeLevels,
            reset: resetDepth < 0 ? CUMULATIVE_RESET_NONE : cumulativeReset,
        },
        ...(lineOnly ? { lineOnly: true } : {}),
        pane: {
            chartType,
            linesDrawn: lineDefs.length > 0,
            anyLine: lines.length > 0,
            secondaryAxis: onSecondary,
            ...(autoSecondary !== undefined ? { autoSecondaryAxis: autoSecondary } : {}),
            legendMarkers: legendEntries.some(
                (e) => e.kind === "line" && legendMarkerDrawn(legendInfo.markerStyle, (lines[e.index]?.marker ?? mk.values()).show)
            ),
            markersUsed: lines.some((_, j) => markerDrawnFor(j)),
            markerTargetsUsed: lineTargets.map((_, j) => markerDrawnFor(j)),
            axisLevels: levelColumns.length - commonDepth,
            pareto: paretoOn,
            lineOnly,
            seriesMode,
            singleSeries: !seriesMode,
            rankOrder,
            continuous: continuousOn,
        },
        ...(noCategory ? { noCategory: true } : {}),
        isEmpty: dataPoints.length === 0 && !lineOnly,
        noData: false,
    };
}

/** カテゴリが上限（30,000 件）に達したときの警告 */
export const TRUNCATED_TITLE = "すべてのカテゴリを読み込めていません";
export const TRUNCATED_NOTICE =
    "カテゴリが 30,000 件に達したので、並べた順の先頭の 30,000 件で描いています。「その他」と積み上げの合計、パレートの累積比は、読み込めたカテゴリで計算しています。フィルターで絞ってください。";
/** 凡例の値が上限（2,000）に達したときの警告 */
export { SERIES_TRUNCATED_TITLE, SERIES_TRUNCATED_NOTICE };

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
    /** ほかのビジュアルで選ばれたときの該当分（ハイライト）。列にハイライトが無ければ持たない。率の線は分子と分母の該当分も持つ */
    highlights?: Array<number | null>;
    highlightRatio?: { numerators: Array<number | null>; denominators: Array<number | null> };
    /**
     * 折れ線だけで凡例があるとき、凡例の値ごとに分けた線の系列（凡例のグループ）。measureKey はメジャーの queryName、measureName はメジャーの表示名。
     * 色・書式は凡例の値ごと（棒の系列と同じ保存先）で、テーマの色も凡例の値で取る
     */
    series?: { group: DataViewValueColumnGroup; measureKey: string; measureName: string };
}

/** 0〜count-1 の行のうち、merged に無い行（「その他」にまとめなかった行）。merged は Set で引く（includes だと行の数の 2 乗） */
function keptRowsOf(count: number, merged: readonly number[]): number[] {
    const mergedSet = new Set(merged);
    const kept: number[] = [];
    for (let i = 0; i < count; i++) if (!mergedSet.has(i)) kept.push(i);
    return kept;
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
