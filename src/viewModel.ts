"use strict";

import powerbi from "powerbi-visuals-api";
import { scaleLinear } from "d3-scale";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;
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
    DETAIL_CONTENTS,
    LINE_STYLES,
    LINE_SHAPE_DEFAULTS,
    ORIENTATIONS,
    Orientation,
} from "./settings";
import {
    resolveUnit,
    formatValue,
    formatDynamicValue,
    resolveBadgeText,
    composeUnitText,
    UnitDefinition,
    UNIT_DEFINITIONS,
} from "./unitUtils";
import { TooltipSource, TooltipStack, TooltipColumn, tooltipColumnOf, formatTooltipValue, BLANK_TEXT } from "./tooltip";

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
    /** リボンでの順位（カテゴリの中で値の大きい順、1 が最大）。リボン以外は null */
    rank: number | null;
    formattedValue: string;
    dataLabelText: string;
    /** データラベルの詳細の行（全体に対する割合か、ラベルの詳細のフィールドの値）。無ければ空 */
    detailText: string;
    selectionId: ISelectionId;
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
    category: string;
    rowIndex: number;
    points: DataPoint[];
    /** 積み上げのときの合計。集合・100% 積み上げでは null */
    totals: StackTotals | null;
}

/** 折れ線の点 1 つ。並びは categoryGroups と同じ（並べ替えのあと） */
export interface LinePoint {
    rowIndex: number;
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
    /** 網掛け領域を出すか（カードの表示と、線ごとの「このシリーズに表示」） */
    areaShow: boolean;
    /** 網掛け領域の下の辺（値 0 の高さ。軸の範囲の外なら端）の軸の比率 */
    baselineRatio: number;
    /** 凡例のクリックで選ぶ ID（メジャー） */
    selectionId: ISelectionId;
    points: LinePoint[];
    /** ツールヒント用。values は DataView の行番号で引く */
    tooltip: TooltipColumn;
}

export interface ValueAxis2Settings {
    /** 第 2 Y 軸を描くか（折れ線を右の軸で描くとき） */
    show: boolean;
    valueShow: boolean;
    ticks: Tick[];
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
    kind: "bar" | "line";
    /** 棒なら series の添字、折れ線なら lines の添字 */
    index: number;
    name: string;
    color: string;
    /** クリックで選ぶ ID。系列 1 本の棒は null（選ばない） */
    selectionId: ISelectionId | null;
}

/** 折れ線のマーカー */
export interface MarkerSettings {
    show: boolean;
    /** 型（"circle" | "square" | "diamond" | "triangle" | "cross" | "shortDash" | "longDash" | "plus"） */
    shape: string;
    size: number;
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
    position: string;
    /** 空ならタイトルなし */
    title: string;
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
}

export interface ValueAxisSettings {
    start: string;
    end: string;
    logarithmic: boolean;
    /** 対数が指定されたが 0・正負混在のデータのため線形で描いている */
    logarithmicFallback: boolean;
    invertRange: boolean;
    roundRange: boolean;
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
    verticalShow: boolean;
    verticalColor: string;
    verticalTransparency: number;
    verticalStyle: string;
    verticalWidth: number;
    verticalScaleWithWidth: boolean;
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
    /** 0 = 上限なし */
    maxBarWidth: number;
    cornerRadius: number;
}

export interface ViewModel {
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
    unitInfo: UnitInfo;
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
    markers: MarkerSettings;
    areas: AreaSettings;
    valueAxis2: ValueAxis2Settings;
    ribbons: RibbonSettings;
    hasHighlights: boolean;
    /** ツールヒントの元データ。データが無いときは null */
    tooltip: TooltipSource | null;
    isEmpty: boolean;
}

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
    maxBarWidth: 0,
    cornerRadius: 0,
};

const clampPercent = (v: number): number => Math.max(0, Math.min(100, v));
const clampBorderWidth = (v: number): number => Math.max(1, Math.min(5, v));

/** 対数目盛りの本数の上限。標準の値軸と同程度（5〜8 本）に収める */
export const MAX_LOG_TICKS = 8;

/** Math.log10 の丸め誤差を吸収する（10 の冪ならちょうどの整数を返す） */
/**
 * 0 を含む範囲 [lo, hi] を広げて、0 が下から zero（0〜1）の割合の高さに来るようにする（第 2 Y 軸の「0 を配置する」、#91）。
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
    titleShow: true,
    titleText: "",
    titleFontFamily: "DIN",
    titleFontSize: 12,
    titleBold: false,
    titleItalic: false,
    titleUnderline: false,
    titleColor: "#252423",
    minCategoryWidth: 20,
};

const EMPTY_VALUE_AXIS: ValueAxisSettings = {
    start: "",
    end: "",
    logarithmic: false,
    logarithmicFallback: false,
    invertRange: false,
    roundRange: true,
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
    verticalShow: false,
    verticalColor: "#E1DFDD",
    verticalTransparency: 0,
    verticalStyle: "dotted",
    verticalWidth: 1,
    verticalScaleWithWidth: false,
};

const EMPTY_DATA_LABELS: DataLabelsSettings = {
    show: false,
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
};

const EMPTY_LEGEND: LegendInfo = {
    show: false,
    position: "topLeft",
    title: "",
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

const EMPTY: ViewModel = {
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
        { displayName: viewModel.tooltip.categoryName, value: group.category },
        { displayName: line.name, value: formatTooltipValue(line.tooltip.values[group.rowIndex], line.tooltip.format) },
    ];
}

/** 棒のツールヒントに足す、積み上げの合計（積み上げ）か割合（100% 積み上げ） */
export function tooltipStackOf(viewModel: ViewModel, d: DataPoint): TooltipStack {
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
    group: DataViewValueColumnGroup;
    /** 系列ごとの書式の保存先。凡例の系列はグループ、メジャーの系列は列のメタデータに入る */
    objects: Array<DataViewObjects | undefined>;
}

const isMeasure = (column: DataViewValueColumn): boolean => !!column.source?.roles?.measure;

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
    return measures.map((column) => ({
        key: column.source.queryName ?? column.source.displayName,
        name: column.source.displayName,
        column,
        tooltips,
        detail,
        group: groups[0],
        objects: [column.source.objects],
    }));
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

export function transform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel
): ViewModel {
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

    // 日付のカテゴリはモデルの書式で出す（String() だと JS の日付の文字列になる）。それ以外は 1.4 までと同じ
    const categoryFormatter = categories.source?.type?.dateTime
        ? valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(categories.source) || "yyyy/MM/dd" })
        : null;
    const categoryText = (i: number): string => {
        const raw = categories.values[i];
        return categoryFormatter && raw !== null && raw !== undefined ? categoryFormatter.format(raw) : String(raw);
    };

    const orientation: Orientation =
        String(settings.columns.orientation?.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal
            ? ORIENTATIONS.horizontal
            : ORIENTATIONS.vertical;
    const horizontal = orientation === ORIENTATIONS.horizontal;
    const ribbonSelected = String(settings.columns.chartType?.value?.value ?? "") === CHART_TYPES.ribbon;

    // 折れ線の値。凡例があると、系列ごとに同じメジャーの列が複製されて届くので queryName でまとめる。
    // 横棒とリボンでは折れ線を描かない（標準の複合は縦棒だけで、リボン グラフに折れ線は無い）
    const lineColumns = new Map<string, DataViewValueColumn[]>();
    for (const group of horizontal || ribbonSelected ? [] : groups) {
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
     * 系列ごとの値がどの行でも同じなら（凡例に左右されないメジャー）その値、違えば足した値を使う
     * （標準はカテゴリ単位で計算するが、カスタムビジュアルは凡例で分けたデータしか受け取れないため）
     */
    const lineDefs = [...lineColumns.entries()].map(([key, columns]) => {
        const numbersAt = (i: number) =>
            columns
                .map((c) => c.values[i])
                .filter((v) => v !== null && v !== undefined)
                .map((v) => (typeof v === "number" ? v : Number(v)))
                .filter((v) => Number.isFinite(v));
        const independent = Array.from({ length: rowCount }, (_, i) => numbersAt(i)).every(
            (vals) => vals.every((v) => v === vals[0])
        );
        const values: Array<number | null> = Array.from({ length: rowCount }, (_, i) => {
            const vals = numbersAt(i);
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

    const chartTypeValue = getDropdownValue(settings.columns.chartType?.value, CHART_TYPES.clustered);
    const chartType: ChartType =
        chartTypeValue === CHART_TYPES.stacked || chartTypeValue === CHART_TYPES.stacked100 || chartTypeValue === CHART_TYPES.ribbon
            ? chartTypeValue
            : CHART_TYPES.clustered;
    const stacked = chartType !== CHART_TYPES.clustered;
    const percent = chartType === CHART_TYPES.stacked100;
    // リボンは積み上げと同じく 1 本に積むが、カテゴリごとに値の大きい順に上から並べる（標準と同じ）
    const ribbon = chartType === CHART_TYPES.ribbon;

    // 複数系列の空白は棒を描かないので数えない（系列 1 本の空白は 1.4 までと同じく 0 として数える）
    const isSkipped = (column: DataViewValueColumn, i: number) => seriesMode && isBlankAt(column, i);

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
    if (lineDefs.length) {
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

    const unitDef = resolveUnit(unitTypeKey, maxAbs, unitNotation, precision);

    const roundRange = valAxis.roundRange.value ?? true;
    const invertRange = valAxis.invertRange.value ?? false;
    const logarithmic = valAxis.logarithmic.value ?? false;

    let userStart: number | undefined = undefined;
    let userEnd: number | undefined = undefined;
    if (valAxis.start.value != null && valAxis.start.value.trim() !== "") {
        const parsed = Number(valAxis.start.value.trim());
        if (!isNaN(parsed)) userStart = parsed;
    }
    if (valAxis.end.value != null && valAxis.end.value.trim() !== "") {
        const parsed = Number(valAxis.end.value.trim());
        if (!isNaN(parsed)) userEnd = parsed;
    }

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
        const headroom = totalLabelsOn ? (Math.max(maxValue, 0) - Math.min(minValue, 0)) * 0.1 : 0;
        const lowerBound = userStart !== undefined ? userStart : (minValue < 0 ? minValue - headroom : 0);
        const upperBound = userEnd !== undefined ? userEnd : (maxValue > 0 ? maxValue + headroom : 1);
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
    const badgeText = percent
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
    const ticks: Tick[] = rawTicks.map((t: number) => {
        let label: string;
        if (percent) {
            label = `${Math.round(t * 100)}%`;
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
    });

    // X軸設定の抽出
    const catAxis = settings.categoryAxis;
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
        titleText: catAxis.titleText.value?.trim() || (categories.source?.displayName ?? ""),
        titleFontFamily: catAxis.titleFont.fontFamily.value ?? "DIN",
        titleFontSize: Math.max(8, Math.min(32, catAxis.titleFont.fontSize.value ?? 12)),
        titleBold: catAxis.titleFont.bold?.value ?? false,
        titleItalic: catAxis.titleFont.italic?.value ?? false,
        titleUnderline: catAxis.titleFont.underline?.value ?? false,
        titleColor: catAxis.titleColor.value?.value || "#252423",
        minCategoryWidth: Math.max(0, Math.min(500, catAxis.minCategoryWidth.value ?? 20)),
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

    const valueAxisSettings: ValueAxisSettings = {
        start: valAxis.start.value ?? "",
        end: valAxis.end.value ?? "",
        logarithmic,
        logarithmicFallback,
        invertRange,
        roundRange,
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
        verticalShow: gl.verticalShow.value ?? false,
        verticalColor: gl.verticalColor.value?.value || "#E1DFDD",
        verticalTransparency: Math.max(0, Math.min(100, gl.verticalTransparency.value ?? 0)),
        verticalStyle: String(gl.verticalStyle.value?.value ?? "dotted"),
        verticalWidth: Math.max(1, Math.min(10, gl.verticalWidth.value ?? 1)),
        verticalScaleWithWidth: gl.verticalScaleWithWidth.value ?? false,
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
    };

    // 列（columns）設定の抽出
    const col = settings.columns;
    const columnsSettings: ColumnsSettings = {
        fill: col.fill.value?.value || "#118DFF",
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
        seriesSpacing: Math.max(0, Math.min(90, col.seriesSpacing.value ?? 0)),
        maxBarWidth: Math.max(0, col.maxBarWidth.value ?? 0),
        cornerRadius: Math.max(0, Math.min(30, col.cornerRadius.value ?? 0)),
    };

    // 凡例（複数系列のときだけ描く）。値が複数で凡例のフィールドが無いときは、標準と同じくタイトルを出さない
    const lg = settings.legend;
    const legendInfo: LegendInfo = {
        show: seriesMode && (lg.show.value ?? true),
        position: getDropdownValue(lg.position.value, "topLeft"),
        title: (lg.titleShow.value ?? true) ? (lg.titleText.value?.trim() || (legendSource?.displayName ?? "")) : "",
        fontFamily: lg.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, lg.font.fontSize.value ?? 10)),
        bold: lg.font.bold?.value ?? false,
        italic: lg.font.italic?.value ?? false,
        underline: lg.font.underline?.value ?? false,
        color: lg.labelColor.value?.value || "#605E5C",
    };

    // 系列。複数系列の色は、個別の指定が無ければ標準と同じくレポートのテーマの色を順に割り当てる。
    // 系列 1 本のときは 1.4 までと同じく「列」のカラー（カテゴリごとの指定があればそちら）
    const series: SeriesInfo[] = slots.map((slot) => {
        if (!seriesMode) return { name: slot.name, color: columnsSettings.fill, selectionId: null };
        // テーマの色は、個別の指定がある系列でも必ず取る。Desktop の colorPalette は呼んだ順に色を割り当てるので、
        // 取らないと後ろの系列の色が 1 つ前にずれる（標準では、ある系列の色を変えても他の系列の色は変わらない）
        const themeColor = host.colorPalette.getColor(slot.key).value;
        const color = firstOf(slot.objects, (o) => customColor(o, "fill")) ?? themeColor;
        const selectionId = legendSource
            ? host.createSelectionIdBuilder().withSeries(categorical!.values!, slot.group).createSelectionId()
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
    // 選んだときはその単位で割って語を付ける（#91）
    const tl = settings.totalLabels;
    const totalPrecisionValue = String(tl.precision.value?.value ?? "auto");
    const totalPrecision = totalPrecisionValue !== "auto" ? totalPrecisionValue : precision;
    const totalUnitKey = getDropdownValue(tl.unitType.value, "auto");
    const totalUnitDef = totalUnitKey === "auto" ? null : resolveUnit(totalUnitKey, maxAbs, unitNotation, totalPrecision);
    const totalText = (v: number): string =>
        totalUnitDef
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

    // リボンの帯（グラフの種類がリボンで縦棒のとき。標準のリボン グラフは縦棒だけ）
    const rb = settings.ribbons;
    const ribbonSettings: RibbonSettings = {
        show: ribbon && !horizontal,
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
    const formatted = (val: number) =>
        // 100% 積み上げの軸は割合なので、データラベルは値ごとに単位を付ける（1,250億 など）
        isLogScaleActive || percent
            ? {
                formattedValue: formatDynamicValue(val, unitNotation, precision, true),
                dataLabelText: formatDynamicValue(val, unitNotation, labelPrecision, true),
            }
            : {
                formattedValue: formatValue(val, unitDef.divisor, precision),
                dataLabelText: formatValue(val, unitDef.divisor, labelPrecision),
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
    const detailFormatters = new Map<DataViewValueColumn, ReturnType<typeof valueFormatter.create>>();
    const detailTextOf = (slot: SeriesSlot, i: number, val: number, skipped: boolean): string => {
        if (skipped) return "";
        if (detailContent !== DETAIL_CONTENTS.custom) {
            const denominator = seriesMode || percent ? absoluteSums[i] : grandAbsolute;
            return percentText(denominator > 0 ? val / denominator : 0);
        }
        const column = slot.detail;
        const raw = column?.values[i];
        if (!column || raw === null || raw === undefined) return "";
        if (typeof raw !== "number") return String(raw);
        if (detailUnitKey !== "auto") {
            const unit = resolveUnit(detailUnitKey, Math.abs(raw), unitNotation, detailPrecision);
            return `${formatValue(raw, unit.divisor, detailPrecision)}${unit.unitWord}`;
        }
        // 自動はフィールドの書式のまま（小数点以下の桁数を選んだら、その桁で出す）
        const format = valueFormatter.getFormatStringByColumn(column.source) ?? "";
        if (detailPrecision !== "auto") return /%/.test(format) ? percentText(raw) : formatValue(raw, 1, detailPrecision);
        let formatter = detailFormatters.get(column);
        if (!formatter) {
            formatter = valueFormatter.create({ format });
            detailFormatters.set(column, formatter);
        }
        return formatter.format(raw);
    };

    // データポイントとカテゴリ別ターゲットの生成
    const categoryGroups: CategoryGroup[] = [];
    const columnTargets: ColumnTarget[] = [];

    for (let i = 0; i < rowCount; i++) {
        const category = categoryText(i);
        const targetObjects = categories.objects?.[i];

        // 系列 1 本のときの、カテゴリごとの見た目（1.4 までと同じ）
        const categoryFill = customColor(targetObjects, "fill") ?? columnsSettings.fill;
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

        // リボン：カテゴリの中で値の小さい順に 0 から積む（いちばん大きい系列が上に来る）。
        // 負の値は 0 に近い順に下へ積む。順位は値の大きい順（1 が最大）
        const ribbonStarts = new Map<number, number>();
        const ribbonRanks = new Map<number, number>();
        if (ribbon) {
            const present = slots
                .map((slot, s) => ({ s, v: isSkipped(slot.column, i) ? null : numberAt(slot.column, i) }))
                .filter((e): e is { s: number; v: number } => e.v !== null);
            let up = 0;
            let down = 0;
            present.filter((e) => e.v >= 0).sort((a, b) => a.v - b.v).forEach((e) => {
                ribbonStarts.set(e.s, up);
                up += e.v;
            });
            present.filter((e) => e.v < 0).sort((a, b) => b.v - a.v).forEach((e) => {
                ribbonStarts.set(e.s, down);
                down += e.v;
            });
            [...present].sort((a, b) => b.v - a.v).forEach((e, k) => ribbonRanks.set(e.s, k + 1));
        }

        const points: DataPoint[] = slots.map((slot, s) => {
            const val = numberAt(slot.column, i);
            const skipped = isSkipped(slot.column, i);
            const share = percent ? (absoluteSum > 0 ? val / absoluteSum : 0) : null;
            // 積む量（100% なら割合）と、この棒の始まり・終わり
            const amount = skipped ? 0 : percent ? share! : val;
            let start = 0;
            if (ribbon) {
                start = ribbonStarts.get(s) ?? 0;
            } else if (stacked) {
                start = amount >= 0 ? positiveEnd : negativeEnd;
                if (amount >= 0) positiveEnd += amount;
                else negativeEnd += amount;
            }
            const end = stacked ? start + amount : val;
            // 系列 1 本の ID はカテゴリだけ（1.4 までと同じ。カテゴリごとの色の保存先がこの ID の selector）
            const builder = host.createSelectionIdBuilder().withCategory(categories, i);
            const selectionId = !seriesMode
                ? builder.createSelectionId()
                : legendSource
                    ? builder.withSeries(categorical!.values!, slot.group).createSelectionId()
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
            return {
                category,
                rowIndex: i,
                seriesIndex: s,
                blank: skipped,
                value: val,
                startRatio: stacked ? calcRatio(start) : zeroRatio,
                valRatio: calcRatio(end),
                share,
                outermost: true,
                rank: ribbon ? ribbonRanks.get(s) ?? null : null,
                ...formatted(val),
                detailText: detailTextOf(slot, i, val, skipped),
                selectionId,
                highlight,
                highlightRatio: highlightEnd === null ? null : calcRatio(highlightEnd),
                color: seriesMode ? series[s].color : categoryFill,
                transparency: seriesMode ? style.transparency : categoryTransparency,
                borderShow: seriesMode ? style.borderShow : categoryShowBorder,
                borderColor: seriesMode ? style.borderColor : categoryBorderFill,
                borderTransparency: seriesMode ? style.borderTransparency : categoryBorderTransparency,
                borderWidth: seriesMode ? style.borderWidth : categoryBorderWidth,
                labelShow: style.labelShow,
                labelColor: style.explicitLabelColor ?? dataLabelsSettings.color,
            };
        });

        // 積み上げでは、角丸は正と負それぞれいちばん外側の棒だけに付ける（リボンは値の大きい順に積むので、最大と最小）
        if (stacked) {
            const pick = (want: (d: DataPoint) => boolean, better: (a: DataPoint, b: DataPoint) => boolean) =>
                points.reduce((best, d, k) => (want(d) && (best < 0 || better(d, points[best])) ? k : best), -1);
            const lastPositive = ribbon
                ? pick((d) => !d.blank && d.value > 0, (a, b) => a.value >= b.value)
                : points.map((d) => !d.blank && d.value > 0).lastIndexOf(true);
            const lastNegative = ribbon
                ? pick((d) => !d.blank && d.value < 0, (a, b) => a.value <= b.value)
                : points.map((d) => !d.blank && d.value < 0).lastIndexOf(true);
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

        categoryGroups.push({ category, rowIndex: i, points, totals });

        if (!seriesMode) {
            columnTargets.push({
                name: category,
                selector: points[0].selectionId.getSelector(),
                color: categoryFill,
                transparency: categoryTransparency,
                borderShow: categoryShowBorder,
                borderMatchColumn: categoryBorderMatchColumn,
                borderColor: categoryBorderFill,
                borderTransparency: categoryBorderTransparency,
                borderWidth: categoryBorderWidth,
            });
        }
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
                color: series[s].color,
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

    // レイアウトの並び替え。値順は、複数系列ではカテゴリの合計（空白は 0）の大きい順
    if (columnsSettings.sortByValue) {
        const totalOf = (g: CategoryGroup) => g.points.reduce((sum, d) => sum + (d.blank ? 0 : d.value), 0);
        categoryGroups.sort((a, b) => totalOf(b) - totalOf(a));
    }
    if (columnsSettings.reverseOrder) {
        categoryGroups.reverse();
    }
    const dataPoints = categoryGroups.flatMap((g) => g.points);

    // --- 折れ線と第 2 Y 軸 ---------------------------------------------------
    const v2 = settings.valueAxis2;
    const parseOptional = (raw: string | undefined | null): number | undefined => {
        if (raw == null || raw.trim() === "") return undefined;
        const n = Number(raw.trim());
        return isNaN(n) ? undefined : n;
    };
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
    // 第 2 Y 軸の範囲：対数目盛り・範囲を丸める・0 を配置する（標準の第 2 Y 軸と同じ項目、#91）。
    // 対数は、折れ線の値がすべて正かすべて負で、0 を配置しないときだけ効かせる（標準も 0 を配置すると対数は押せない）
    const alignZeros2 = v2.alignZeros.value ?? false;
    const logSign2 = lineMin > 0 ? 1 : lineMax < 0 ? -1 : 0;
    const log2Active = onSecondary && (v2.logarithmic.value ?? false) && logSign2 !== 0 && !alignZeros2;
    // 対数では、全体一律の表示単位の語を単位ラベルとタイトルに付けない（Y 軸と同じ）
    const unitWord2 = log2Active ? "" : unitDef2.unitWord;
    let calcRatio2 = calcRatio;
    let ticks2: Tick[] = [];
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
            calcRatio2 = (v: number) => (span2 > 0 ? Math.max(0, Math.min(1, (v - lo2) / span2)) : 0);
            raw2 = scaleLinear().domain([lo2, hi2]).ticks(5);
        }
        const step2 = raw2.length > 1 ? Math.abs(raw2[1] - raw2[0]) : 1;
        const percentDigits = v2Precision !== "auto" ? Number(v2Precision) : step2 * 100 < 1 ? 1 : 0;
        ticks2 = raw2.map((t) => {
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
        const measureBuilder = () => host.createSelectionIdBuilder();
        // 棒の系列と同じく、テーマの色は色の指定があっても取る（後ろの線の色がずれないように）
        const themeColor = host.colorPalette.getColor(def.key).value;
        return {
            name: def.name,
            color: ownFill ? String(ownFill) : themeColor,
            width: Math.max(1, Math.min(10, ownWidth ?? lineCard.width.value ?? 3)),
            lineStyle: ownStyle !== undefined && ownStyle !== null ? String(ownStyle) : defaultLineStyle,
            lineJoin: ownText("lineJoin", defaultShape.lineJoin),
            interpolation: ownText("interpolation", defaultShape.interpolation),
            smoothing: ownText("smoothing", defaultShape.smoothing),
            tension: clampPercent(typeof own?.tension === "number" ? own.tension : defaultShape.tension),
            stepPosition: ownText("stepPosition", defaultShape.stepPosition),
            areaShow: areasOn && (typeof ownAreaShow === "boolean" ? ownAreaShow : true),
            baselineRatio,
            selectionId: measureBuilder().withMeasure(def.key).createSelectionId(),
            points: categoryGroups.map((g) => {
                const value = def.values[g.rowIndex];
                return {
                    rowIndex: g.rowIndex,
                    value,
                    ratio: value === null ? null : onSecondary ? calcRatio2(value) : calcRatio(value),
                    selectionId: measureBuilder().withCategory(categories, g.rowIndex).withMeasure(def.key).createSelectionId(),
                };
            }),
            tooltip: { displayName: def.name, format: def.format, values: def.values },
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
        areaShow: typeof lineDefs[j].objects?.areas?.show === "boolean" ? Boolean(lineDefs[j].objects?.areas?.show) : true,
    }));

    const valueAxis2Settings: ValueAxis2Settings = {
        show: onSecondary,
        valueShow: v2.valueShow.value ?? true,
        ticks: ticks2,
        fontFamily: v2.font.fontFamily.value ?? "Segoe UI",
        fontSize: Math.max(8, Math.min(32, v2.font.fontSize.value ?? 9)),
        bold: v2.font.bold?.value ?? false,
        italic: v2.font.italic?.value ?? false,
        underline: v2.font.underline?.value ?? false,
        labelColor: v2.labelColor.value?.value || "#605E5C",
        titleShow: v2.titleShow.value ?? true,
        // 自動のタイトルは折れ線の名前（標準と同じ。複数なら「および」でつなぐ）
        titleText: styledTitle(
            v2.titleText.value?.trim() || lines.map((l) => l.name).join(" および "),
            getDropdownValue(v2.titleStyle.value, "showTitleOnly"),
            composeUnitText(unitWord2, v2UnitText, v2IncludeDisplayUnit)
        ),
        titleFontFamily: v2.titleFont.fontFamily.value ?? "DIN",
        titleFontSize: Math.max(8, Math.min(32, v2.titleFont.fontSize.value ?? 12)),
        titleBold: v2.titleFont.bold?.value ?? false,
        titleItalic: v2.titleFont.italic?.value ?? false,
        titleUnderline: v2.titleFont.underline?.value ?? false,
        titleColor: v2.titleColor.value?.value || "#252423",
        badgeText: onSecondary
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
    const legendEntries: LegendItemInfo[] = [
        ...(seriesMode
            ? series.map((s, index) => ({ kind: "bar" as const, index, name: s.name, color: s.color, selectionId: s.selectionId }))
            : lines.length
                ? [{ kind: "bar" as const, index: 0, name: slots[0].name, color: columnsSettings.fill, selectionId: null }]
                : []),
        ...lines.map((l, index) => ({ kind: "line" as const, index, name: l.name, color: l.color, selectionId: l.selectionId })),
    ];
    legendInfo.show = (seriesMode || legendEntries.length > 1) && (lg.show.value ?? true);

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
        unitInfo: {
            unitDef: isLogScaleActive ? { ...unitDef, unitWord: "" } : unitDef,
            badgeText,
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
        markers: {
            show: mk.show.value ?? false,
            shape: getDropdownValue(mk.shape.value, "circle"),
            size: Math.max(1, Math.min(20, mk.size.value ?? 5)),
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
            measure: tooltipColumnOf(slots[0].column),
            extras: slots[0].tooltips.map(tooltipColumnOf),
            ...(seriesMode
                ? {
                    series: slots.map((slot) => ({
                        legendName: legendSource?.displayName ?? null,
                        seriesName: slot.name,
                        measure: tooltipColumnOf(slot.column),
                        extras: slot.tooltips.map(tooltipColumnOf),
                    })),
                }
                : {}),
        },
        isEmpty: dataPoints.length === 0,
    };
}
