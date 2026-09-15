"use strict";

import powerbi from "powerbi-visuals-api";
import { scaleLinear } from "d3-scale";
import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import { VisualFormattingSettingsModel, ColumnTarget } from "./settings";
import {
    resolveUnit,
    formatValue,
    formatDynamicValue,
    resolveBadgeText,
    composeUnitText,
    UnitDefinition,
    UNIT_DEFINITIONS,
} from "./unitUtils";
import { TooltipSource, tooltipColumnOf } from "./tooltip";

export interface DataPoint {
    category: string;
    /** DataView の行番号。並べ替え後もツールヒント等で元の行を引くのに使う */
    rowIndex: number;
    value: number;
    valRatio: number;
    formattedValue: string;
    dataLabelText: string;
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
    verticalShow: boolean;
    verticalColor: string;
    verticalTransparency: number;
    verticalStyle: string;
    verticalWidth: number;
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
    /** 0 = 上限なし */
    maxBarWidth: number;
    cornerRadius: number;
}

export interface ViewModel {
    dataPoints: DataPoint[];
    maxValue: number;
    minValue: number;
    niceMin: number;
    niceMax: number;
    zeroRatio: number;
    ticks: Tick[];
    unitInfo: UnitInfo;
    columns: ColumnsSettings;
    columnTargets: ColumnTarget[];
    dataLabels: DataLabelsSettings;
    categoryAxis: CategoryAxisSettings;
    valueAxis: ValueAxisSettings;
    gridlines: GridlinesSettings;
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
    maxBarWidth: 0,
    cornerRadius: 0,
};

const clampPercent = (v: number): number => Math.max(0, Math.min(100, v));
const clampBorderWidth = (v: number): number => Math.max(1, Math.min(5, v));

/** 対数目盛りの本数の上限。標準の値軸と同程度（5〜8 本）に収める */
export const MAX_LOG_TICKS = 8;

/** Math.log10 の丸め誤差を吸収する（10 の冪ならちょうどの整数を返す） */
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
    verticalShow: false,
    verticalColor: "#E1DFDD",
    verticalTransparency: 0,
    verticalStyle: "dotted",
    verticalWidth: 1,
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
};

const EMPTY: ViewModel = {
    dataPoints: [],
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
    dataLabels: EMPTY_DATA_LABELS,
    categoryAxis: EMPTY_CATEGORY_AXIS,
    valueAxis: EMPTY_VALUE_AXIS,
    gridlines: EMPTY_GRIDLINES,
    hasHighlights: false,
    tooltip: null,
    isEmpty: true,
};

export function transform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel
): ViewModel {
    const categorical: DataViewCategorical | undefined = dataView?.categorical;
    const categories = categorical?.categories?.[0];
    // 「ツールヒント」の列も values に並ぶ。値の列は位置ではなくロールで探す
    // （値を空にしてツールヒントだけ入れたときに、ツールヒントの列を棒として描かない）
    const valueColumns: powerbi.DataViewValueColumn[] = categorical?.values ?? [];
    const values = valueColumns.find((column) => column.source?.roles?.measure);
    const tooltipColumns = valueColumns.filter((column) => column !== values && column.source?.roles?.tooltips);

    if (!categories || !values || !categories.values.length) {
        return EMPTY;
    }

    const highlights = values.highlights;
    const rawValues: number[] = [];
    let maxAbs = 0;
    let maxValue = -Infinity;
    let minValue = Infinity;

    for (let i = 0; i < categories.values.length; i++) {
        const raw = values.values[i];
        const val = typeof raw === "number" ? raw : Number(raw) || 0;
        rawValues.push(val);
        if (val > maxValue) maxValue = val;
        if (val < minValue) minValue = val;
        const abs = Math.abs(val);
        if (abs > maxAbs) maxAbs = abs;
    }

    if (!isFinite(maxValue)) maxValue = 0;
    if (!isFinite(minValue)) minValue = 0;

    const getDropdownValue = (sliceVal: any, fallback: string): string => {
        if (sliceVal == null) return fallback;
        if (typeof sliceVal === "object" && sliceVal !== null && "value" in sliceVal) {
            return String(sliceVal.value);
        }
        return String(sliceVal);
    };

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
    const isLogScaleActive = logarithmic && logSign !== 0;
    const logarithmicFallback = logarithmic && !isLogScaleActive;
    // 対数軸の範囲（大きさ）。負のデータでは 0 に近い側 (-logLower) が上端になる
    let logLower = 1;
    let logUpper = 10;

    if (isLogScaleActive) {
        let magMin = Infinity;
        let magMax = 0;
        for (const val of rawValues) {
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
    } else {
        const lowerBound = userStart !== undefined ? userStart : (minValue < 0 ? minValue : 0);
        const upperBound = userEnd !== undefined ? userEnd : (maxValue > 0 ? maxValue : 1);
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
    const badgeText = resolveBadgeText({
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
        if (isLogScaleActive) {
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

    // ユーザー設定の単位ラベル（例: "億円", "円", "bn"）を反映したタイトル生成
    const composedUnit = composeUnitText(effectiveUnitWord, unitText, unitIncludeDisplayUnit);
    const rawValueTitle = valAxis.titleText.value?.trim() || (values.source?.displayName ?? "");
    const valueTitleStyle = getDropdownValue(valAxis.titleStyle.value, "showTitleOnly");
    let formattedValueTitle = rawValueTitle;
    if (valueTitleStyle === "showUnitOnly") {
        formattedValueTitle = composedUnit ? `(${composedUnit})` : rawValueTitle;
    } else if (valueTitleStyle === "showBoth") {
        formattedValueTitle = composedUnit ? `${rawValueTitle} (${composedUnit})` : rawValueTitle;
    }

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
        verticalShow: gl.verticalShow.value ?? false,
        verticalColor: gl.verticalColor.value?.value || "#E1DFDD",
        verticalTransparency: Math.max(0, Math.min(100, gl.verticalTransparency.value ?? 0)),
        verticalStyle: String(gl.verticalStyle.value?.value ?? "dotted"),
        verticalWidth: Math.max(1, Math.min(10, gl.verticalWidth.value ?? 1)),
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
        // 範囲: 間隔 0〜50%、角丸 0〜30px、罫線 1〜5px。範囲は描画側でクランプする
        categorySpacing: Math.max(0, Math.min(50, col.categorySpacing.value ?? 20)),
        maxBarWidth: Math.max(0, col.maxBarWidth.value ?? 0),
        cornerRadius: Math.max(0, Math.min(30, col.cornerRadius.value ?? 0)),
    };

    /** 個別に指定された色を読む。無ければ null */
    const customColor = (
        objects: powerbi.DataViewObjects | undefined,
        property: "fill" | "borderFill"
    ): string | null => {
        const fill = objects?.columns?.[property] as powerbi.Fill | undefined;
        return fill?.solid?.color ? String(fill.solid.color) : null;
    };

    /** 個別に指定された数値を読む。無ければ null */
    const customNumber = (
        objects: powerbi.DataViewObjects | undefined,
        property: "transparency" | "borderTransparency" | "borderWidth"
    ): number | null => {
        const raw = objects?.columns?.[property];
        return typeof raw === "number" ? raw : null;
    };

    /** 個別に指定された真偽値を読む。無ければ null */
    const customFlag = (
        objects: powerbi.DataViewObjects | undefined,
        property: "showBorder" | "borderMatchColumn"
    ): boolean | null => {
        const raw = objects?.columns?.[property];
        return typeof raw === "boolean" ? raw : null;
    };

    // データポイントとカテゴリ別ターゲットの生成
    const dataPoints: DataPoint[] = [];
    const columnTargets: ColumnTarget[] = [];

    for (let i = 0; i < categories.values.length; i++) {
        const val = rawValues[i];
        const category = String(categories.values[i]);
        const targetObjects = categories.objects?.[i];

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

        const selectionId = host.createSelectionIdBuilder().withCategory(categories, i).createSelectionId();
        const labelPrecision = dataLabelsSettings.precision !== "auto" ? dataLabelsSettings.precision : precision;
        let formattedValue: string;
        let dataLabelText: string;
        if (isLogScaleActive) {
            formattedValue = formatDynamicValue(val, unitNotation, precision, true);
            dataLabelText = formatDynamicValue(val, unitNotation, labelPrecision, true);
        } else {
            formattedValue = formatValue(val, unitDef.divisor, precision);
            dataLabelText = formatValue(val, unitDef.divisor, labelPrecision);
        }

        // ハイライトは該当しない行が null。数値で持ち、棒の高さは軸と同じ calcRatio で出す。
        // 軸の範囲・目盛りは全体の値 (rawValues) だけで決め、ハイライトでは動かさない（標準と同じ）
        const rawHighlight = highlights?.[i];
        const highlightNumber = rawHighlight === null || rawHighlight === undefined ? NaN : Number(rawHighlight);
        const highlight = isFinite(highlightNumber) ? highlightNumber : null;

        dataPoints.push({
            category,
            rowIndex: i,
            value: val,
            valRatio: calcRatio(val),
            formattedValue,
            dataLabelText,
            selectionId,
            highlight,
            highlightRatio: highlight === null ? null : calcRatio(highlight),
            color: categoryFill,
            transparency: categoryTransparency,
            borderShow: categoryShowBorder,
            borderColor: categoryBorderFill,
            borderTransparency: categoryBorderTransparency,
            borderWidth: categoryBorderWidth,
        });

        columnTargets.push({
            name: category,
            selector: selectionId.getSelector(),
            color: categoryFill,
            transparency: categoryTransparency,
            borderShow: categoryShowBorder,
            borderMatchColumn: categoryBorderMatchColumn,
            borderColor: categoryBorderFill,
            borderTransparency: categoryBorderTransparency,
            borderWidth: categoryBorderWidth,
        });
    }

    // レイアウトの並び替え
    if (columnsSettings.sortByValue) {
        dataPoints.sort((a, b) => b.value - a.value);
    }
    if (columnsSettings.reverseOrder) {
        dataPoints.reverse();
    }

    const unitFontSize = Math.max(6, Math.min(32, valAxis.unitFontSize.value ?? 9));
    const unitColor = valAxis.unitColor.value?.value || "#605E5C";

    return {
        dataPoints,
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
        dataLabels: dataLabelsSettings,
        categoryAxis: categoryAxisSettings,
        valueAxis: valueAxisSettings,
        gridlines: gridlinesSettings,
        hasHighlights: !!highlights,
        tooltip: {
            categoryName: categories.source?.displayName ?? "",
            measure: tooltipColumnOf(values),
            extras: tooltipColumns.map(tooltipColumnOf),
        },
        isEmpty: dataPoints.length === 0,
    };
}
