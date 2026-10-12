import powerbi from "powerbi-visuals-api";

/** 線の種類（グリッド線・折れ線・つなぎの線・目標・定数線で共通） */
export const LINE_STYLES = {
    solid: "solid",
    dashed: "dashed",
    dotted: "dotted",
} as const;

/** 線の種類の選択肢。並びは実線・破線・点線（既定は並びの位置ではなく値で選ぶ） */
export const LINE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: LINE_STYLES.solid, displayName: "実線" },
    { value: LINE_STYLES.dashed, displayName: "破線" },
    { value: LINE_STYLES.dotted, displayName: "点線" },
];

/** 線の端の形（標準の「ダッシュ キャップ」）。保存値は標準と同じ none・round・square */
export const DASH_CAPS = { none: "none", round: "round", square: "square" } as const;

export const DASH_CAP_ITEMS: powerbi.IEnumMember[] = [
    { value: DASH_CAPS.none, displayName: "フラット" },
    { value: DASH_CAPS.round, displayName: "丸" },
    { value: DASH_CAPS.square, displayName: "四角" },
];

/** 線のスタイルの「カスタム」。選ぶと、模様（ダッシュ配列）と端の形を自分で決める */
export const CUSTOM_LINE_STYLE = "custom";
export const CUSTOM_LINE_STYLE_ITEM: powerbi.IEnumMember = { value: CUSTOM_LINE_STYLE, displayName: "カスタム" };

/** カスタムの模様。dashArray は線と隙間の長さ（px）を空白で区切って並べた文字（例 "5 5 0 5"） */
export interface CustomDash {
    dashArray: string;
    dashCap: string;
}

/**
 * ダッシュ配列の文字を stroke-dasharray にする。空白・カンマで区切った 0 以上の数だけを使い、
 * 数が 1 つも無い・負の数や数でないものが混ざる・すべて 0 なら null（実線で描く）。
 * 「幅で拡大縮小」がオンなら、線の幅に比例させる
 */
export function customDashOf(dashArray: string, width: number, scaleWithWidth: boolean): string | null {
    const parts = dashArray.trim().split(/[\s,]+/).filter((p) => p !== "");
    if (!parts.length) return null;
    const values = parts.map(Number);
    if (values.some((v) => !Number.isFinite(v) || v < 0) || values.every((v) => v === 0)) return null;
    const w = scaleWithWidth ? Math.max(1, width) : 1;
    return values.map((v) => String(v * w)).join(" ");
}

/**
 * グリッド線の線種の模様（stroke-dasharray）。標準に合わせ、点線は細かい点（丸い端と組み合わせる）、破線は 4px 刻み。
 * 「幅で拡大縮小」がオンなら、点線・破線の模様を線の幅に比例させる（標準と同じく、細い線では模様が細かくなる）。
 * カスタムは custom の模様。実線は null
 */
export function gridDashOf(style: string, width: number, scaleWithWidth: boolean, custom?: CustomDash): string | null {
    const w = Math.max(1, width);
    if (style === "dotted") return scaleWithWidth ? `${w} ${2 * w}` : "1 3";
    if (style === "dashed") return scaleWithWidth ? `${3 * w} ${3 * w}` : "4 4";
    if (style === CUSTOM_LINE_STYLE && custom) return customDashOf(custom.dashArray, width, scaleWithWidth);
    return null;
}

/** 線の端の形（stroke-linecap）。点線は丸い端で細かい点にする。カスタムはダッシュ キャップ（フラットは butt） */
export function lineCapOf(style: string, custom?: CustomDash): "round" | "square" | "butt" {
    if (style === "dotted") return "round";
    if (style === CUSTOM_LINE_STYLE && custom) return custom.dashCap === DASH_CAPS.round ? "round" : custom.dashCap === DASH_CAPS.square ? "square" : "butt";
    return "butt";
}

/**
 * グリッド線の線種。標準に合わせ、点線は細かい点（長さ 1 の線に丸い端）、破線は 4px 刻み。
 * crispEdges は掛けない（標準と同じ）。画面の拡大率が整数でない（150% など）と、1px の線が
 * 物理画素 1.5 個ぶんになり、crispEdges では線ごとに 1 個か 2 個に丸められて太さがそろわない（1.29.4.0 まで）
 */
export function gridLineStroke(style: string, width = 1, scaleWithWidth = false, custom?: CustomDash): {
    dashArray?: string;
    lineCap?: "round" | "square" | "butt";
    shapeRendering: "crispEdges" | "auto";
} {
    // 模様は gridDashOf（「幅で拡大縮小」がオンなら線の幅に比例、オフなら幅によらず同じ模様）。点線は丸い端で細かい点にする。
    // カスタムは、ダッシュ配列の模様とダッシュ キャップの端（模様が読めなければ実線）
    const dashArray = gridDashOf(style, width, scaleWithWidth, custom);
    if (!dashArray) return { shapeRendering: "auto" };
    if (style === CUSTOM_LINE_STYLE) return { dashArray, lineCap: lineCapOf(style, custom), shapeRendering: "auto" };
    return style === "dotted" ? { dashArray, lineCap: "round", shapeRendering: "auto" } : { dashArray, shapeRendering: "auto" };
}

/**
 * 折れ線の模様（stroke-dasharray）と端の形。破線・点線は線の幅に比例させる（「幅で拡大縮小」をオフにすると幅 1 のときの模様）。
 * カスタムはダッシュ配列の模様とダッシュ キャップの端（模様が読めなければ実線）
 */
export function lineStrokeOf(
    line: { lineStyle: string; dashArray: string; dashCap: string; scaleWithWidth: boolean },
    width: number
): { dashArray?: string; lineCap: "round" | "square" | "butt" } {
    const w = line.scaleWithWidth ? width : 1;
    // 模様の長さは小数 2 桁に丸める（0.1 × 3 が 0.30000000000000004 にならないように）
    const n = (v: number) => Math.round(v * 100) / 100;
    if (line.lineStyle === "dashed") return { dashArray: `${n(w * 3)} ${n(w * 2)}`, lineCap: "butt" };
    if (line.lineStyle === "dotted") return { dashArray: `${n(w * 0.1)} ${n(w * 2)}`, lineCap: "round" };
    if (line.lineStyle === CUSTOM_LINE_STYLE) {
        const custom = { dashArray: line.dashArray, dashCap: line.dashCap };
        return { dashArray: customDashOf(line.dashArray, width, line.scaleWithWidth) ?? undefined, lineCap: lineCapOf(line.lineStyle, custom) };
    }
    return { lineCap: "butt" };
}
