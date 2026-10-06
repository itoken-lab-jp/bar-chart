import powerbi from "powerbi-visuals-api";

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
