import powerbi from "powerbi-visuals-api";

/** ハイコントラストで使える 3 色（前景・背景・選んだものの色） */
export interface HighContrastColors {
    foreground: string;
    background: string;
    selected: string;
}

/**
 * ハイコントラストのテーマなら、使える色を返す。そうでなければ null。
 * ハイコントラストでは、この 3 色だけで描く（テーマの色・書式で選んだ色は読めないことがある）
 */
export function highContrastOf(palette: powerbi.extensibility.ISandboxExtendedColorPalette | undefined): HighContrastColors | null {
    if (!palette?.isHighContrast) return null;
    const foreground = palette.foreground?.value ?? "#000000";
    return {
        foreground,
        background: palette.background?.value ?? "#FFFFFF",
        selected: palette.foregroundSelected?.value ?? foreground,
    };
}

/** 色を使えないときに、塗りと線の模様で系列を見分ける。並びの順に使い、足りなければ先頭から繰り返す */
export interface HighContrastPattern {
    /** 塗りを前景色にするか（false なら背景色で塗り、前景色の線で囲む） */
    solid: boolean;
    /** 囲む線の模様（SVG の stroke-dasharray）。実線なら空 */
    dash: string;
}

const PATTERNS: HighContrastPattern[] = [
    { solid: false, dash: "" },
    { solid: true, dash: "" },
    { solid: false, dash: "4 2" },
    { solid: false, dash: "1 2" },
    { solid: false, dash: "6 2 1 2" },
];

export function highContrastPattern(index: number): HighContrastPattern {
    const n = PATTERNS.length;
    return PATTERNS[((index % n) + n) % n];
}

/** 線（折れ線など）を見分ける模様。1 本目は実線 */
const LINE_DASHES = ["", "6 3", "2 2", "8 3 2 3"];

export function highContrastLineDash(index: number): string {
    const n = LINE_DASHES.length;
    return LINE_DASHES[((index % n) + n) % n];
}

const COLOR_KEY = /(color|fill)$/i;
/** 背景色にする項目（背景・文字の縁）。その下の color も背景色にする。網掛けは透かして塗るので前景色 */
const BACKGROUND_KEY = /(background|halo)/i;
const COLOR_VALUE = /^(#[0-9a-f]{3,8}|rgba?\(.*\))$/i;

/**
 * 描くための設定の色を、まとめてハイコントラストの色に置き換えた写しを返す（元は変えない）。
 * 名前が color・fill で終わる項目のうち、色の文字列が入っているものが対象。背景・文字の縁の色は背景色に、
 * それ以外（文字・線・軸・塗り・網掛け）は前景色にする。空（自動）はそのまま残す。
 * 写すのは普通のオブジェクトと配列だけ（選択の ID などのクラスや関数は同じものを指す）。同じものを指していた所は、写しでも同じものを指す。
 * skip に挙げた名前の項目は写さずにそのまま使う（DataView を持つツールヒントの元データなど）
 */
export function withHighContrastColors<T>(target: T, hc: HighContrastColors, skip: readonly string[] = []): T {
    const copies = new Map<object, unknown>();
    const copy = (node: unknown, background: boolean): unknown => {
        if (!node || typeof node !== "object") return node;
        const done = copies.get(node);
        if (done) return done;
        if (Array.isArray(node)) {
            const out: unknown[] = [];
            copies.set(node, out);
            node.forEach((v) => out.push(copy(v, background)));
            return out;
        }
        const proto = Object.getPrototypeOf(node);
        if (proto !== Object.prototype && proto !== null) return node;
        const out: Record<string, unknown> = {};
        copies.set(node, out);
        for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
            if (skip.includes(key)) out[key] = value;
            else if (typeof value === "string")
                out[key] = COLOR_KEY.test(key) && COLOR_VALUE.test(value.trim()) ? (background || BACKGROUND_KEY.test(key) ? hc.background : hc.foreground) : value;
            else out[key] = copy(value, background || BACKGROUND_KEY.test(key));
        }
        return out;
    };
    return copy(target, false) as T;
}
