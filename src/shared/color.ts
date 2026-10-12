/**
 * 色の計算。色は #RRGGBB・#RGB・#RRGGBBAA（透明度は無視）を読み、読めなければ呼ぶ側の既定に落とす
 */

/** 読める文字の濃い色（白の反対） */
export const DARK_TEXT = "#252423";

/** WCAG の相対輝度（#RRGGBB・#RGB・#RRGGBBAA）。透明度は無視する。読めなければ null */
export function luminanceOf(color: string): number | null {
    const hex = color.trim().replace("#", "");
    const full = hex.length === 3 ? hex.split("").map((c) => c + c).join("") : hex.length === 8 ? hex.slice(0, 6) : hex;
    if (!/^[0-9a-f]{6}$/i.test(full)) return null;
    const value = Number.parseInt(full, 16);
    const channel = (shift: number) => {
        const c = ((value >> shift) & 0xff) / 255;
        return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    };
    return 0.2126 * channel(16) + 0.7152 * channel(8) + 0.0722 * channel(0);
}

/** 2 つの色のコントラスト比（WCAG、1〜21）。読めない色があれば 21 とみなす */
export function contrastRatio(a: string, b: string): number {
    const la = luminanceOf(a);
    const lb = luminanceOf(b);
    if (la === null || lb === null) return 21;
    return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/** 背景の色に対して読みやすい文字の色（白か濃い灰色）。WCAG の相対輝度が 0.4 を超えれば濃い灰色。読めない色は濃い灰色 */
export function contrastingText(backgroundColor: string): string {
    const luminance = luminanceOf(backgroundColor);
    if (luminance === null) return DARK_TEXT;
    return luminance > 0.4 ? DARK_TEXT : "#FFFFFF";
}

/**
 * 下の色 base の上で読める文字色。prefer（既定の色）のコントラスト比が minimum 以上なら prefer、
 * 足りなければ白と濃い灰色のうちコントラストの高いほう
 */
export function readableText(base: string, prefer: string, minimum: number): string {
    if (contrastRatio(prefer, base) >= minimum) return prefer;
    return contrastRatio("#FFFFFF", base) >= contrastRatio(DARK_TEXT, base) ? "#FFFFFF" : DARK_TEXT;
}

/**
 * 色を toward に amount だけ寄せる（#RRGGBB・#RGB。読めなければそのまま）。
 * 透かして描いたものの見た目の色（背景と混ざった色）を出し、上に置く文字の色を決めるのに使う
 */
export function blend(color: string, toward: string, amount: number): string {
    const parse = (c: string) => {
        // #RGB は #RRGGBB に広げて読む（背景が #FFF のときにも混ぜる）
        const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c.trim());
        const hex = short ? short.slice(1).map((d) => d + d).join("") : /^#([0-9a-f]{6})$/i.exec(c.trim())?.[1];
        return hex ? parseInt(hex, 16) : null;
    };
    const from = parse(color);
    const to = parse(toward);
    if (from === null || to === null) return color;
    const channel = (shift: number) => {
        const a = (from >> shift) & 0xff;
        const b = (to >> shift) & 0xff;
        return Math.round(a + (b - a) * amount);
    };
    return `#${[16, 8, 0].map((shift) => channel(shift).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/**
 * 文字に使う色を、色相と彩度を保ったまま明るさだけ動かして、地に対して minimum のコントラスト比にする（足りていればそのまま）。
 * 地が明るければ暗く、暗ければ明るくする。淡い面の色（センチメントの色など）から文字の色を作るとき、字の色へ混ぜると
 * 灰に近づいて黒い字と見分けにくくなるので、色味を残す。#RRGGBB・#RGB 以外は fallback
 */
export function readableKeepingHue(color: string, background: string, minimum: number, fallback: string): string {
    const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color.trim())?.[1];
    if (!hex || luminanceOf(background) === null) return fallback;
    if (contrastRatio(color, background) >= minimum) return color;
    const full = hex.length === 3 ? hex.split("").map((d) => d + d).join("") : hex;
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16) / 255);
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const l = (max + min) / 2;
    const d = max - min;
    const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
    const h = d === 0 ? 0 : max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    const toHex = (lightness: number) => {
        const c = (1 - Math.abs(2 * lightness - 1)) * s;
        const x = c * (1 - Math.abs((h % 2) - 1));
        const m = lightness - c / 2;
        const [r1, g1, b1] = h < 1 ? [c, x, 0] : h < 2 ? [x, c, 0] : h < 3 ? [0, c, x] : h < 4 ? [0, x, c] : h < 5 ? [x, 0, c] : [c, 0, x];
        return `#${[r1, g1, b1].map((v) => Math.round((v + m) * 255).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
    };
    const darker = (luminanceOf(background) ?? 1) > 0.18;
    for (let step = 1; step <= 200; step++) {
        const lightness = darker ? l - step * 0.005 : l + step * 0.005;
        if (lightness < 0 || lightness > 1) break;
        const candidate = toHex(lightness);
        if (contrastRatio(candidate, background) >= minimum) return candidate;
    }
    return fallback;
}

/** 同じ色か（#FFF と #FFFFFF、white と #FFFFFF を同じとみなす。大文字・小文字は問わない） */
export function sameColor(a: string, b: string): boolean {
    return colorKeyOf(a) === colorKeyOf(b);
}

const NAMED_COLORS: Record<string, string> = { white: "#ffffff", black: "#000000" };

function colorKeyOf(color: string): string {
    const c = color.trim().toLowerCase();
    if (NAMED_COLORS[c]) return NAMED_COLORS[c];
    const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(c);
    return short ? `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}` : c;
}

/** データ ラベルの自動の色で、棒の外に出すラベル（と背景を付けたラベル）に優先する灰色 */
export const LABEL_OUTSIDE_COLOR = "#605E5C";
/**
 * 自動の文字色に求めるコントラスト比。棒の中は白を優先し 3 以上（標準と同じく、テーマの青 #118DFF の上は白）、
 * 外は灰色を優先し 4.5 以上
 */
export const INSIDE_TEXT_CONTRAST = 3;
export const OUTSIDE_TEXT_CONTRAST = 4.5;

/**
 * データ ラベルの自動の文字色。base は文字のすぐ下の色（棒の中なら棒の色、外ならビジュアルの背景。ラベルの背景を出すなら
 * その色を下の色と混ぜた色）。棒の中は白、外は灰色を優先し、読めなければ白か濃い灰色の読めるほう
 */
export function autoLabelColor(inside: boolean, base: string): string {
    return inside ? readableText(base, "#FFFFFF", INSIDE_TEXT_CONTRAST) : readableText(base, LABEL_OUTSIDE_COLOR, OUTSIDE_TEXT_CONTRAST);
}
