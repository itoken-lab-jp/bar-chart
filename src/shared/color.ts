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
