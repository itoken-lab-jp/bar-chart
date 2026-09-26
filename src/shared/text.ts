import { textMeasurementService } from "powerbi-visuals-utils-formattingutils";

/** pt -> px 換算比率 (1pt = 4/3 px) */
export const PT_TO_PX = 4 / 3;

export interface FontSpec {
    family: string;
    /** ピクセル */
    size: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
}

/**
 * 文字の幅。Power BI の文字幅の計測（DOM）で測り、測れなければ全角・半角の近似で返す。
 * テストは textMeasurementService を決定的な近似に差し替えるが、DOM が無いので近似の側を通る
 */
export function measureTextWidth(text: string, font: FontSpec): number {
    try {
        if (typeof document !== "undefined") {
            const width = textMeasurementService.measureSvgTextWidth({
                text,
                fontFamily: font.family,
                fontSize: `${font.size}px`,
                fontWeight: font.bold ? "bold" : "normal",
                fontStyle: font.italic ? "italic" : "normal",
            });
            if (width > 0) return width;
        }
    } catch {
        // 近似に落とす
    }
    let w = 0;
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        w += code >= 0x20 && code <= 0x7e ? font.size * 0.6 : font.size;
    }
    return font.bold ? w * 1.05 : w;
}

/**
 * 幅に収まるように先頭を「…」で切る。上の階層から並べた名前で、一番下のレベルの名前を残すのに使う。
 * 1 文字も入らなければ空
 */
export function truncateStartToWidth(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const ellipsis = "…";
    for (let n = 1; n < text.length; n++) {
        const candidate = ellipsis + text.slice(n);
        if (measureTextWidth(candidate, font) <= maxWidth) return candidate;
    }
    return "";
}

/** 幅に収まるように末尾を「…」で切る。1 文字も入らなければ空 */
export function truncateToWidth(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const ellipsis = "…";
    for (let n = text.length - 1; n > 0; n--) {
        const candidate = text.slice(0, n) + ellipsis;
        if (measureTextWidth(candidate, font) <= maxWidth) return candidate;
    }
    return "";
}
