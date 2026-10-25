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

const graphemeSegmenter: Intl.Segmenter | null =
    typeof Intl !== "undefined" && typeof Intl.Segmenter === "function" ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : null;

/**
 * 字を、見た目の 1 文字（書記素）ごとに分けたもの。サロゲート対・結合文字（濁点など）・ZWJ でつないだ絵文字・異体字セレクターは
 * 前の字と 1 つにまとめる（途中で切ると、欠けた字や □ が出る）。Intl.Segmenter が無ければコード ポイントごと（サロゲート対だけまとめる）。
 * 和文・英字・数字は 1 字ずつで、文字列の 1 字（UTF-16 の 1 単位）と同じ
 */
export function graphemesOf(text: string): string[] {
    if (graphemeSegmenter) return Array.from(graphemeSegmenter.segment(text), (s) => s.segment);
    return Array.from(text);
}

const ELLIPSIS = "…";

/** 幅に収まるように末尾を「…」で切る。字の途中（書記素の中）では切らない。1 文字も入らなければ空 */
export function truncateToWidth(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const parts = graphemesOf(text);
    // 頭から n 字の字（n = 1〜parts.length − 1）
    const heads: string[] = [];
    for (let n = 1, head = ""; n < parts.length; n++) heads.push((head += parts[n - 1]));
    for (let n = parts.length - 1; n > 0; n--) {
        // 「…」の前の空白は詰める（「製品 …」ではなく「製品…」）
        const candidate = heads[n - 1].trimEnd() + ELLIPSIS;
        if (measureTextWidth(candidate, font) <= maxWidth) return candidate;
    }
    return "";
}

/**
 * 頭の 1 文字（書記素）と「…」。省いても名前の頭を残すときの一番短い形。元の字のほうが狭ければ（1 文字の名前など。
 * 「東…」は「東」より広い）元の字。同じ幅なら「…」の形のまま
 */
export function headWithEllipsis(text: string, font: FontSpec): string {
    if (!text) return "";
    const short = graphemesOf(text)[0] + ELLIPSIS;
    return measureTextWidth(text, font) < measureTextWidth(short, font) ? text : short;
}
