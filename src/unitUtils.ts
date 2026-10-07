import { pickUnitIndex } from "./shared/units";
import { formatSigned, shownSignOf, SignStyle } from "./shared/numberFormat";

/**
 * 日本語ビジネスレポート向け単位ユーティリティ
 * 0〜12の桁数（指数）ベースでスケーリングと表記を解決する。単位の核と文字の幅はほかのビジュアルと共通（shared/）
 */

export type { UnitDefinition } from "./shared/units";
export { UNIT_DEFINITIONS, resolveAutoUnitKey, toStandardUnitKey, resolveUnit, formatValue } from "./shared/units";
export type { FontSpec } from "./shared/text";
export { measureTextWidth } from "./shared/text";
export { contrastingText } from "./shared/color";
import type { FontSpec } from "./shared/text";
import { measureTextWidth } from "./shared/text";

/** 値ごとに単位を選ぶときの段（小さい順） */
const DYNAMIC_LADDER_JA = [
    { divisor: 1, word: "" },
    { divisor: 1e3, word: "千" },
    { divisor: 1e4, word: "万" },
    { divisor: 1e8, word: "億" },
    { divisor: 1e12, word: "兆" },
];
const DYNAMIC_LADDER_STD = [
    { divisor: 1, word: "" },
    { divisor: 1e3, word: "K" },
    { divisor: 1e6, word: "M" },
    { divisor: 1e9, word: "bn" },
    { divisor: 1e12, word: "T" },
];

/**
 * 対数スケールなど桁が大きく跨がる場合に、値ごとの桁に応じた動的スケーリング・単位付与を行う。
 * 単位は丸めた後の値で決め直す（99,999,999 → 「1億」、9,999.9 → 「1万」、標準 999,999 → 「1M」）
 */
export function formatDynamicValue(
    value: number,
    notation: string = "japanese",
    precision: string = "auto",
    withUnit: boolean = true,
    /** マイナスと 0 の書き方（▲・±0 など）。省略すると - と 0 */
    style: Partial<SignStyle> = {}
): string {
    const abs = Math.abs(value);
    if (abs === 0) return formatSigned(0, 1, "0", style);

    if (!withUnit) {
        return formatSigned(value, 1, precision, style);
    }

    const rung = dynamicRungOf(abs, notation, precision);
    return formatSigned(value, rung.divisor, precision, style, rung.word);
}

/** 値ごとに選ぶ単位の段（formatDynamicValue と同じ選び方） */
function dynamicRungOf(abs: number, notation: string, precision: string): { divisor: number; word: string } {
    const ladder = notation === "standard" ? DYNAMIC_LADDER_STD : DYNAMIC_LADDER_JA;
    return ladder[pickUnitIndex(abs, ladder.map((r) => r.divisor), precision)];
}

/** formatDynamicValue（単位を付ける）で書いた文字の見える符号（-1・0・1）。色を決めるのに使う */
export function dynamicShownSign(value: number, notation: string = "japanese", precision: string = "auto", style: Partial<SignStyle> = {}): -1 | 0 | 1 {
    const abs = Math.abs(value);
    if (abs === 0) return shownSignOf(0, 1, "0", style);
    return shownSignOf(value, dynamicRungOf(abs, notation, precision).divisor, precision, style);
}

/**
 * スケーリング語とユーザー入力単位を合成する
 * @param unitWord 指数倍率に対応する語（例: "百万", "億"）
 * @param unitText ユーザー入力の単位（例: "円", "人", "件"）
 * @param includeDisplayUnit スケーリング語を含めるかどうか（既定: true）
 */
export function composeUnitText(
    unitWord: string,
    unitText: string,
    includeDisplayUnit: boolean = true
): string {
    const word = includeDisplayUnit ? (unitWord || "") : "";
    const custom = (unitText || "").trim();
    return `${word}${custom}`;
}

/**
 * 合成された単位文字列からバッジ用表記を組み立てる
 * @param composed 合成単位（例: "百万円", "人"）
 * @param style 表記スタイル ("parentheses" | "withPrefix")
 */
export function formatUnitBadge(
    composed: string,
    style: string = "parentheses"
): string {
    const trimmed = (composed || "").trim();
    if (!trimmed) return "";
    if (style === "withPrefix") {
        return `(単位: ${trimmed})`;
    }
    return `(${trimmed})`;
}

/**
 * 軸上バッジのテキストを解決する
 */
export function resolveBadgeText(spec: {
    unitShow: boolean;
    unitPosition: string; // "valueAxisTop" | "plotTopRight" | "none"
    unitIncludeDisplayUnit: boolean;
    unitStyle: string;
    unitWord: string;
    unitText: string;
}): string {
    if (!spec.unitShow || spec.unitPosition === "none") {
        return "";
    }
    const composed = composeUnitText(spec.unitWord, spec.unitText, spec.unitIncludeDisplayUnit);
    return formatUnitBadge(composed, spec.unitStyle);
}

// --- データラベルの配置 --------------------------------------------------------

export interface PlacedLabel {
    fits: boolean;
    fitsVertically: boolean;
    fitsAlongBar: boolean;
    outside: boolean;
    y: number; // 基準点 anchorY
    width: number;
    height: number;
    box: { x: number; y: number; width: number; height: number };
    /** 基準点に対する各行のベースライン */
    baselines: number[];
    /** 各行を描く x（まとまりの中心から）と text-anchor（labelBlock） */
    xs: number[];
    anchor: "start" | "middle" | "end";
}

/** フォントサイズに対するベースラインより上/下の割合。背景の高さを出すのに使う */
const ASCENT_RATIO = 0.8;
const DESCENT_RATIO = 0.22;
/** 背景と文字のあいだの余白 */
export const LABEL_PADDING = 3;

/** ラベルの 1 行（値の行と詳細の行で文字の設定が違う） */
export interface LabelLine {
    text: string;
    font: FontSpec;
}

/** 行と行のあいだ (px) */
const LABEL_LINE_GAP = 1;

/** ラベルの行の並べ方。単一行は行を 1 行に横に並べる。align は複数行のとき行をそろえる位置 */
export interface LabelLayout {
    singleLine?: boolean;
    align?: "left" | "center" | "right";
}

/** 単一行で並べるときの、行と行のあいだ（文字サイズに対する比。空白 1 つほど） */
const SINGLE_LINE_GAP_RATIO = 0.3;

/**
 * 複数行のラベルの寸法。基準点（0, 0）に対する各行のベースライン、上端と下端、まとまりの幅。
 * xs と anchor は、各行を描く x（まとまりの中心から）と text-anchor。
 * 1 行なら 1.10 までと同じ位置（ベースラインは文字サイズの 0.35 下）
 */
export function labelBlock(
    lines: LabelLine[],
    layout: LabelLayout = {}
): { width: number; topEdge: number; bottomEdge: number; baselines: number[]; xs: number[]; anchor: "start" | "middle" | "end" } {
    const widths = lines.map((l) => measureTextWidth(l.text, l.font));
    if (layout.singleLine && lines.length > 1) {
        // 1 行に並べる：いちばん大きい文字で高さを決め、ベースラインをそろえる
        const size = Math.max(...lines.map((l) => l.font.size));
        const gap = size * SINGLE_LINE_GAP_RATIO;
        const width = widths.reduce((sum, w) => sum + w, 0) + gap * (lines.length - 1);
        const lineY = size * 0.35;
        const xs: number[] = [];
        let cursor = -width / 2;
        widths.forEach((w) => {
            xs.push(cursor);
            cursor += w + gap;
        });
        return {
            width,
            topEdge: lineY - size * ASCENT_RATIO,
            bottomEdge: lineY + size * DESCENT_RATIO,
            baselines: lines.map(() => lineY),
            xs,
            anchor: "start",
        };
    }
    const width = Math.max(0, ...widths);
    const align = lines.length > 1 ? layout.align ?? "center" : "center";
    const anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
    const x = align === "left" ? -width / 2 : align === "right" ? width / 2 : 0;
    const xs = lines.map(() => x);
    if (lines.length <= 1) {
        const size = lines[0]?.font.size ?? 0;
        const lineY = size * 0.35;
        return { width, topEdge: lineY - size * ASCENT_RATIO, bottomEdge: lineY + size * DESCENT_RATIO, baselines: [lineY], xs, anchor };
    }
    const total =
        lines.reduce((sum, l) => sum + l.font.size * (ASCENT_RATIO + DESCENT_RATIO), 0) + LABEL_LINE_GAP * (lines.length - 1);
    const topEdge = -total / 2;
    const baselines: number[] = [];
    let cursor = topEdge;
    for (const line of lines) {
        baselines.push(cursor + line.font.size * ASCENT_RATIO);
        cursor += line.font.size * (ASCENT_RATIO + DESCENT_RATIO) + LABEL_LINE_GAP;
    }
    return { width, topEdge, bottomEdge: topEdge + total, baselines, xs, anchor };
}

/**
 * データラベルの配置（位置・向き・入りきらないときの外側への逃がし）を決める。
 * lines を渡すと複数行（値と詳細）のまとまりで決める
 */
export function placeLabel(spec: {
    position: string; // "auto" | "insideTop" | "insideCenter" | "insideBottom" | "outsideEnd"
    top: number;
    bottom: number;
    barWidth: number;
    value: string;
    font: FontSpec;
    lines?: LabelLine[];
    layout?: LabelLayout;
    /** 値の端が下（負の値の棒など）。省くと上 */
    endAtBottom?: boolean;
    vertical: boolean;
    overflow: boolean;
}): PlacedLabel {
    const {
        position,
        top,
        bottom,
        barWidth,
        value,
        font,
        vertical,
        overflow,
    } = spec;

    const block = labelBlock(spec.lines ?? [{ text: value, font }], spec.layout);

    const left = -block.width / 2;
    const right = block.width / 2;
    const topEdge = block.topEdge;
    const bottomEdge = block.bottomEdge;

    const box = {
        x: left - LABEL_PADDING,
        y: topEdge - LABEL_PADDING,
        width: right - left + LABEL_PADDING * 2,
        height: bottomEdge - topEdge + LABEL_PADDING * 2,
    };

    // 棒の高さの向きは、背景の余白を除いた文字の大きさで判定する（標準は、文字がちょうど入る細い積み上げにもラベルを出す）
    const neededY = vertical ? right - left : bottomEdge - topEdge;
    const neededX = (vertical ? box.height : box.width) + 2;
    const fitsAlongBar = barWidth >= neededX;
    const barHeight = Math.max(0, bottom - top);
    const fitsVertically = barHeight >= neededY;
    const fits = fitsVertically && fitsAlongBar;

    const requestedOutside = position === "outsideEnd";
    // 入りきらなければ外の端へ（標準と同じく、棒の幅より長いラベルも外へ。中に置くと棒からはみ出した白い文字が背景に溶ける）
    const outside = requestedOutside || (!fits && overflow);

    // 値の端が下の棒（負の値。範囲の反転では正の値）は、端・外・根元を上下で入れ替える（標準と同じ）
    const down = spec.endAtBottom ?? false;
    let anchorY: number;
    if (outside) {
        anchorY = down ? bottom + LABEL_PADDING - topEdge : top - LABEL_PADDING - bottomEdge;
    } else if (position === "insideTop" || position === "auto") {
        anchorY = down ? bottom - neededY / 2 : top + neededY / 2;
    } else if (position === "insideBottom") {
        anchorY = down ? top + neededY / 2 : bottom - neededY / 2;
    } else {
        anchorY = (top + bottom) / 2;
    }

    return {
        fits,
        fitsVertically,
        fitsAlongBar,
        outside,
        y: anchorY,
        width: box.width,
        height: box.height,
        box,
        baselines: block.baselines,
        xs: block.xs,
        anchor: block.anchor,
    };
}

