import { pickUnitIndex, type Grouping } from "./shared/units";
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
import { measureTextWidth, truncateToWidth } from "./shared/text";
import { maxOf } from "./shared/minMax";
import { H_LABEL_BACKGROUND_PAD } from "./shared/axisBands";

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
    style: Partial<SignStyle> = {},
    /** 桁区切りを付けるか（値の書式で決める。groupingOf。g の書式は、割っても付かない）。省略すると付ける */
    grouping?: Grouping
): string {
    const abs = Math.abs(value);
    if (abs === 0) return formatSigned(0, 1, "0", style);

    if (!withUnit) {
        return formatSigned(value, 1, precision, style, "", grouping);
    }

    const rung = dynamicRungOf(abs, notation, precision);
    return formatSigned(value, rung.divisor, precision, style, rung.word, grouping);
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

export interface PlacedLabel<L extends LabelLine = LabelLine> {
    /** 出すか。値の向きに入らない内側のラベルと、省いても「…」すら入らないラベルは出さない */
    show: boolean;
    /** 描く行。オーバーフロー テキストがオフで文字の幅の向きに入らない行は、末尾を「…」で省いたもの */
    lines: L[];
    /** 省く前のラベルが、値の向きにも棒の幅の向きにも入るか */
    fits: boolean;
    /** 省く前のラベルが、値の向き（棒の高さ）に入るか */
    fitsVertically: boolean;
    /** 省く前のラベルが、棒の幅の向きに入るか（背景の余白を両側に残して） */
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
        const size = maxOf(lines.map((l) => l.font.size));
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
            ...labelBlockEdges(lines, layout),
            baselines: lines.map(() => lineY),
            xs,
            anchor: "start",
        };
    }
    const width = Math.max(0, maxOf(widths));
    const align = lines.length > 1 ? layout.align ?? "center" : "center";
    const anchor = align === "left" ? "start" : align === "right" ? "end" : "middle";
    const x = align === "left" ? -width / 2 : align === "right" ? width / 2 : 0;
    const xs = lines.map(() => x);
    const { topEdge, bottomEdge } = labelBlockEdges(lines, layout);
    if (lines.length <= 1) return { width, topEdge, bottomEdge, baselines: [(lines[0]?.font.size ?? 0) * 0.35], xs, anchor };
    const baselines: number[] = [];
    let cursor = topEdge;
    for (const line of lines) {
        baselines.push(cursor + line.font.size * ASCENT_RATIO);
        cursor += line.font.size * (ASCENT_RATIO + DESCENT_RATIO) + LABEL_LINE_GAP;
    }
    return { width, topEdge, bottomEdge, baselines, xs, anchor };
}

/**
 * labelBlock の上下の端（topEdge・bottomEdge）。文字の幅は測らない（行の数と文字の大きさだけで決まる）。
 * 横書きのラベルの値の向きの長さ（高さ）を、幅を測らずに出すのに使う
 */
export function labelBlockEdges(lines: Array<Pick<LabelLine, "font">>, layout: LabelLayout = {}): { topEdge: number; bottomEdge: number } {
    if (lines.length <= 1 || layout.singleLine) {
        // 1 行（1 行に並べるときは、いちばん大きい文字で高さを決める）
        const size = lines.length > 1 ? maxOf(lines.map((l) => l.font.size)) : lines[0]?.font.size ?? 0;
        const lineY = size * 0.35;
        return { topEdge: lineY - size * ASCENT_RATIO, bottomEdge: lineY + size * DESCENT_RATIO };
    }
    const total =
        lines.reduce((sum, l) => sum + l.font.size * (ASCENT_RATIO + DESCENT_RATIO), 0) + LABEL_LINE_GAP * (lines.length - 1);
    return { topEdge: -total / 2, bottomEdge: total / 2 };
}

// --- データ ラベルを棒に置く ---------------------------------------------------

/**
 * データ ラベルを棒に置く数。標準（Power BI Desktop の集合縦棒・積み上げ縦棒・集合横棒。レスポンシブはオフ）を 2026-10-19 に測って決めた。
 *
 * 縦棒の横書きのラベル（文字の高さが値の向き）は、文字の四角（labelBlock の上下の端）ではなく、数字にほぼ沿った「数字の四角」
 * （1 行目のベースライン − digitAscent×大きさ 〜 最後の行のベースライン ＋ digitDescent×大きさ）で置く。標準の、棒の端から数字の画素の
 * 端までのすき間は、文字が大きいほど少し広がる（内側上 9pt で 9.1px・24pt で 10.7px）。数字の画素の高さを 0.7em とみると、数字の四角を
 * columnGap 離せば、すき間は内側上 9.26・10.86、内側下 24pt 10.86、外側上 9pt 9.26 px で、測った 9.1・10.7・11.1・9.5 px と 1 つの数で合う。
 * 細い区分に出す線（内側上 9pt 22.3・24pt 42.3、内側下 9pt 20.86、内側中央 9pt 10.32 px）も、測った線（21.2〜24.1・36.9〜44.2・
 * 16.9〜21.2・8.7〜11.3 px の間）に入る。同じ 21.2px の区分に内側上は出さず内側下は出すのも、数字の四角が文字の四角の上に寄っているから。
 *
 * 横棒と、縦棒の縦向きのラベル（文字の幅が値の向き）は、文字の端（字の送りの幅の端）を barGap 離す。数字の画素の端は、字の送りと形の差
 * （約 0.04em）だけ内側になり、24pt で 8.8px・9pt で 8.0px と、測った 8.3（内側上 24pt）・7.5（内側上 9pt）・8.3（内側下 9pt）と 1px 以内。
 * 入るかの線は、9pt で 60（棒の長さ 19.1px）に出さず 80（25.4px）に出し、20pt で -180 を「-1…」に省く標準と合う
 */
export const LABEL_PLACEMENT = {
    /** 数字の四角の上の端：1 行目のベースラインから、1 行目の文字の大きさのこれだけ上 */
    digitAscent: 0.78,
    /** 数字の四角の下の端：最後の行のベースラインから、最後の行の文字の大きさのこれだけ下 */
    digitDescent: 0.08,
    /** 縦棒の横書き：数字の四角の、つなぐ端（内側上・外側上は値の端、内側下は根元）に近い側の端と、つなぐ端のすき間 (px) */
    columnGap: 8.3,
    /** 縦棒の横書き：内側上・内側下は「すき間 ＋ 数字の四角の近い端から文字の四角の遠い端まで ＋ これ ≤ 区分の長さ」なら入る (px) */
    columnSlack: 2,
    /** 文字の幅が値の向きのラベル（横棒・縦棒の縦向き）：文字の端と、つなぐ端のすき間 (px) */
    barGap: 7.5,
    /** 文字の幅が値の向きのラベル：内側上・内側下は「すき間 ＋ 文字の幅 ＋ これ ≤ 棒の長さ」なら入る (px) */
    barSlack: 1,
    /**
     * 描くもの（背景を描くなら背景の四角、描かないなら文字の四角）の近い側の端を、つなぐ端から少なくともこれだけ離す (px)。
     * 大きい文字で背景を描くと、下りる字（カンマ・括弧・g・p・y）の分だけ背景が数字の四角より外へ出て、根元や値の端を越えるため
     */
    minClearance: 1,
} as const;

/**
 * データラベルを棒の外の端に置くか。外側の上と、外側の上として置く自動（集合の棒の「自動」。積み上げの「自動」は App で内側の中央に読み替える）。
 * 内側の位置は、入らなくても外へ逃がさない（標準と同じ。値の向きに入らなければ出さない）
 */
export function labelGoesOutside(position: string): boolean {
    return position === "outsideEnd" || position === "auto";
}

/** 値の向きの、ラベルをつなぐ端（edge）と、ラベルがそこから伸びる向き（dir。+1 は座標が増える向き）。内側の中央は null */
function connectionOf(position: string, lo: number, hi: number, endAtHi: boolean): { edge: number; dir: 1 | -1 } | null {
    if (labelGoesOutside(position)) return endAtHi ? { edge: hi, dir: 1 } : { edge: lo, dir: -1 };
    if (position === "insideTop") return endAtHi ? { edge: hi, dir: -1 } : { edge: lo, dir: 1 };
    if (position === "insideBottom") return endAtHi ? { edge: lo, dir: 1 } : { edge: hi, dir: -1 };
    return null;
}

/**
 * 文字の幅が値の向きのラベル（横棒・縦棒の縦向き）を、値の向きに置く。lo〜hi は棒の値の向きの範囲、endAtHi は値の端が hi の側か、
 * length は文字の幅、pad は描くものが文字の端より外へ出る幅（背景の余白。描かなければ 0）。
 * center は文字の真ん中、fits は値の向きに入るか（外は入るとする）、room は省くときに文字に使える幅
 */
export function placeAlongLength(
    position: string,
    lo: number,
    hi: number,
    endAtHi: boolean,
    length: number,
    pad = 0,
    boxPad = 0
): { center: number; fits: boolean; room: number } {
    const span = Math.max(0, hi - lo);
    const at = connectionOf(position, lo, hi, endAtHi);
    // boxPad は、棒の中に入るかを描くものの四角で見るときの、文字の端から四角の端まで（横棒の背景の左右の余白）。四角が棒の両端の内に入れば入る
    if (!at) return { center: (lo + hi) / 2, fits: length + boxPad * 2 <= span, room: span - boxPad * 2 };
    const { barGap, barSlack, minClearance } = LABEL_PLACEMENT;
    // つなぐ端から文字の近い端まで
    const reach = Math.max(barGap, minClearance + pad);
    const outside = labelGoesOutside(position);
    const farSlack = Math.max(barSlack, boxPad);
    return {
        center: at.edge + at.dir * (reach + length / 2),
        fits: outside || reach + length + farSlack <= span,
        room: outside ? Infinity : span - reach - farSlack,
    };
}

/**
 * 縦棒の横書きのラベルの数字の四角（基準点に対する上下の端）。labelBlock と同じ並べ方で、文字の幅は測らない。
 * 1 行に並べるときは、いちばん大きい文字で決める
 */
export function digitEdgesOf(lines: Array<Pick<LabelLine, "font">>, layout: LabelLayout = {}): { top: number; bottom: number } {
    const { topEdge, bottomEdge } = labelBlockEdges(lines, layout);
    if (!lines.length) return { top: topEdge, bottom: bottomEdge };
    const one = lines.length === 1 || layout.singleLine;
    const first = one ? maxOf(lines.map((l) => l.font.size)) : lines[0].font.size;
    const last = one ? first : lines[lines.length - 1].font.size;
    // 1 行目のベースラインは上の端から ASCENT_RATIO、最後の行のベースラインは下の端から DESCENT_RATIO（labelBlock）
    return {
        top: topEdge + (ASCENT_RATIO - LABEL_PLACEMENT.digitAscent) * first,
        bottom: bottomEdge - (DESCENT_RATIO - LABEL_PLACEMENT.digitDescent) * last,
    };
}

/**
 * 縦棒の横書きのラベル（文字の高さが値の向き）を、値の向きに置く。text は文字の四角、digit は数字の四角（基準点に対する上下の端）、
 * pad は描くものが文字の四角より外へ出る幅。anchor は基準点、fits は値の向きに入るか（外は入るとする）
 */
export function placeAcrossLength(
    position: string,
    lo: number,
    hi: number,
    endAtHi: boolean,
    text: { top: number; bottom: number },
    digit: { top: number; bottom: number },
    pad = 0
): { anchor: number; fits: boolean } {
    const span = Math.max(0, hi - lo);
    const at = connectionOf(position, lo, hi, endAtHi);
    // 内側の中央は、基準点を真ん中に置く（数字の四角が真ん中に来る）。数字の四角が入れば出す（文字の四角より低い区分にも出す。標準と同じ）
    if (!at) return { anchor: (lo + hi) / 2, fits: digit.bottom - digit.top <= span };
    const { columnGap, columnSlack, minClearance } = LABEL_PLACEMENT;
    const outside = labelGoesOutside(position);
    if (at.dir > 0) {
        const anchor = at.edge + Math.max(columnGap - digit.top, minClearance + pad - text.top);
        return { anchor, fits: outside || anchor + text.bottom - at.edge + columnSlack <= span };
    }
    const anchor = at.edge - Math.max(columnGap + digit.bottom, minClearance + pad + text.bottom);
    return { anchor, fits: outside || at.edge - (anchor + text.top) + columnSlack <= span };
}

/** 幅に入るように末尾を「…」で省く。1 文字も入らなければ「…」だけ、「…」も入らなければ null */
export function fitTextToWidth(text: string, maxWidth: number, font: FontSpec): string | null {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const cut = truncateToWidth(text, maxWidth, font);
    if (cut) return cut;
    return measureTextWidth(LABEL_ELLIPSIS, font) <= maxWidth ? LABEL_ELLIPSIS : null;
}

const LABEL_ELLIPSIS = "…";

/** 行ごとに幅に入るように省く。どれかの行が「…」も入らなければ null */
function fitLinesToWidth<L extends LabelLine>(lines: L[], maxWidth: number): L[] | null {
    const fitted: L[] = [];
    for (const line of lines) {
        const text = fitTextToWidth(line.text, maxWidth, line.font);
        if (text === null) return null;
        fitted.push(text === line.text ? line : { ...line, text });
    }
    return fitted;
}

/**
 * 縦棒のデータラベルの配置（位置・向き・入らないときの省き方）を決める。lines を渡すと複数行（値と詳細）のまとまりで決める。
 * 内側の位置は、値の向きに入らなければ出さない（外へ逃がさない）。棒の幅の向きは、オーバーフロー テキストがオンならはみ出して描き、
 * オフなら、横書きは行の末尾を「…」で省いて入れる。縦向きは文字の幅が値の向きなので、オンなら値の向きに入らなければ出さず、
 * オフなら棒の幅に入らなければ出さず、値の向きに入らなければ「…」で省く（標準と同じ。LABEL_PLACEMENT）
 */
export function placeLabel<L extends LabelLine = LabelLine>(spec: {
    position: string; // "auto" | "insideTop" | "insideCenter" | "insideBottom" | "outsideEnd"
    top: number;
    bottom: number;
    barWidth: number;
    value: string;
    font: FontSpec;
    lines?: L[];
    layout?: LabelLayout;
    /** 値の端が下（負の値の棒など）。省くと上 */
    endAtBottom?: boolean;
    vertical: boolean;
    overflow: boolean;
    /** 背景を描く（描くものの端を棒の端から離すのに使う）。省くと描かない */
    background?: boolean;
}): PlacedLabel<L> {
    const { position, top, bottom, barWidth, vertical, overflow, layout } = spec;
    const original = spec.lines ?? ([{ text: spec.value, font: spec.font }] as L[]);
    const outside = labelGoesOutside(position);
    // 値の端が下の棒（負の値。範囲の反転では正の値）は、端・外・根元を上下で入れ替える（標準と同じ）
    const down = spec.endAtBottom ?? false;
    const pad = spec.background ? LABEL_PADDING : 0;
    const { topEdge, bottomEdge } = labelBlockEdges(original, layout);
    let lines = original;
    let block = labelBlock(lines, layout);
    // 棒の幅の向き（背景の余白を両側に残して、さらに 2px）
    const fitsAlongBar = barWidth >= (vertical ? bottomEdge - topEdge : block.width) + LABEL_PADDING * 2 + 2;
    let show = true;
    let fitsVertically: boolean;
    let anchorY: number;
    if (vertical) {
        // 縦向きは rotate(-90) で回す：文字の幅の向きが上下になる（右の端が上）
        let along = placeAlongLength(position, top, bottom, down, block.width, pad);
        fitsVertically = along.fits;
        if (!outside && !along.fits) {
            const fitted = overflow || !fitsAlongBar ? null : fitLinesToWidth(original, along.room);
            if (fitted) {
                lines = fitted;
                block = labelBlock(lines, layout);
                along = placeAlongLength(position, top, bottom, down, block.width, pad);
            } else show = false;
        } else if (!outside && !overflow && !fitsAlongBar) show = false;
        anchorY = along.center;
    } else {
        const across = placeAcrossLength(position, top, bottom, down, { top: topEdge, bottom: bottomEdge }, digitEdgesOf(original, layout), pad);
        fitsVertically = across.fits;
        if (!outside && !across.fits) show = false;
        else if (!outside && !overflow && !fitsAlongBar) {
            const fitted = fitLinesToWidth(original, barWidth - LABEL_PADDING * 2 - 2);
            if (fitted) {
                lines = fitted;
                block = labelBlock(lines, layout);
            } else show = false;
        }
        anchorY = across.anchor;
    }

    const box = {
        x: -block.width / 2 - LABEL_PADDING,
        y: block.topEdge - LABEL_PADDING,
        width: block.width + LABEL_PADDING * 2,
        height: block.bottomEdge - block.topEdge + LABEL_PADDING * 2,
    };

    return {
        show,
        lines,
        fits: fitsVertically && fitsAlongBar,
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

/**
 * 横棒のデータラベルの配置。left〜right は棒の横の範囲、rightward は値の端が右か、thickness は棒の太さ。
 * position は読み替えたあとの位置（集合の自動は外側の端、積み上げの自動は内側の中央）。
 * 内側の位置は、オーバーフロー テキストがオンなら、棒の長さに入らなければ出さず（0 の線を越えて描かない）、太さはみ出して描く。
 * オフなら、太さに入らなければ出さず、長さに入らなければ行の末尾を「…」で省く（標準と同じ）。textLeft は文字の左の端
 */
export function placeBarLabel<L extends LabelLine>(spec: {
    position: string;
    left: number;
    right: number;
    rightward: boolean;
    thickness: number;
    lines: L[];
    layout?: LabelLayout;
    overflow: boolean;
    background?: boolean;
}): { show: boolean; lines: L[]; block: ReturnType<typeof labelBlock>; textLeft: number; outside: boolean; fitsLength: boolean; fitsThickness: boolean } {
    const { position, left, right, rightward, thickness, layout, overflow } = spec;
    // 背景は字の左右に H_LABEL_BACKGROUND_PAD 広げて描くので、棒の中に入るかは背景の四角で見る（棒の外へ出さない）。
    // 1.36.11.55 の最初の作りまでは字の四角で見ていて、背景が棒の端から最大 3px 出た
    const pad = spec.background ? H_LABEL_BACKGROUND_PAD : 0;
    const outside = labelGoesOutside(position);
    let lines = spec.lines;
    let block = labelBlock(lines, layout);
    let along = placeAlongLength(position, left, right, rightward, block.width, pad, pad);
    const fitsLength = along.fits;
    // 太さは文字の大きさ（複数行はまとまりの高さ）＋ 2 で見る
    const blockHeight = lines.length > 1 ? block.bottomEdge - block.topEdge : (lines[0]?.font.size ?? 0);
    const fitsThickness = thickness >= blockHeight + 2;
    let show = true;
    if (!outside && overflow) show = fitsLength;
    else if (!outside && !fitsThickness) show = false;
    else if (!outside && !fitsLength) {
        const fitted = fitLinesToWidth(lines, along.room);
        if (fitted) {
            lines = fitted;
            block = labelBlock(lines, layout);
            along = placeAlongLength(position, left, right, rightward, block.width, pad, pad);
        } else show = false;
    }
    return { show, lines, block, textLeft: along.center - block.width / 2, outside, fitsLength, fitsThickness };
}

/**
 * 棒の外の端に置いたデータ ラベルの、値の向きの長さに入れる字の外の余白 (px)。文字の送りの端から、長さの端まで。
 * ふつうは背景の余白（背景を描かなくても足す。文字の縁の分）。横棒の右へ伸びる棒（rightward。範囲の反転では負の値の棒）の
 * ラベルは 0 にする。
 * 標準の集合横棒（2026-10-20 の Desktop）は、外のラベルの字の送りの右の端を、ラベルの無いときの描く範囲の右の端に付けて置いた
 * （字の画素の右の端は、送りの端の 0.04em 内側で、9〜36pt でビューの右の端から同じ所）。背景を描いても字の位置を変えず、背景の四角を
 * 描く範囲の右の外（右の余白の中）へ H_LABEL_BACKGROUND_PAD 広げた（1.36.11.54 までは、背景を描くと背景の余白の分だけ字を内側へ寄せていた）。
 * 左へ伸びる棒は測っていないので、背景の四角が描く範囲に収まるだけとる（背景なしは前のまま。背景ありは横棒の背景の余白）
 */
export function outsideLabelSlack(horizontalBars: boolean, background: boolean, rightward: boolean): number {
    if (!horizontalBars) return LABEL_PADDING;
    return rightward ? 0 : background ? H_LABEL_BACKGROUND_PAD : LABEL_PADDING;
}

/**
 * 棒の外の端に置いたデータ ラベルの、値の向きの長さ (px)。つなぐ端から、文字の遠い側の端に字の外の余白（outsideLabelSlack）を足した所まで。
 * alongLength は文字の幅が値の向き（横棒・縦棒の縦向き）、horizontalBars は横棒か、
 * up は値の端が値の軸の大きい側の端の向き（縦棒で上、横棒で右。範囲の反転では逆）か。縦棒の横書きは数字の四角で置くので、上と下で長さが違う
 */
export function outsideLabelLength(
    lines: Array<Pick<LabelLine, "font">>,
    layout: LabelLayout,
    spec: { alongLength: boolean; width: number; background: boolean; up: boolean; horizontalBars?: boolean }
): number {
    const pad = spec.background ? LABEL_PADDING : 0;
    if (spec.alongLength) {
        const { center } = placeAlongLength("outsideEnd", 0, 0, false, spec.width, pad);
        return -center + spec.width / 2 + outsideLabelSlack(spec.horizontalBars ?? false, spec.background, spec.up);
    }
    const { topEdge, bottomEdge } = labelBlockEdges(lines, layout);
    const text = { top: topEdge, bottom: bottomEdge };
    const { anchor } = placeAcrossLength("outsideEnd", 0, 0, !spec.up, text, digitEdgesOf(lines, layout), pad);
    return spec.up ? -(anchor + topEdge) + LABEL_PADDING : anchor + bottomEdge + LABEL_PADDING;
}
