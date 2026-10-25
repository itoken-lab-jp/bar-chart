"use strict";

import type { DataLabelsSettings, DataPoint } from "./viewModel";
import { FontSpec, labelBlock, labelGoesOutside, LabelLine, LABEL_PADDING, measureTextWidth, outsideLabelLength } from "./unitUtils";
import { PT_TO_PX, truncateToWidth } from "./shared/text";
import { maxOf } from "./shared/minMax";
import { outsideTopHeadroomOf } from "./shared/plotTop";

/** データラベルの文字を作るのに要る、棒の項目 */
export type LabelTexts = Pick<DataPoint, "titleText" | "dataLabelText" | "detailText" | "labelShow" | "blank">;

/** データラベルの 1 行。どの行か（タイトル・値・詳細）で文字の設定と class が違う */
export interface DataLabelLine extends LabelLine {
    kind: "title" | "value" | "detail";
}

/** データラベルの行（タイトル・値・詳細の行。標準と同じくこの順に上から）。ラベルの表示を最適化なら、最大幅を超える行を「…」で切る */
export function labelLinesOf(dl: DataLabelsSettings, d: LabelTexts): DataLabelLine[] {
    const lines = rawLabelLinesOf(dl, d);
    if (!dl.optimizeLabelDisplay) return lines;
    return lines.map((line) => ({ ...line, text: truncateToWidth(line.text, dl.labelMaxWidth, line.font) }));
}

function rawLabelLinesOf(dl: DataLabelsSettings, d: LabelTexts): DataLabelLine[] {
    const lines: DataLabelLine[] = [];
    if (dl.titleShow && d.titleText) {
        lines.push({
            kind: "title",
            text: d.titleText,
            font: {
                family: dl.titleFontFamily,
                size: dl.titleFontSize * PT_TO_PX,
                bold: dl.titleBold,
                italic: dl.titleItalic,
                underline: dl.titleUnderline,
            },
        });
    }
    if (dl.valueShow) {
        lines.push({
            kind: "value",
            text: d.dataLabelText,
            font: { family: dl.fontFamily, size: dl.fontSize * PT_TO_PX, bold: dl.bold, italic: dl.italic, underline: dl.underline },
        });
    }
    if (dl.detailShow && d.detailText) {
        lines.push({
            kind: "detail",
            text: d.detailText,
            font: {
                family: dl.detailFontFamily,
                size: dl.detailFontSize * PT_TO_PX,
                bold: dl.detailBold,
                italic: dl.detailItalic,
                underline: dl.detailUnderline,
            },
        });
    }
    return lines;
}

/**
 * 値の軸の端にデータ ラベルの余白を取るか（棒の外の端に置くか）。外側の上と、外側の上として置く「自動」（集合の縦棒・横棒は、
 * 入るかやオーバーフロー テキストによらず外に置く。App の labelPositionOf・hLabelPositionOf）。積み上げは外に置かない（呼ぶ側で除く）。
 * 内側の位置は、入らなくても外へ逃がさず出さないので、余白は取らない
 */
export function labelMayGoOutside(dl: DataLabelsSettings): boolean {
    return dl.show && labelGoesOutside(dl.position);
}

/**
 * データラベルを棒の外の端に出したときの、値の向きの長さ (px)。棒の端から、文字の遠い側の端に字の外の余白を足した所まで
 * （unitUtils の outsideLabelLength。placeLabel・placeBarLabel と同じ置き方）。
 * 縦棒は、縦向き（rotate(-90)）なら行の幅、横書きなら数字の四角で置いた行の高さ。横棒は、文字が横書きで値の向きが横なので行の幅。
 * up は値の端が上（横棒は右。範囲の反転では逆）か。縦棒の横書きは、上と下で長さが少し違う。棒の外に置きうるラベルだけ。ほかは 0。
 * 縦棒の横書きは文字の幅を測らない（高さは行の数と文字の大きさだけで決まる。最大幅で切っても行は減らない）
 */
export function outsideLabelLengthOf(dl: DataLabelsSettings, d: LabelTexts, horizontalBars = false, up = true): number {
    if (!labelMayGoOutside(dl) || !d.labelShow || d.blank) return 0;
    const layout = { singleLine: dl.singleLine, align: dl.horizontalAlignment };
    const background = dl.backgroundShow;
    if (!horizontalBars && dl.orientation !== "vertical") {
        const lines = rawLabelLinesOf(dl, d);
        if (!lines.length) return 0;
        return outsideLabelLength(lines, layout, { alongLength: false, width: 0, background, up });
    }
    const lines = labelLinesOf(dl, d);
    if (!lines.length) return 0;
    return outsideLabelLength(lines, layout, { alongLength: true, width: labelBlock(lines, layout).width, background, up, horizontalBars });
}

/**
 * 縦棒の上の端の外に置く横書きのデータ ラベルの、文字の四角（と背景の余白）の上に足して空ける高さ (px)。値の軸の上の余白に足す
 * （viewModel の measureLabels。字の外の余白として扱う）。文字の大きさは一番上の行（1 行に並べるなら一番大きい字）で見て、0 より小さければ 0。
 * 下の端（棒の下に置くラベル）・横棒・縦向きのラベルは標準を測っていないので、足さない
 */
export function outsideLabelHeadroomOf(dl: DataLabelsSettings, d: LabelTexts, horizontalBars = false, up = true): number {
    if (horizontalBars || !up || dl.orientation === "vertical" || !labelMayGoOutside(dl) || !d.labelShow || d.blank) return 0;
    const lines = rawLabelLinesOf(dl, d);
    if (!lines.length) return 0;
    const size = dl.singleLine ? maxOf(lines.map((line) => line.font.size)) : lines[0].font.size;
    return outsideTopHeadroomOf(size);
}

/** 合計ラベルと棒の端のすき間 (px) */
export const TOTAL_LABEL_GAP = 4;
/** 合計ラベルの、ベースラインから背景の箱の上端までの高さ（余白の半分を除く。文字の大きさに対する割合） */
export const TOTAL_LABEL_ASCENT = 0.85;
/**
 * 合計ラベルの字がベースラインより下りる深さ（文字の大きさに対する割合）。括弧や g・p・y は、背景の箱の下の端（1 - TOTAL_LABEL_ASCENT）
 * より深く下りる（0.21 くらい）。データ ラベルの見積もり（unitUtils の DESCENT_RATIO）と同じ
 */
export const TOTAL_LABEL_DESCENT = 0.22;

/**
 * 上へ置く合計ラベルの、棒の端からベースラインまでの高さ (px)。ふつうは TOTAL_LABEL_GAP。文字が大きいと、ベースラインより下
 * （背景の下の余白と、カンマや括弧が下りる所）が棒に入るので、深い方の分だけ上げて、背景も文字も棒の端より上に置く
 * （標準と同じ。棒の中のデータ ラベルに掛からない）。下へ置くときは、背景の上の端がもとから棒の端より下にある
 */
export function totalLabelRiseOf(fontPx: number): number {
    return Math.max(TOTAL_LABEL_GAP, fontPx * (1 - TOTAL_LABEL_ASCENT) + LABEL_PADDING / 2, fontPx * TOTAL_LABEL_DESCENT);
}

/** 合計ラベルの置き場所 1 つ。kind は出す合計（正の合計・負の合計・正と負を足した合計）、up は正の棒の端（値の大きい側）に置くか */
export interface TotalLabelSpot {
    kind: "positive" | "negative" | "net";
    up: boolean;
}

/**
 * 合計ラベルの置き場所。「正と負の分割」なら、正の合計を正の棒の端に、負の合計を負の棒の端に分けて出す。
 * 分割しなければ、正と負を足した合計を 1 つ（正なら正の棒の端、負なら負の棒の端）。描く側（App）と値の軸の余白（viewModel）で同じものを使う
 */
export function totalLabelSpotsOf(positive: number, negative: number, split: boolean): TotalLabelSpot[] {
    const hasPositive = positive > 0;
    const hasNegative = negative < 0;
    const spots: TotalLabelSpot[] = [];
    if (split) {
        if (hasPositive) spots.push({ kind: "positive", up: true });
        if (hasNegative) spots.push({ kind: "negative", up: false });
    } else if (hasPositive || hasNegative) {
        spots.push({ kind: "net", up: positive + negative >= 0 ? hasPositive : !hasNegative });
    }
    return spots;
}

/**
 * 合計ラベルを棒の端の外に置いたときの、値の向きの長さ (px)。棒の端から背景の箱の外の端まで（背景が無くても同じ箱で測る。
 * App の renderTotalLabels・renderHTotals の箱と同じ）。縦棒は、上へ置くとベースラインの高さ（totalLabelRiseOf）＋ベースラインより上＋余白の半分、
 * 下へ置くとすき間＋文字の高さ＋余白の半分（文字の幅を測らない）。横棒は、すき間＋文字の幅＋余白
 */
export function totalLabelLengthOf(text: string, font: FontSpec, placement: { horizontal: boolean; screenUp: boolean }): number {
    if (placement.horizontal) return TOTAL_LABEL_GAP + measureTextWidth(text, font) + LABEL_PADDING;
    if (placement.screenUp) return totalLabelRiseOf(font.size) + font.size * TOTAL_LABEL_ASCENT + LABEL_PADDING / 2;
    return TOTAL_LABEL_GAP + font.size + LABEL_PADDING / 2;
}
