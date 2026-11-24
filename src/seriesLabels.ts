/**
 * 系列ラベル（線の端の外の列に線の名前を置く。標準の折れ線・複合グラフの「系列ラベル」）の幅と置き場所。
 * 2026-10-26 の Desktop で見た標準の置き方：列はプロットと値の軸のあいだ（第 2 Y 軸があれば軸の内側）、
 * ラベルは線の端の高さに置き、近いものは上下にずらして重ねず、引き出し線を折って線の端とつなぐ
 */
import type { SeriesLabelValues } from "./settings";
import { FontSpec, PT_TO_PX, graphemesOf, measureTextWidth, truncateToWidth } from "./shared/text";
import { maxOf } from "./shared/minMax";

/** ラベルの字と背景の枠のあいだ（左右・上下 px）。標準の地は余白が大きめ（2026-10-26 の Desktop） */
export const SERIES_LABEL_PAD_X = 6;
export const SERIES_LABEL_PAD_Y = 3;
/** 背景の地の角の丸み（px） */
export const SERIES_LABEL_RADIUS = 3;
/** プロットの端と列のあいだ（px。引き出し線を折る場所。標準より列が右に寄って見えたので 1.36.15.6 で 8 から詰めた） */
export const SERIES_LABEL_GAP = 4;
/** 列と値の軸の目盛りのあいだ（px） */
export const SERIES_LABEL_AXIS_GAP = 6;
/** 上下にずらしたラベルどうしのあいだ（px。地が接しないように） */
export const SERIES_LABEL_STACK_GAP = 2;
/** 引き出し線を線の端から離して始める長さと、斜めに折る部分の横の長さ（px） */
export const SERIES_LABEL_LEADER_START = 3;
export const SERIES_LABEL_LEADER_BEND = 6;
/** 幅が「自動」のときの列の上限（ビジュアルの幅に対する割合） */
export const SERIES_LABEL_AUTO_MAX = 0.25;
/** 行の高さ（文字の大きさに対する倍率） */
const LINE_HEIGHT = 1.25;
const ELLIPSIS = "…";

export interface SeriesLabelInput {
    key: string;
    text: string;
    values: SeriesLabelValues;
    /** 線の端の点の高さ（px） */
    anchorY: number;
}

export interface PlacedSeriesLabel {
    key: string;
    lines: string[];
    values: SeriesLabelValues;
    font: FontSpec;
    anchorY: number;
    /** 枠の上の端と高さ・幅（px）。左右の位置は描く側で列に合わせる */
    top: number;
    height: number;
    width: number;
}

export function seriesLabelFontOf(values: SeriesLabelValues): FontSpec {
    return { family: values.fontFamily, size: values.fontSize * PT_TO_PX, bold: values.bold, italic: values.italic, underline: values.underline };
}

/**
 * 列の幅（px。字の幅と左右の枠の余白を含み、プロットとのあいだは含まない）。fixed（ビジュアルの幅に対する割合）があればその幅、
 * 無ければいちばん長い名前に合わせ、ビジュアルの幅の SERIES_LABEL_AUTO_MAX までにする
 */
export function seriesLabelColumnWidth(inputs: SeriesLabelInput[], fixed: number | null, visualWidth: number): number {
    if (!inputs.length) return 0;
    if (fixed !== null) return Math.max(2 * SERIES_LABEL_PAD_X + 1, fixed * visualWidth);
    const need = maxOf(inputs.map((input) => measureTextWidth(input.text, seriesLabelFontOf(input.values)))) + 2 * SERIES_LABEL_PAD_X;
    return Math.min(need, Math.max(2 * SERIES_LABEL_PAD_X + 1, visualWidth * SERIES_LABEL_AUTO_MAX));
}

/** 幅に入るように行に分ける。折り返さないなら 1 行で「…」で切る（「…」も入らなければ「…」だけ）。折り返すなら字の切れ目で折る（1 文字も入らない行は 1 文字置く） */
export function seriesLabelLines(text: string, maxWidth: number, font: FontSpec, wrap: boolean): string[] {
    if (measureTextWidth(text, font) <= maxWidth) return [text];
    if (!wrap) {
        // 「…」も入らない幅でもラベルは捨てない（標準は狭い列でも地と引き出し線を残す）
        return [truncateToWidth(text, maxWidth, font) || ELLIPSIS];
    }
    const lines: string[] = [];
    let current = "";
    for (const g of graphemesOf(text)) {
        if (current && measureTextWidth(current + g, font) > maxWidth) {
            lines.push(current);
            current = g.trimStart();
        } else {
            current += g;
        }
    }
    if (current) lines.push(current);
    return lines;
}

/**
 * ラベルを列に置く。線の端の高さを真ん中にし、重なるものは上下に振り分けて描く範囲（top〜bottom）に収める。
 * ずらす量は maxOffset（px。標準の「最大オフセット」）までにとどめ、それでも重なるものは重ねたまま出す。
 * 標準は列が狭くて名前が入りきらないときもラベルを捨てない（2026-10-26 の Desktop。1.36.15.5 では超えたものを捨てて、幅 5% で消えた）
 */
export function placeSeriesLabels(
    inputs: SeriesLabelInput[],
    columnWidth: number,
    span: [number, number],
    maxOffset: number
): PlacedSeriesLabel[] {
    const textWidth = Math.max(1, columnWidth - 2 * SERIES_LABEL_PAD_X);
    const items = inputs
        .map((input) => {
            const font = seriesLabelFontOf(input.values);
            const lines = seriesLabelLines(input.text, textWidth, font, input.values.wrap);
            const height = lines.length * font.size * LINE_HEIGHT + 2 * SERIES_LABEL_PAD_Y;
            const width = Math.min(columnWidth, Math.max(0, maxOf(lines.map((line) => measureTextWidth(line, font)))) + 2 * SERIES_LABEL_PAD_X);
            return { key: input.key, lines, values: input.values, font, anchorY: input.anchorY, height, width };
        })
        .sort((a, b) => a.anchorY - b.anchorY);
    // 枠の上下に STACK_GAP の半分ずつの空きを足して積み、枠はその内側に置く
    const half = SERIES_LABEL_STACK_GAP / 2;
    const tops = stackWithin(items.map((item) => ({ center: item.anchorY, height: item.height + SERIES_LABEL_STACK_GAP })), span);
    return items.map((item, i) => {
        const shifted = tops[i] + half + item.height / 2 - item.anchorY;
        const shift = Math.max(-maxOffset, Math.min(maxOffset, shifted));
        return { ...item, top: item.anchorY + shift - item.height / 2 };
    });
}

/**
 * 高さ順に並べた枠を、重ねずに範囲に収める上の端（並びは入力と同じ）。重なる枠はひとかたまりにし、かたまりの真ん中を
 * 線の端の高さの平均に置く（標準と同じく上下に振り分ける）。範囲の端を超えるかたまりは内側へ押し戻す。
 * 全部の高さが範囲より高ければ、上の端からそのまま積む
 */
function stackWithin(boxes: Array<{ center: number; height: number }>, [top, bottom]: [number, number]): number[] {
    type Cluster = { first: number; last: number; height: number; top: number };
    const placeOf = (first: number, last: number, height: number): number => {
        // 各枠の真ん中と線の端の高さの差の 2 乗の和が最小になる上の端（かたまりの中の位置の分を引いた平均）
        let offset = 0;
        let sum = 0;
        for (let i = first; i <= last; i++) {
            sum += boxes[i].center - (offset + boxes[i].height / 2);
            offset += boxes[i].height;
        }
        const ideal = sum / (last - first + 1);
        return Math.max(top, Math.min(bottom - height, ideal));
    };
    const clusters: Cluster[] = [];
    boxes.forEach((box, i) => {
        let cluster: Cluster = { first: i, last: i, height: box.height, top: placeOf(i, i, box.height) };
        for (;;) {
            const prev = clusters[clusters.length - 1];
            if (!prev || prev.top + prev.height <= cluster.top + 1e-9) break;
            clusters.pop();
            const height = prev.height + cluster.height;
            cluster = { first: prev.first, last: cluster.last, height, top: placeOf(prev.first, cluster.last, height) };
        }
        clusters.push(cluster);
    });
    const tops: number[] = [];
    for (const cluster of clusters) {
        // 範囲より高いかたまりは上の端から積む
        let y = cluster.height > bottom - top ? top : cluster.top;
        for (let i = cluster.first; i <= cluster.last; i++) {
            tops.push(y);
            y += boxes[i].height;
        }
    }
    return tops;
}
