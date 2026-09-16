"use strict";

import { FontSpec, measureTextWidth } from "./unitUtils";

/**
 * 凡例の配置。グラフの外側（上下左右のどれか）に置き、グラフに使えない幅を reserve で返す。
 * 座標はすべて凡例の枠 (box) の左上が原点。
 */

export type LegendSide = "top" | "bottom" | "left" | "right";
export type LegendAlign = "start" | "center" | "end";

/** "topLeft" のような位置の値を、置く辺と寄せ方に分ける。知らない値は上詰め (左) */
export function legendPlacement(position: string): { side: LegendSide; align: LegendAlign } {
    const m = /^(top|bottom|left|right)(Left|Center|Right|Top|Bottom)$/.exec(position);
    if (!m) return { side: "top", align: "start" };
    const align: LegendAlign = m[2] === "Center" ? "center" : m[2] === "Left" || m[2] === "Top" ? "start" : "end";
    return { side: m[1] as LegendSide, align };
}

export interface LegendEntry {
    name: string;
    color: string;
    /** 棒は四角、折れ線は短い線（とマーカー）の印を描く。印の幅が違う。省略すると棒 */
    kind?: "bar" | "line";
}

export interface LegendItemBox {
    /** 系列の番号（entries の添字） */
    index: number;
    /** 入りきらなければ末尾を「…」にした名前 */
    text: string;
    name: string;
    color: string;
    /** 印の左端 */
    x: number;
    /** 行の中央 */
    y: number;
    /** 印の幅（棒の四角は markerRadius × 2、折れ線は長め。横棒の折れ線は縦の線なので四角と同じ） */
    glyphWidth: number;
    /** 印と文字を合わせた幅 */
    width: number;
}

export interface LegendLayout {
    side: LegendSide;
    box: { x: number; y: number; width: number; height: number };
    title: { text: string; x: number; y: number } | null;
    items: LegendItemBox[];
    reserve: { top: number; bottom: number; left: number; right: number };
    markerRadius: number;
    rowHeight: number;
}

const PAD = 4;
const ITEM_GAP = 12;
/** 印と名前のあいだ。描画側も同じ値で文字を置く */
export const LEGEND_MARKER_GAP = 4;
/** 折れ線の印（短い線）の長さ。文字サイズに対する比 */
const LINE_GLYPH_RATIO = 1.6;
const MARKER_GAP = LEGEND_MARKER_GAP;
const TITLE_GAP = 8;
/** 左右に置くとき、凡例が取ってよい幅の上限（ビジュアルの幅に対する比） */
const SIDE_MAX_RATIO = 0.3;

function fitText(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const ellipsis = "…";
    for (let n = text.length - 1; n > 0; n--) {
        const cut = text.slice(0, n) + ellipsis;
        if (measureTextWidth(cut, font) <= maxWidth) return cut;
    }
    return text.charAt(0) + ellipsis;
}

export function layoutLegend(spec: {
    entries: LegendEntry[];
    /** 空ならタイトルなし */
    title: string;
    font: FontSpec;
    position: string;
    width: number;
    height: number;
    /** 横棒のとき true。折れ線の印を縦の短い線にするので、印の幅を四角と同じにする */
    horizontal?: boolean;
}): LegendLayout | null {
    const { entries, font, width, height } = spec;
    if (!entries.length || width <= 0 || height <= 0) return null;

    const { side, align } = legendPlacement(spec.position);
    const titleFont: FontSpec = { ...font, bold: true };
    const rowHeight = Math.ceil(font.size * 1.5);
    const markerRadius = Math.max(3, font.size * 0.32);
    const glyphWidthOf = (entry: LegendEntry) =>
        entry.kind === "line" && !spec.horizontal ? Math.max(markerRadius * 2, Math.round(font.size * LINE_GLYPH_RATIO)) : markerRadius * 2;
    const markerWidthOf = (entry: LegendEntry) => glyphWidthOf(entry) + MARKER_GAP;
    const reserve = { top: 0, bottom: 0, left: 0, right: 0 };

    if (side === "top" || side === "bottom") {
        // 横に並べ、入りきらなければ次の行へ折り返す
        const maxRowWidth = Math.max(40, width - PAD * 2);
        const titleText = spec.title ? fitText(spec.title, maxRowWidth / 2, titleFont) : "";
        const titleWidth = titleText ? measureTextWidth(titleText, titleFont) + TITLE_GAP : 0;

        interface Placed { index: number; text: string; width: number }
        const rows: { items: Placed[]; width: number }[] = [];
        let row = { items: [] as Placed[], width: titleWidth };
        entries.forEach((entry, index) => {
            const text = fitText(entry.name, maxRowWidth - markerWidthOf(entry), font);
            const itemWidth = markerWidthOf(entry) + measureTextWidth(text, font);
            const need = (row.items.length ? ITEM_GAP : 0) + itemWidth;
            if (row.items.length > 0 && row.width + need > maxRowWidth) {
                rows.push(row);
                row = { items: [], width: 0 };
                row.items.push({ index, text, width: itemWidth });
                row.width = itemWidth;
            } else {
                row.items.push({ index, text, width: itemWidth });
                row.width += need;
            }
        });
        rows.push(row);

        const boxHeight = rows.length * rowHeight + PAD * 2;
        const box = { x: 0, y: side === "top" ? 0 : height - boxHeight, width, height: boxHeight };
        const rowStart = (rowWidth: number) =>
            align === "start" ? PAD : align === "center" ? (width - rowWidth) / 2 : width - PAD - rowWidth;

        const items: LegendItemBox[] = [];
        let title: LegendLayout["title"] = null;
        rows.forEach((r, ri) => {
            const y = PAD + ri * rowHeight + rowHeight / 2;
            let x = rowStart(r.width);
            if (ri === 0 && titleText) {
                title = { text: titleText, x, y };
                x += titleWidth;
            }
            r.items.forEach((p, k) => {
                if (k > 0) x += ITEM_GAP;
                const entry = entries[p.index];
                items.push({ index: p.index, text: p.text, name: entry.name, color: entry.color, x, y, glyphWidth: glyphWidthOf(entry), width: p.width });
                x += p.width;
            });
        });

        if (side === "top") reserve.top = boxHeight;
        else reserve.bottom = boxHeight;
        return { side, box, title, items, reserve, markerRadius, rowHeight };
    }

    // 左右に置くときは縦に並べる。幅はビジュアルの 3 割まで
    const maxColumnWidth = Math.max(40, width * SIDE_MAX_RATIO) - PAD * 2;
    const titleText = spec.title ? fitText(spec.title, maxColumnWidth, titleFont) : "";
    const texts = entries.map((e) => fitText(e.name, maxColumnWidth - markerWidthOf(e), font));
    const contentWidth = Math.max(
        titleText ? measureTextWidth(titleText, titleFont) : 0,
        ...texts.map((t, i) => markerWidthOf(entries[i]) + measureTextWidth(t, font))
    );
    const boxWidth = Math.min(maxColumnWidth, contentWidth) + PAD * 2;
    const lines = entries.length + (titleText ? 1 : 0);
    const contentHeight = lines * rowHeight;
    const top = align === "start" ? PAD : align === "center" ? (height - contentHeight) / 2 : height - PAD - contentHeight;
    const y0 = Math.max(PAD, top);
    const box = { x: side === "left" ? 0 : width - boxWidth, y: 0, width: boxWidth, height };

    const title = titleText ? { text: titleText, x: PAD, y: y0 + rowHeight / 2 } : null;
    const first = titleText ? 1 : 0;
    const items = entries.map((entry, index) => ({
        index,
        text: texts[index],
        name: entry.name,
        color: entry.color,
        x: PAD,
        y: y0 + (first + index) * rowHeight + rowHeight / 2,
        glyphWidth: glyphWidthOf(entry),
        width: markerWidthOf(entry) + measureTextWidth(texts[index], font),
    }));

    if (side === "left") reserve.left = boxWidth + PAD;
    else reserve.right = boxWidth + PAD;
    return { side, box, title, items, reserve, markerRadius, rowHeight };
}
