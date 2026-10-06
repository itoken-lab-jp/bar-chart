"use strict";

import { FontSpec, measureTextWidth } from "./unitUtils";
import { LegendSide, LegendAlign, legendPlacement } from "./shared/legend";

/**
 * 凡例の配置。グラフの外側（上下左右のどれか）に置き、グラフに使えない幅を reserve で返す。
 * 座標はすべて凡例の枠 (box) の左上が原点。
 */

export type { LegendSide, LegendAlign };
export { legendPlacement };

export interface LegendEntry {
    name: string;
    color: string;
    /** 棒は四角、折れ線は短い線（とマーカー）、比較レイヤーはずらして重ねた四角の印を描く。印の幅が違う。省略すると棒 */
    kind?: "bar" | "line" | "layer";
    /**
     * まとまり（省略すると 0）。まとまりが変わる所で行を改め（左右に置くときは間を空け）、groupTitles の見出しを付ける。
     * 比較レイヤーを系列と別のまとまりにして、どこからがレイヤーかを読めるようにする
     */
    group?: number;
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
    /** まとまりの見出し（groupTitles）。無ければ空 */
    groupTitles: Array<{ text: string; x: number; y: number }>;
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
    /** まとまりの見出し（キーは group）。0 のまとまりの見出しは title */
    groupTitles?: Record<number, string>;
}): LegendLayout | null {
    const { entries, font, width, height } = spec;
    const groupOf = (index: number) => entries[index].group ?? 0;
    /** index の項目でまとまりが変わるか（先頭は変わらない） */
    const startsGroup = (index: number) => index > 0 && groupOf(index) !== groupOf(index - 1);
    const groupTitleOf = (index: number) => spec.groupTitles?.[groupOf(index)] ?? "";
    if (!entries.length || width <= 0 || height <= 0) return null;

    const { side, align } = legendPlacement(spec.position);
    const titleFont: FontSpec = { ...font, bold: true };
    const rowHeight = Math.ceil(font.size * 1.5);
    const markerRadius = Math.max(3, font.size * 0.32);
    const glyphWidthOf = (entry: LegendEntry) =>
        entry.kind === "layer"
            ? layerGlyphWidth(markerRadius, spec.horizontal)
            : entry.kind === "line" && !spec.horizontal ? Math.max(markerRadius * 2, Math.round(font.size * LINE_GLYPH_RATIO)) : markerRadius * 2;
    const markerWidthOf = (entry: LegendEntry) => glyphWidthOf(entry) + MARKER_GAP;
    const reserve = { top: 0, bottom: 0, left: 0, right: 0 };

    if (side === "top" || side === "bottom") {
        // 横に並べ、入りきらなければ次の行へ折り返す
        const maxRowWidth = Math.max(40, width - PAD * 2);
        const titleText = spec.title ? fitText(spec.title, maxRowWidth / 2, titleFont) : "";
        const titleWidth = titleText ? measureTextWidth(titleText, titleFont) + TITLE_GAP : 0;

        interface Placed { index: number; text: string; width: number }
        interface Row { items: Placed[]; width: number; title: string; titleWidth: number }
        const rows: Row[] = [];
        let row: Row = { items: [], width: titleWidth, title: titleText, titleWidth };
        entries.forEach((entry, index) => {
            const text = fitText(entry.name, maxRowWidth - markerWidthOf(entry), font);
            const itemWidth = markerWidthOf(entry) + measureTextWidth(text, font);
            // まとまりが変わったら行を改め、その行の頭に見出しを置く
            if (startsGroup(index)) {
                rows.push(row);
                const groupTitle = fitText(groupTitleOf(index), maxRowWidth / 2, titleFont);
                const groupTitleWidth = groupTitle ? measureTextWidth(groupTitle, titleFont) + TITLE_GAP : 0;
                row = { items: [], width: groupTitleWidth, title: groupTitle, titleWidth: groupTitleWidth };
            }
            const need = (row.items.length ? ITEM_GAP : 0) + itemWidth;
            if (row.items.length > 0 && row.width + need > maxRowWidth) {
                rows.push(row);
                row = { items: [], width: 0, title: "", titleWidth: 0 };
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
        const groupTitles: LegendLayout["groupTitles"] = [];
        rows.forEach((r, ri) => {
            const y = PAD + ri * rowHeight + rowHeight / 2;
            let x = rowStart(r.width);
            if (r.title) {
                if (ri === 0) title = { text: r.title, x, y };
                else groupTitles.push({ text: r.title, x, y });
                x += r.titleWidth;
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
        return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight };
    }

    // 左右に置くときは縦に並べる。幅はビジュアルの 3 割まで
    const maxColumnWidth = Math.max(40, width * SIDE_MAX_RATIO) - PAD * 2;
    const titleText = spec.title ? fitText(spec.title, maxColumnWidth, titleFont) : "";
    const texts = entries.map((e) => fitText(e.name, maxColumnWidth - markerWidthOf(e), font));
    // まとまりが変わる所に、見出しの行（見出しが無ければ半行の間）を足す
    const breaks = entries.map((_, i) => (startsGroup(i) ? fitText(groupTitleOf(i), maxColumnWidth, titleFont) : null));
    const contentWidth = Math.max(
        titleText ? measureTextWidth(titleText, titleFont) : 0,
        ...breaks.map((t) => (t ? measureTextWidth(t, titleFont) : 0)),
        ...texts.map((t, i) => markerWidthOf(entries[i]) + measureTextWidth(t, font))
    );
    const boxWidth = Math.min(maxColumnWidth, contentWidth) + PAD * 2;
    const breakLines = breaks.reduce((sum, t) => sum + (t === null ? 0 : t ? 1 : 0.5), 0);
    const lines = entries.length + (titleText ? 1 : 0) + breakLines;
    const contentHeight = lines * rowHeight;
    const top = align === "start" ? PAD : align === "center" ? (height - contentHeight) / 2 : height - PAD - contentHeight;
    const y0 = Math.max(PAD, top);
    const box = { x: side === "left" ? 0 : width - boxWidth, y: 0, width: boxWidth, height };

    const title = titleText ? { text: titleText, x: PAD, y: y0 + rowHeight / 2 } : null;
    const groupTitles: LegendLayout["groupTitles"] = [];
    let line = titleText ? 1 : 0;
    const items = entries.map((entry, index) => {
        const groupTitle = breaks[index];
        if (groupTitle) {
            groupTitles.push({ text: groupTitle, x: PAD, y: y0 + line * rowHeight + rowHeight / 2 });
            line += 1;
        } else if (groupTitle === "") {
            line += 0.5;
        }
        const y = y0 + line * rowHeight + rowHeight / 2;
        line += 1;
        return {
            index,
            text: texts[index],
            name: entry.name,
            color: entry.color,
            x: PAD,
            y,
            glyphWidth: glyphWidthOf(entry),
            width: markerWidthOf(entry) + measureTextWidth(texts[index], font),
        };
    });

    if (side === "left") reserve.left = boxWidth + PAD;
    else reserve.right = boxWidth + PAD;
    return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight };
}

/**
 * 比較レイヤーの印（縦長の棒）の幅と高さ。棒の系列の四角と形で見分けられるよう、幅は四角の半分、高さは四角の 1.3 倍。
 * 横棒では横長にする（棒の向きに合わせる）
 */
export const LAYER_GLYPH_THICKNESS = 1;
export const LAYER_GLYPH_LENGTH = 2.6;
export function layerGlyphWidth(markerRadius: number, horizontal = false): number {
    return markerRadius * (horizontal ? LAYER_GLYPH_LENGTH : LAYER_GLYPH_THICKNESS);
}
