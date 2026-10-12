"use strict";

import { FontSpec, measureTextWidth } from "./unitUtils";
import { LegendSide, LegendAlign, legendPlacement } from "./shared/legend";
import type { LegendArrow } from "./shared/LegendArrowButton";

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
     * まとまり（省略すると 0）。まとまりが変わる所に groupTitles の見出しを付ける（見出しが無ければ間を広げる）。
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
    /** 入りきらないときのページの数と、いま出しているページ（0 から）。入りきれば 1 と 0 */
    pageCount: number;
    page: number;
    /** ページを送る矢印（標準と同じ。上下に置けば左右、左右に置けば上下）。入りきれば無し */
    arrows: { prev: LegendArrow; next: LegendArrow } | null;
}


const PAD = 4;
const ITEM_GAP = 12;
/** 印と名前のあいだ。描画側も同じ値で文字を置く */
export const LEGEND_MARKER_GAP = 4;
/** 折れ線の印（短い線）の長さ。文字サイズに対する比 */
const LINE_GLYPH_RATIO = 1.6;
const MARKER_GAP = LEGEND_MARKER_GAP;
const TITLE_GAP = 8;
/** ページを送る矢印と項目のあいだ */
const ARROW_GAP = 6;
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
    /** 出すページ（入りきらないとき）。範囲の外なら端のページ */
    page?: number;
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
        // 標準と同じく 1 行に並べ、入りきらなければページに分けて両端の矢印で送る。タイトルはどのページでも左に出す
        const maxRowWidth = Math.max(40, width - PAD * 2);
        const titleText = spec.title ? fitText(spec.title, maxRowWidth / 2, titleFont) : "";
        const titleWidth = titleText ? measureTextWidth(titleText, titleFont) + TITLE_GAP : 0;
        const arrowSize = rowHeight;
        const arrowRoom = arrowSize + ARROW_GAP;

        // 並べる物：まとまりの見出し（あれば）と項目。まとまりが変わる所は見出しか少し広い間を空ける
        interface Token { index: number | null; text: string; width: number; title: boolean; groupStart: boolean }
        const tokensFor = (room: number): Token[] => {
            const tokens: Token[] = [];
            entries.forEach((entry, index) => {
                if (startsGroup(index)) {
                    // 見出しが無ければ、文字の無い見出しで間を広げる（項目の間の 2 倍）
                    const groupTitle = fitText(groupTitleOf(index), room / 2, titleFont);
                    const titleWidthHere = groupTitle ? measureTextWidth(groupTitle, titleFont) + TITLE_GAP - ITEM_GAP : 0;
                    tokens.push({ index: null, text: groupTitle, width: titleWidthHere, title: true, groupStart: true });
                }
                const text = fitText(entry.name, room - markerWidthOf(entry), font);
                tokens.push({ index, text, width: markerWidthOf(entry) + measureTextWidth(text, font), title: false, groupStart: startsGroup(index) });
            });
            return tokens;
        };
        // 見出しの幅は「見出し＋TITLE_GAP − ITEM_GAP」にしてあるので、どの間も ITEM_GAP で足せる
        const widthOf = (tokens: Token[]) => tokens.reduce((sum, t, k) => sum + (k > 0 ? ITEM_GAP : 0) + t.width, 0);
        let tokens = tokensFor(maxRowWidth - titleWidth);
        let pages: Token[][] = [tokens];
        if (titleWidth + widthOf(tokens) > maxRowWidth) {
            const room = Math.max(20, maxRowWidth - titleWidth - arrowRoom * 2);
            tokens = tokensFor(room);
            pages = [];
            let current: Token[] = [];
            for (const t of tokens) {
                if (current.length && widthOf([...current, t]) > room) {
                    pages.push(current);
                    current = [];
                }
                current.push(t);
            }
            if (current.length) pages.push(current);
            // まとまりの見出しがページの最後に残ったら、次のページの頭へ送る（見出しだけが前のページに取り残されない）
            for (let k = 0; k < pages.length - 1; k++) {
                const last = pages[k][pages[k].length - 1];
                if (pages[k].length > 1 && last.title) pages[k + 1].unshift(pages[k].pop()!);
            }
        }
        const pageCount = pages.length;
        const page = Math.max(0, Math.min(pageCount - 1, spec.page ?? 0));
        const shown = pages[page];
        const paged = pageCount > 1;

        const boxHeight = rowHeight + PAD * 2;
        const box = { x: 0, y: side === "top" ? 0 : height - boxHeight, width, height: boxHeight };
        const y = PAD + rowHeight / 2;
        // 矢印の位置はどのページでも同じにする（いちばん広いページの幅で並べる。押した所に次の矢印がある）
        const pageWidth = Math.max(...pages.map(widthOf));
        const contentWidth = titleWidth + (paged ? pageWidth + arrowRoom * 2 : widthOf(shown));
        const rowStart = align === "start" ? PAD : align === "center" ? (width - contentWidth) / 2 : width - PAD - contentWidth;

        const items: LegendItemBox[] = [];
        const groupTitles: LegendLayout["groupTitles"] = [];
        let x = rowStart;
        const title: LegendLayout["title"] = titleText ? { text: titleText, x, y } : null;
        x += titleWidth;
        let arrows: LegendLayout["arrows"] = null;
        if (paged) {
            arrows = {
                prev: { x: x + arrowSize / 2, y, size: arrowSize, direction: "left", enabled: page > 0 },
                next: { x: 0, y, size: arrowSize, direction: "right", enabled: page < pageCount - 1 },
            };
            x += arrowRoom;
        }
        shown.forEach((t, k) => {
            if (k > 0) x += ITEM_GAP;
            if (t.title || t.index === null) {
                if (t.text) groupTitles.push({ text: t.text, x, y });
            } else {
                const entry = entries[t.index];
                items.push({ index: t.index, text: t.text, name: entry.name, color: entry.color, x, y, glyphWidth: glyphWidthOf(entry), width: t.width });
            }
            x += t.width;
        });
        if (arrows) arrows.next.x = rowStart + contentWidth - arrowSize / 2;

        if (side === "top") reserve.top = boxHeight;
        else reserve.bottom = boxHeight;
        return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight, pageCount, page, arrows };
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
    const box = { x: side === "left" ? 0 : width - boxWidth, y: 0, width: boxWidth, height };

    // 行：まとまりの見出し（見出しが無ければ半行の間）と項目。タイトルはどのページでも上に出す
    interface Line { index: number | null; text: string; height: number }
    const lineList: Line[] = [];
    entries.forEach((_, index) => {
        const groupTitle = breaks[index];
        if (groupTitle) lineList.push({ index: null, text: groupTitle, height: 1 });
        else if (groupTitle === "") lineList.push({ index: null, text: "", height: 0.5 });
        lineList.push({ index, text: texts[index], height: 1 });
    });
    const available = Math.max(rowHeight, height - PAD * 2);
    const fitLines = Math.floor(available / rowHeight);
    const totalLinesOf = (title: number) => title + lineList.reduce((sum, l) => sum + l.height, 0);
    // 入りきらなければ、上下の矢印の行を取ってページに分ける（標準と同じ）。矢印 2 行と項目 1 行も入らないほど低ければ
    // タイトルを外し、それでも入らなければ矢印を出さずに入るだけ出す（矢印が描く範囲の外に出て押せなくならないように）
    let titleLines = titleText ? 1 : 0;
    let paged = totalLinesOf(titleLines) * rowHeight > available;
    if (paged && fitLines < titleLines + 3) titleLines = 0;
    if (paged && fitLines < 3) paged = false;
    let pages: Line[][] = [lineList];
    if (paged) {
        const perPage = Math.max(1, fitLines - titleLines - 2);
        pages = [];
        let current: Line[] = [];
        let used = 0;
        for (const l of lineList) {
            if (current.length && used + l.height > perPage) {
                pages.push(current);
                current = [];
                used = 0;
            }
            // ページの頭に半行の間は置かない
            if (!current.length && l.index === null && !l.text) continue;
            current.push(l);
            used += l.height;
        }
        if (current.length) pages.push(current);
        // まとまりの見出しがページの最後に残ったら、次のページの頭へ送る
        for (let k = 0; k < pages.length - 1; k++) {
            const last = pages[k][pages[k].length - 1];
            if (pages[k].length > 1 && last.index === null && last.text) pages[k + 1].unshift(pages[k].pop()!);
        }
    }
    const pageCount = pages.length;
    const page = Math.max(0, Math.min(pageCount - 1, spec.page ?? 0));
    const shownAll = pages[page];
    // 矢印を出さないほど低いときは、入る行だけ出す
    const shown = paged || totalLinesOf(titleLines) <= fitLines ? shownAll : shownAll.slice(0, Math.max(1, fitLines - titleLines));
    // 矢印の位置はどのページでも同じにする（いちばん行の多いページで並べる）
    const pageLines = Math.max(...pages.map((pg) => pg.reduce((sum, l) => sum + l.height, 0)));
    const shownLines = titleLines + (paged ? pageLines + 2 : shown.reduce((sum, l) => sum + l.height, 0));
    const contentHeight = shownLines * rowHeight;
    const top = align === "start" ? PAD : align === "center" ? (height - contentHeight) / 2 : height - PAD - contentHeight;
    const y0 = Math.max(PAD, top);

    const title = titleText && titleLines ? { text: titleText, x: PAD, y: y0 + rowHeight / 2 } : null;
    const groupTitles: LegendLayout["groupTitles"] = [];
    const arrowSize = rowHeight * 0.8;
    let line = titleLines;
    let arrows: LegendLayout["arrows"] = null;
    if (paged) {
        arrows = {
            prev: { x: PAD + arrowSize / 2, y: y0 + line * rowHeight + rowHeight / 2, size: arrowSize, direction: "up", enabled: page > 0 },
            next: { x: PAD + arrowSize / 2, y: 0, size: arrowSize, direction: "down", enabled: page < pageCount - 1 },
        };
        line += 1;
    }
    const items: LegendItemBox[] = [];
    for (const l of shown) {
        if (l.index === null) {
            if (l.text) groupTitles.push({ text: l.text, x: PAD, y: y0 + line * rowHeight + rowHeight / 2 });
            line += l.height;
            continue;
        }
        const entry = entries[l.index];
        items.push({
            index: l.index,
            text: l.text,
            name: entry.name,
            color: entry.color,
            x: PAD,
            y: y0 + line * rowHeight + rowHeight / 2,
            glyphWidth: glyphWidthOf(entry),
            width: markerWidthOf(entry) + measureTextWidth(l.text, font),
        });
        line += 1;
    }
    if (arrows) arrows.next.y = y0 + (titleLines + 1 + pageLines) * rowHeight + rowHeight / 2;

    if (side === "left") reserve.left = boxWidth + PAD;
    else reserve.right = boxWidth + PAD;
    return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight, pageCount, page, arrows };
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
