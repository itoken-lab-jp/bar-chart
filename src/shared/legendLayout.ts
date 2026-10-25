"use strict";

import { FontSpec, measureTextWidth } from "./text";
import {
    LEGEND_ARROW_GAP,
    LEGEND_PAD,
    LEGEND_SIDE_ARROW_RATIO,
    LEGEND_SIDE_LIFT,
    LEGEND_SIDE_MAX_RATIO,
    LEGEND_TITLE_GAP,
    LegendSide,
    legendMarkerRadius,
    legendPlacement,
    legendDescentRoom,
    legendItemGap,
    legendMarkerGap,
    legendRowHeight,
    legendSideRowHeight,
    legendSideTitleHeight,
    legendTextOverhang,
} from "./legend";
import { linesByWidth, pagesOfLines, type LegendArrow } from "./LegendArrowButton";
import { maxOf } from "./minMax";

/**
 * 凡例の組み方。グラフの外側（上下左右のどれか）に置く箱の中に、タイトル・項目・ページを送る矢印を並べ、グラフに使えない幅を reserve で返す。
 * 座標はすべて凡例の箱 (box) の左上が原点。行の高さ（左右に置いたときは行の間隔とタイトルの行の高さ）・余白・字の位置の決まりは legend.ts。
 * 印の形と幅は作品の側で決めて渡す（四角の一辺は legendMarkerRadius の 2 倍、横に引く線は legendLineGlyphWidth）
 */

export interface LegendEntry {
    name: string;
    /**
     * 印の幅 (px)。描いた印の見える左右の端まで（線の模様と端の形、マーカーの見える幅で決まる。決まった印の幅より広くも狭くもなる。
     * 線の印は legend.ts の lineInkSpan）。項目のあいだ（前の名前 → 次の印の左の端）と、印と名前のあいだ（印の右の端 → 名前）は、
     * この幅の端から取る。作品は、印の見える左の端が項目の左の端（x）に来るように描く
     */
    glyphWidth: number;
    /**
     * まとまり（省略すると 0）。まとまりが変わる所に groupTitles の見出しを付ける（見出しが無ければ間を広げる）。
     * 別の種類の項目を別のまとまりにして、どこからが別の種類かを読めるようにする
     */
    group?: number;
}

export interface LegendItemBox {
    /** 項目の番号（entries の添字） */
    index: number;
    /** 入りきらなければ末尾を「…」にした名前 */
    text: string;
    name: string;
    /** 印の左端 */
    x: number;
    /** 行の縦の真ん中（印の真ん中。字のベースラインは legendBaselineOf） */
    y: number;
    /** 印の幅 */
    glyphWidth: number;
    /** 名前の字の送りの左の端（印の見える右の端から legendMarkerGap） */
    textX: number;
    /** 印と文字を合わせた幅 */
    width: number;
}

export interface LegendLayout<I extends LegendItemBox = LegendItemBox> {
    side: LegendSide;
    box: { x: number; y: number; width: number; height: number };
    /** y は行の縦の真ん中 */
    title: { text: string; x: number; y: number } | null;
    /** まとまりの見出し（groupTitles）。無ければ空。y は行の縦の真ん中 */
    groupTitles: Array<{ text: string; x: number; y: number }>;
    items: I[];
    reserve: { top: number; bottom: number; left: number; right: number };
    markerRadius: number;
    /** 行の高さ。左右に置いたときは行の間隔（タイトルの行の高さは titleHeight） */
    rowHeight: number;
    /** タイトルの行の高さ。上下に置いたときは rowHeight と同じ */
    titleHeight: number;
    /** 入りきらないときのページの数と、いま出しているページ（0 から）。入りきれば 1 と 0 */
    pageCount: number;
    page: number;
    /** ページを送る矢印（標準と同じ。上下に置けば左右、左右に置けば上下）。入りきれば無し */
    arrows: { prev: LegendArrow; next: LegendArrow } | null;
}

const PAD = LEGEND_PAD;
const TITLE_GAP = LEGEND_TITLE_GAP;
const ARROW_GAP = LEGEND_ARROW_GAP;
const SIDE_MAX_RATIO = LEGEND_SIDE_MAX_RATIO;

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
    /** まとまりの見出し（キーは group）。0 のまとまりの見出しは title */
    groupTitles?: Record<number, string>;
    /** 出すページ（入りきらないとき）。範囲の外なら端のページ */
    page?: number;
    /** 上下に置いたときに並べる行の数（省略すると 1）。左右に置いたときは効かない */
    rows?: number;
}): LegendLayout | null {
    const { entries, font, width, height } = spec;
    const groupOf = (index: number) => entries[index].group ?? 0;
    /** index の項目でまとまりが変わるか（先頭は変わらない） */
    const startsGroup = (index: number) => index > 0 && groupOf(index) !== groupOf(index - 1);
    const groupTitleOf = (index: number) => spec.groupTitles?.[groupOf(index)] ?? "";
    if (!entries.length || width <= 0 || height <= 0) return null;

    const { side, align } = legendPlacement(spec.position);
    const titleFont: FontSpec = { ...font, bold: true };
    const markerRadius = legendMarkerRadius(font.size);
    // 項目のあいだ・印と名前のあいだは文字の大きさで決める（legend.ts）
    const ITEM_GAP = legendItemGap(font.size);
    const MARKER_GAP = legendMarkerGap(font.size);
    const glyphWidthOf = (entry: LegendEntry) => entry.glyphWidth;
    const markerWidthOf = (entry: LegendEntry) => glyphWidthOf(entry) + MARKER_GAP;
    const reserve = { top: 0, bottom: 0, left: 0, right: 0 };

    if (side === "top" || side === "bottom") {
        // 標準と同じく 1 行に並べ、入りきらなければページに分けて両端の矢印で送る。タイトルはどのページでも 1 行目の左に出す。
        // 行数が 2 以上なら、入りきらない項目を次の行へ折り返し、その行数でも入らなければ、その行数ずつのページに分けて送る
        const rows = Math.max(1, Math.floor(spec.rows ?? 1));
        const rowHeight = legendRowHeight(font.size);
        const maxRowWidth = Math.max(40, width - PAD * 2);
        const arrowSize = rowHeight;
        const arrowRoom = arrowSize + ARROW_GAP;
        // タイトルは行の半分まで
        let titleText = spec.title ? fitText(spec.title, maxRowWidth / 2, titleFont) : "";
        let titleWidth = titleText ? measureTextWidth(titleText, titleFont) + TITLE_GAP : 0;

        // 並べる物：まとまりの見出し（あれば）と項目。まとまりが変わる所は見出しか少し広い間を空ける
        interface Token { index: number | null; text: string; width: number; title: boolean; groupStart: boolean }
        const tokensFor = (room: number): Token[] => {
            const tokens: Token[] = [];
            entries.forEach((entry, index) => {
                // まとまりの頭の項目は見出しと同じ行に置くので、見出しと合わせて room に入るよう名前を縮める
                let groupRoom = 0;
                if (startsGroup(index)) {
                    // 見出しは room の半分まで、項目の名前の 1 文字と「…」を残す長さまで。それも入らなければ見出しの文字を出さない。
                    // 見出しが無ければ、文字の無い見出しで間を広げる（項目の間の 2 倍）
                    const shortest = Math.min(measureTextWidth(entry.name, font), measureTextWidth(fitText(entry.name, 0, font), font));
                    const titleLimit = Math.min(room / 2, room - TITLE_GAP - markerWidthOf(entry) - shortest);
                    const fitted = fitText(groupTitleOf(index), titleLimit, titleFont);
                    const groupTitle = measureTextWidth(fitted, titleFont) <= titleLimit ? fitted : "";
                    const titleWidthHere = groupTitle ? measureTextWidth(groupTitle, titleFont) + TITLE_GAP - ITEM_GAP : 0;
                    tokens.push({ index: null, text: groupTitle, width: titleWidthHere, title: true, groupStart: true });
                    groupRoom = titleWidthHere + ITEM_GAP;
                }
                const text = fitText(entry.name, room - groupRoom - markerWidthOf(entry), font);
                tokens.push({ index, text, width: markerWidthOf(entry) + measureTextWidth(text, font), title: false, groupStart: startsGroup(index) });
            });
            return tokens;
        };
        // 見出しの幅は「見出し＋TITLE_GAP − ITEM_GAP」にしてあるので、どの間も ITEM_GAP で足せる
        const widthOf = (tokens: Token[]) => tokens.reduce((sum, t, k) => sum + (k > 0 ? ITEM_GAP : 0) + t.width, 0);
        // room に収まる行に分ける。まとまりの見出しは続く項目とひとまとめにして同じ行に置き（見出しだけが行の末に取り残されない）、
        // 行の頭に文字の無い見出し（広げた間）は置かない
        const linesOf = (tokens: Token[], room: number): Token[][] =>
            linesByWidth(
                tokens.map((t) => t.width),
                room,
                ITEM_GAP,
                { joinNext: (i) => tokens[i].title, dropAtStart: (i) => tokens[i].title && !tokens[i].text }
            ).map((line) => line.map((i) => tokens[i]));
        let tokens = tokensFor(maxRowWidth - titleWidth);
        /** ページ → 行 → 並べる物 */
        let pages: Token[][][] = [[tokens]];
        if (titleWidth + widthOf(tokens) > maxRowWidth) {
            // 2 行以上なら、まず矢印なしで折り返して、その行数に収まるかを見る（2 行目以降はタイトルの右の始まりにそろえる）
            const wrapped = rows > 1 ? linesOf(tokens, maxRowWidth - titleWidth) : [];
            if (rows > 1 && wrapped.length <= rows) {
                pages = [wrapped];
            } else {
                // ページに分けるときのタイトルは、両端の矢印を除いた残りの半分まで（項目を置く所を残す）。1 文字と「…」も入らなければ外す
                const titleRoom = (maxRowWidth - arrowRoom * 2) / 2 - TITLE_GAP;
                titleText = spec.title ? fitText(spec.title, titleRoom, titleFont) : "";
                if (measureTextWidth(titleText, titleFont) > titleRoom) titleText = "";
                titleWidth = titleText ? measureTextWidth(titleText, titleFont) + TITLE_GAP : 0;
                const room = Math.max(20, maxRowWidth - titleWidth - arrowRoom * 2);
                tokens = tokensFor(room);
                // 送った先のページも room に収まる。見出しは続く項目と同じページに置く
                pages = pagesOfLines(linesOf(tokens, room), rows);
            }
        }
        const pageCount = pages.length;
        const page = Math.max(0, Math.min(pageCount - 1, spec.page ?? 0));
        const shown = pages[page];
        const paged = pageCount > 1;
        // ページに分けるときは、どのページでも行数ぶんの高さを取る（ページを送っても描く範囲が動かない）。分けないときは使う行だけ
        const lineCount = paged ? rows : shown.length;

        // 文字が大きいと、最後の行の字の下の端が下の PAD を越えるので、越える分だけ箱の下を延ばす（行・字・印の位置は変えない）
        const boxHeight = rowHeight * lineCount + PAD * 2 + legendDescentRoom(font.size);
        const box = { x: 0, y: side === "top" ? 0 : height - boxHeight, width, height: boxHeight };
        const rowY = (line: number) => PAD + rowHeight / 2 + line * rowHeight;
        // 矢印の位置はどのページでも同じにする（いちばん広い行の幅で並べる。押した所に次の矢印がある）。縦は使う行のまとまりの真ん中
        const pageWidth = maxOf(pages.flat().map(widthOf));
        const arrowY = PAD + (rowHeight * lineCount) / 2;

        const items: LegendItemBox[] = [];
        const groupTitles: LegendLayout["groupTitles"] = [];
        let title: LegendLayout["title"] = null;
        let arrows: LegendLayout["arrows"] = null;
        shown.forEach((line, row) => {
            const y = rowY(row);
            // 寄せ方は行ごと。どの行もタイトルの幅を空けてから項目を並べる（2 行目以降の項目は 1 行目の項目の始まりにそろう）。
            // 名前を縮めても入らないほど狭いときも、矢印は枠の中に置く
            const contentWidth = Math.min(maxRowWidth, titleWidth + (paged ? pageWidth + arrowRoom * 2 : widthOf(line)));
            const rowStart = align === "start" ? PAD : align === "center" ? (width - contentWidth) / 2 : width - PAD - contentWidth;
            let x = rowStart;
            if (row === 0 && titleText) title = { text: titleText, x, y };
            x += titleWidth;
            if (paged) {
                if (row === 0)
                    arrows = {
                        prev: { x: x + arrowSize / 2, y: arrowY, size: arrowSize, direction: "left", enabled: page > 0 },
                        next: { x: rowStart + contentWidth - arrowSize / 2, y: arrowY, size: arrowSize, direction: "right", enabled: page < pageCount - 1 },
                    };
                x += arrowRoom;
            }
            line.forEach((t, k) => {
                if (k > 0) x += ITEM_GAP;
                if (t.title || t.index === null) {
                    if (t.text) groupTitles.push({ text: t.text, x, y });
                } else {
                    const entry = entries[t.index];
                    items.push({ index: t.index, text: t.text, name: entry.name, x, y, glyphWidth: glyphWidthOf(entry), textX: x + markerWidthOf(entry), width: t.width });
                }
                x += t.width;
            });
        });

        if (side === "top") reserve.top = boxHeight;
        else reserve.bottom = boxHeight;
        return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight, titleHeight: rowHeight, pageCount, page, arrows };
    }

    // 左右に置くときは縦に並べる。箱の幅の上限はビジュアルの幅の 3 割（LEGEND_SIDE_MAX_RATIO）。3 割が 40px に満たない狭いビジュアルでも、
    // 40px までは取る（箱の中の左右 4px を含む。描く範囲の側にさらに 4px 空けるので、グラフの幅は最大 44px 減る）。
    // 名前・タイトルは「1 文字 ＋ …」まで縮める。それでも列に入らない字は箱からはみ出すので、作品の側で箱の大きさで切り抜いて描く
    const maxColumnWidth = Math.max(40, width * SIDE_MAX_RATIO) - PAD * 2;
    const titleText = spec.title ? fitText(spec.title, maxColumnWidth, titleFont) : "";
    const texts = entries.map((e) => fitText(e.name, maxColumnWidth - markerWidthOf(e), font));
    // まとまりが変わる所に、見出しの行（見出しが無ければ半行の間）を足す
    const breaks = entries.map((_, i) => (startsGroup(i) ? fitText(groupTitleOf(i), maxColumnWidth, titleFont) : null));
    const contentWidth = Math.max(
        titleText ? measureTextWidth(titleText, titleFont) : 0,
        maxOf(breaks.map((t) => (t ? measureTextWidth(t, titleFont) : 0))),
        maxOf(texts.map((t, i) => markerWidthOf(entries[i]) + measureTextWidth(t, font)))
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
    // 行の間隔（項目・まとまりの見出し・矢印の行）と、タイトルの行の高さは別に決める（legend.ts）
    const rowHeight = legendSideRowHeight(font.size);
    const titleHeight = legendSideTitleHeight(font.size);
    const available = Math.max(rowHeight, height - PAD * 2);
    /** タイトルの行を除いて、入る行の数 */
    const fitLinesBelow = (title: number) => Math.floor((available - title * titleHeight) / rowHeight);
    const totalLinesOf = () => lineList.reduce((sum, l) => sum + l.height, 0);
    const heightOf = (title: number, lines: number) => title * titleHeight + lines * rowHeight;
    // 入りきらなければ、上下の矢印の行を取ってページに分ける（標準と同じ）。タイトルを外せば入るなら、タイトルを外してページに分けない。
    // 矢印 2 行と項目 1 行も入らないほど低ければタイトルを外し、それでも入らなければ矢印を出さずに入るだけ出す（矢印が描く範囲の外に出て押せなくならないように）
    let titleLines = titleText ? 1 : 0;
    let paged = heightOf(titleLines, totalLinesOf()) > available;
    if (paged && (heightOf(0, totalLinesOf()) <= available || fitLinesBelow(titleLines) < 3)) {
        titleLines = 0;
        paged = heightOf(0, totalLinesOf()) > available;
    }
    if (paged && fitLinesBelow(0) < 3) paged = false;
    let pages: Line[][] = [lineList];
    if (paged) {
        const perPage = Math.max(1, fitLinesBelow(titleLines) - 2);
        // まとまりの見出しは続く項目とひとまとめにして、同じページに置く（見出しだけが前のページに取り残されず、送った先のページも perPage に収まる）。
        // 1 ページに 1 行しか入らないときだけ、見出しと項目を別のページにする
        const units: Line[][] = [];
        for (const l of lineList) {
            const last = units[units.length - 1];
            const lastLine = last?.[last.length - 1];
            if (lastLine && lastLine.index === null && lastLine.text && perPage >= 2) last.push(l);
            else units.push([l]);
        }
        pages = [];
        let current: Line[] = [];
        let used = 0;
        for (const unit of units) {
            const unitHeight = unit.reduce((sum, l) => sum + l.height, 0);
            if (current.length && used + unitHeight > perPage) {
                pages.push(current);
                current = [];
                used = 0;
            }
            for (const l of unit) {
                // ページの頭に半行の間は置かない
                if (!current.length && l.index === null && !l.text) continue;
                current.push(l);
                used += l.height;
            }
        }
        if (current.length) pages.push(current);
    }
    const pageCount = pages.length;
    const page = Math.max(0, Math.min(pageCount - 1, spec.page ?? 0));
    const shownAll = pages[page];
    // 矢印を出さないほど低いときは、入る行だけ出す
    const shown = paged || heightOf(titleLines, totalLinesOf()) <= available ? shownAll : shownAll.slice(0, Math.max(1, fitLinesBelow(titleLines)));
    // 矢印の位置はどのページでも同じにする（いちばん行の多いページで並べる）
    const pageLines = maxOf(pages.map((pg) => pg.reduce((sum, l) => sum + l.height, 0)));
    const contentHeight = heightOf(titleLines, paged ? pageLines + 2 : shown.reduce((sum, l) => sum + l.height, 0));
    // 箱の上下の空き（PAD）を取らずに寄せ、中身を LEGEND_SIDE_LIFT だけ上げる（legend.ts）。入る行の数は今までどおり上下の PAD を除いて数える
    const startTop = -LEGEND_SIDE_LIFT;
    const top = align === "start" ? startTop : align === "center" ? (height - contentHeight) / 2 - LEGEND_SIDE_LIFT : height - contentHeight - LEGEND_SIDE_LIFT;
    // 最後の行が字の行（項目・まとまりの見出し）なら、字の下の端（文字が大きいと行の下に出る）が箱の下の端を越えないところまで上げる
    // （上下の凡例が箱を延ばすのと同じく、字を箱の下の端で切らない）。矢印の行は行の中に収まる
    const lastLine = shown[shown.length - 1];
    const lastIsText = !paged && lastLine !== undefined && (lastLine.index !== null || lastLine.text !== "");
    const textBottomTop = lastIsText ? height - contentHeight - legendTextOverhang(font.size, rowHeight) : Infinity;
    const y0 = Math.min(Math.max(startTop, top), textBottomTop);

    const title = titleText && titleLines ? { text: titleText, x: PAD, y: y0 + titleHeight / 2 } : null;
    const groupTitles: LegendLayout["groupTitles"] = [];
    const arrowSize = rowHeight * LEGEND_SIDE_ARROW_RATIO;
    /** タイトルの下の行の上の端 */
    const rowsTop = y0 + titleLines * titleHeight;
    let line = 0;
    let arrows: LegendLayout["arrows"] = null;
    if (paged) {
        arrows = {
            prev: { x: PAD + arrowSize / 2, y: rowsTop + rowHeight / 2, size: arrowSize, direction: "up", enabled: page > 0 },
            next: { x: PAD + arrowSize / 2, y: 0, size: arrowSize, direction: "down", enabled: page < pageCount - 1 },
        };
        line += 1;
    }
    const items: LegendItemBox[] = [];
    for (const l of shown) {
        if (l.index === null) {
            if (l.text) groupTitles.push({ text: l.text, x: PAD, y: rowsTop + line * rowHeight + rowHeight / 2 });
            line += l.height;
            continue;
        }
        const entry = entries[l.index];
        items.push({
            index: l.index,
            text: l.text,
            name: entry.name,
            x: PAD,
            y: rowsTop + line * rowHeight + rowHeight / 2,
            glyphWidth: glyphWidthOf(entry),
            textX: PAD + markerWidthOf(entry),
            width: markerWidthOf(entry) + measureTextWidth(l.text, font),
        });
        line += 1;
    }
    if (arrows) arrows.next.y = rowsTop + (1 + pageLines) * rowHeight + rowHeight / 2;

    if (side === "left") reserve.left = boxWidth + PAD;
    else reserve.right = boxWidth + PAD;
    return { side, box, title, groupTitles, items, reserve, markerRadius, rowHeight, titleHeight, pageCount, page, arrows };
}
