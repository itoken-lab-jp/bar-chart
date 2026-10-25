"use strict";

import { LineSeriesInfo, MarkerSettings, legendMarkerDrawn } from "./viewModel";
import { markerInkSpan } from "./linePath";
import { CUSTOM_LINE_STYLE, lineStrokeOf } from "./shared/gridlines";
import { legendLineGlyphWidth, legendMarkerRadius, lineInkSpan } from "./shared/legend";
import { maxOf, minOf } from "./shared/minMax";

/**
 * 凡例の折れ線の印（短い線と、真ん中に重ねるマーカー）の形。描く所（App の凡例）と、組み方に渡す印の幅（layoutLegend の inkWidth）の
 * 両方がこれを読むので、マーカーの大きさや線の模様・端の形を変えても、名前は描いた印の見える右の端から印と名前のあいだを空けて置かれ、
 * 印の見える左の端は前の名前から項目のあいだを空けて置かれる
 */
export interface LegendLineGlyph {
    drawLine: boolean;
    drawMarker: boolean;
    /** マーカーの型（「マーカー（円）」のスタイルなら円） */
    markerShape: string;
    /** マーカーのサイズ（markerPath に渡す値） */
    markerSize: number;
    /** 線の幅 */
    lineWidth: number;
    dash?: string;
    cap: "round" | "square" | "butt";
    /** 線の長さの半分（縦棒は横の線、横棒は縦の線） */
    half: number;
    /**
     * 線の真ん中（＝マーカーの真ん中）を 0 とした、見える形の左右の端（横の向き）。線は模様（破線・点線・カスタム）と端の形、
     * マーカーは回転と罫線を含めた左右の端（左右を別々に。回した三角は左右で違う）。どこも見えなければ null
     */
    ink: { left: number; right: number } | null;
}

/**
 * 凡例の折れ線の印の形。line はその線（無ければ「すべて」の書式だけで描く）、markers は「すべて」のマーカーの書式、
 * style は凡例のマーカーのスタイル、fontPx は凡例の文字の大きさ
 */
export function legendLineGlyphOf(spec: { line: LineSeriesInfo | undefined; markers: MarkerSettings; style: string; horizontal: boolean; fontPx: number }): LegendLineGlyph {
    const { line, style, horizontal } = spec;
    const mk = line?.marker ?? spec.markers;
    const r = legendMarkerRadius(spec.fontPx);
    const lineWidth = Math.max(1, Math.min(3, line?.width ?? 2));
    // スタイル：線とマーカー（既定・1.35 まで）・線・マーカー・マーカー（円）
    const markerOnly = style === "markerOnly" || style === "markerCircleDefault";
    const drawLine = style === "lineOnly" || (!markerOnly && (!line || line.lineShow || !mk.show));
    const drawMarker = legendMarkerDrawn(style, mk.show);
    const markerShape = style === "markerCircleDefault" ? "circle" : mk.shape;
    const custom = line?.lineStyle === CUSTOM_LINE_STYLE ? lineStrokeOf(line, lineWidth) : null;
    const dash = custom
        ? custom.dashArray
        : line?.lineStyle === "dashed" ? `${lineWidth * 2} ${lineWidth * 1.5}` : line?.lineStyle === "dotted" ? `${lineWidth * 0.1} ${lineWidth * 2}` : undefined;
    const cap = custom ? custom.lineCap : line?.lineStyle === "dotted" ? "round" : "butt";
    // 横棒の縦の線は、行の高さに収まる長さ。縦棒の横の線は線の印の長さ
    const half = horizontal ? r * 1.4 : legendLineGlyphWidth(spec.fontPx) / 2;
    // マーカーは行に収まる大きさまで（半径はサイズの 0.75 倍）。マーカーだけのときは四角の印より小さくしない
    const size = Math.min(mk.size, (r * 1.3) / 0.75);
    const markerSize = markerOnly ? Math.min(Math.max(size, r / 0.75), (r * 1.3) / 0.75) : size;
    // 見える横の端。縦棒の横の線は、線の左の端から模様と端の形で見える範囲（線の終わりより手前で最後のダッシュが終わることがある）。
    // 横棒の縦の線は、どこかが見えれば線の幅の半分ずつ。マーカーは回転と罫線を含めた端
    const spans: Array<{ left: number; right: number }> = [];
    const lineSpan = drawLine ? lineInkSpan(half * 2, lineWidth, cap, dash) : null;
    if (lineSpan) spans.push(horizontal ? { left: -lineWidth / 2, right: lineWidth / 2 } : { left: lineSpan.start - half, right: lineSpan.end - half });
    if (drawMarker) {
        // 左右は別々（回した三角は左右で長さが違う）
        spans.push(markerInkSpan(markerShape, markerSize, horizontal, mk.rotation && markerShape !== "circle" ? mk.rotation : 0, mk.borderShow ? Math.min(2, mk.borderWidth) : 0));
    }
    const ink = spans.length ? { left: minOf(spans.map((x) => x.left)), right: maxOf(spans.map((x) => x.right)) } : null;
    return { drawLine, drawMarker, markerShape, markerSize, lineWidth, dash, cap, half, ink };
}
