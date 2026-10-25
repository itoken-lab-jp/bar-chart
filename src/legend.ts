"use strict";

import { FontSpec } from "./unitUtils";
import { LegendSide, LegendAlign, legendPlacement, legendMarkerRadius, legendLineGlyphWidth } from "./shared/legend";
import { layoutLegend as layoutLegendBox, LegendItemBox as SharedLegendItemBox, LegendLayout as SharedLegendLayout } from "./shared/legendLayout";

/**
 * 凡例の配置。グラフの外側（上下左右のどれか）に置き、グラフに使えない幅を reserve で返す。
 * 座標はすべて凡例の枠 (box) の左上が原点。組み方は共通（shared/legendLayout）で、ここでは印の幅と色を足す
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
    /**
     * 描いた印の見える幅（折れ線の印の、線の模様と端の形・マーカーで決まる左右の端の差。legendGlyph.ts の ink）。
     * 渡せば印の幅はこれにして（決まった幅より広くも狭くもなる）、印と名前のあいだ・項目のあいだを見える端から取る。
     * 省くと決まった幅（どこも見えない印も決まった幅のまま）
     */
    inkWidth?: number;
}

/**
 * 印の幅は、棒の四角は markerRadius × 2、比較レイヤーは layerGlyphWidth。折れ線は描いた印の見える幅（inkWidth）で、
 * 渡されなければ決まった幅（縦棒は線の印の長さ、横棒は縦の線なので四角と同じ）。y は行の中央
 */
export interface LegendItemBox extends SharedLegendItemBox {
    color: string;
}

export type LegendLayout = SharedLegendLayout<LegendItemBox>;

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
    /** 上下に置いたときに並べる行の数（省略すると 1）。左右に置いたときは効かない */
    rows?: number;
}): LegendLayout | null {
    const markerRadius = legendMarkerRadius(spec.font.size);
    const glyphWidthOf = (entry: LegendEntry) =>
        entry.kind === "layer"
            ? layerGlyphWidth(markerRadius, spec.horizontal)
            : entry.kind === "line" && entry.inkWidth !== undefined
              ? entry.inkWidth
              : entry.kind === "line" && !spec.horizontal ? legendLineGlyphWidth(spec.font.size) : markerRadius * 2;
    const layout = layoutLegendBox({
        ...spec,
        entries: spec.entries.map((entry) => ({ name: entry.name, glyphWidth: glyphWidthOf(entry), group: entry.group })),
    });
    return layout && { ...layout, items: layout.items.map((item) => ({ ...item, color: spec.entries[item.index].color })) };
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
