/**
 * プロット領域の背景色。
 * 標準の 11 月版（Desktop 2.160）の書式 > ビジュアル >「プロットエリアの背景」の色の 2 欄に合わせる：
 * 背景色の既定は「塗りつぶしなし」、背景色の透過性 (%) は既定 0（標準は塗りつぶしなしのあいだ灰色。ここでは出さない）。
 * 保存のプロパティ名も標準と同じ（plotArea の backgroundColor・backgroundColorTransparency）。
 * 標準にある画像（画像のソース・イメージ・サイズ調整・画像の透過性）と背景色の条件付き書式（fx）は入れない
 */
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

/** 塗る色と不透明度。塗らないなら null */
export interface PlotAreaFill {
    color: string;
    opacity: number;
}

export class PlotAreaCardSettings extends formattingSettings.SimpleCard {
    name = "plotArea";
    displayName = "プロットエリアの背景";

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色",
        value: { value: "" },
        isNoFillItemSupported: true,
    });

    backgroundColorTransparency = new formattingSettings.NumUpDown({
        name: "backgroundColorTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    slices = [this.backgroundColor, this.backgroundColorTransparency];

    /**
     * 背景色が塗りつぶしなしのあいだは、透過性を出さない。標準は灰色で残すが、
     * この版の書式の部品は欄ごとに灰色にできない（まとまりごとだけ）ので、出さない形にした
     */
    syncVisibility(): void {
        this.backgroundColorTransparency.visible = !!this.backgroundColor.value?.value;
    }
}

/** 書式の値から塗る色を決める。色が無い（塗りつぶしなし）なら null。透過性は 0〜100 に収める */
export function plotAreaFillOf(card: PlotAreaCardSettings): PlotAreaFill | null {
    const color = card.backgroundColor.value?.value;
    if (!color) return null;
    const transparency = Math.max(0, Math.min(100, card.backgroundColorTransparency.value ?? 0));
    return { color, opacity: (100 - transparency) / 100 };
}

