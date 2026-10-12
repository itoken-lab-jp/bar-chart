import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { LEGEND_POSITION_ITEMS, standardLegendPosition } from "./legend";

/** グリッド線の 1 本ぶん（自作の「グリッド線」カードの横 horizontal* か縦 vertical*）と、テーマが値を持つ軸のカード */
export interface ThemeGridline {
    /** テーマが gridlineShow などを持つ軸のカード（valueAxis・categoryAxis） */
    axis: string;
    /** 自作の「グリッド線」カードの保存名の前に付く語 */
    prefix: "horizontal" | "vertical";
    show: formattingSettings.ToggleSwitch;
    color: formattingSettings.ColorPicker;
    style: formattingSettings.ItemDropdown;
    width: formattingSettings.NumUpDown;
}

export interface ThemeTargets {
    categoryTitleShow: formattingSettings.ToggleSwitch;
    valueTitleShow: formattingSettings.ToggleSwitch;
    legendTitleShow: formattingSettings.ToggleSwitch;
    legendPosition: formattingSettings.ItemDropdown;
    /** カテゴリの軸の「高さの最大値 (%)」と、描画と同じ丸めの下限 */
    categoryMaxHeight: formattingSettings.NumUpDown;
    maxHeightMin: number;
    gridlines: ThemeGridline[];
    /** グリッド線の線のスタイルの選択肢（テーマの値で探す） */
    lineStyleItems: powerbi.IEnumMember[];
    /** グリッド線の幅の丸めの下限（描画と同じ） */
    gridlineWidthMin: number;
}

/**
 * 基本テーマ・カスタムテーマに合わせる。テーマは標準のビジュアルの名前（showAxisTitle・showTitle・maxMarginFactor・gridline*）で
 * 値を持つので、capabilities にその名前も置いて受け取り、作り手が自作の設定を保存していないときの既定にする。
 * 書式ペインにも同じ値を出す（「既定値にリセット」でテーマの値に戻る）。populate の直後に呼ぶ。
 * 凡例の位置は、古い保存値（topLeft など）を標準の値に読み替える
 */
export function applyThemeBasics(objects: powerbi.DataViewObjects | undefined, t: ThemeTargets): void {
    const raw = (card: string, prop: string): unknown => objects?.[card]?.[prop];
    const inherit = (slice: formattingSettings.ToggleSwitch, card: string, own: string, standard: string) => {
        const themeValue = raw(card, standard);
        if (raw(card, own) == null && typeof themeValue === "boolean") slice.value = themeValue;
    };
    inherit(t.categoryTitleShow, "categoryAxis", "titleShow", "showAxisTitle");
    inherit(t.valueTitleShow, "valueAxis", "titleShow", "showAxisTitle");
    inherit(t.legendTitleShow, "legend", "titleShow", "showTitle");
    const position = standardLegendPosition(raw("legend", "position"));
    if (position) t.legendPosition.value = LEGEND_POSITION_ITEMS.find((i) => i.value === position)!;

    // 項目名の欄の上限：標準は categoryAxis の maxMarginFactor で持つ（Fluent 2 は 50）
    const marginFactor = raw("categoryAxis", "maxMarginFactor");
    if (raw("categoryAxis", "maxHeight") == null && typeof marginFactor === "number") {
        t.categoryMaxHeight.value = Math.max(t.maxHeightMin, Math.min(100, marginFactor));
    }

    // グリッド線：標準は軸のカードの中（gridlineShow・gridlineColor・gridlineStyle・gridlineThickness）
    for (const g of t.gridlines) {
        const own = (prop: string) => raw("gridlines", `${g.prefix}${prop}`) != null;
        const show = raw(g.axis, "gridlineShow");
        if (!own("Show") && typeof show === "boolean") g.show.value = show;
        const color = (raw(g.axis, "gridlineColor") as powerbi.Fill | undefined)?.solid?.color;
        if (!own("Color") && typeof color === "string" && color) g.color.value = { value: color };
        const style = t.lineStyleItems.find((i) => i.value === raw(g.axis, "gridlineStyle"));
        if (!own("Style") && style) g.style.value = style;
        const thickness = raw(g.axis, "gridlineThickness");
        // 描画と同じ範囲に丸める（書式ペインの値と描画の幅をずらさない）
        if (!own("Width") && typeof thickness === "number" && thickness > 0) g.width.value = Math.max(t.gridlineWidthMin, Math.min(10, thickness));
    }
}
