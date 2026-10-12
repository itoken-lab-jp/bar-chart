/**
 * 定数線（標準の分析ペインの「定数線」にあたる線）の書式・見た目・名前の置き方。
 * 線を引く場所（値・メジャー・カテゴリ）はビジュアルごとに決め、線・網掛け領域・データ ラベルの作りはここで共通にする。
 * 項目と既定は標準の定数線に合わせた（2026-10-06、Desktop 2.158 の分析ペインで確かめた）
 */
import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { CUSTOM_LINE_STYLE, CUSTOM_LINE_STYLE_ITEM, DASH_CAP_ITEMS, LINE_STYLES, LINE_STYLE_ITEMS, lineStrokeOf } from "./gridlines";
import { dropdownValue } from "./formatting";

/** 線を、棒の後ろに描くか前に描くか（標準の「位置」。標準の表示は「遅延」「前面」で、遅延は Behind の訳し違えなので背面と書く） */
export const REFERENCE_LINE_LAYERS = { back: "back", front: "front" } as const;
export const REFERENCE_LINE_LAYER_ITEMS: powerbi.IEnumMember[] = [
    { value: REFERENCE_LINE_LAYERS.back, displayName: "背面" },
    { value: REFERENCE_LINE_LAYERS.front, displayName: "前面" },
];

/** データ ラベルを線のどちら側に置くか（標準の「水平方向の位置」の「左へ移動」「右へ移動」） */
export const REFERENCE_LABEL_SIDES = { left: "left", right: "right" } as const;
export const REFERENCE_LABEL_SIDE_ITEMS: powerbi.IEnumMember[] = [
    { value: REFERENCE_LABEL_SIDES.left, displayName: "左へ移動" },
    { value: REFERENCE_LABEL_SIDES.right, displayName: "右へ移動" },
];

/** データ ラベルを上に置くか下に置くか（標準の「縦位置」） */
export const REFERENCE_LABEL_ENDS = { top: "top", bottom: "bottom" } as const;
export const REFERENCE_LABEL_END_ITEMS: powerbi.IEnumMember[] = [
    { value: REFERENCE_LABEL_ENDS.top, displayName: "上" },
    { value: REFERENCE_LABEL_ENDS.bottom, displayName: "下" },
];

/** データ ラベルの中身（標準の「スタイル」。標準の「双方向」は Both の訳し違えなので両方と書く） */
export const REFERENCE_LABEL_TEXTS = { value: "value", name: "name", both: "both" } as const;
export const REFERENCE_LABEL_TEXT_ITEMS: powerbi.IEnumMember[] = [
    { value: REFERENCE_LABEL_TEXTS.value, displayName: "データ値" },
    { value: REFERENCE_LABEL_TEXTS.name, displayName: "名前" },
    { value: REFERENCE_LABEL_TEXTS.both, displayName: "両方" },
];

/** 値の軸の定数線の網掛け領域を、値の小さい側に塗るか大きい側に塗るか（標準の「次の値より前:」「次の値より後:」） */
export const VALUE_SHADE_REGIONS = { below: "below", above: "above" } as const;
export const VALUE_SHADE_REGION_ITEMS: powerbi.IEnumMember[] = [
    { value: VALUE_SHADE_REGIONS.below, displayName: "値より小さい側" },
    { value: VALUE_SHADE_REGIONS.above, displayName: "値より大きい側" },
];

/** 線のスタイルの選択肢。最後に「カスタム」（ダッシュ配列とダッシュ キャップを自分で決める） */
export const REFERENCE_LINE_STYLE_ITEMS: powerbi.IEnumMember[] = [...LINE_STYLE_ITEMS, CUSTOM_LINE_STYLE_ITEM];

const DEFAULT_COLOR = "#605E5C";
/** 網掛けの色。棒の後ろに広く塗るので、既定は淡い灰色（標準は線の色と同じ青。透過性 40% でも棒と見分けにくい） */
const DEFAULT_SHADE_COLOR = "#E1DFDD";

/**
 * 定数線のカードの部品（線・網掛け領域・データ ラベル）。カードは部品を自分の欄に並べ、まとまり（Group）を組む。
 * labelHorizontal は「水平方向の位置」の既定（標準の Y 軸の定数線は左）。labelTextName は「スタイル」の保存名（既定 labelText）
 */
export function referenceLineParts(options: { labelHorizontal: "left" | "right"; labelTextName?: string }) {
    return {
        /** 「水平方向の位置」の既定（読み取りで保存値が選択肢に無いときにも使う） */
        labelHorizontalDefault: options.labelHorizontal,
        color: new formattingSettings.ColorPicker({ name: "color", displayName: "カラー", value: { value: DEFAULT_COLOR } }),
        transparency: new formattingSettings.NumUpDown({ name: "transparency", displayName: "透過性 (%)", value: 0 }),
        lineStyle: new formattingSettings.ItemDropdown({
            name: "lineStyle",
            displayName: "線のスタイル",
            items: REFERENCE_LINE_STYLE_ITEMS,
            value: REFERENCE_LINE_STYLE_ITEMS.find((item) => item.value === LINE_STYLES.dashed)!,
        }),
        dashArray: new formattingSettings.TextInput({ name: "dashArray", displayName: "ダッシュ配列", value: "", placeholder: "例: 5 5 0 5" }),
        scaleWithWidth: new formattingSettings.ToggleSwitch({ name: "scaleWithWidth", displayName: "幅で拡大縮小", value: true }),
        dashCap: new formattingSettings.ItemDropdown({ name: "dashCap", displayName: "ダッシュ キャップ", items: DASH_CAP_ITEMS, value: DASH_CAP_ITEMS[0] }),
        width: new formattingSettings.NumUpDown({ name: "width", displayName: "幅 (px)", value: 1 }),
        layer: new formattingSettings.ItemDropdown({ name: "layer", displayName: "位置", items: REFERENCE_LINE_LAYER_ITEMS, value: REFERENCE_LINE_LAYER_ITEMS[1] }),
        shadeShow: new formattingSettings.ToggleSwitch({ name: "shadeShow", displayName: "網掛け領域", value: false }),
        shadeMatchLine: new formattingSettings.ToggleSwitch({ name: "shadeMatchLine", displayName: "線の色を一致させる", value: false }),
        shadeColor: new formattingSettings.ColorPicker({ name: "shadeColor", displayName: "カラー", value: { value: DEFAULT_SHADE_COLOR } }),
        shadeTransparency: new formattingSettings.NumUpDown({ name: "shadeTransparency", displayName: "透過性 (%)", value: 40 }),
        labelShow: new formattingSettings.ToggleSwitch({ name: "labelShow", displayName: "データ ラベル", value: true }),
        labelHorizontal: new formattingSettings.ItemDropdown({
            name: "labelHorizontal",
            displayName: "水平方向の位置",
            items: REFERENCE_LABEL_SIDE_ITEMS,
            value: REFERENCE_LABEL_SIDE_ITEMS.find((item) => item.value === options.labelHorizontal)!,
        }),
        labelVertical: new formattingSettings.ItemDropdown({
            name: "labelVertical",
            displayName: "縦位置",
            items: REFERENCE_LABEL_END_ITEMS,
            value: REFERENCE_LABEL_END_ITEMS[0],
        }),
        labelText: new formattingSettings.ItemDropdown({
            name: options.labelTextName ?? "labelText",
            displayName: "スタイル",
            items: REFERENCE_LABEL_TEXT_ITEMS,
            value: REFERENCE_LABEL_TEXT_ITEMS[0],
        }),
        labelColor: new formattingSettings.ColorPicker({ name: "labelColor", displayName: "カラー", value: { value: DEFAULT_COLOR } }),
        labelFontSize: new formattingSettings.NumUpDown({ name: "labelFontSize", displayName: "テキストのサイズ", value: 9 }),
        /** 線が近くてラベルがぶつかるとき、ずらす（縦の線は段を変え、横の線は横へ）。オフなら重ねたまま。標準に無い項目 */
        labelAvoidOverlap: new formattingSettings.ToggleSwitch({ name: "labelAvoidOverlap", displayName: "ラベルの重なりを避ける", value: true }),
    };
}

export type ReferenceLineParts = ReturnType<typeof referenceLineParts>;

/** カスタムの線の項目は、線のスタイルがカスタムのときだけ。網掛けの色は、線の色に合わせないときだけ */
export function syncReferenceLineVisibility(parts: ReferenceLineParts): void {
    const custom = String(parts.lineStyle.value?.value) === CUSTOM_LINE_STYLE;
    parts.dashArray.visible = custom;
    parts.dashCap.visible = custom;
    parts.shadeColor.visible = !(parts.shadeMatchLine.value ?? false);
}

/** 定数線の見た目（網掛けの向きは軸ごとに違うので、ビジュアルが足す） */
export interface ReferenceLineStyle {
    show: boolean;
    color: string;
    transparency: number;
    lineStyle: string;
    dashArray: string;
    scaleWithWidth: boolean;
    dashCap: string;
    width: number;
    /** 棒の後ろ（back）か前（front）か */
    layer: "back" | "front";
    shadeShow: boolean;
    shadeColor: string;
    shadeTransparency: number;
    labelShow: boolean;
    labelHorizontal: "left" | "right";
    labelVertical: "top" | "bottom";
    labelColor: string;
    labelFontSize: number;
    labelAvoidOverlap: boolean;
}

const clampTo = (value: number | undefined, min: number, max: number, fallback: number) => Math.max(min, Math.min(max, value ?? fallback));

/** カードの部品から見た目を読む。show は線を引くか（ビジュアルが決める） */
export function referenceLineStyleOf(parts: ReferenceLineParts, show: boolean): ReferenceLineStyle {
    const color = parts.color.value?.value || DEFAULT_COLOR;
    return {
        show,
        color,
        transparency: clampTo(parts.transparency.value, 0, 100, 0),
        lineStyle: dropdownValue(parts.lineStyle, LINE_STYLES.dashed),
        dashArray: parts.dashArray.value ?? "",
        scaleWithWidth: parts.scaleWithWidth.value ?? true,
        dashCap: dropdownValue(parts.dashCap, "none"),
        width: clampTo(parts.width.value, 1, 10, 1),
        layer: dropdownValue(parts.layer, REFERENCE_LINE_LAYERS.front) === REFERENCE_LINE_LAYERS.back ? "back" : "front",
        shadeShow: parts.shadeShow.value ?? false,
        shadeColor: (parts.shadeMatchLine.value ?? false) ? color : parts.shadeColor.value?.value || DEFAULT_SHADE_COLOR,
        shadeTransparency: clampTo(parts.shadeTransparency.value, 0, 100, 40),
        labelShow: parts.labelShow.value ?? true,
        labelHorizontal: dropdownValue(parts.labelHorizontal, parts.labelHorizontalDefault) === "left" ? "left" : "right",
        labelVertical: dropdownValue(parts.labelVertical, "top") === "bottom" ? "bottom" : "top",
        labelColor: parts.labelColor.value?.value || DEFAULT_COLOR,
        labelFontSize: clampTo(parts.labelFontSize.value, 6, 32, 9),
        labelAvoidOverlap: parts.labelAvoidOverlap.value ?? true,
    };
}

/** データ ラベルの中身（スタイルで選んだデータ値・名前・両方） */
export function referenceLabelOf(style: string, name: string, value: string): string {
    return style === REFERENCE_LABEL_TEXTS.name ? name : style === REFERENCE_LABEL_TEXTS.both ? `${name} ${value}` : value;
}

/** 線の SVG のスタイル（色・透過性・幅・模様・端の形） */
export function referenceLineSvgStyle(style: ReferenceLineStyle) {
    const stroke = lineStrokeOf(style, style.width);
    return {
        stroke: style.color,
        strokeOpacity: Math.max(0, Math.min(1, 1 - style.transparency / 100)),
        strokeWidth: style.width,
        strokeDasharray: stroke.dashArray,
        strokeLinecap: stroke.lineCap,
    };
}

/** 網掛け領域の SVG のスタイル */
export function referenceShadeSvgStyle(style: ReferenceLineStyle) {
    return { fill: style.shadeColor, fillOpacity: Math.max(0, Math.min(1, 1 - style.shadeTransparency / 100)) };
}

/**
 * データ ラベルの文字の SVG のスタイル。棒や隣の線の上でも読めるよう、文字のまわりに背景の色の縁を付ける
 * （縁を先に塗って、文字を上に描く）。background はビジュアルの背景（ハイコントラストならその背景）
 */
export function referenceLabelSvgStyle(style: ReferenceLineStyle, fontFamily: string, background: string) {
    return {
        fontSize: `${style.labelFontSize}pt`,
        fontFamily,
        fill: style.labelColor,
        stroke: background,
        strokeWidth: 3,
        strokeLinejoin: "round" as const,
        paintOrder: "stroke",
    };
}

/**
 * データ ラベルの置き場所（標準の「水平方向の位置」「縦位置」）。
 * 線が縦（horizontal = false）：side は線の左右、end は描く範囲の上下の端。入らない側なら反対へ。avoid なら、ぶつかるラベルを端から内へ 1 段ずつずらす。
 * 線が横（horizontal = true）：side は線の左端・右端、end は線の上・下。入らない側なら反対へ。avoid なら、上下が近くてぶつかるラベルを端から内へずらす。
 * at は線の位置（線に直交する方向の座標）、clamp はその範囲、span は線に沿う方向の範囲 [始め, 終わり]
 */
export function placeReferenceLabels(
    items: Array<{ label: string; at: number }>,
    horizontal: boolean,
    clamp: [number, number],
    span: [number, number],
    fontPx: number,
    widthOf: (text: string) => number,
    avoid: boolean,
    side: "left" | "right" = "right",
    end: "top" | "bottom" = "top"
): Array<{ x: number; y: number; anchor: "start" | "end" }> {
    const GAP = 3;
    const rowHeight = fontPx * 1.25;
    const order = items.map((_, k) => k).sort((a, b) => items[a].at - items[b].at);
    const out = new Array<{ x: number; y: number; anchor: "start" | "end" }>(items.length);
    if (!horizontal) {
        const rows: Array<Array<[number, number]>> = [];
        for (const k of order) {
            const { label, at } = items[k];
            const width = widthOf(label);
            const fitsRight = at + GAP + width <= clamp[1];
            const fitsLeft = at - GAP - width >= clamp[0];
            const toLeft = side === "left" ? fitsLeft || !fitsRight : !fitsRight && fitsLeft;
            const left = toLeft ? at - GAP - width : at + GAP;
            const range: [number, number] = [left, left + width];
            let row = 0;
            if (avoid) {
                while (rows[row]?.some(([a, b]) => range[0] < b + GAP && a < range[1] + GAP)) row++;
            }
            (rows[row] ??= []).push(range);
            out[k] = {
                x: toLeft ? at - GAP : at + GAP,
                y: end === "top" ? span[0] + fontPx + row * rowHeight : span[1] - GAP - row * rowHeight,
                anchor: toLeft ? "end" : "start",
            };
        }
        return out;
    }
    const placedBoxes: Array<{ top: number; bottom: number; from: number; to: number }> = [];
    for (const k of order) {
        const { label, at } = items[k];
        const width = widthOf(label);
        // 線の上か下。入らない側（描く範囲の上下の端の線）なら反対へ
        const aboveFits = at - GAP - fontPx >= clamp[0];
        const belowFits = at + GAP + fontPx <= clamp[1];
        const below = end === "bottom" ? belowFits || !aboveFits : !aboveFits;
        const top = below ? at + GAP : at - GAP - fontPx;
        const bottom = top + fontPx;
        // 線の右端なら右から左へ、左端なら左から右へ。ぶつかる前のラベルがあれば、その先へ
        let from = side === "right" ? span[1] - GAP : span[0] + GAP;
        if (avoid) {
            for (const box of placedBoxes) {
                if (top < box.bottom + 1 && box.top < bottom + 1) from = side === "right" ? Math.min(from, box.to - GAP * 2) : Math.max(from, box.to + GAP * 2);
            }
        }
        placedBoxes.push({ top, bottom, from, to: side === "right" ? from - width : from + width });
        out[k] = { x: from, y: bottom, anchor: side === "right" ? "end" : "start" };
    }
    return out;
}
