/**
 * 書式ペインのカードの部品（凡例・グリッド線・フォント）。
 *
 * どのスライスも、保存先の名前（name）・表示名・既定値を作品ごとに渡す。作品ごとに違う所
 * （見出しのトグルの表示名、タイトルの既定、選択肢の並び）は引数で受け、部品の側では決めない。
 * 保存先の名前と既定値はレポートに保存される値そのものなので、変えると公開済みのレポートの設定と見た目が変わる。
 */
import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import { LEGEND_POSITION_ITEMS } from "./legend";
import { CUSTOM_LINE_STYLE, CUSTOM_LINE_STYLE_ITEM, DASH_CAP_ITEMS } from "./gridlines";

export const DEFAULT_FONT_FAMILY = "Segoe UI";

export interface FontOptions {
    /** FontControl 自体の名前（書式ペインの部品の識別。保存先ではない） */
    controlName: string;
    /** 子のスライスの保存先の前に付ける語。空なら fontFamily・fontSize・bold・italic・underline */
    prefix?: string;
    displayName: string;
    size: number;
    bold?: boolean;
    family?: string;
}

/** フォント（種類・サイズ・太字・斜体・下線）。子のスライスの name が capabilities のプロパティになる */
export function fontControl(options: FontOptions): formattingSettings.FontControl {
    const prefix = options.prefix ?? "";
    const n = (suffix: string) => (prefix ? `${prefix}${suffix}` : suffix[0].toLowerCase() + suffix.slice(1));
    return new formattingSettings.FontControl({
        name: options.controlName,
        displayName: options.displayName,
        fontFamily: new formattingSettings.FontPicker({ name: n("FontFamily"), displayName: "フォント", value: options.family ?? DEFAULT_FONT_FAMILY }),
        fontSize: new formattingSettings.NumUpDown({ name: n("FontSize"), displayName: "文字サイズ", value: options.size }),
        bold: new formattingSettings.ToggleSwitch({ name: n("Bold"), displayName: "太字", value: options.bold ?? false }),
        italic: new formattingSettings.ToggleSwitch({ name: n("Italic"), displayName: "斜体", value: false }),
        underline: new formattingSettings.ToggleSwitch({ name: n("Underline"), displayName: "下線", value: false }),
    });
}

export interface LegendOptions {
    /** 見出しのトグルの表示名 */
    showDisplayName: string;
    /** タイトルの表示の既定 */
    titleShow: boolean;
    font: FontOptions;
}

/** 凡例カードのスライスとグループ（オプション・テキスト・タイトル） */
export function legendParts(options: LegendOptions) {
    const show = new formattingSettings.ToggleSwitch({ name: "show", displayName: options.showDisplayName, value: true });
    const position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "位置",
        items: LEGEND_POSITION_ITEMS,
        value: LEGEND_POSITION_ITEMS[0],
    });
    const font = fontControl(options.font);
    const labelColor = new formattingSettings.ColorPicker({ name: "labelColor", displayName: "カラー", value: { value: "#605E5C" } });
    const titleShow = new formattingSettings.ToggleSwitch({ name: "titleShow", displayName: "タイトル", value: options.titleShow });
    const titleText = new formattingSettings.TextInput({ name: "titleText", displayName: "タイトル テキスト", value: "", placeholder: "自動" });
    const optionsGroup = new formattingSettings.Group({ name: "legendOptions", displayName: "オプション", slices: [position] });
    const textGroup = new formattingSettings.Group({ name: "legendText", displayName: "テキスト", slices: [font, labelColor] });
    const titleGroup = new formattingSettings.Group({ name: "legendTitle", displayName: "タイトル", topLevelSlice: titleShow, slices: [titleText] });
    return { show, position, font, labelColor, titleShow, titleText, optionsGroup, textGroup, titleGroup };
}

export interface GridlineOptions {
    /** 保存先の前に付ける語（horizontal・vertical） */
    prefix: string;
    showDisplayName: string;
    show: boolean;
    /** 線の種類の選択肢（並びは作品ごとに違う）と既定 */
    styleItems: powerbi.IEnumMember[];
    style: powerbi.IEnumMember;
}

/**
 * グリッド線 1 方向ぶんのスライス（表示・カラー・透過性・線のスタイル・ダッシュ配列・幅で拡大縮小・ダッシュ キャップ・幅）。
 * 線のスタイルの選択肢の最後に「カスタム」を足す。ダッシュ配列とダッシュ キャップは、カスタムのときだけ出す（sync で切り替える）
 */
export function gridlineParts(options: GridlineOptions) {
    const n = (suffix: string) => `${options.prefix}${suffix}`;
    const show = new formattingSettings.ToggleSwitch({ name: n("Show"), displayName: options.showDisplayName, value: options.show });
    const color = new formattingSettings.ColorPicker({ name: n("Color"), displayName: "カラー", value: { value: "#E1DFDD" } });
    const transparency = new formattingSettings.NumUpDown({ name: n("Transparency"), displayName: "透過性 (%)", value: 0 });
    const style = new formattingSettings.ItemDropdown({
        name: n("Style"),
        displayName: "線のスタイル",
        items: [...options.styleItems, CUSTOM_LINE_STYLE_ITEM],
        value: options.style,
    });
    // 線と隙間の長さ（px）を空白で区切って並べる（標準と同じ）
    const dashArray = new formattingSettings.TextInput({ name: n("DashArray"), displayName: "ダッシュ配列", value: "", placeholder: "例: 5 5 0 5", visible: false });
    // 点線・破線の模様を線の幅に合わせて伸び縮みさせる（標準の「幅で拡大縮小」）
    const scaleWithWidth = new formattingSettings.ToggleSwitch({ name: n("ScaleWithWidth"), displayName: "幅で拡大縮小", value: false });
    const dashCap = new formattingSettings.ItemDropdown({ name: n("DashCap"), displayName: "ダッシュ キャップ", items: DASH_CAP_ITEMS, value: DASH_CAP_ITEMS[0], visible: false });
    const width = new formattingSettings.NumUpDown({ name: n("Width"), displayName: "幅 (px)", value: 1 });
    /** 保存値を読んだあとに呼ぶ。カスタムのときだけダッシュ配列とダッシュ キャップを出す */
    const sync = () => {
        const custom = String((style.value as powerbi.IEnumMember | undefined)?.value) === CUSTOM_LINE_STYLE;
        dashArray.visible = custom;
        dashCap.visible = custom;
    };
    return {
        show,
        color,
        transparency,
        style,
        dashArray,
        scaleWithWidth,
        dashCap,
        width,
        sync,
        slices: [color, transparency, style, dashArray, scaleWithWidth, dashCap, width],
    };
}

export interface CategoryAxisOptions {
    font: FontOptions;
    titleFont: FontOptions;
    titleStyleItems: powerbi.IEnumMember[];
    /** スクロールの最初の位置の選択肢 */
    scrollStartItems: powerbi.IEnumMember[];
    /** 高さの最大値・カテゴリの最小幅を、スライダーで出すか（false なら数の入力） */
    sliders: boolean;
}

/**
 * カテゴリの軸カードの共通のスライスと、タイトル・レイアウトのグループ。
 * 値のグループは作品ごとに足すスライスがあるので、作品の側で組む（valueSlices を先頭に並べる）
 */
export function categoryAxisParts(options: CategoryAxisOptions) {
    const Num = options.sliders ? formattingSettings.Slider : formattingSettings.NumUpDown;
    const show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "値", value: true });
    const font = fontControl(options.font);
    const labelColor = new formattingSettings.ColorPicker({ name: "labelColor", displayName: "カラー", value: { value: "#605E5C" } });
    // 項目名に使う大きさの上限（ビュー全体に対する %）。縦向きは高さ、横向きは幅
    const maxHeight = new Num({ name: "maxHeight", displayName: "高さの最大値 (%)", value: 25 });
    // カテゴリの軸のタイトルは初期オフ。項目名で何の軸か読めることが多い。縦横どちらの向きでもこのカードに当てる
    const titleShow = new formattingSettings.ToggleSwitch({ name: "titleShow", displayName: "タイトル", value: false });
    const titleText = new formattingSettings.TextInput({ name: "titleText", displayName: "タイトル テキスト", value: "", placeholder: "自動" });
    // 標準の X 軸と同じ項目。カテゴリの軸には単位が無いので、どれを選んでもタイトルのまま
    const titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: options.titleStyleItems,
        value: options.titleStyleItems[0],
    });
    const titleFont = fontControl(options.titleFont);
    const titleColor = new formattingSettings.ColorPicker({ name: "titleColor", displayName: "カラー", value: { value: "#252423" } });
    // 帯がこれより狭くなるときは、スクロールする（標準と同じ）
    const minCategoryWidth = new Num({ name: "minCategoryWidth", displayName: "カテゴリの最小幅 (px)", value: 20 });
    // はみ出してスクロールするとき、開いたときにどこから見せるか。末尾は最後のカテゴリの側
    const scrollStart = new formattingSettings.ItemDropdown({
        name: "scrollStart",
        displayName: "スクロールの最初の位置",
        items: options.scrollStartItems,
        value: options.scrollStartItems[0],
    });
    const titleGroup = new formattingSettings.Group({
        name: "categoryTitle",
        displayName: "タイトル",
        topLevelSlice: titleShow,
        slices: [titleText, titleStyle, titleFont, titleColor],
    });
    const layoutGroup = new formattingSettings.Group({ name: "categoryLayout", displayName: "レイアウト", slices: [minCategoryWidth, scrollStart] });
    return {
        show,
        font,
        labelColor,
        maxHeight,
        titleShow,
        titleText,
        titleStyle,
        titleFont,
        titleColor,
        minCategoryWidth,
        scrollStart,
        valueSlices: [font, labelColor, maxHeight],
        titleGroup,
        layoutGroup,
    };
}

export interface ValueAxisOptions {
    font: FontOptions;
    titleFont: FontOptions;
    titleStyleItems: powerbi.IEnumMember[];
    unitTypeItems: powerbi.IEnumMember[];
    unitNotationItems: powerbi.IEnumMember[];
    precisionItems: powerbi.IEnumMember[];
    /** 書式ペインに出す説明（作品ごとに効き方が違うもの） */
    startDescription?: string;
    tickCountDescription: string;
    unitTypeDescription?: string;
}

/**
 * 値の軸カードの共通のスライスと、タイトルのグループ。範囲・値・単位ラベルのグループは
 * 作品ごとに足すスライスが間に入るので、作品の側で組む
 */
export function valueAxisParts(options: ValueAxisOptions) {
    const text = (name: string, displayName: string, description?: string, placeholder = "自動") =>
        new formattingSettings.TextInput({ name, displayName, value: "", placeholder, ...(description ? { description } : {}) });
    const toggle = (name: string, displayName: string, value: boolean) => new formattingSettings.ToggleSwitch({ name, displayName, value });
    const dropdown = (name: string, displayName: string, items: powerbi.IEnumMember[], description?: string) =>
        new formattingSettings.ItemDropdown({ name, displayName, items, value: items[0], ...(description ? { description } : {}) });
    const color = (name: string, value: string) => new formattingSettings.ColorPicker({ name, displayName: "カラー", value: { value } });

    const titleShow = toggle("titleShow", "タイトル", true);
    const titleText = text("titleText", "タイトル テキスト");
    const titleStyle = dropdown("titleStyle", "スタイル", options.titleStyleItems);
    const titleFont = fontControl(options.titleFont);
    const titleColor = color("titleColor", "#252423");
    return {
        start: text("start", "最小値", options.startDescription),
        end: text("end", "最大値"),
        invertRange: toggle("invertRange", "範囲の反転", false),
        roundRange: toggle("roundRange", "範囲を丸める", true),
        // 目盛り（グリッド線）の本数の目安。空なら自動（描く範囲の長さで決める）
        tickCount: text("tickCount", "目盛りの本数 (目安)", options.tickCountDescription),
        show: toggle("show", "値", true),
        font: fontControl(options.font),
        labelColor: color("labelColor", "#605E5C"),
        unitType: dropdown("unitType", "表示単位", options.unitTypeItems, options.unitTypeDescription),
        unitNotation: dropdown("unitNotation", "単位の表記", options.unitNotationItems),
        precision: dropdown("precision", "小数点以下の桁数", options.precisionItems),
        switchPosition: toggle("switchPosition", "軸の位置を切り替える", false),
        titleShow,
        titleText,
        titleStyle,
        titleFont,
        titleColor,
        unitShow: toggle("unitShow", "単位ラベルの表示", true),
        unitText: text("unitText", "単位の追加文字", undefined, "例: 円, 人, 件"),
        titleGroup: new formattingSettings.Group({
            name: "valueTitle",
            displayName: "タイトル",
            topLevelSlice: titleShow,
            slices: [titleText, titleStyle, titleFont, titleColor],
        }),
    };
}

/**
 * データ ラベルの値の色・桁数と、符号の書き方（マイナス・0・丸めて 0 のマイナス・符号の色）。
 * 選択肢は数値の書式（numberFormat）と同じ
 */
export function labelValueParts(items: {
    precisions: powerbi.IEnumMember[];
    negativeStyles: powerbi.IEnumMember[];
    zeroStyles: powerbi.IEnumMember[];
    toneModes: powerbi.IEnumMember[];
    goodColor: string;
    badColor: string;
}) {
    return {
        color: new formattingSettings.ColorPicker({ name: "color", displayName: "カラー", value: { value: "" } }),
        precision: new formattingSettings.ItemDropdown({ name: "precision", displayName: "小数点以下の桁数", items: items.precisions, value: items.precisions[0] }),
        negativeStyle: new formattingSettings.ItemDropdown({
            name: "negativeStyle",
            displayName: "マイナス",
            items: items.negativeStyles,
            value: items.negativeStyles[0],
        }),
        zeroStyle: new formattingSettings.ItemDropdown({ name: "zeroStyle", displayName: "0", items: items.zeroStyles, value: items.zeroStyles[0] }),
        negativeZero: new formattingSettings.ToggleSwitch({ name: "negativeZero", displayName: "丸めて 0 のマイナスに符号", value: true }),
        toneMode: new formattingSettings.ItemDropdown({
            name: "toneMode",
            displayName: "符号の色",
            description: "棒の外のラベルと、背景を付けたラベルに効く。棒の中のラベルは、棒の色に合わせて読める色を自動で選ぶ",
            items: items.toneModes,
            value: items.toneModes[0],
        }),
        positiveColor: new formattingSettings.ColorPicker({ name: "positiveColor", displayName: "プラスの色", value: { value: items.goodColor } }),
        negativeColor: new formattingSettings.ColorPicker({ name: "negativeColor", displayName: "マイナスの色", value: { value: items.badColor } }),
    };
}

/** データ ラベルの背景（表示・カラー・透過性）。表示名と既定は作品ごとに渡す */
export function labelBackgroundParts(options: { showDisplayName: string; colorDisplayName: string; color: string; transparency: number }) {
    return {
        backgroundShow: new formattingSettings.ToggleSwitch({ name: "backgroundShow", displayName: options.showDisplayName, value: false }),
        backgroundColor: new formattingSettings.ColorPicker({ name: "backgroundColor", displayName: options.colorDisplayName, value: { value: options.color } }),
        backgroundTransparency: new formattingSettings.NumUpDown({ name: "backgroundTransparency", displayName: "透過性 (%)", value: options.transparency }),
    };
}
