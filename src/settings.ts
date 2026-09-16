"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsCompositeCard = formattingSettings.CompositeCard;
import FormattingSettingsGroup = formattingSettings.Group;
import FormattingSettingsContainer = formattingSettings.Container;
import FormattingSettingsContainerItem = formattingSettings.ContainerItem;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

export interface ColumnTarget {
    name: string;
    selector: powerbi.data.Selector;
    color: string;
    transparency: number;
    borderShow: boolean;
    borderMatchColumn: boolean;
    borderColor: string;
    borderTransparency: number;
    borderWidth: number;
}

/**
 * 「設定の適用先」に並べるカテゴリの上限。カテゴリ数ぶん項目が増えると書式ペインが
 * 重くなり選べなくなるため頭を打つ。
 * 上限を超えたカテゴリは「すべて」の設定で描かれる（個別に保存済みの値は描画に効く）。
 */
export const MAX_COLUMN_TARGETS = 30;

/**
 * 表示単位。value (0〜12 の桁数) はレポートに保存されるので変えない。表示名だけを単位の語にする。
 * 英語表記に K/M/bn/T の語が無い桁（万・億など）は、1 つ下の K/M/bn/T（十・百は単位なし）に
 * 読み替える（unitUtils.toStandardUnitKey）。
 */
export const UNIT_TYPES = [
    { value: "auto", displayName: "自動" },
    { value: "0", displayName: "なし" },
    { value: "1", displayName: "十" },
    { value: "2", displayName: "百" },
    { value: "3", displayName: "千" },
    { value: "4", displayName: "万" },
    { value: "5", displayName: "十万" },
    { value: "6", displayName: "百万" },
    { value: "7", displayName: "千万" },
    { value: "8", displayName: "億" },
    { value: "9", displayName: "十億" },
    { value: "10", displayName: "百億" },
    { value: "11", displayName: "千億" },
    { value: "12", displayName: "兆" },
];

export const UNIT_POSITIONS = [
    { value: "valueAxisTop", displayName: "Y 軸の上 (左上)" },
    { value: "plotTopRight", displayName: "プロット エリアの右上" },
    { value: "none", displayName: "非表示" },
];

/** value は保存済みレポートとの互換のため変えない（"standard" = K・M・bn・T の表記） */
export const UNIT_NOTATIONS = [
    { value: "japanese", displayName: "日本語（万・億）" },
    { value: "standard", displayName: "英語（K・M・bn）" },
];

/** 表示される形そのものを選択肢の名前にする。value は保存済みレポートとの互換のため変えない */
export const UNIT_STYLES = [
    { value: "parentheses", displayName: "(百万円)" },
    { value: "withPrefix", displayName: "(単位: 百万円)" },
];

export const PRECISIONS = [
    { value: "auto", displayName: "自動" },
    { value: "0", displayName: "0" },
    { value: "1", displayName: "1" },
    { value: "2", displayName: "2" },
    { value: "3", displayName: "3" },
];


/** 値が空のとき入力欄に出す文字。標準の「外側のパディング」と同じ */
export const AUTO_PLACEHOLDER = "自動";

/**
 * 値を空（undefined）にでき、空のとき入力欄に「自動」と出す NumUpDown。
 * API の visuals.NumUpDown は placeholderText を持つが、formattingmodel 6.0.4 の
 * NumUpDown は options しか渡さないので、ここで足す。
 * 空のまま = 保存値なし = 自動。書式ペインの「既定値に戻す」でも空に戻る。
 */
export class AutoNumUpDown extends formattingSettings.NumUpDown {
    getFormattingComponent(objectName: string): powerbi.visuals.NumUpDown {
        return {
            ...super.getFormattingComponent(objectName),
            placeholderText: AUTO_PLACEHOLDER,
        };
    }
}

/**
 * グラフの種類。標準では別々のビジュアルだが、barChart は書式ペインで切り替える。
 * ribbon は 1.18 までの保存値だけ（リボンは種類ではなく「リボン」カードの見せ方にした、#108）
 */
export const CHART_TYPES = {
    clustered: "clustered",
    stacked: "stacked",
    stacked100: "stacked100",
    ribbon: "ribbon",
} as const;

export type ChartType = typeof CHART_TYPES.clustered | typeof CHART_TYPES.stacked | typeof CHART_TYPES.stacked100;

export const CHART_TYPE_ITEMS: powerbi.IEnumMember[] = [
    { value: CHART_TYPES.clustered, displayName: "集合" },
    { value: CHART_TYPES.stacked, displayName: "積み上げ" },
    { value: CHART_TYPES.stacked100, displayName: "100% 積み上げ" },
];

/** リボンを出すときの積む順。凡例の順は標準の積み上げ＋リボン、値の大きい順は標準のリボン グラフと同じ */
export const RIBBON_ORDERS = {
    legend: "legend",
    value: "value",
} as const;

export const RIBBON_ORDER_ITEMS: powerbi.IEnumMember[] = [
    { value: RIBBON_ORDERS.legend, displayName: "凡例の順" },
    { value: RIBBON_ORDERS.value, displayName: "値の大きい順" },
];

/** 棒の向き。標準では縦棒と横棒は別のビジュアルだが、barChart は書式ペインで切り替える */
export const ORIENTATIONS = {
    vertical: "vertical",
    horizontal: "horizontal",
} as const;

export type Orientation = (typeof ORIENTATIONS)[keyof typeof ORIENTATIONS];

export const ORIENTATION_ITEMS: powerbi.IEnumMember[] = [
    { value: ORIENTATIONS.vertical, displayName: "縦棒" },
    { value: ORIENTATIONS.horizontal, displayName: "横棒" },
];

class TargetItem extends FormattingSettingsCard {
    name = "columnsTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

/**
 * グラフの種類と向き。いちばん最初に選ぶものなので、書式ペインのいちばん上のカードにする（#100）。
 * 1.14 までは「列」（いまの「棒」）カードにあり、保存先も columns だった。新しいカードに保存が無いレポートは、
 * 古い保存を読んで同じ種類で開く（VisualFormattingSettingsModel.applyChartTypeDefaults）
 */
export class ChartCardSettings extends FormattingSettingsCard {
    name = "chart";
    displayName = "グラフの種類";

    /** 既定は集合（系列 1 本なら 1.4 までと同じ見た目） */
    chartType = new formattingSettings.ItemDropdown({
        name: "chartType",
        displayName: "種類",
        items: CHART_TYPE_ITEMS,
        value: CHART_TYPE_ITEMS[0],
    });

    /** 既定は縦棒（1.7 までと同じ） */
    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "向き",
        items: ORIENTATION_ITEMS,
        value: ORIENTATION_ITEMS[0],
    });

    slices = [this.chartType, this.orientation];
}

export class ColumnsCardSettings extends FormattingSettingsCompositeCard {
    name = "columns";
    /** 標準の日本語の表示は「列」（Columns の訳）だが、縦棒・横棒のどちらでも読めるよう「棒」にする（保存先は columns のまま） */
    displayName = "棒";
    analyticsPane = false;

    // --- 対象ごとに指定できるもの (「すべて」用の実体) ---
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "カラー",
        value: { value: "#118DFF" },
    });

    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    showBorder = new formattingSettings.ToggleSwitch({
        name: "showBorder",
        displayName: "罫線",
        value: false,
    });

    borderMatchColumn = new formattingSettings.ToggleSwitch({
        name: "borderMatchColumn",
        displayName: "棒の色を一致させる",
        value: false,
    });

    borderFill = new formattingSettings.ColorPicker({
        name: "borderFill",
        displayName: "罫線のカラー",
        value: { value: "#605E5C" },
    });

    borderTransparency = new formattingSettings.NumUpDown({
        name: "borderTransparency",
        displayName: "罫線の透過性 (%)",
        value: 0,
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "罫線の幅 (px)",
        value: 1,
    });

    // --- レイアウト (カード全体) ---
    reverseOrder = new formattingSettings.ToggleSwitch({
        name: "reverseOrder",
        displayName: "順序を逆にする",
        value: false,
    });

    sortByValue = new formattingSettings.ToggleSwitch({
        name: "sortByValue",
        displayName: "値順で並べ替え",
        value: false,
    });

    /**
     * 最初のカテゴリの前と最後のカテゴリの後ろの余白（カテゴリ 1 つ分の幅に対する %）。
     * 既定は空 = 「自動」（カテゴリ間のスペースの半分。1.3.1.0 までと同じ配置）。
     * 範囲 0〜100 は描画側でクランプする
     */
    outerPadding = new AutoNumUpDown({
        name: "outerPadding",
        displayName: "外側のパディング (%)",
        value: undefined,
    });

    categorySpacing = new formattingSettings.NumUpDown({
        name: "categorySpacing",
        displayName: "カテゴリ間のスペース (%)",
        value: 20,
    });

    /**
     * 集合で、同じカテゴリの中の系列と系列のすき間（系列 1 本ぶんの幅に対する %）。
     * 標準と同じく既定は 0（すき間なし）。範囲は描画側でクランプする
     */
    seriesSpacing = new formattingSettings.NumUpDown({
        name: "seriesSpacing",
        displayName: "系列間のスペース (%)",
        value: 0,
    });

    /** 0 = 上限なし。棒が太くなりすぎるのを抑える */
    maxBarWidth = new formattingSettings.NumUpDown({
        name: "maxBarWidth",
        displayName: "最大幅 (px)",
        value: 0,
    });

    cornerRadius = new formattingSettings.NumUpDown({
        name: "cornerRadius",
        displayName: "角丸 (px)",
        value: 0,
    });

    /**
     * 「すべて」の項目。複数系列では色を系列ごとに決めるので、カラーを出さない
     * （標準も、系列があるときは「すべて」でカラーを選べない）
     */
    private allItem(seriesMode = false): FormattingSettingsCard {
        return new TargetItem("すべて", [
            ...(seriesMode ? [] : [this.fill]),
            this.transparency,
            this.showBorder,
            this.borderMatchColumn,
            this.borderFill,
            this.borderTransparency,
            this.borderWidth,
        ]);
    }

    targetGroup = new FormattingSettingsGroup({
        name: "columnsTargets",
        slices: [],
        container: new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: [this.allItem()],
        }),
    });

    layoutGroup = new FormattingSettingsGroup({
        name: "columnsLayout",
        displayName: "レイアウト",
        slices: [
            this.reverseOrder,
            this.sortByValue,
            this.outerPadding,
            this.categorySpacing,
            this.seriesSpacing,
            this.maxBarWidth,
            this.cornerRadius,
        ],
    });

    groups = [this.targetGroup, this.layoutGroup];

    /**
     * 「設定の適用先」に対象を並べる。系列が 1 本ならカテゴリ、複数なら系列が対象になる
     * （標準の集合縦棒と同じ）。対象の selector はどちらも viewModel が作る
     */
    applyTargets(targets: ColumnTarget[], seriesMode = false): void {
        this.targetGroup.container = new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: [
                this.allItem(seriesMode),
                ...targets.slice(0, MAX_COLUMN_TARGETS).map(
                    (target) =>
                        new TargetItem(target.name, [
                            new formattingSettings.ColorPicker({
                                name: "fill",
                                displayName: "カラー",
                                value: { value: target.color },
                                selector: target.selector,
                            }),
                            new formattingSettings.NumUpDown({
                                name: "transparency",
                                displayName: "透過性 (%)",
                                value: target.transparency,
                                selector: target.selector,
                            }),
                            new formattingSettings.ToggleSwitch({
                                name: "showBorder",
                                displayName: "罫線",
                                value: target.borderShow,
                                selector: target.selector,
                            }),
                            new formattingSettings.ToggleSwitch({
                                name: "borderMatchColumn",
                                displayName: "棒の色を一致させる",
                                value: target.borderMatchColumn,
                                selector: target.selector,
                            }),
                            new formattingSettings.ColorPicker({
                                name: "borderFill",
                                displayName: "罫線のカラー",
                                value: { value: target.borderColor },
                                selector: target.selector,
                            }),
                            new formattingSettings.NumUpDown({
                                name: "borderTransparency",
                                displayName: "罫線の透過性 (%)",
                                value: target.borderTransparency,
                                selector: target.selector,
                            }),
                            new formattingSettings.NumUpDown({
                                name: "borderWidth",
                                displayName: "罫線の幅 (px)",
                                value: target.borderWidth,
                                selector: target.selector,
                            }),
                        ])
                ),
            ],
        });
    }
}

export const LABEL_POSITIONS = {
    auto: "auto",
    outsideEnd: "outsideEnd",
    insideTop: "insideTop",
    insideCenter: "insideCenter",
    insideBottom: "insideBottom",
} as const;

export const LABEL_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: LABEL_POSITIONS.auto, displayName: "自動" },
    { value: LABEL_POSITIONS.outsideEnd, displayName: "外側の上" },
    { value: LABEL_POSITIONS.insideTop, displayName: "内側の上" },
    { value: LABEL_POSITIONS.insideCenter, displayName: "内側の中央" },
    { value: LABEL_POSITIONS.insideBottom, displayName: "内側の下" },
];

export const LABEL_ORIENTATIONS = {
    horizontal: "horizontal",
    vertical: "vertical",
} as const;

export const LABEL_ORIENTATION_ITEMS: powerbi.IEnumMember[] = [
    { value: LABEL_ORIENTATIONS.horizontal, displayName: "横" },
    { value: LABEL_ORIENTATIONS.vertical, displayName: "縦" },
];

/** データラベルを系列ごとに変えるときの対象（複数系列のときだけ） */
export const DETAIL_CONTENTS = {
    percentOfTotal: "percentOfTotal",
    custom: "custom",
} as const;

/** データ ラベルの詳細のコンテンツ。標準の「カスタム」はフィールドを選ぶので、ここでは「ラベルの詳細」に入れたフィールド */
export const DETAIL_CONTENT_ITEMS: powerbi.IEnumMember[] = [
    { value: DETAIL_CONTENTS.percentOfTotal, displayName: "全体に対する割合" },
    { value: DETAIL_CONTENTS.custom, displayName: "カスタム（ラベルの詳細）" },
];

export interface LabelTarget {
    name: string;
    selector: powerbi.data.Selector;
    show: boolean;
    /** 空 = 自動（棒の色に合わせて白か黒） */
    color: string;
}

class LabelTargetItem extends FormattingSettingsCard {
    name = "labelTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

export const LEGEND_POSITIONS = {
    topLeft: "topLeft",
    topCenter: "topCenter",
    topRight: "topRight",
    bottomLeft: "bottomLeft",
    bottomCenter: "bottomCenter",
    bottomRight: "bottomRight",
    leftTop: "leftTop",
    leftCenter: "leftCenter",
    leftBottom: "leftBottom",
    rightTop: "rightTop",
    rightCenter: "rightCenter",
    rightBottom: "rightBottom",
} as const;

/** 標準の凡例の「位置」と同じ 12 通り。既定は上詰め (左) */
export const LEGEND_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: LEGEND_POSITIONS.topLeft, displayName: "上詰め (左)" },
    { value: LEGEND_POSITIONS.topCenter, displayName: "上詰め (中央)" },
    { value: LEGEND_POSITIONS.topRight, displayName: "上詰め (右)" },
    { value: LEGEND_POSITIONS.bottomLeft, displayName: "下詰め (左)" },
    { value: LEGEND_POSITIONS.bottomCenter, displayName: "下詰め (中央)" },
    { value: LEGEND_POSITIONS.bottomRight, displayName: "下詰め (右)" },
    { value: LEGEND_POSITIONS.leftTop, displayName: "左上" },
    { value: LEGEND_POSITIONS.leftCenter, displayName: "左中央" },
    { value: LEGEND_POSITIONS.leftBottom, displayName: "左下" },
    { value: LEGEND_POSITIONS.rightTop, displayName: "右上" },
    { value: LEGEND_POSITIONS.rightCenter, displayName: "右中央" },
    { value: LEGEND_POSITIONS.rightBottom, displayName: "右下" },
];

/** 凡例。系列が複数のとき（値が 2 つ以上か、凡例にフィールドがあるとき）だけ描く */
export class LegendCardSettings extends FormattingSettingsCompositeCard {
    name = "legend";
    displayName = "凡例";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "表示",
        value: true,
    });

    topLevelSlice = this.show;

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "位置",
        items: LEGEND_POSITION_ITEMS,
        value: LEGEND_POSITION_ITEMS[0],
    });

    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "文字サイズ",
            value: 10,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: true,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    optionsGroup = new FormattingSettingsGroup({
        name: "legendOptions",
        displayName: "オプション",
        slices: [this.position],
    });

    textGroup = new FormattingSettingsGroup({
        name: "legendText",
        displayName: "テキスト",
        slices: [this.font, this.labelColor],
    });

    titleGroup = new FormattingSettingsGroup({
        name: "legendTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [this.titleText],
    });

    groups = [this.optionsGroup, this.textGroup, this.titleGroup];
}

export class DataLabelsCardSettings extends FormattingSettingsCompositeCard {
    name = "dataLabels";
    displayName = "データ ラベル";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "表示",
        value: false,
    });

    topLevelSlice = this.show;

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "位置",
        items: LABEL_POSITION_ITEMS,
        value: LABEL_POSITION_ITEMS[0],
    });

    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "方向",
        items: LABEL_ORIENTATION_ITEMS,
        value: LABEL_ORIENTATION_ITEMS[0],
    });

    overflow = new formattingSettings.ToggleSwitch({
        name: "overflow",
        displayName: "オーバーフロー テキスト",
        value: true,
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "文字サイズ",
        value: 9,
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "フォント",
        value: "Segoe UI",
    });

    bold = new formattingSettings.ToggleSwitch({
        name: "bold",
        displayName: "太字",
        value: false,
    });

    italic = new formattingSettings.ToggleSwitch({
        name: "italic",
        displayName: "斜体",
        value: false,
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "" },
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    /** 値の行を出すか。既定はオンで、100% 積み上げでは保存していなければオフ（標準と同じ、applyChartTypeDefaults） */
    valueShow = new formattingSettings.ToggleSwitch({
        name: "valueShow",
        displayName: "値",
        value: true,
    });

    /** 詳細の行を出すか。既定はオフで、100% 積み上げでは保存していなければオン（標準と同じ） */
    detailShow = new formattingSettings.ToggleSwitch({
        name: "detailShow",
        displayName: "詳細",
        value: false,
    });

    detailContent = new formattingSettings.ItemDropdown({
        name: "detailContent",
        displayName: "コンテンツ",
        items: DETAIL_CONTENT_ITEMS,
        value: DETAIL_CONTENT_ITEMS[0],
    });

    detailFont = new formattingSettings.FontControl({
        name: "detailFont",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "detailFontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "detailFontSize",
            displayName: "文字サイズ",
            value: 9,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "detailBold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "detailItalic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "detailUnderline",
            displayName: "下線",
            value: false,
        }),
    });

    /** 空 = 自動（値の行と同じく、棒の色に合わせて白か黒） */
    detailColor = new formattingSettings.ColorPicker({
        name: "detailColor",
        displayName: "カラー",
        value: { value: "" },
    });

    detailTransparency = new formattingSettings.NumUpDown({
        name: "detailTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    /** カスタム（ラベルの詳細のフィールド）の数値の表示単位。「自動」はフィールドの書式のまま */
    detailUnitType = new formattingSettings.ItemDropdown({
        name: "detailUnitType",
        displayName: "表示単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0],
    });

    detailPrecision = new formattingSettings.ItemDropdown({
        name: "detailPrecision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    backgroundShow = new formattingSettings.ToggleSwitch({
        name: "backgroundShow",
        displayName: "背景の表示",
        value: false,
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色",
        value: { value: "#FFFFFF" },
    });

    backgroundTransparency = new formattingSettings.NumUpDown({
        name: "backgroundTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    optionsGroup = new FormattingSettingsGroup({
        name: "labelOptions",
        displayName: "オプション",
        slices: [
            this.orientation,
            this.position,
            this.overflow,
        ],
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "labelValues",
        displayName: "値",
        topLevelSlice: this.valueShow,
        slices: [
            this.fontFamily,
            this.fontSize,
            this.bold,
            this.italic,
            this.color,
            this.precision,
        ],
    });

    detailGroup = new FormattingSettingsGroup({
        name: "labelDetail",
        displayName: "詳細",
        topLevelSlice: this.detailShow,
        slices: [
            this.detailContent,
            this.detailFont,
            this.detailColor,
            this.detailTransparency,
            this.detailUnitType,
            this.detailPrecision,
        ],
    });

    backgroundGroup = new FormattingSettingsGroup({
        name: "labelBackground",
        displayName: "背景",
        topLevelSlice: this.backgroundShow,
        slices: [
            this.backgroundColor,
            this.backgroundTransparency,
        ],
    });

    groups = [
        this.optionsGroup,
        this.valuesGroup,
        this.detailGroup,
        this.backgroundGroup,
    ];

    /**
     * グラフの種類による既定（標準は 100% 積み上げだけ、値がオフで詳細＝全体に対する割合がオン）。
     * 保存していない項目だけ変える。saved はレポートに保存された dataLabels の値。
     * hasDetailField は「ラベルの詳細」にフィールドがあるとき true（コンテンツの既定をカスタムにする）
     */
    applyChartTypeDefaults(chartType: string, saved: powerbi.DataViewObject | undefined, hasDetailField: boolean): void {
        const percent = chartType === CHART_TYPES.stacked100;
        if (saved?.valueShow === undefined) this.valueShow.value = !percent;
        if (saved?.detailShow === undefined) this.detailShow.value = percent;
        if (saved?.detailContent === undefined) {
            this.detailContent.value = itemOf(DETAIL_CONTENT_ITEMS, !percent && hasDetailField ? DETAIL_CONTENTS.custom : DETAIL_CONTENTS.percentOfTotal);
        }
        // 表示単位は数値のフィールドにだけ効くので、割合のときは出さない
        this.detailUnitType.visible = String(this.detailContent.value?.value) === DETAIL_CONTENTS.custom;
    }

    /**
     * 複数系列のとき、「設定の適用先」に系列を並べ、系列ごとに表示とカラーを変えられるようにする。
     * 系列が 1 本なら何も足さない（1.4 までと同じ）
     */
    applyTargets(targets: LabelTarget[]): void {
        const base = [this.optionsGroup, this.valuesGroup, this.detailGroup, this.backgroundGroup];
        if (!targets.length) {
            this.groups = base;
            return;
        }
        const container = new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: targets.slice(0, MAX_COLUMN_TARGETS).map(
                (target) =>
                    new LabelTargetItem(target.name, [
                        new formattingSettings.ToggleSwitch({
                            name: "show",
                            displayName: "このシリーズに表示",
                            value: target.show,
                            selector: target.selector,
                        }),
                        new formattingSettings.ColorPicker({
                            name: "color",
                            displayName: "カラー",
                            value: { value: target.color },
                            selector: target.selector,
                        }),
                    ])
            ),
        });
        this.groups = [new FormattingSettingsGroup({ name: "labelTargets", slices: [], container }), ...base];
    }
}

/**
 * 合計ラベル。積み上げのときだけ効く（棒の端に合計を出す）。
 * 表示単位は Y 軸の表示単位に従う（標準は合計ラベル側でも選べる）
 */
export class TotalLabelsCardSettings extends FormattingSettingsCompositeCard {
    name = "totalLabels";
    displayName = "合計ラベル";
    description = "積み上げのときに、棒の端に合計を出す";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "表示",
        value: false,
    });

    topLevelSlice = this.show;

    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "文字サイズ",
            value: 9,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    /** 「自動」は Y 軸の表示単位に従う。選ぶと、その単位で割って語を付ける（#91） */
    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0], // auto
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    /** オフ（既定）: 正と負を足した合計を 1 つ。オン: 正の合計を上の端、負の合計を下の端に分けて出す */
    splitPositiveNegative = new formattingSettings.ToggleSwitch({
        name: "splitPositiveNegative",
        displayName: "正と負の分割",
        value: false,
    });

    backgroundShow = new formattingSettings.ToggleSwitch({
        name: "backgroundShow",
        displayName: "背景",
        value: false,
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "カラー",
        value: { value: "#FFFFFF" },
    });

    backgroundTransparency = new formattingSettings.NumUpDown({
        name: "backgroundTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "totalValues",
        displayName: "値",
        slices: [this.font, this.color, this.unitType, this.precision, this.splitPositiveNegative],
    });

    backgroundGroup = new FormattingSettingsGroup({
        name: "totalBackground",
        displayName: "背景",
        topLevelSlice: this.backgroundShow,
        slices: [this.backgroundColor, this.backgroundTransparency],
    });

    groups = [this.valuesGroup, this.backgroundGroup];
}

export const LINE_STYLES = {
    solid: "solid",
    dashed: "dashed",
    dotted: "dotted",
} as const;

export const LINE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: LINE_STYLES.dotted, displayName: "点線" },
    { value: LINE_STYLES.dashed, displayName: "破線" },
    { value: LINE_STYLES.solid, displayName: "実線" },
];

/** 線の結合の種類。標準の日本語の表示は「マイター」「四捨五入」「ベベル」（Round の訳）だが、読みやすさを優先して「ラウンド」にする */
export const LINE_JOIN_ITEMS: powerbi.IEnumMember[] = [
    { value: "miter", displayName: "マイター" },
    { value: "round", displayName: "ラウンド" },
    { value: "bevel", displayName: "ベベル" },
];

export const INTERPOLATIONS = {
    linear: "linear",
    smooth: "smooth",
    step: "step",
} as const;

export const INTERPOLATION_ITEMS: powerbi.IEnumMember[] = [
    { value: INTERPOLATIONS.linear, displayName: "線形" },
    { value: INTERPOLATIONS.smooth, displayName: "スムーズ" },
    { value: INTERPOLATIONS.step, displayName: "ステップ" },
];

/** スムーズの種類。標準の日本語の表示は「モノトーン」「基数」（Cardinal の訳）だが、読みやすさを優先して「カーディナル」にする */
export const SMOOTHING_ITEMS: powerbi.IEnumMember[] = [
    { value: "monotone", displayName: "モノトーン" },
    { value: "cardinal", displayName: "カーディナル" },
];

export const STEP_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "before", displayName: "次の値より前" },
    { value: "center", displayName: "中央" },
    { value: "after", displayName: "次の値より後" },
];

/** マーカーの型。標準のドロップダウンは記号だけだが、見分けやすいよう名前を添える */
export const MARKER_SHAPE_ITEMS: powerbi.IEnumMember[] = [
    { value: "circle", displayName: "● 円" },
    { value: "square", displayName: "■ 正方形" },
    { value: "diamond", displayName: "◆ ひし形" },
    { value: "triangle", displayName: "▲ 三角形" },
    { value: "cross", displayName: "× バツ" },
    { value: "shortDash", displayName: "- 短いダッシュ" },
    { value: "longDash", displayName: "― 長いダッシュ" },
    { value: "plus", displayName: "＋ プラス" },
];

const itemOf = (items: powerbi.IEnumMember[], value: string): powerbi.IEnumMember =>
    items.find((i) => i.value === value) ?? items[0];

export const TITLE_STYLES = {
    showTitleOnly: "showTitleOnly",
    showUnitOnly: "showUnitOnly",
    showBoth: "showBoth",
} as const;

export const TITLE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: TITLE_STYLES.showTitleOnly, displayName: "タイトルのみを表示" },
    { value: TITLE_STYLES.showUnitOnly, displayName: "単位のみを表示" },
    { value: TITLE_STYLES.showBoth, displayName: "両方を表示" },
];

export class CategoryAxisCardSettings extends FormattingSettingsCompositeCard {
    name = "categoryAxis";
    displayName = "X 軸";

    // --- 値グループ ---
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "値",
        value: true,
    });

    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "文字サイズ",
            value: 9,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    maxHeight = new formattingSettings.Slider({
        name: "maxHeight",
        displayName: "高さの最大値 (%)",
        value: 25,
    });

    // 「ラベルの連結」は階層（複数フィールド）の軸用、タイトルの「スタイル」は表示単位のある値軸用。
    // このビジュアルのX軸はカテゴリ1列だけで意味を持たないため置かない

    // --- タイトルグループ ---
    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: true,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "titleFontFamily",
            displayName: "フォント",
            value: "DIN",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "titleFontSize",
            displayName: "文字サイズ",
            value: 12,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "titleBold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "titleItalic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "titleUnderline",
            displayName: "下線",
            value: false,
        }),
    });

    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    /** 標準の X 軸と同じ項目。カテゴリの軸には単位が無いので、どれを選んでもタイトルのまま（#91） */
    titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: TITLE_STYLE_ITEMS,
        value: TITLE_STYLE_ITEMS[0],
    });

    // --- レイアウトグループ ---
    minCategoryWidth = new formattingSettings.Slider({
        name: "minCategoryWidth",
        displayName: "カテゴリの最小幅 (px)",
        value: 20,
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "categoryValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: [
            this.font,
            this.labelColor,
            this.maxHeight,
        ],
    });

    titleGroup = new FormattingSettingsGroup({
        name: "categoryTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [
            this.titleText,
            this.titleStyle,
            this.titleFont,
            this.titleColor,
        ],
    });

    layoutGroup = new FormattingSettingsGroup({
        name: "categoryLayout",
        displayName: "レイアウト",
        slices: [
            this.minCategoryWidth,
        ],
    });

    groups = [
        this.valuesGroup,
        this.titleGroup,
        this.layoutGroup,
    ];
}

export class ValueAxisCardSettings extends FormattingSettingsCompositeCard {
    name = "valueAxis";
    displayName = "Y 軸";

    // --- 範囲グループ ---
    start = new formattingSettings.TextInput({
        name: "start",
        displayName: "最小値",
        value: "",
        placeholder: "自動",
    });

    end = new formattingSettings.TextInput({
        name: "end",
        displayName: "最大値",
        value: "",
        placeholder: "自動",
    });

    logarithmic = new formattingSettings.ToggleSwitch({
        name: "logarithmic",
        displayName: "対数目盛り",
        value: false,
    });

    invertRange = new formattingSettings.ToggleSwitch({
        name: "invertRange",
        displayName: "範囲の反転",
        value: false,
    });

    roundRange = new formattingSettings.ToggleSwitch({
        name: "roundRange",
        displayName: "範囲を丸める",
        value: true,
    });

    // --- 値グループ ---
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "値",
        value: true,
    });

    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "文字サイズ",
            value: 9,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0], // auto
    });

    unitNotation = new formattingSettings.ItemDropdown({
        name: "unitNotation",
        displayName: "単位の表記",
        items: UNIT_NOTATIONS,
        value: UNIT_NOTATIONS[0], // japanese
    });

    showUnitOnAxis = new formattingSettings.ToggleSwitch({
        name: "showUnitOnAxis",
        displayName: "軸ラベルに単位を表示",
        value: false,
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0], // auto
    });

    switchPosition = new formattingSettings.ToggleSwitch({
        name: "switchPosition",
        displayName: "軸の位置を切り替える",
        value: false,
    });

    // --- タイトルグループ ---
    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: true,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: TITLE_STYLE_ITEMS,
        value: TITLE_STYLE_ITEMS[0],
    });

    titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "titleFontFamily",
            displayName: "フォント",
            value: "DIN",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "titleFontSize",
            displayName: "文字サイズ",
            value: 12,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "titleBold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "titleItalic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "titleUnderline",
            displayName: "下線",
            value: false,
        }),
    });

    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    // --- 単位ラベルグループ (実務向け独自機能) ---
    unitShow = new formattingSettings.ToggleSwitch({
        name: "unitShow",
        displayName: "単位ラベルの表示",
        value: true,
    });

    unitText = new formattingSettings.TextInput({
        name: "unitText",
        displayName: "単位の追加文字",
        value: "",
        placeholder: "例: 円, 人, 件",
    });

    unitPosition = new formattingSettings.ItemDropdown({
        name: "unitPosition",
        displayName: "位置",
        items: UNIT_POSITIONS,
        value: UNIT_POSITIONS[0], // valueAxisTop
    });

    unitStyle = new formattingSettings.ItemDropdown({
        name: "unitStyle",
        displayName: "スタイル",
        items: UNIT_STYLES,
        value: UNIT_STYLES[0], // parentheses
    });

    unitIncludeDisplayUnit = new formattingSettings.ToggleSwitch({
        name: "unitIncludeDisplayUnit",
        displayName: "表示単位（百万など）を付ける",
        value: true,
    });

    unitFontSize = new formattingSettings.NumUpDown({
        name: "unitFontSize",
        displayName: "単位文字サイズ",
        value: 9,
    });

    unitColor = new formattingSettings.ColorPicker({
        name: "unitColor",
        displayName: "単位カラー",
        value: { value: "#605E5C" },
    });

    rangeGroup = new FormattingSettingsGroup({
        name: "valueRange",
        displayName: "範囲",
        slices: [
            this.start,
            this.end,
            this.logarithmic,
            this.invertRange,
            this.roundRange,
        ],
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "valueValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: [
            this.font,
            this.labelColor,
            this.unitType,
            this.unitNotation,
            this.showUnitOnAxis,
            this.precision,
            this.switchPosition,
        ],
    });

    titleGroup = new FormattingSettingsGroup({
        name: "valueTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [
            this.titleText,
            this.titleStyle,
            this.titleFont,
            this.titleColor,
        ],
    });

    unitGroup = new FormattingSettingsGroup({
        name: "valueUnit",
        displayName: "単位ラベル",
        topLevelSlice: this.unitShow,
        slices: [
            this.unitText,
            this.unitPosition,
            this.unitStyle,
            this.unitIncludeDisplayUnit,
            this.unitFontSize,
            this.unitColor,
        ],
    });

    groups = [
        this.rangeGroup,
        this.valuesGroup,
        this.titleGroup,
        this.unitGroup,
    ];
}

/**
 * 第 2 Y 軸（折れ線の軸）。標準と同じく、表示を保存していなければ自動
 * （棒と範囲が近ければ隠して左の軸を共有、違えば右に出す）。オン・オフを保存するとそれに従う
 */
export class ValueAxis2CardSettings extends FormattingSettingsCompositeCard {
    name = "valueAxis2";
    displayName = "第 2 Y 軸";
    description = "折れ線の軸。棒と尺度が違うときに右に出す";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "表示",
        value: true,
    });

    topLevelSlice = this.show;

    start = new formattingSettings.TextInput({
        name: "start",
        displayName: "最小値",
        value: "",
        placeholder: "自動",
    });

    end = new formattingSettings.TextInput({
        name: "end",
        displayName: "最大値",
        value: "",
        placeholder: "自動",
    });

    valueShow = new formattingSettings.ToggleSwitch({
        name: "valueShow",
        displayName: "値",
        value: true,
    });

    font = new formattingSettings.FontControl({
        name: "font",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "fontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            displayName: "文字サイズ",
            value: 9,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "bold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "italic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "underline",
            displayName: "下線",
            value: false,
        }),
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: true,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "titleFontFamily",
            displayName: "フォント",
            value: "DIN",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "titleFontSize",
            displayName: "文字サイズ",
            value: 12,
        }),
        bold: new formattingSettings.ToggleSwitch({
            name: "titleBold",
            displayName: "太字",
            value: false,
        }),
        italic: new formattingSettings.ToggleSwitch({
            name: "titleItalic",
            displayName: "斜体",
            value: false,
        }),
        underline: new formattingSettings.ToggleSwitch({
            name: "titleUnderline",
            displayName: "下線",
            value: false,
        }),
    });

    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    // --- 表示単位・タイトルのスタイル・単位ラベル（Y 軸と同じ。単位ラベルは第 2 Y 軸の上に出す） ---
    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0], // auto
    });

    unitNotation = new formattingSettings.ItemDropdown({
        name: "unitNotation",
        displayName: "単位の表記",
        items: UNIT_NOTATIONS,
        value: UNIT_NOTATIONS[0], // japanese
    });

    showUnitOnAxis = new formattingSettings.ToggleSwitch({
        name: "showUnitOnAxis",
        displayName: "軸ラベルに単位を表示",
        value: false,
    });

    titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: TITLE_STYLE_ITEMS,
        value: TITLE_STYLE_ITEMS[0],
    });

    unitShow = new formattingSettings.ToggleSwitch({
        name: "unitShow",
        displayName: "単位ラベルの表示",
        value: true,
    });

    unitText = new formattingSettings.TextInput({
        name: "unitText",
        displayName: "単位の追加文字",
        value: "",
        placeholder: "例: 円, 人, 件",
    });

    unitStyle = new formattingSettings.ItemDropdown({
        name: "unitStyle",
        displayName: "スタイル",
        items: UNIT_STYLES,
        value: UNIT_STYLES[0], // parentheses
    });

    unitIncludeDisplayUnit = new formattingSettings.ToggleSwitch({
        name: "unitIncludeDisplayUnit",
        displayName: "表示単位（百万など）を付ける",
        value: true,
    });

    unitFontSize = new formattingSettings.NumUpDown({
        name: "unitFontSize",
        displayName: "単位文字サイズ",
        value: 9,
    });

    unitColor = new formattingSettings.ColorPicker({
        name: "unitColor",
        displayName: "単位カラー",
        value: { value: "#605E5C" },
    });

    // --- 範囲（標準の第 2 Y 軸と同じ項目、#91） ---
    logarithmic = new formattingSettings.ToggleSwitch({
        name: "logarithmic",
        displayName: "対数目盛り",
        value: false,
    });

    roundRange = new formattingSettings.ToggleSwitch({
        name: "roundRange",
        displayName: "範囲を丸める",
        value: true,
    });

    /** 第 2 Y 軸に 0 を含め、その高さを Y 軸の 0 にそろえる（英語 UI の Align zeros） */
    alignZeros = new formattingSettings.ToggleSwitch({
        name: "alignZeros",
        displayName: "0 を配置する",
        value: false,
    });

    rangeGroup = new FormattingSettingsGroup({
        name: "value2Range",
        displayName: "範囲",
        slices: [this.start, this.end, this.logarithmic, this.roundRange, this.alignZeros],
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "value2Values",
        displayName: "値",
        topLevelSlice: this.valueShow,
        slices: [this.font, this.labelColor, this.unitType, this.unitNotation, this.showUnitOnAxis, this.precision],
    });

    titleGroup = new FormattingSettingsGroup({
        name: "value2Title",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [this.titleText, this.titleStyle, this.titleFont, this.titleColor],
    });

    unitGroup = new FormattingSettingsGroup({
        name: "value2Unit",
        displayName: "単位ラベル",
        topLevelSlice: this.unitShow,
        slices: [this.unitText, this.unitStyle, this.unitIncludeDisplayUnit, this.unitFontSize, this.unitColor],
    });

    groups = [this.rangeGroup, this.valuesGroup, this.titleGroup, this.unitGroup];
}

/** 折れ線 1 本ぶんの書式の対象 */
export interface LineTarget {
    name: string;
    selector: powerbi.data.Selector;
    color: string;
    width: number;
    lineStyle: string;
    lineJoin: string;
    interpolation: string;
    smoothing: string;
    /** テンション (%) */
    tension: number;
    stepPosition: string;
    /** ステップの段と段をつなぐ線を出すか。オフなら値ごとの水平な線だけ（横棒では垂直な線だけ） */
    stepConnect: boolean;
    /** 網掛け領域を出すか（網掛け領域の「このシリーズに表示」） */
    areaShow: boolean;
    /** 線を出すか（線の「このシリーズに表示」） */
    lineShow: boolean;
}

/** 線の形の値（「すべて」と線ごとに同じ項目） */
export type LineShapeValues = Pick<LineTarget, "lineJoin" | "interpolation" | "smoothing" | "tension" | "stepPosition" | "stepConnect">;

/** 線の形の既定（標準と同じ。テンションは Desktop のスライダーの位置から読んだ値。段のつなぎは標準に無い項目で、既定は標準と同じくつなぐ） */
export const LINE_SHAPE_DEFAULTS = {
    lineJoin: "round",
    interpolation: INTERPOLATIONS.linear,
    smoothing: "monotone",
    tension: 60,
    stepPosition: "center",
    stepConnect: true,
} as const;

class LineTargetItem extends FormattingSettingsCard {
    name = "lineTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

/** 線の形のスライス。スムーズの種類とテンション・ステップの位置は、補間の種類に合うときだけ出す（標準と同じ） */
function lineShapeSlices(
    values: LineShapeValues,
    selector?: powerbi.data.Selector
): FormattingSettingsSlice[] {
    const smooth = values.interpolation === INTERPOLATIONS.smooth;
    return [
        new formattingSettings.ItemDropdown({
            name: "lineJoin",
            displayName: "結合の種類",
            items: LINE_JOIN_ITEMS,
            value: itemOf(LINE_JOIN_ITEMS, values.lineJoin),
            selector,
        }),
        new formattingSettings.ItemDropdown({
            name: "interpolation",
            displayName: "補間の種類",
            items: INTERPOLATION_ITEMS,
            value: itemOf(INTERPOLATION_ITEMS, values.interpolation),
            selector,
        }),
        new formattingSettings.ItemDropdown({
            name: "smoothing",
            displayName: "スムーズの種類",
            items: SMOOTHING_ITEMS,
            value: itemOf(SMOOTHING_ITEMS, values.smoothing),
            selector,
            visible: smooth,
        }),
        new formattingSettings.Slider({
            name: "tension",
            displayName: "テンション (%)",
            value: values.tension,
            selector,
            visible: smooth && values.smoothing === "cardinal",
        }),
        new formattingSettings.ItemDropdown({
            name: "stepPosition",
            displayName: "ステップの位置",
            items: STEP_POSITION_ITEMS,
            value: itemOf(STEP_POSITION_ITEMS, values.stepPosition),
            selector,
            visible: values.interpolation === INTERPOLATIONS.step,
        }),
        new formattingSettings.ToggleSwitch({
            name: "stepConnect",
            displayName: "段のつなぎを表示",
            value: values.stepConnect,
            selector,
            visible: values.interpolation === INTERPOLATIONS.step,
        }),
    ];
}

/**
 * 折れ線の見た目。「設定の適用先」に折れ線を並べ、線ごとにカラー・幅・線のスタイル・結合と補間を変えられる。
 * 標準の日本語の表示は「行」（Lines の訳）だが、読みやすさを優先して「線」にする
 */
export class LinesCardSettings extends FormattingSettingsCompositeCard {
    name = "lines";
    displayName = "線";

    /**
     * 線を出すか（標準の「すべての系列に表示」）。既定はオンで、横棒では保存していなければオフ
     * （横棒のカテゴリは順番に意味の無いものが多いので、線でつながずマーカーだけ出す、#103）
     */
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "すべての系列に表示",
        value: true,
    });

    /** 線ごとの色の実体。「すべて」には出さない（色は線ごとにテーマの色を割り当てる） */
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "カラー",
        value: { value: "#12239E" },
    });

    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 3,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: LINE_STYLE_ITEMS.find((i) => i.value === LINE_STYLES.solid) ?? LINE_STYLE_ITEMS[0],
    });

    lineJoin = new formattingSettings.ItemDropdown({
        name: "lineJoin",
        displayName: "結合の種類",
        items: LINE_JOIN_ITEMS,
        value: itemOf(LINE_JOIN_ITEMS, LINE_SHAPE_DEFAULTS.lineJoin),
    });

    interpolation = new formattingSettings.ItemDropdown({
        name: "interpolation",
        displayName: "補間の種類",
        items: INTERPOLATION_ITEMS,
        value: itemOf(INTERPOLATION_ITEMS, LINE_SHAPE_DEFAULTS.interpolation),
    });

    smoothing = new formattingSettings.ItemDropdown({
        name: "smoothing",
        displayName: "スムーズの種類",
        items: SMOOTHING_ITEMS,
        value: itemOf(SMOOTHING_ITEMS, LINE_SHAPE_DEFAULTS.smoothing),
    });

    tension = new formattingSettings.Slider({
        name: "tension",
        displayName: "テンション (%)",
        value: LINE_SHAPE_DEFAULTS.tension,
    });

    stepPosition = new formattingSettings.ItemDropdown({
        name: "stepPosition",
        displayName: "ステップの位置",
        items: STEP_POSITION_ITEMS,
        value: itemOf(STEP_POSITION_ITEMS, LINE_SHAPE_DEFAULTS.stepPosition),
    });

    /** ステップの段と段をつなぐ線を出すか（ユーザーの提案。オフなら値ごとの水準の線だけが並ぶ） */
    stepConnect = new formattingSettings.ToggleSwitch({
        name: "stepConnect",
        displayName: "段のつなぎを表示",
        value: LINE_SHAPE_DEFAULTS.stepConnect,
    });

    /** 「すべて」の値（保存値か既定） */
    shapeValues(): LineShapeValues {
        const dropdown = (slice: formattingSettings.ItemDropdown, fallback: string) => String(slice.value?.value ?? fallback);
        return {
            lineJoin: dropdown(this.lineJoin, LINE_SHAPE_DEFAULTS.lineJoin),
            interpolation: dropdown(this.interpolation, LINE_SHAPE_DEFAULTS.interpolation),
            smoothing: dropdown(this.smoothing, LINE_SHAPE_DEFAULTS.smoothing),
            tension: typeof this.tension.value === "number" ? this.tension.value : LINE_SHAPE_DEFAULTS.tension,
            stepPosition: dropdown(this.stepPosition, LINE_SHAPE_DEFAULTS.stepPosition),
            stepConnect: this.stepConnect.value ?? LINE_SHAPE_DEFAULTS.stepConnect,
        };
    }

    private allItem(): FormattingSettingsCard {
        const values = this.shapeValues();
        this.smoothing.visible = values.interpolation === INTERPOLATIONS.smooth;
        this.tension.visible = this.smoothing.visible && values.smoothing === "cardinal";
        this.stepPosition.visible = values.interpolation === INTERPOLATIONS.step;
        this.stepConnect.visible = this.stepPosition.visible;
        return new LineTargetItem("すべて", [
            this.show,
            this.lineStyle,
            this.lineJoin,
            this.width,
            this.interpolation,
            this.smoothing,
            this.tension,
            this.stepPosition,
            this.stepConnect,
        ]);
    }

    targetGroup = new FormattingSettingsGroup({
        name: "lineTargets",
        slices: [],
        container: new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: [this.allItem()],
        }),
    });

    groups = [this.targetGroup];

    applyTargets(targets: LineTarget[]): void {
        this.targetGroup.container = new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: [
                this.allItem(),
                ...targets.slice(0, MAX_COLUMN_TARGETS).map((target) => {
                    const [lineJoin, ...interpolation] = lineShapeSlices(target, target.selector);
                    // 並びは標準と同じ（線のスタイル・結合の種類・幅・補間の種類…）
                    return new LineTargetItem(target.name, [
                        new formattingSettings.ToggleSwitch({
                            name: "show",
                            displayName: "このシリーズに表示",
                            value: target.lineShow,
                            selector: target.selector,
                        }),
                        new formattingSettings.ColorPicker({
                            name: "fill",
                            displayName: "カラー",
                            value: { value: target.color },
                            selector: target.selector,
                        }),
                        new formattingSettings.ItemDropdown({
                            name: "lineStyle",
                            displayName: "線のスタイル",
                            items: LINE_STYLE_ITEMS,
                            value: LINE_STYLE_ITEMS.find((i) => i.value === target.lineStyle) ?? LINE_STYLE_ITEMS[0],
                            selector: target.selector,
                        }),
                        lineJoin,
                        new formattingSettings.NumUpDown({
                            name: "width",
                            displayName: "幅 (px)",
                            value: target.width,
                            selector: target.selector,
                        }),
                        ...interpolation,
                    ]);
                }),
            ],
        });
    }
}

/**
 * リボン。積み上げ・100% 積み上げで、同じ系列を隣のカテゴリの棒と帯でつなぐ（標準の積み上げ縦棒の「リボン」カードと同じく、
 * 見出しのトグルで出す、#108）。「積む順」を値の大きい順にすると、標準のリボン グラフと同じく順位の入れ替わりが帯で見える。
 * 標準は系列ごとにも変えられるが、ここでは「すべて」だけ（系列ごとは見送り）
 */
export class RibbonsCardSettings extends FormattingSettingsCompositeCard {
    name = "ribbons";
    displayName = "リボン";
    description = "積み上げの棒を、隣のカテゴリの同じ系列と帯でつなぐ";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "リボン",
        value: false,
    });

    topLevelSlice = this.show;

    order = new formattingSettings.ItemDropdown({
        name: "order",
        displayName: "積む順",
        items: RIBBON_ORDER_ITEMS,
        value: RIBBON_ORDER_ITEMS[0],
    });

    matchSeriesColor = new formattingSettings.ToggleSwitch({
        name: "matchSeriesColor",
        displayName: "データ系列の色を一致させる",
        value: true,
    });

    /** 系列の色に合わせないときの帯の色（標準はオフで灰色） */
    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "カラー",
        value: { value: "#C8C6C4" },
    });

    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 30,
    });

    borderShow = new formattingSettings.ToggleSwitch({
        name: "borderShow",
        displayName: "罫線",
        value: false,
    });

    borderMatchRibbon = new formattingSettings.ToggleSwitch({
        name: "borderMatchRibbon",
        displayName: "リボンの色を合わせる",
        value: false,
    });

    borderFill = new formattingSettings.ColorPicker({
        name: "borderFill",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    borderTransparency = new formattingSettings.NumUpDown({
        name: "borderTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "幅 (px)",
        value: 1,
    });

    /** 棒の端と帯の端のすき間（カテゴリの間のすき間に対する %）。既定 0 で棒に接する */
    spacing = new formattingSettings.NumUpDown({
        name: "spacing",
        displayName: "リボンと棒の間のスペース (%)",
        value: 0,
    });

    colorGroup = new FormattingSettingsGroup({
        name: "ribbonColor",
        displayName: "カラー",
        slices: [this.matchSeriesColor, this.fill, this.transparency],
    });

    borderGroup = new FormattingSettingsGroup({
        name: "ribbonBorder",
        displayName: "罫線",
        topLevelSlice: this.borderShow,
        slices: [this.borderMatchRibbon, this.borderFill, this.borderTransparency, this.borderWidth],
    });

    layoutGroup = new FormattingSettingsGroup({
        name: "ribbonLayout",
        displayName: "レイアウト",
        slices: [this.order, this.spacing],
    });

    groups = [this.colorGroup, this.borderGroup, this.layoutGroup];
}

/**
 * 折れ線の点の印。標準と同じく既定は出さず、「すべてのカテゴリに表示」で出す。
 * 標準の「設定の適用先」はカテゴリごとにも選べるが、ここでは「すべて」だけ（カテゴリごとは見送り）
 */
export class MarkersCardSettings extends FormattingSettingsCompositeCard {
    name = "markers";
    displayName = "マーカー";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "すべてのカテゴリに表示",
        value: false,
    });

    shape = new formattingSettings.ItemDropdown({
        name: "shape",
        displayName: "型",
        items: MARKER_SHAPE_ITEMS,
        value: MARKER_SHAPE_ITEMS[0],
    });

    size = new formattingSettings.NumUpDown({
        name: "size",
        displayName: "サイズ (px)",
        value: 5,
    });

    /** 空 = 線の色（標準の既定と同じ） */
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "" },
    });

    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    borderShow = new formattingSettings.ToggleSwitch({
        name: "borderShow",
        displayName: "罫線",
        value: false,
    });

    borderMatchLine = new formattingSettings.ToggleSwitch({
        name: "borderMatchLine",
        displayName: "線の色を一致させる",
        value: false,
    });

    borderFill = new formattingSettings.ColorPicker({
        name: "borderFill",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    borderTransparency = new formattingSettings.NumUpDown({
        name: "borderTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "幅 (px)",
        value: 1,
    });

    optionsGroup = new FormattingSettingsGroup({
        name: "markerOptions",
        displayName: "設定の適用先",
        slices: [this.show],
    });

    shapeGroup = new FormattingSettingsGroup({
        name: "markerShape",
        displayName: "シェイプ",
        slices: [this.shape, this.size],
    });

    colorGroup = new FormattingSettingsGroup({
        name: "markerColor",
        displayName: "カラー",
        slices: [this.color, this.transparency],
    });

    borderGroup = new FormattingSettingsGroup({
        name: "markerBorder",
        displayName: "罫線",
        topLevelSlice: this.borderShow,
        slices: [this.borderMatchLine, this.borderFill, this.borderTransparency, this.borderWidth],
    });

    groups = [this.optionsGroup, this.shapeGroup, this.colorGroup, this.borderGroup];
}

class AreaTargetItem extends FormattingSettingsCard {
    name = "areaTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

/**
 * 網掛け領域。折れ線と値 0 の高さのあいだを、線の色（既定）を透かして塗る。
 * 「設定の適用先」に折れ線を並べ、線ごとに出す・出さないを選べる（標準と同じ）
 */
export class AreasCardSettings extends FormattingSettingsCompositeCard {
    name = "areas";
    displayName = "網掛け領域";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "網掛け領域",
        value: false,
    });

    topLevelSlice = this.show;

    matchLineColor = new formattingSettings.ToggleSwitch({
        name: "matchLineColor",
        displayName: "線の色を一致させる",
        value: true,
    });

    fill = new formattingSettings.ColorPicker({
        name: "fill",
        displayName: "カラー",
        value: { value: "#118DFF" },
    });

    transparency = new formattingSettings.Slider({
        name: "transparency",
        displayName: "領域の透過性 (%)",
        value: 60,
    });

    colorGroup = new FormattingSettingsGroup({
        name: "areaColor",
        displayName: "カラー",
        slices: [this.matchLineColor, this.fill, this.transparency],
    });

    groups: FormattingSettingsGroup[] = [this.colorGroup];

    applyTargets(targets: LineTarget[]): void {
        if (!targets.length) {
            this.groups = [this.colorGroup];
            return;
        }
        const container = new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: targets.slice(0, MAX_COLUMN_TARGETS).map(
                (target) =>
                    new AreaTargetItem(target.name, [
                        new formattingSettings.ToggleSwitch({
                            name: "show",
                            displayName: "このシリーズに表示",
                            value: target.areaShow,
                            selector: target.selector,
                        }),
                    ])
            ),
        });
        this.groups = [new FormattingSettingsGroup({ name: "areaTargets", slices: [], container }), this.colorGroup];
    }
}

export class GridlinesCardSettings extends FormattingSettingsCompositeCard {
    name = "gridlines";
    displayName = "グリッド線";

    // --- 横 (水平・Y軸目盛線) ---
    horizontalShow = new formattingSettings.ToggleSwitch({
        name: "horizontalShow",
        displayName: "横",
        value: true,
    });

    horizontalColor = new formattingSettings.ColorPicker({
        name: "horizontalColor",
        displayName: "カラー",
        value: { value: "#E1DFDD" },
    });

    horizontalTransparency = new formattingSettings.NumUpDown({
        name: "horizontalTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    horizontalStyle = new formattingSettings.ItemDropdown({
        name: "horizontalStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: LINE_STYLE_ITEMS[0], // 点線
    });

    /** 点線・破線の模様を線の幅に合わせて伸び縮みさせる（標準の「幅で拡大縮小」、#91） */
    horizontalScaleWithWidth = new formattingSettings.ToggleSwitch({
        name: "horizontalScaleWithWidth",
        displayName: "幅で拡大縮小",
        value: false,
    });

    horizontalWidth = new formattingSettings.NumUpDown({
        name: "horizontalWidth",
        displayName: "幅 (px)",
        value: 1,
    });

    // --- 縦 (垂直・X軸目盛線) ---
    verticalShow = new formattingSettings.ToggleSwitch({
        name: "verticalShow",
        displayName: "縦",
        value: false,
    });

    verticalColor = new formattingSettings.ColorPicker({
        name: "verticalColor",
        displayName: "カラー",
        value: { value: "#E1DFDD" },
    });

    verticalTransparency = new formattingSettings.NumUpDown({
        name: "verticalTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    verticalStyle = new formattingSettings.ItemDropdown({
        name: "verticalStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: LINE_STYLE_ITEMS[0], // 点線
    });

    verticalScaleWithWidth = new formattingSettings.ToggleSwitch({
        name: "verticalScaleWithWidth",
        displayName: "幅で拡大縮小",
        value: false,
    });

    verticalWidth = new formattingSettings.NumUpDown({
        name: "verticalWidth",
        displayName: "幅 (px)",
        value: 1,
    });

    horizontalGroup = new FormattingSettingsGroup({
        name: "horizontalGridlines",
        displayName: "横",
        topLevelSlice: this.horizontalShow,
        slices: [
            this.horizontalColor,
            this.horizontalTransparency,
            this.horizontalStyle,
            this.horizontalScaleWithWidth,
            this.horizontalWidth,
        ],
    });

    verticalGroup = new FormattingSettingsGroup({
        name: "verticalGridlines",
        displayName: "縦",
        topLevelSlice: this.verticalShow,
        slices: [
            this.verticalColor,
            this.verticalTransparency,
            this.verticalStyle,
            this.verticalScaleWithWidth,
            this.verticalWidth,
        ],
    });

    groups = [this.horizontalGroup, this.verticalGroup];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chart = new ChartCardSettings();
    categoryAxis = new CategoryAxisCardSettings();
    valueAxis = new ValueAxisCardSettings();
    valueAxis2 = new ValueAxis2CardSettings();
    legend = new LegendCardSettings();
    gridlines = new GridlinesCardSettings();
    columns = new ColumnsCardSettings();
    lines = new LinesCardSettings();
    areas = new AreasCardSettings();
    markers = new MarkersCardSettings();
    ribbons = new RibbonsCardSettings();
    dataLabels = new DataLabelsCardSettings();
    totalLabels = new TotalLabelsCardSettings();

    // 標準の複合グラフと同じ並び（グラフの種類のあと、X 軸・Y 軸・第 2 Y 軸・凡例・グリッド線・棒・線・網掛け領域・マーカー・データ ラベル・合計ラベル）。
    // リボンは標準のリボン グラフと同じく棒のあと
    cards = [
        this.chart,
        this.categoryAxis,
        this.valueAxis,
        this.valueAxis2,
        this.legend,
        this.gridlines,
        this.columns,
        this.ribbons,
        this.lines,
        this.areas,
        this.markers,
        this.dataLabels,
        this.totalLabels,
    ];

    /**
     * データが来たあとで「設定の適用先」を作る。seriesMode は複数系列のとき true
     * （列は系列が対象になり、データラベルにも系列ごとの設定が付く）。lineTargets は折れ線
     */
    applyTargets(
        targets: ColumnTarget[],
        options: { seriesMode?: boolean; labelTargets?: LabelTarget[]; lineTargets?: LineTarget[] } = {}
    ): void {
        this.columns.applyTargets(targets, options.seriesMode ?? false);
        this.dataLabels.applyTargets(options.labelTargets ?? []);
        this.lines.applyTargets(options.lineTargets ?? []);
        this.areas.applyTargets(options.lineTargets ?? []);
        this.applyOrientation();
    }

    /**
     * 横棒のときは、標準と同じくカードの名前を画面の向きで付け直す（値の軸が「X 軸」、カテゴリの軸が「Y 軸」）。
     * カテゴリの軸の「高さの最大値」「カテゴリの最小幅」も、横棒では「最大幅」「最小カテゴリの高さ」になる。
     * 書式の名前（保存先）は変えないので、向きを切り替えても設定はそのまま
     */
    applyOrientation(): void {
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        this.categoryAxis.displayName = horizontal ? "Y 軸" : "X 軸";
        this.valueAxis2.displayName = horizontal ? "第 2 X 軸" : "第 2 Y 軸";
        this.valueAxis.displayName = horizontal ? "X 軸" : "Y 軸";
        this.categoryAxis.maxHeight.displayName = horizontal ? "最大幅 (%)" : "高さの最大値 (%)";
        this.categoryAxis.minCategoryWidth.displayName = horizontal ? "最小カテゴリの高さ (px)" : "カテゴリの最小幅 (px)";
        const [first, second] = horizontal ? [this.valueAxis, this.categoryAxis] : [this.categoryAxis, this.valueAxis];
        // グラフの種類はいつもいちばん上
        const rest = this.cards.filter((c) => c !== this.chart && c !== this.categoryAxis && c !== this.valueAxis);
        this.cards = [this.chart, first, second, ...rest];
    }

    /**
     * グラフの種類による既定を、レポートに保存していない項目に当てる（#92）。
     * populateFormattingSettingsModel のあと、transform の前に呼ぶ
     */
    applyChartTypeDefaults(dataView: powerbi.DataView | undefined): void {
        // 1.14 までは種類と向きを columns に保存していた。新しいカードに保存が無ければ、古い保存を使う
        const objects = dataView?.metadata?.objects;
        const legacy = (property: "chartType" | "orientation", slice: formattingSettings.ItemDropdown) => {
            const old = objects?.columns?.[property];
            if (objects?.chart?.[property] === undefined && old !== undefined && old !== null) {
                const item = slice.items.find((i) => i.value === String(old));
                if (item) slice.value = item;
            }
        };
        legacy("chartType", this.chart.chartType);
        legacy("orientation", this.chart.orientation);
        // 1.18 までは種類に「リボン」があった。積み上げ＋リボン オン＋値の大きい順で開く（見た目は同じ、#108）
        const savedType = objects?.chart?.chartType ?? objects?.columns?.chartType;
        if (String(savedType) === CHART_TYPES.ribbon) {
            this.chart.chartType.value = itemOf(CHART_TYPE_ITEMS, CHART_TYPES.stacked);
            if (objects?.ribbons?.show === undefined) this.ribbons.show.value = true;
            if (objects?.ribbons?.order === undefined) this.ribbons.order.value = itemOf(RIBBON_ORDER_ITEMS, RIBBON_ORDERS.value);
        }

        // 横棒では、線はつながずマーカーだけ出すのを既定にする（保存していなければ、#103）
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        if (objects?.lines?.show === undefined) this.lines.show.value = !horizontal;
        if (objects?.markers?.show === undefined) this.markers.show.value = horizontal;

        const chartType = String(this.chart.chartType.value?.value ?? CHART_TYPES.clustered);
        const saved = objects?.dataLabels;
        const hasDetailField = !!dataView?.metadata?.columns?.some((c) => c.roles?.labelDetail);
        this.dataLabels.applyChartTypeDefaults(chartType, saved, hasDetailField);
    }

    /**
     * グラフの種類とデータに関係ないカードを隠す（標準はその種類で使うカードだけを出す、#90）。
     * 隠すのは書式ペインの表示だけで、保存済みの値は残る。hasLines は「折れ線の値」にフィールドがあるとき true
     */
    applyCardVisibility(hasLines: boolean): void {
        const chartType = String(this.chart.chartType.value?.value ?? CHART_TYPES.clustered);
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        // 折れ線の値は、どの種類・向きでも描く（横棒は既定でマーカーだけ、#103）
        const linesDrawn = hasLines;
        this.totalLabels.visible = chartType === CHART_TYPES.stacked;
        // リボンは積み上げ・100% 積み上げで出せる（縦棒・横棒とも。集合は棒が横に並ぶので帯でつながない、#108）
        this.ribbons.visible = chartType === CHART_TYPES.stacked || chartType === CHART_TYPES.stacked100;
        this.valueAxis2.visible = linesDrawn;
        this.lines.visible = linesDrawn;
        this.areas.visible = linesDrawn;
        this.markers.visible = linesDrawn;
    }
}
