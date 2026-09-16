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

/** グラフの種類。標準では別々のビジュアルだが、barChart は書式ペインで切り替える */
export const CHART_TYPES = {
    clustered: "clustered",
    stacked: "stacked",
    stacked100: "stacked100",
    ribbon: "ribbon",
} as const;

export type ChartType = (typeof CHART_TYPES)[keyof typeof CHART_TYPES];

export const CHART_TYPE_ITEMS: powerbi.IEnumMember[] = [
    { value: CHART_TYPES.clustered, displayName: "集合" },
    { value: CHART_TYPES.stacked, displayName: "積み上げ" },
    { value: CHART_TYPES.stacked100, displayName: "100% 積み上げ" },
    { value: CHART_TYPES.ribbon, displayName: "リボン" },
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

export class ColumnsCardSettings extends FormattingSettingsCompositeCard {
    name = "columns";
    displayName = "列";
    analyticsPane = false;

    /** 既定は集合（系列 1 本なら 1.4 までと同じ見た目） */
    chartType = new formattingSettings.ItemDropdown({
        name: "chartType",
        displayName: "グラフの種類",
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
        displayName: "列の色を一致させる",
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

    typeGroup = new FormattingSettingsGroup({
        name: "columnsType",
        displayName: "グラフの種類",
        slices: [this.chartType, this.orientation],
    });

    groups = [this.typeGroup, this.targetGroup, this.layoutGroup];

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
                                displayName: "列の色を一致させる",
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
        slices: [
            this.fontFamily,
            this.fontSize,
            this.bold,
            this.italic,
            this.color,
            this.precision,
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
        this.backgroundGroup,
    ];

    /**
     * 複数系列のとき、「設定の適用先」に系列を並べ、系列ごとに表示とカラーを変えられるようにする。
     * 系列が 1 本なら何も足さない（1.4 までと同じ）
     */
    applyTargets(targets: LabelTarget[]): void {
        const base = [this.optionsGroup, this.valuesGroup, this.backgroundGroup];
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
        slices: [this.font, this.color, this.precision, this.splitPositiveNegative],
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

    rangeGroup = new FormattingSettingsGroup({
        name: "value2Range",
        displayName: "範囲",
        slices: [this.start, this.end],
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
}

class LineTargetItem extends FormattingSettingsCard {
    name = "lineTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

/**
 * 折れ線の見た目。「設定の適用先」に折れ線を並べ、線ごとにカラー・幅・線のスタイルを変えられる。
 * 標準の日本語の表示は「行」（Lines の訳）だが、読みやすさを優先して「線」にする
 */
export class LinesCardSettings extends FormattingSettingsCompositeCard {
    name = "lines";
    displayName = "線";

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

    private allItem(): FormattingSettingsCard {
        return new LineTargetItem("すべて", [this.width, this.lineStyle]);
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
                ...targets.slice(0, MAX_COLUMN_TARGETS).map(
                    (target) =>
                        new LineTargetItem(target.name, [
                            new formattingSettings.ColorPicker({
                                name: "fill",
                                displayName: "カラー",
                                value: { value: target.color },
                                selector: target.selector,
                            }),
                            new formattingSettings.NumUpDown({
                                name: "width",
                                displayName: "幅 (px)",
                                value: target.width,
                                selector: target.selector,
                            }),
                            new formattingSettings.ItemDropdown({
                                name: "lineStyle",
                                displayName: "線のスタイル",
                                items: LINE_STYLE_ITEMS,
                                value: LINE_STYLE_ITEMS.find((i) => i.value === target.lineStyle) ?? LINE_STYLE_ITEMS[0],
                                selector: target.selector,
                            }),
                        ])
                ),
            ],
        });
    }
}

/**
 * リボン（グラフの種類がリボンのときだけ効く）。同じ系列を隣のカテゴリとつなぐ帯の見た目。
 * 標準は系列ごとにも変えられるが、ここでは「すべて」だけ（系列ごとは見送り）
 */
export class RibbonsCardSettings extends FormattingSettingsCompositeCard {
    name = "ribbons";
    displayName = "リボン";
    description = "グラフの種類がリボンのときに、同じ系列を隣のカテゴリとつなぐ帯";

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

    /** 列の端と帯の端のすき間（カテゴリの間のすき間に対する %）。既定 0 で列に接する */
    spacing = new formattingSettings.NumUpDown({
        name: "spacing",
        displayName: "リボンと列の間のスペース (%)",
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
        slices: [this.spacing],
    });

    groups = [this.colorGroup, this.borderGroup, this.layoutGroup];
}

/** 折れ線の点の印。標準と同じく既定は出さない */
export class MarkersCardSettings extends FormattingSettingsCompositeCard {
    name = "markers";
    displayName = "マーカー";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "表示",
        value: false,
    });

    topLevelSlice = this.show;

    size = new formattingSettings.NumUpDown({
        name: "size",
        displayName: "サイズ (px)",
        value: 5,
    });

    shapeGroup = new FormattingSettingsGroup({
        name: "markerShape",
        displayName: "シェイプ",
        slices: [this.size],
    });

    groups = [this.shapeGroup];
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
            this.verticalWidth,
        ],
    });

    groups = [this.horizontalGroup, this.verticalGroup];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    categoryAxis = new CategoryAxisCardSettings();
    valueAxis = new ValueAxisCardSettings();
    valueAxis2 = new ValueAxis2CardSettings();
    legend = new LegendCardSettings();
    gridlines = new GridlinesCardSettings();
    columns = new ColumnsCardSettings();
    lines = new LinesCardSettings();
    markers = new MarkersCardSettings();
    ribbons = new RibbonsCardSettings();
    dataLabels = new DataLabelsCardSettings();
    totalLabels = new TotalLabelsCardSettings();

    // 標準の複合グラフと同じ並び（X 軸・Y 軸・第 2 Y 軸・凡例・グリッド線・列・線・マーカー・データ ラベル・合計ラベル）。
    // リボンは標準のリボン グラフと同じく列のあと
    cards = [
        this.categoryAxis,
        this.valueAxis,
        this.valueAxis2,
        this.legend,
        this.gridlines,
        this.columns,
        this.ribbons,
        this.lines,
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
        this.applyOrientation();
    }

    /**
     * 横棒のときは、標準と同じくカードの名前を画面の向きで付け直す（値の軸が「X 軸」、カテゴリの軸が「Y 軸」）。
     * カテゴリの軸の「高さの最大値」「カテゴリの最小幅」も、横棒では「最大幅」「最小カテゴリの高さ」になる。
     * 書式の名前（保存先）は変えないので、向きを切り替えても設定はそのまま
     */
    applyOrientation(): void {
        const horizontal = String(this.columns.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        this.categoryAxis.displayName = horizontal ? "Y 軸" : "X 軸";
        this.valueAxis.displayName = horizontal ? "X 軸" : "Y 軸";
        this.categoryAxis.maxHeight.displayName = horizontal ? "最大幅 (%)" : "高さの最大値 (%)";
        this.categoryAxis.minCategoryWidth.displayName = horizontal ? "最小カテゴリの高さ (px)" : "カテゴリの最小幅 (px)";
        const [first, second] = horizontal ? [this.valueAxis, this.categoryAxis] : [this.categoryAxis, this.valueAxis];
        this.cards = [first, second, ...this.cards.filter((c) => c !== this.categoryAxis && c !== this.valueAxis)];
    }
}
