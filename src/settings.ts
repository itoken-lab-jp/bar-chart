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

    private allItem(): FormattingSettingsCard {
        return new TargetItem("すべて", [
            this.fill,
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
            this.maxBarWidth,
            this.cornerRadius,
        ],
    });

    groups = [this.targetGroup, this.layoutGroup];

    applyTargets(targets: ColumnTarget[]): void {
        this.targetGroup.container = new FormattingSettingsContainer({
            displayName: "設定の適用先",
            containerItems: [
                this.allItem(),
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
    gridlines = new GridlinesCardSettings();
    columns = new ColumnsCardSettings();
    dataLabels = new DataLabelsCardSettings();

    cards = [
        this.categoryAxis,
        this.valueAxis,
        this.gridlines,
        this.columns,
        this.dataLabels,
    ];

    applyTargets(targets: ColumnTarget[]): void {
        this.columns.applyTargets(targets);
    }
}
