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

import { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS } from "./shared/units";
import { LEGEND_POSITIONS, LEGEND_POSITION_ITEMS, standardLegendPosition, legendPlacementValue } from "./shared/legend";
import { AUTO_PLACEHOLDER, AutoNumUpDown, itemOf } from "./shared/formatting";
import { NEGATIVE_STYLE_ITEMS, ZERO_STYLE_ITEMS } from "./shared/numberFormat";

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

export { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS, AUTO_PLACEHOLDER, AutoNumUpDown };

export const UNIT_POSITIONS = [
    { value: "valueAxisTop", displayName: "Y 軸の上 (左上)" },
    { value: "plotTopRight", displayName: "プロット エリアの右上" },
    { value: "none", displayName: "非表示" },
];

/** 表示される形そのものを選択肢の名前にする。value は保存済みレポートとの互換のため変えない */
export const UNIT_STYLES = [
    { value: "parentheses", displayName: "(百万円)" },
    { value: "withPrefix", displayName: "(単位: 百万円)" },
];

/**
 * グラフの種類。標準では別々のビジュアルだが、barChart は書式ペインで切り替える。
 * ribbon は 1.18 までの保存値だけ（リボンは種類ではなく「リボン」カードの見せ方にした）
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
 * グラフの種類と向き。いちばん最初に選ぶものなので、書式ペインのいちばん上のカードにする。
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

    /** ドリルダウンしたとき、今いる位置（事業A ＞ 製品A1 など）を左上に出す。標準に無い項目 */
    drillPathShow = new formattingSettings.ToggleSwitch({
        name: "drillPathShow",
        displayName: "ドリルの位置",
        description: "ドリルダウンや絞り込みで、表示している項目の上の階層が 1 つに決まるとき、その位置（事業A ＞ 製品A1 など）を左上に出す",
        value: true,
    });

    slices = [this.chartType, this.orientation, this.drillPathShow];
}

/** 棒の値の計算。なし（素の値）・累計・パレート */
export const CALCULATION_MODES = {
    none: "none",
    cumulative: "cumulative",
    pareto: "pareto",
} as const;

export const CALCULATION_ITEMS: powerbi.IEnumMember[] = [
    { value: CALCULATION_MODES.none, displayName: "なし" },
    { value: CALCULATION_MODES.cumulative, displayName: "累計" },
    { value: CALCULATION_MODES.pareto, displayName: "パレート" },
];

/** 累計を 0 に戻さない（区切りの選択肢の先頭）。ほかの選択肢は階層のレベルの queryName */
export const CUMULATIVE_RESET_NONE = "none";

const CUMULATIVE_RESET_NONE_ITEM: powerbi.IEnumMember = { value: CUMULATIVE_RESET_NONE, displayName: "区切らない" };

/**
 * 累計・パレート。棒の値を並んだ順に足していく。区切り（階層のレベル）が変わったところで 0 に戻す。
 * 「切り替えボタンを出す」をオンにすると、閲覧者がグラフの上のボタンで累計を切り替えられる（押した状態はレポートに保存する）
 */
export class CalculationCardSettings extends FormattingSettingsCard {
    name = "calculation";
    displayName = "累計・パレート";

    mode = new formattingSettings.ItemDropdown({
        name: "mode",
        displayName: "計算",
        items: CALCULATION_ITEMS,
        value: CALCULATION_ITEMS[0],
    });

    /** 選択肢はデータ次第（X 軸の階層のレベル）なので、update のたびに applyCumulativeLevels で組み直す */
    cumulativeReset = new formattingSettings.ItemDropdown({
        name: "cumulativeReset",
        displayName: "区切り",
        items: [CUMULATIVE_RESET_NONE_ITEM],
        value: CUMULATIVE_RESET_NONE_ITEM,
    });

    cumulativeToggle = new formattingSettings.ToggleSwitch({
        name: "cumulativeToggle",
        displayName: "切り替えボタンを出す",
        value: false,
    });

    // --- パレート ---
    // options は付けない（素の numeric に付けると書式ペインが空になった記録がある）。範囲は viewModel でクランプする
    /** 累積比がここまでを A にする（「上位 2 割で 8 割」の 8 割）。境目ちょうどは A に入れる */
    boundaryAB = new formattingSettings.NumUpDown({
        name: "boundaryAB",
        displayName: "A と B の境目 (%)",
        value: 80,
    });

    boundaryBC = new formattingSettings.NumUpDown({
        name: "boundaryBC",
        displayName: "B と C の境目 (%)",
        value: 95,
    });

    /** 棒をランクの色で塗る。凡例（系列）があるときは系列の色のまま */
    colorByRank = new formattingSettings.ToggleSwitch({
        name: "colorByRank",
        displayName: "ランクで色分け",
        value: true,
    });

    colorA = new formattingSettings.ColorPicker({
        name: "colorA",
        displayName: "A の色",
        value: { value: "#118DFF" },
    });

    colorB = new formattingSettings.ColorPicker({
        name: "colorB",
        displayName: "B の色",
        value: { value: "#74B9FF" },
    });

    colorC = new formattingSettings.ColorPicker({
        name: "colorC",
        displayName: "C の色",
        value: { value: "#C4E1FF" },
    });

    ratioColor = new formattingSettings.ColorPicker({
        name: "ratioColor",
        displayName: "累積比の線の色",
        value: { value: "#E66C37" },
    });

    /** 累積比の軸に、境目の高さの破線を引く */
    showThresholds = new formattingSettings.ToggleSwitch({
        name: "showThresholds",
        displayName: "境目に線を引く",
        value: true,
    });

    /** X 軸の下に A・B・C の帯を出す。帯を押すと、そのランクの棒をまとめて選ぶ */
    showRankBand = new formattingSettings.ToggleSwitch({
        name: "showRankBand",
        displayName: "ランクの帯",
        value: true,
    });

    labelA = new formattingSettings.TextInput({ name: "labelA", displayName: "A の名前", value: "A", placeholder: "A" });
    labelB = new formattingSettings.TextInput({ name: "labelB", displayName: "B の名前", value: "B", placeholder: "B" });
    labelC = new formattingSettings.TextInput({ name: "labelC", displayName: "C の名前", value: "C", placeholder: "C" });

    slices = [
        this.mode,
        this.cumulativeReset,
        this.cumulativeToggle,
        this.boundaryAB,
        this.boundaryBC,
        this.colorByRank,
        this.colorA,
        this.colorB,
        this.colorC,
        this.ratioColor,
        this.showThresholds,
        this.showRankBand,
        this.labelA,
        this.labelB,
        this.labelC,
    ];

    /** 区切りの選択肢を階層のレベル（いちばん下を除く）で組み直し、累計のときだけ区切りとボタンを出す */
    applyCumulativeLevels(levels: powerbi.IEnumMember[], current: string): void {
        const items = [CUMULATIVE_RESET_NONE_ITEM, ...levels];
        this.cumulativeReset.items = items;
        this.cumulativeReset.value = items.find((item) => item.value === current) ?? items[0];
        const mode = String(this.mode.value?.value ?? CALCULATION_MODES.none);
        const cumulative = mode === CALCULATION_MODES.cumulative;
        this.cumulativeReset.visible = cumulative && levels.length > 0;
        this.cumulativeToggle.visible = cumulative;
        // パレートの項目はパレートのときだけ。色はランクで色分けするときだけ
        const pareto = mode === CALCULATION_MODES.pareto;
        [this.boundaryAB, this.boundaryBC, this.colorByRank, this.ratioColor, this.showThresholds, this.showRankBand].forEach((s) => (s.visible = pareto));
        // 色は棒の色分けにもランクの帯にも使う。名前は帯にもツールヒントにも使うので、パレートならいつも出す
        [this.colorA, this.colorB, this.colorC].forEach((s) => (s.visible = pareto && ((this.colorByRank.value ?? true) || (this.showRankBand.value ?? true))));
        [this.labelA, this.labelB, this.labelC].forEach((s) => (s.visible = pareto));
    }
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

    // --- その他にまとめる ---
    /**
     * 値の大きい順に上位の件数だけ棒を出し、残りを 1 本の「その他」にまとめて最後に置く。0 ならまとめない。
     * X 軸の階層を展開しているときは使わない。options は付けない（範囲は viewModel でクランプする）
     */
    otherCount = new formattingSettings.NumUpDown({
        name: "otherCount",
        displayName: "上位の件数",
        value: 0,
    });

    otherLabel = new formattingSettings.TextInput({
        name: "otherLabel",
        displayName: "名前",
        value: "その他",
        placeholder: "その他",
    });

    /** 系列が 1 本のときの「その他」の棒の色（系列があれば系列の色） */
    otherFill = new formattingSettings.ColorPicker({
        name: "otherFill",
        displayName: "色",
        value: { value: "#A0A0A0" },
    });

    otherGroup = new FormattingSettingsGroup({
        name: "columnsOther",
        displayName: "その他にまとめる",
        slices: [this.otherCount, this.otherLabel, this.otherFill],
    });

    groups = [this.targetGroup, this.layoutGroup, this.otherGroup];

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

export { LEGEND_POSITIONS, LEGEND_POSITION_ITEMS, standardLegendPosition, legendPlacementValue };

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

    /** マイナスの書き方（-・▲・△・括弧）。既定は - */
    negativeStyle = new formattingSettings.ItemDropdown({
        name: "negativeStyle",
        displayName: "マイナス",
        items: NEGATIVE_STYLE_ITEMS,
        value: NEGATIVE_STYLE_ITEMS[0],
    });

    /** 0（丸めて 0 になる値を含む）の書き方（0・±0・-）。既定は 0 */
    zeroStyle = new formattingSettings.ItemDropdown({
        name: "zeroStyle",
        displayName: "0",
        items: ZERO_STYLE_ITEMS,
        value: ZERO_STYLE_ITEMS[0],
    });

    /** 丸めて 0 になるマイナスに符号を残す（▲0）。切ると 0 の書き方にそろえる */
    negativeZero = new formattingSettings.ToggleSwitch({
        name: "negativeZero",
        displayName: "丸めて 0 のマイナスに符号",
        value: true,
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
            this.negativeStyle,
            this.zeroStyle,
            this.negativeZero,
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

    /** 「自動」は Y 軸の表示単位に従う。選ぶと、その単位で割って語を付ける */
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

/** 段のつなぎを出さないステップの、値ごとの線の長さ（標準に無い項目） */
export const STEP_WIDTHS = {
    /** カテゴリの間隔いっぱい（隣のカテゴリとのすき間の真ん中から真ん中まで）。1.21 までと同じ */
    step: "step",
    /** そのカテゴリの棒のまとまりの幅。棒ごとの目標の印に使う */
    bar: "bar",
} as const;

export const STEP_WIDTH_ITEMS: powerbi.IEnumMember[] = [
    { value: STEP_WIDTHS.step, displayName: "カテゴリの間隔" },
    { value: STEP_WIDTHS.bar, displayName: "棒の幅" },
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

/** 階層の上のレベルの見せ方。区切り線は標準と同じ、囲みは標準に無い見せ方 */
export const HIERARCHY_STYLES = {
    lines: "lines",
    boxed: "boxed",
} as const;

export const HIERARCHY_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: HIERARCHY_STYLES.lines, displayName: "区切り線" },
    { value: HIERARCHY_STYLES.boxed, displayName: "囲み" },
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

    /**
     * 階層を全部展開したときのラベルの見せ方。既定（オフ）は標準と同じく、いちばん下のレベルの下に
     * 上のレベルを段に重ねる。オンなら「FY26 H1 Q1」のように 1 行につなぐ
     */
    concatenateLabels = new formattingSettings.ToggleSwitch({
        name: "concatenateLabels",
        displayName: "ラベルの連結",
        value: false,
    });

    /**
     * 段に重ねた上のレベルの見せ方。区切り線（既定）は標準と同じく点線で区切る。
     * 囲みは区切りごとに角の丸い淡い枠で囲む（標準に無い）
     */
    hierarchyStyle = new formattingSettings.ItemDropdown({
        name: "hierarchyStyle",
        displayName: "階層の見せ方",
        items: HIERARCHY_STYLE_ITEMS,
        value: HIERARCHY_STYLE_ITEMS[0],
    });

    // --- タイトルグループ ---
    /**
     * カテゴリの軸のタイトルは初期オフ（2026-09-23 決定）。項目名で何の軸か読めることが多く、
     * 縦横どちらの向きでもカテゴリの軸（この card）に当てる。数値の軸のタイトルは初期オンのまま
     */
    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: false,
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

    /** 標準の X 軸と同じ項目。カテゴリの軸には単位が無いので、どれを選んでもタイトルのまま */
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
            this.concatenateLabels,
            this.hierarchyStyle,
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

    /**
     * 目盛り（グリッド線）の本数の目安。空なら自動（描く範囲の長さで決める。標準と同じ）。
     * 数を入れると、その本数以内で切りのいい目盛りを選ぶ。データが変わっても間隔が追従する
     */
    tickCount = new formattingSettings.TextInput({
        name: "tickCount",
        displayName: "目盛りの本数 (目安)",
        description: "空なら自動。数を入れると、その本数以内で切りのいい目盛りにする（数字が重なるなら減らす）。第 2 Y 軸も同じ上限。対数の軸では効かない（10 の累乗で決まる）",
        value: "",
        placeholder: "自動",
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
            this.tickCount,
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

    // --- 範囲（標準の第 2 Y 軸と同じ項目） ---
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
    /** 段のつなぎを出さないときの、値ごとの線の長さ（STEP_WIDTHS） */
    stepWidth: string;
    /** 網掛け領域を出すか（網掛け領域の「このシリーズに表示」） */
    areaShow: boolean;
    /** 線を出すか（線の「このシリーズに表示」） */
    lineShow: boolean;
    /** 棒を累計にしているとき、この線も累計にするか */
    includeCumulative: boolean;
}

/** 線の形の値（「すべて」と線ごとに同じ項目） */
export type LineShapeValues = Pick<LineTarget, "lineJoin" | "interpolation" | "smoothing" | "tension" | "stepPosition" | "stepConnect" | "stepWidth">;

/**
 * 線の形の既定（標準と同じ。テンションは Desktop のスライダーの位置から読んだ値。
 * 段のつなぎと段の幅は標準に無い項目で、既定は標準と同じくつなぎ、1.21 までと同じくカテゴリの間隔いっぱい）
 */
export const LINE_SHAPE_DEFAULTS = {
    lineJoin: "round",
    interpolation: INTERPOLATIONS.linear,
    smoothing: "monotone",
    tension: 60,
    stepPosition: "center",
    stepConnect: true,
    stepWidth: STEP_WIDTHS.step,
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
        new formattingSettings.ItemDropdown({
            name: "stepWidth",
            displayName: "段の幅",
            items: STEP_WIDTH_ITEMS,
            value: itemOf(STEP_WIDTH_ITEMS, values.stepWidth),
            selector,
            visible: values.interpolation === INTERPOLATIONS.step && !values.stepConnect,
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
     * （横棒のカテゴリは順番に意味の無いものが多いので、線でつながずマーカーだけ出す）
     */
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "すべての系列に表示",
        value: true,
    });

    /**
     * 棒を累計にしているとき、線も累計にするか（既定はオフ。累計と実績を並べて見ることが多いため）。
     * 棒が累計のときだけ出す（applyCardVisibility）
     */
    includeCumulative = new formattingSettings.ToggleSwitch({
        name: "includeCumulative",
        displayName: "累計に含める",
        value: false,
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

    /** 段のつなぎを出さないときの線の長さ。「棒の幅」なら棒ごとの目標の印になる */
    stepWidth = new formattingSettings.ItemDropdown({
        name: "stepWidth",
        displayName: "段の幅",
        items: STEP_WIDTH_ITEMS,
        value: itemOf(STEP_WIDTH_ITEMS, LINE_SHAPE_DEFAULTS.stepWidth),
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
            stepWidth: dropdown(this.stepWidth, LINE_SHAPE_DEFAULTS.stepWidth),
        };
    }

    private allItem(): FormattingSettingsCard {
        const values = this.shapeValues();
        this.smoothing.visible = values.interpolation === INTERPOLATIONS.smooth;
        this.tension.visible = this.smoothing.visible && values.smoothing === "cardinal";
        this.stepPosition.visible = values.interpolation === INTERPOLATIONS.step;
        this.stepConnect.visible = this.stepPosition.visible;
        this.stepWidth.visible = this.stepPosition.visible && !values.stepConnect;
        return new LineTargetItem("すべて", [
            this.show,
            this.includeCumulative,
            this.lineStyle,
            this.lineJoin,
            this.width,
            this.interpolation,
            this.smoothing,
            this.tension,
            this.stepPosition,
            this.stepConnect,
            this.stepWidth,
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
                        ...(this.includeCumulative.visible
                            ? [
                                new formattingSettings.ToggleSwitch({
                                    name: "includeCumulative",
                                    displayName: "累計に含める",
                                    value: target.includeCumulative,
                                    selector: target.selector,
                                }),
                            ]
                            : []),
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
 * 見出しのトグルで出す）。「積む順」を値の大きい順にすると、標準のリボン グラフと同じく順位の入れ替わりが帯で見える。
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

    /** 点線・破線の模様を線の幅に合わせて伸び縮みさせる（標準の「幅で拡大縮小」） */
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
    calculation = new CalculationCardSettings();
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
        this.calculation,
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
     * 系列 1 本の棒の色は、保存が無ければテーマの 1 番目で塗る。書式ペインの「カラー」にも実際に塗る色を出す（保存はしない）。
     * null も保存なしとして扱う（viewModel の customColor と同じ）
     */
    applySingleSeriesFill(seriesMode: boolean, fill: string, objects: powerbi.DataViewObjects | undefined): void {
        if (!seriesMode && objects?.columns?.fill == null) this.columns.fill.value = { value: fill };
    }

    /**
     * 基本テーマ・カスタムテーマに合わせる。テーマは標準のビジュアルの名前（showAxisTitle・showTitle）で値を持つので、
     * capabilities にその名前も置いて受け取り、作り手が自作の設定（titleShow）を保存していないときの既定にする。
     * 書式ペインにも同じ値を出す（「既定値にリセット」でテーマの値に戻る）。populate の直後に呼ぶ。
     * 凡例の位置は 1.26 までの保存値（topLeft など）を標準の値に読み替える
     */
    applyThemeDefaults(objects: powerbi.DataViewObjects | undefined): void {
        const raw = (card: string, prop: string): unknown => objects?.[card]?.[prop];
        const inherit = (slice: formattingSettings.ToggleSwitch, card: string, own: string, standard: string) => {
            const themeValue = raw(card, standard);
            if (raw(card, own) == null && typeof themeValue === "boolean") slice.value = themeValue;
        };
        inherit(this.categoryAxis.titleShow, "categoryAxis", "titleShow", "showAxisTitle");
        inherit(this.valueAxis.titleShow, "valueAxis", "titleShow", "showAxisTitle");
        inherit(this.legend.titleShow, "legend", "titleShow", "showTitle");
        const position = standardLegendPosition(raw("legend", "position"));
        if (position) this.legend.position.value = LEGEND_POSITION_ITEMS.find((i) => i.value === position)!;

        // 項目名の欄の上限：標準は categoryAxis の maxMarginFactor で持つ（Fluent 2 は 50）。描画と同じ 0〜100 に丸める
        const marginFactor = raw("categoryAxis", "maxMarginFactor");
        if (raw("categoryAxis", "maxHeight") == null && typeof marginFactor === "number") {
            this.categoryAxis.maxHeight.value = Math.max(0, Math.min(100, marginFactor));
        }

        // グリッド線：標準は軸のカードの中（gridlineShow・gridlineColor・gridlineStyle・gridlineThickness）。
        // 自作の「グリッド線」カードの横（horizontal*）は数値の軸の線、縦（vertical*）はカテゴリの軸の線（縦棒・横棒で入れ替わらない）
        const gridlines: Array<[axis: string, prefix: "horizontal" | "vertical"]> = [
            ["valueAxis", "horizontal"],
            ["categoryAxis", "vertical"],
        ];
        for (const [axis, prefix] of gridlines) {
            const own = (prop: string) => raw("gridlines", `${prefix}${prop}`) != null;
            const show = raw(axis, "gridlineShow");
            if (!own("Show") && typeof show === "boolean") this.gridlines[`${prefix}Show`].value = show;
            const color = (raw(axis, "gridlineColor") as powerbi.Fill | undefined)?.solid?.color;
            if (!own("Color") && typeof color === "string" && color) this.gridlines[`${prefix}Color`].value = { value: color };
            const style = LINE_STYLE_ITEMS.find((i) => i.value === raw(axis, "gridlineStyle"));
            if (!own("Style") && style) this.gridlines[`${prefix}Style`].value = style;
            const thickness = raw(axis, "gridlineThickness");
            // 描画と同じ 1〜10 に丸める（書式ペインの値と描画の幅をずらさない）
            if (!own("Width") && typeof thickness === "number" && thickness > 0) this.gridlines[`${prefix}Width`].value = Math.max(1, Math.min(10, thickness));
        }
    }

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
     * グラフの種類による既定を、レポートに保存していない項目に当てる。
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
        // 1.18 までは種類に「リボン」があった。積み上げ＋リボン オン＋値の大きい順で開く（見た目は同じ）
        const savedType = objects?.chart?.chartType ?? objects?.columns?.chartType;
        if (String(savedType) === CHART_TYPES.ribbon) {
            this.chart.chartType.value = itemOf(CHART_TYPE_ITEMS, CHART_TYPES.stacked);
            if (objects?.ribbons?.show === undefined) this.ribbons.show.value = true;
            if (objects?.ribbons?.order === undefined) this.ribbons.order.value = itemOf(RIBBON_ORDER_ITEMS, RIBBON_ORDERS.value);
        }

        // 横棒では、線はつながずマーカーだけ出すのを既定にする（保存していなければ）
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        if (objects?.lines?.show === undefined) this.lines.show.value = !horizontal;
        if (objects?.markers?.show === undefined) this.markers.show.value = horizontal;

        const chartType = String(this.chart.chartType.value?.value ?? CHART_TYPES.clustered);
        const saved = objects?.dataLabels;
        const hasDetailField = !!dataView?.metadata?.columns?.some((c) => c.roles?.labelDetail);
        this.dataLabels.applyChartTypeDefaults(chartType, saved, hasDetailField);
    }

    /**
     * グラフの種類とデータに関係ないカードを隠す（標準はその種類で使うカードだけを出す）。
     * 隠すのは書式ペインの表示だけで、保存済みの値は残る。hasLines は「折れ線の値」にフィールドがあるとき true
     */
    applyCardVisibility(hasLines: boolean): void {
        const chartType = String(this.chart.chartType.value?.value ?? CHART_TYPES.clustered);
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        // 折れ線の値は、どの種類・向きでも描く（横棒は既定でマーカーだけ）
        const linesDrawn = hasLines;
        this.totalLabels.visible = chartType === CHART_TYPES.stacked;
        // リボンは積み上げ・100% 積み上げで出せる（縦棒・横棒とも。集合は棒が横に並ぶので帯でつながない）
        this.ribbons.visible = chartType === CHART_TYPES.stacked || chartType === CHART_TYPES.stacked100;
        this.valueAxis2.visible = linesDrawn;
        this.lines.visible = linesDrawn;
        this.areas.visible = linesDrawn;
        this.markers.visible = linesDrawn;
    }
}
