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
import { SCROLL_START_ITEMS } from "./shared/scrollStart";
import { NEGATIVE_STYLE_ITEMS, ZERO_STYLE_ITEMS, SIGN_TONE_MODE_ITEMS, DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR } from "./shared/numberFormat";
import { CUSTOM_LINE_STYLE, CUSTOM_LINE_STYLE_ITEM, DASH_CAP_ITEMS } from "./shared/gridlines";
import { legendParts, gridlineParts, categoryAxisParts, valueAxisParts, labelValueParts, labelBackgroundParts, labelHaloParts } from "./shared/formatCards";

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

/** 比較レイヤーの手前の棒を置く側。縦棒は右・左、横棒は下・上 */
export const COMPARE_DIRECTIONS = {
    rightFront: "rightFront",
    leftFront: "leftFront",
} as const;

export const COMPARE_DIRECTION_ITEMS: powerbi.IEnumMember[] = [
    { value: COMPARE_DIRECTIONS.rightFront, displayName: "右（横棒は下）" },
    { value: COMPARE_DIRECTIONS.leftFront, displayName: "左（横棒は上）" },
];

/** 「比較の列」で受けたとき、どの値を手前にするか（「手前にする値の名前」が空か、名前が見つからないとき） */
export const COMPARE_ORDERS = {
    lastFront: "lastFront",
    firstFront: "firstFront",
} as const;

export const COMPARE_ORDER_ITEMS: powerbi.IEnumMember[] = [
    { value: COMPARE_ORDERS.lastFront, displayName: "並びの最後の値" },
    { value: COMPARE_ORDERS.firstFront, displayName: "並びの最初の値" },
];

/** X 軸の定数線を、印の付いたカテゴリのどこに引くか。前は期の境目（そのカテゴリの始まり） */
export const CATEGORY_LINE_POSITIONS = {
    before: "before",
    center: "center",
    after: "after",
} as const;

export const CATEGORY_LINE_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_POSITIONS.before, displayName: "カテゴリの前" },
    { value: CATEGORY_LINE_POSITIONS.center, displayName: "カテゴリの中央" },
    { value: CATEGORY_LINE_POSITIONS.after, displayName: "カテゴリの後ろ" },
];

/** 比較レイヤーの枚数の上限（「値」1 枚と「比較値」4 枚。「比較の列」なら値 5 つ） */
export const MAX_COMPARE_LAYERS = 5;

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

    /** マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。標準に無い項目 */
    copyButton = new formattingSettings.ToggleSwitch({
        name: "copyButton",
        displayName: "画像のコピー",
        description: "マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。押すと、見えているグラフを画像としてクリップボードに入れ、PowerPoint などに貼れる",
        value: true,
    });

    slices = [this.chartType, this.orientation, this.drillPathShow, this.copyButton];
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

    /**
     * 積み上げで、積んだ棒と棒のすき間（標準の積み上げの「系列間のスペース」）。系列の展開がオフなら px
     * （下の棒との間を空け、値の軸はそのまま）、オンなら %（プロットの長さの半分に対する %）。既定は 0
     */
    stackedSpacing = new formattingSettings.NumUpDown({
        name: "stackedSpacing",
        displayName: "系列間のスペース (px)",
        value: 0,
    });

    /** 積み上げで、凡例の順と逆の順に積む（標準の積み上げの複合の「積み上げ順を逆にする」） */
    reverseStackOrder = new formattingSettings.ToggleSwitch({
        name: "reverseStackOrder",
        displayName: "積み上げ順を逆にする",
        value: false,
    });

    /** 系列の展開：積んだ棒を離して並べ、値の軸を消す（標準と同じ。折れ線があるときは使えない） */
    stackedExplode = new formattingSettings.ToggleSwitch({
        name: "stackedExplode",
        displayName: "系列の展開",
        value: false,
    });

    /** 積み上げで、罫線を積んだ棒の外側だけに引く（標準の「内側の罫線を非表示」） */
    borderOutlineOnly = new formattingSettings.ToggleSwitch({
        name: "borderOutlineOnly",
        displayName: "内側の罫線を非表示",
        value: false,
    });

    /** 集合で、系列の棒を重ねる（標準の「重複」）。オンなら系列間のスペースを重ねる割合（%）として読む */
    overlap = new formattingSettings.ToggleSwitch({
        name: "overlap",
        displayName: "重複",
        value: false,
    });

    /** 重なった棒の前後を入れ替える（既定は凡例の後ろの系列が手前。標準の「重複を反転する」） */
    overlapReverse = new formattingSettings.ToggleSwitch({
        name: "overlapReverse",
        displayName: "重複を反転する",
        value: false,
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
            this.borderOutlineOnly,
        ]);
    }

    /**
     * グラフの種類で、レイアウトの項目を出し分ける。集合の系列間のスペース（%）と、積み上げの系列間のスペース・系列の展開は別の項目
     * （標準も積み上げでは単位が px か % になる）。系列の展開は折れ線があると使えない
     */
    applyChartType(stacked: boolean, hasLines: boolean): void {
        this.seriesSpacing.visible = !stacked;
        this.overlap.visible = !stacked;
        this.overlapReverse.visible = !stacked && (this.overlap.value ?? false);
        this.stackedSpacing.visible = stacked;
        this.stackedExplode.visible = stacked && !hasLines;
        this.borderOutlineOnly.visible = stacked;
        this.reverseStackOrder.visible = stacked;
        const explode = stacked && !hasLines && (this.stackedExplode.value ?? false);
        this.stackedSpacing.displayName = explode ? "系列間のスペース (%)" : "系列間のスペース (px)";
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
            this.overlap,
            this.overlapReverse,
            this.stackedSpacing,
            this.stackedExplode,
            this.reverseStackOrder,
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

/** データ ラベルのタイトルのコンテンツ。標準の「カスタム」はフィールドを選ぶので、ここでは「ラベルのタイトル」に入れたフィールド */
export const TITLE_CONTENTS = {
    seriesName: "seriesName",
    custom: "custom",
} as const;

export const TITLE_CONTENT_ITEMS: powerbi.IEnumMember[] = [
    { value: TITLE_CONTENTS.seriesName, displayName: "系列名" },
    { value: TITLE_CONTENTS.custom, displayName: "カスタム（ラベルのタイトル）" },
];

/** データ ラベルの行（タイトル・値・詳細）を縦に積むか、1 行に並べるか（標準の「レイアウト」） */
export const LABEL_LAYOUTS = {
    multiLine: "multiLine",
    singleLine: "singleLine",
} as const;

export const LABEL_LAYOUT_ITEMS: powerbi.IEnumMember[] = [
    { value: LABEL_LAYOUTS.multiLine, displayName: "複数行" },
    { value: LABEL_LAYOUTS.singleLine, displayName: "単一行" },
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

    private parts = legendParts({ showDisplayName: "表示", titleShow: true, font: { controlName: "font", displayName: "フォント", size: 10 } });
    show = this.parts.show;
    topLevelSlice = this.show;
    position = this.parts.position;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;

    /** 折れ線の凡例の印（標準の複合の「スタイル」）。折れ線があるときだけ出す */
    markerStyle = new formattingSettings.ItemDropdown({
        name: "legendMarkerRendering",
        displayName: "スタイル",
        items: LEGEND_MARKER_STYLE_ITEMS,
        value: itemOf(LEGEND_MARKER_STYLE_ITEMS, LEGEND_MARKER_STYLES.lineAndMarker),
    });

    /** 凡例の折れ線のマーカーを、マーカーの色ではなく線の色で塗る（標準の「線の色に合わせる」） */
    matchLineColor = new formattingSettings.ToggleSwitch({
        name: "matchLineColor",
        displayName: "線の色に合わせる",
        value: false,
    });

    /** 凡例の項目を逆の順に並べる（棒の並び・積む順は変えない。積み上げの上下と凡例の上下を合わせるとき） */
    reverseOrder = new formattingSettings.ToggleSwitch({
        name: "reverseOrder",
        displayName: "表示順を反転",
        value: false,
    });

    /** 棒の項目と折れ線の項目を別々に出し入れする（2026-10-08 ユーザー「棒の凡例と線の凡例で個別で表示するしないの設定がほしい」） */
    showBars = new formattingSettings.ToggleSwitch({
        name: "showBars",
        displayName: "棒の項目",
        description: "凡例に棒（系列）の項目を出す。切ると折れ線の項目だけ",
        value: true,
    });
    showLines = new formattingSettings.ToggleSwitch({
        name: "showLines",
        displayName: "折れ線の項目",
        description: "凡例に折れ線の項目を出す。切ると棒の項目だけ",
        value: true,
    });

    optionsGroup = new FormattingSettingsGroup({
        name: "legendOptions",
        displayName: "オプション",
        slices: [this.position, this.showBars, this.showLines, this.markerStyle, this.matchLineColor, this.reverseOrder],
    });
    textGroup = this.parts.textGroup;
    titleGroup = this.parts.titleGroup;

    groups = [this.optionsGroup, this.textGroup, this.titleGroup];

    /** 折れ線の印の項目は、折れ線があるときだけ出す */
    applyLines(hasLines: boolean): void {
        // 棒と折れ線の両方があるときだけ、どちらを出すかを選べる
        this.showBars.visible = hasLines;
        this.showLines.visible = hasLines;
        this.markerStyle.visible = hasLines;
        this.matchLineColor.visible = hasLines;
    }
}

export class DataLabelsCardSettings extends FormattingSettingsCompositeCard {
    name = "dataLabels";
    displayName = "データ ラベル";

    private values = labelValueParts({
        precisions: PRECISIONS,
        negativeStyles: NEGATIVE_STYLE_ITEMS,
        zeroStyles: ZERO_STYLE_ITEMS,
        toneModes: SIGN_TONE_MODE_ITEMS,
        goodColor: DEFAULT_GOOD_COLOR,
        badColor: DEFAULT_BAD_COLOR,
    });
    private background = labelBackgroundParts({ showDisplayName: "背景の表示", colorDisplayName: "背景色", color: "#FFFFFF", transparency: 0 });

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

    /** 標準と同じ。オンにすると、最大幅より長いラベルの行を「…」で切る */
    optimizeLabelDisplay = new formattingSettings.ToggleSwitch({
        name: "optimizeLabelDisplay",
        displayName: "ラベルの表示を最適化",
        description: "切り捨てと幅の制限を使用して可読性を高める",
        value: false,
    });

    /** ラベルの表示を最適化のときの、ラベルの行の幅の上限。標準と同じ既定 200 px */
    labelMaxWidth = new formattingSettings.NumUpDown({
        name: "labelMaxWidth",
        displayName: "最大幅 (px)",
        value: 200,
    });

    /** 詳細がカスタム（ラベルの詳細のフィールド）で、その値が空白のときに出す文字。空なら詳細の行を出さない */
    detailShowBlankAs = new formattingSettings.TextInput({
        name: "detailShowBlankAs",
        displayName: "空白の表示方法",
        value: "",
        placeholder: "",
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

    underline = new formattingSettings.ToggleSwitch({
        name: "underline",
        displayName: "下線",
        value: false,
    });

    color = this.values.color;

    /** 値の文字の透過性（標準の「値 > 透過性」） */
    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 0,
    });
    precision = this.values.precision;
    negativeStyle = this.values.negativeStyle;
    zeroStyle = this.values.zeroStyle;
    negativeZero = this.values.negativeZero;
    toneMode = this.values.toneMode;
    positiveColor = this.values.positiveColor;
    negativeColor = this.values.negativeColor;

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

    /** 値の行の上に、系列名かフィールド（ラベルのタイトル）の行を出す（標準の「タイトル」）。既定はオフ */
    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: false,
    });

    titleContent = new formattingSettings.ItemDropdown({
        name: "titleContent",
        displayName: "コンテンツ",
        items: TITLE_CONTENT_ITEMS,
        value: TITLE_CONTENT_ITEMS[0],
    });

    titleFont = new formattingSettings.FontControl({
        name: "titleFont",
        displayName: "フォント",
        fontFamily: new formattingSettings.FontPicker({
            name: "titleFontFamily",
            displayName: "フォント",
            value: "Segoe UI",
        }),
        fontSize: new formattingSettings.NumUpDown({
            name: "titleFontSize",
            displayName: "文字サイズ",
            value: 9,
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

    /** 空 = 自動（値の行と同じく、棒の色に合わせて白か黒） */
    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "" },
    });

    titleTransparency = new formattingSettings.NumUpDown({
        name: "titleTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    /** カスタム（ラベルのタイトルのフィールド）の数値の表示単位。「自動」はフィールドの書式のまま */
    titleUnitType = new formattingSettings.ItemDropdown({
        name: "titleUnitType",
        displayName: "表示単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0],
    });

    titlePrecision = new formattingSettings.ItemDropdown({
        name: "titlePrecision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    /** タイトルがカスタムで、フィールドの値が空白のときに出す文字。空ならタイトルの行を出さない */
    titleShowBlankAs = new formattingSettings.TextInput({
        name: "titleShowBlankAs",
        displayName: "空白の表示方法",
        value: "",
        placeholder: "",
    });

    /** 「ラベルの値」のフィールドが空白のときに出す文字。空なら棒の値を出す */
    valueShowBlankAs = new formattingSettings.TextInput({
        name: "valueShowBlankAs",
        displayName: "空白の表示方法",
        value: "",
        placeholder: "",
    });

    /** タイトル・値・詳細を縦に積む（複数行）か、1 行に並べる（単一行）か */
    labelContentLayout = new formattingSettings.ItemDropdown({
        name: "labelContentLayout",
        displayName: "レイアウト",
        items: LABEL_LAYOUT_ITEMS,
        value: LABEL_LAYOUT_ITEMS[0],
    });

    /** 複数行のとき、行どうしを左・中央・右のどれでそろえるか */
    horizontalAlignment = new formattingSettings.AlignmentGroup({
        name: "horizontalAlignment",
        displayName: "水平方向の配置",
        mode: "horizontalAlignment" as powerbi.visuals.AlignmentGroupMode,
        value: "center",
    });

    /** 文字のまわりに縁を付ける（標準に無い項目。棒や線の上でも読める） */
    private halo = labelHaloParts();
    haloShow = this.halo.haloShow;
    haloColor = this.halo.haloColor;
    haloWidth = this.halo.haloWidth;
    haloGroup = this.halo.haloGroup;

    backgroundShow = this.background.backgroundShow;
    backgroundColor = this.background.backgroundColor;
    backgroundTransparency = this.background.backgroundTransparency;

    titleGroup = new FormattingSettingsGroup({
        name: "labelTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [
            this.titleContent,
            this.titleFont,
            this.titleColor,
            this.titleTransparency,
            this.titleUnitType,
            this.titlePrecision,
            this.titleShowBlankAs,
        ],
    });

    layoutGroup = new FormattingSettingsGroup({
        name: "labelLayout",
        displayName: "レイアウト",
        slices: [this.labelContentLayout, this.horizontalAlignment],
    });

    optionsGroup = new FormattingSettingsGroup({
        name: "labelOptions",
        displayName: "オプション",
        slices: [
            this.orientation,
            this.position,
            this.overflow,
            this.optimizeLabelDisplay,
            this.labelMaxWidth,
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
            this.underline,
            this.color,
            this.transparency,
            this.precision,
            this.negativeStyle,
            this.zeroStyle,
            this.negativeZero,
            this.toneMode,
            this.positiveColor,
            this.negativeColor,
            this.valueShowBlankAs,
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
            this.detailShowBlankAs,
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
        this.titleGroup,
        this.valuesGroup,
        this.detailGroup,
        this.backgroundGroup,
        this.haloGroup,
        this.layoutGroup,
    ];

    /**
     * グラフの種類による既定（標準は 100% 積み上げだけ、値がオフで詳細＝全体に対する割合がオン）。
     * 保存していない項目だけ変える。saved はレポートに保存された dataLabels の値。
     * hasDetailField は「ラベルの詳細」にフィールドがあるとき true（コンテンツの既定をカスタムにする）
     */
    applyChartTypeDefaults(chartType: string, saved: powerbi.DataViewObject | undefined, hasDetailField: boolean, hasValueField = false, hasTitleField = false): void {
        const percent = chartType === CHART_TYPES.stacked100;
        if (saved?.valueShow === undefined) this.valueShow.value = !percent;
        if (saved?.detailShow === undefined) this.detailShow.value = percent;
        if (saved?.detailContent === undefined) {
            this.detailContent.value = itemOf(DETAIL_CONTENT_ITEMS, !percent && hasDetailField ? DETAIL_CONTENTS.custom : DETAIL_CONTENTS.percentOfTotal);
        }
        // 表示単位は数値のフィールドにだけ効くので、割合のときは出さない
        this.detailUnitType.visible = String(this.detailContent.value?.value) === DETAIL_CONTENTS.custom;
        this.detailShowBlankAs.visible = this.detailUnitType.visible;
        this.labelMaxWidth.visible = this.optimizeLabelDisplay.value ?? false;
        // 「ラベルのタイトル」にフィールドを入れたら、コンテンツの既定をカスタムにする（詳細と同じ）
        if (saved?.titleContent === undefined) {
            this.titleContent.value = itemOf(TITLE_CONTENT_ITEMS, hasTitleField ? TITLE_CONTENTS.custom : TITLE_CONTENTS.seriesName);
        }
        // タイトルの表示単位・桁数・空白は、フィールド（カスタム）のときだけ効く
        const titleCustom = String(this.titleContent.value?.value) === TITLE_CONTENTS.custom;
        this.titleUnitType.visible = titleCustom;
        this.titlePrecision.visible = titleCustom;
        this.titleShowBlankAs.visible = titleCustom;
        // 値の空白の表示方法は、「ラベルの値」にフィールドがあるときだけ効く
        this.valueShowBlankAs.visible = hasValueField;
        // 水平方向の配置は、行を縦に積むときだけ効く
        this.horizontalAlignment.visible = String(this.labelContentLayout.value?.value) !== LABEL_LAYOUTS.singleLine;
    }

    /**
     * 複数系列のとき、「設定の適用先」に系列を並べ、系列ごとに表示とカラーを変えられるようにする。
     * 系列が 1 本なら何も足さない（1.4 までと同じ）
     */
    applyTargets(targets: LabelTarget[]): void {
        const base = [this.optionsGroup, this.titleGroup, this.valuesGroup, this.detailGroup, this.backgroundGroup, this.haloGroup, this.layoutGroup];
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
/** 折れ線の線のスタイル。最後に「カスタム」（ダッシュ配列とダッシュ キャップを自分で決める）を足す */
export const LINE_STROKE_STYLE_ITEMS: powerbi.IEnumMember[] = [...LINE_STYLE_ITEMS, CUSTOM_LINE_STYLE_ITEM];

/** 凡例の折れ線の印（標準の複合の「スタイル」と同じ値）。棒グラフの既定は線とマーカー（1.35 までの描き方） */
export const LEGEND_MARKER_STYLES = {
    markerCircleDefault: "markerCircleDefault",
    markerOnly: "markerOnly",
    lineOnly: "lineOnly",
    lineAndMarker: "lineAndMarker",
} as const;

export const LEGEND_MARKER_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: LEGEND_MARKER_STYLES.markerCircleDefault, displayName: "マーカー（円）" },
    { value: LEGEND_MARKER_STYLES.markerOnly, displayName: "マーカー" },
    { value: LEGEND_MARKER_STYLES.lineOnly, displayName: "線" },
    { value: LEGEND_MARKER_STYLES.lineAndMarker, displayName: "線とマーカー" },
];

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

/**
 * 折れ線の値の書式（ツールヒント。第 2 Y 軸の目盛りは、1 本目の線の書式に % があればパーセントで出し、それ以外は第 2 Y 軸の表示単位と小数で出す）。
 * 自動は、率の分子・分母で作った線ならパーセント、
 * それ以外はメジャーの書式。分子・分母の線にはメジャーの書式が無い（分子の書式は金額などの #,0）ので、自動ではパーセントに寄せる。
 * 客単価（売上 ÷ 客数）のようにパーセントでない率は、カスタム書式で書式文字列を入れる
 */
export const LINE_FORMAT_MODES = {
    auto: "auto",
    percent: "percent",
    custom: "custom",
} as const;

export const LINE_FORMAT_MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: LINE_FORMAT_MODES.auto, displayName: "自動" },
    { value: LINE_FORMAT_MODES.percent, displayName: "パーセント" },
    { value: LINE_FORMAT_MODES.custom, displayName: "カスタム書式" },
];

/** 折れ線の書式で「パーセント」にしたとき（自動で率の線に当てるとき）の書式文字列 */
export const LINE_PERCENT_FORMAT = "0.0%";

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

/** 階層の上のレベルの見せ方。区切り線は標準と同じ、帯は標準に無い見せ方（保存値は 1.36 までの「囲み」と同じ boxed） */
export const HIERARCHY_STYLES = {
    lines: "lines",
    boxed: "boxed",
} as const;

export const HIERARCHY_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: HIERARCHY_STYLES.lines, displayName: "区切り線" },
    { value: HIERARCHY_STYLES.boxed, displayName: "帯" },
];

/** 段に重ねた上のレベルの段の高さ。文字の上下の余白（px）を、区切り線と帯で分けて持つ */
export const LEVEL_ROW_PADDINGS: Record<string, { lines: number; boxed: number }> = {
    normal: { lines: 8, boxed: 12 },
    compact: { lines: 5, boxed: 8 },
    tight: { lines: 3, boxed: 5 },
};

export const LEVEL_ROW_HEIGHT_ITEMS: powerbi.IEnumMember[] = [
    { value: "normal", displayName: "標準" },
    { value: "compact", displayName: "詰める" },
    { value: "tight", displayName: "もっと詰める" },
];

export class CategoryAxisCardSettings extends FormattingSettingsCompositeCard {
    name = "categoryAxis";
    displayName = "X 軸";

    private parts = categoryAxisParts({
        font: { controlName: "font", displayName: "フォント", size: 9 },
        titleFont: { controlName: "titleFont", prefix: "title", displayName: "フォント", size: 12, family: "DIN" },
        titleStyleItems: TITLE_STYLE_ITEMS,
        scrollStartItems: SCROLL_START_ITEMS,
        sliders: true,
    });
    show = this.parts.show;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    maxHeight = this.parts.maxHeight;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;
    titleStyle = this.parts.titleStyle;
    titleFont = this.parts.titleFont;
    titleColor = this.parts.titleColor;
    minCategoryWidth = this.parts.minCategoryWidth;
    scrollStart = this.parts.scrollStart;

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
     * 帯は区切りごとに角の丸い淡い面を敷く（枠線は引かない。標準に無い）
     */
    hierarchyStyle = new formattingSettings.ItemDropdown({
        name: "hierarchyStyle",
        displayName: "階層の見せ方",
        items: HIERARCHY_STYLE_ITEMS,
        value: HIERARCHY_STYLE_ITEMS[0],
    });

    /** 段に重ねた上のレベルの段の高さ。既定（標準）は今までと同じ */
    levelRowHeight = new formattingSettings.ItemDropdown({
        name: "levelRowHeight",
        displayName: "段の高さ",
        description: "縦棒で階層を段に重ねたとき、上のレベルの段の上下の余白を詰める（横棒の段の幅は変わらない）",
        items: LEVEL_ROW_HEIGHT_ITEMS,
        value: LEVEL_ROW_HEIGHT_ITEMS[0],
    });

    valuesGroup = new FormattingSettingsGroup({
        name: "categoryValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: [...this.parts.valueSlices, this.concatenateLabels, this.hierarchyStyle, this.levelRowHeight],
    });

    titleGroup = this.parts.titleGroup;
    layoutGroup = this.parts.layoutGroup;

    groups = [this.valuesGroup, this.titleGroup, this.layoutGroup];
}

export class ValueAxisCardSettings extends FormattingSettingsCompositeCard {
    name = "valueAxis";
    displayName = "Y 軸";

    private parts = valueAxisParts({
        font: { controlName: "font", displayName: "フォント", size: 9 },
        titleFont: { controlName: "titleFont", prefix: "title", displayName: "フォント", size: 12, family: "DIN" },
        titleStyleItems: TITLE_STYLE_ITEMS,
        unitTypeItems: UNIT_TYPES,
        unitNotationItems: UNIT_NOTATIONS,
        precisionItems: PRECISIONS,
        tickCountDescription: "空なら自動。数を入れると、その本数以内で切りのいい目盛りにする（数字が重なるなら減らす）。第 2 Y 軸も同じ上限。対数の軸では効かない（10 の累乗で決まる）",
    });

    start = this.parts.start;
    end = this.parts.end;

    logarithmic = new formattingSettings.ToggleSwitch({
        name: "logarithmic",
        displayName: "対数目盛り",
        value: false,
    });

    invertRange = this.parts.invertRange;
    roundRange = this.parts.roundRange;
    tickCount = this.parts.tickCount;
    show = this.parts.show;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    unitType = this.parts.unitType;
    unitNotation = this.parts.unitNotation;

    showUnitOnAxis = new formattingSettings.ToggleSwitch({
        name: "showUnitOnAxis",
        displayName: "軸ラベルに単位を表示",
        value: false,
    });

    precision = this.parts.precision;
    switchPosition = this.parts.switchPosition;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;
    titleStyle = this.parts.titleStyle;
    titleFont = this.parts.titleFont;
    titleColor = this.parts.titleColor;
    unitShow = this.parts.unitShow;
    unitText = this.parts.unitText;

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

    titleGroup = this.parts.titleGroup;

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
/** マーカーの書式（「すべて」か、線ごとに保存した値） */
export interface LineMarkerValues {
    show: boolean;
    /** 型（"circle" | "square" | "diamond" | "triangle" | "cross" | "shortDash" | "longDash" | "plus"） */
    shape: string;
    size: number;
    /** 中心で回す角度（度、時計回り。0〜359） */
    rotation: number;
    /** 空 = 線の色 */
    color: string;
    transparency: number;
    borderShow: boolean;
    borderMatchLine: boolean;
    borderColor: string;
    borderTransparency: number;
    borderWidth: number;
}

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
    /** 値の書式（LINE_FORMAT_MODES） */
    formatMode: string;
    /** 書式がカスタム書式のときの書式文字列 */
    customFormat: string;
    /** 線のスタイルがカスタムのときの模様（線と隙間の長さ px を空白で区切る）と端の形 */
    dashArray: string;
    dashCap: string;
    /** 破線・点線・カスタムの模様を線の幅に比例させるか */
    scaleWithWidth: boolean;
    /** 線の透過性 (%) */
    transparency: number;
    /** この線のマーカーの書式 */
    marker: LineMarkerValues;
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
        items: LINE_STROKE_STYLE_ITEMS,
        value: LINE_STROKE_STYLE_ITEMS.find((i) => i.value === LINE_STYLES.solid) ?? LINE_STROKE_STYLE_ITEMS[0],
    });

    /** 線のスタイルがカスタムのときの模様。線と隙間の長さ（px）を空白で区切って並べる（標準と同じ） */
    dashArray = new formattingSettings.TextInput({
        name: "dashArray",
        displayName: "ダッシュ配列",
        value: "",
        placeholder: "例: 5 5 0 5",
    });

    /** 破線・点線・カスタムの模様を線の幅に比例させる。既定はオン（1.35 までと同じ描き方） */
    scaleWithWidth = new formattingSettings.ToggleSwitch({
        name: "scaleWithWidth",
        displayName: "幅で拡大縮小",
        value: true,
    });

    dashCap = new formattingSettings.ItemDropdown({
        name: "dashCap",
        displayName: "ダッシュ キャップ",
        items: DASH_CAP_ITEMS,
        value: DASH_CAP_ITEMS[0],
    });

    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 0,
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

    formatMode = new formattingSettings.ItemDropdown({
        name: "formatMode",
        displayName: "書式",
        description: "ツールヒントの書式。自動は、率の分子・分母の線ならパーセント、それ以外はメジャーの書式。1 本目の線がパーセントなら、第 2 Y 軸の目盛りもパーセントにする",
        items: LINE_FORMAT_MODE_ITEMS,
        value: LINE_FORMAT_MODE_ITEMS[0],
    });

    customFormat = new formattingSettings.TextInput({
        name: "customFormat",
        displayName: "カスタム書式",
        value: "",
        placeholder: "例: #,0 または 0.0%",
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
        this.customFormat.visible = String(this.formatMode.value?.value ?? LINE_FORMAT_MODES.auto) === LINE_FORMAT_MODES.custom;
        const style = String(this.lineStyle.value?.value ?? LINE_STYLES.solid);
        this.dashArray.visible = style === CUSTOM_LINE_STYLE;
        this.dashCap.visible = style === CUSTOM_LINE_STYLE;
        this.scaleWithWidth.visible = style !== LINE_STYLES.solid;
        return new LineTargetItem("すべて", [
            this.show,
            this.includeCumulative,
            this.lineStyle,
            this.dashArray,
            this.scaleWithWidth,
            this.dashCap,
            this.lineJoin,
            this.width,
            this.transparency,
            this.interpolation,
            this.smoothing,
            this.tension,
            this.stepPosition,
            this.stepConnect,
            this.stepWidth,
            this.formatMode,
            this.customFormat,
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
                            items: LINE_STROKE_STYLE_ITEMS,
                            value: LINE_STROKE_STYLE_ITEMS.find((i) => i.value === target.lineStyle) ?? LINE_STROKE_STYLE_ITEMS[0],
                            selector: target.selector,
                        }),
                        new formattingSettings.TextInput({
                            name: "dashArray",
                            displayName: "ダッシュ配列",
                            value: target.dashArray,
                            placeholder: "例: 5 5 0 5",
                            selector: target.selector,
                            visible: target.lineStyle === CUSTOM_LINE_STYLE,
                        }),
                        new formattingSettings.ToggleSwitch({
                            name: "scaleWithWidth",
                            displayName: "幅で拡大縮小",
                            value: target.scaleWithWidth,
                            selector: target.selector,
                            visible: target.lineStyle !== LINE_STYLES.solid,
                        }),
                        new formattingSettings.ItemDropdown({
                            name: "dashCap",
                            displayName: "ダッシュ キャップ",
                            items: DASH_CAP_ITEMS,
                            value: itemOf(DASH_CAP_ITEMS, target.dashCap),
                            selector: target.selector,
                            visible: target.lineStyle === CUSTOM_LINE_STYLE,
                        }),
                        lineJoin,
                        new formattingSettings.NumUpDown({
                            name: "width",
                            displayName: "幅 (px)",
                            value: target.width,
                            selector: target.selector,
                        }),
                        new formattingSettings.NumUpDown({
                            name: "transparency",
                            displayName: "透過性 (%)",
                            value: target.transparency,
                            selector: target.selector,
                        }),
                        ...interpolation,
                        new formattingSettings.ItemDropdown({
                            name: "formatMode",
                            displayName: "書式",
                            items: LINE_FORMAT_MODE_ITEMS,
                            value: itemOf(LINE_FORMAT_MODE_ITEMS, target.formatMode),
                            selector: target.selector,
                        }),
                        new formattingSettings.TextInput({
                            name: "customFormat",
                            displayName: "カスタム書式",
                            value: target.customFormat,
                            placeholder: "例: #,0 または 0.0%",
                            selector: target.selector,
                            visible: target.formatMode === LINE_FORMAT_MODES.custom,
                        }),
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
 * 比較レイヤー（標準に無い項目）。「比較値」に入れたメジャーを、「値」の棒の奥に同じ幅で少しずつずらして重ね、奥ほど薄くする。
 * 予定・見通し・実績を集合棒で並べると 1 カテゴリの棒が細く割れるので、1 本の位置で前後に重ねて比べる（旧棒グラフの比較レイヤー）
 */
export class CompareCardSettings extends FormattingSettingsCard {
    name = "compare";
    displayName = "比較";
    description = "「比較値」の棒を「値」の棒の奥に重ねる";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "比較",
        value: true,
    });

    topLevelSlice = this.show;

    /**
     * 「比較の列」で受けたとき、どの値を手前にするか。予定・見通し・実績の順に並べた列なら、最後（実績）が手前（旧棒グラフと同じ既定）。
     * 手前から奥へは、並びを手前の値から離れる向きにたどる
     */
    order = new formattingSettings.ItemDropdown({
        name: "order",
        displayName: "手前にする値",
        items: COMPARE_ORDER_ITEMS,
        value: COMPARE_ORDER_ITEMS[0],
    });

    /** 名前で手前の値を決める（並びに意味を持たせない。「実績」と入れれば並びが変わっても実績が手前）。空なら「手前にする値」 */
    frontValue = new formattingSettings.TextInput({
        name: "frontValue",
        displayName: "手前にする値の名前",
        value: "",
        placeholder: "例：実績",
    });

    /** 0 で隣り合わせ（集合と同じ並び）、100 で完全に重なる。0〜100 に丸める */
    overlap = new formattingSettings.NumUpDown({
        name: "overlap",
        displayName: "重なり (%)",
        value: 60,
    });

    direction = new formattingSettings.ItemDropdown({
        name: "direction",
        displayName: "手前の棒",
        items: COMPARE_DIRECTION_ITEMS,
        value: COMPARE_DIRECTION_ITEMS[0],
    });

    /** いちばん奥の棒の濃さ。手前（100%）からここまで同じ割合で薄くする。枚数はデータで変わるので、1 枚ずつの設定は置かない */
    backOpacity = new formattingSettings.NumUpDown({
        name: "backOpacity",
        displayName: "いちばん奥の不透明度 (%)",
        value: 30,
    });

    /** 薄い塗りだけでは重なった棒の境目が見えにくいので、奥の棒の輪郭を棒の色で濃く引けるようにする */
    outline = new formattingSettings.ToggleSwitch({
        name: "outline",
        displayName: "輪郭線",
        value: false,
    });

    showInLegend = new formattingSettings.ToggleSwitch({
        name: "showInLegend",
        displayName: "凡例に出す",
        value: true,
    });

    slices = [this.order, this.frontValue, this.overlap, this.direction, this.backOpacity, this.outline, this.showInLegend];

    /** 「手前にする値」は「比較の列」で受けたときだけ出す（「比較値」はメジャーを入れた順で前後が決まる） */
    applyCompareBy(hasCompareBy: boolean): void {
        this.order.visible = hasCompareBy;
        this.frontValue.visible = hasCompareBy;
    }
}

/** X 軸の定数線を、棒の後ろに描くか前に描くか（標準の「位置」。標準の表示は「遅延」「前面」で、遅延は Behind の訳し違えなので背面と書く） */
export const CATEGORY_LINE_LAYERS = { back: "back", front: "front" } as const;
export const CATEGORY_LINE_LAYER_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_LAYERS.back, displayName: "背面" },
    { value: CATEGORY_LINE_LAYERS.front, displayName: "前面" },
];

/** 網掛け領域を、線のどちら側に塗るか（標準の「次の値より前:」「次の値より後:」。カテゴリの軸なので線を基準に書く） */
export const CATEGORY_LINE_SHADE_REGIONS = { before: "before", after: "after" } as const;
export const CATEGORY_LINE_SHADE_REGION_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_SHADE_REGIONS.before, displayName: "線より前" },
    { value: CATEGORY_LINE_SHADE_REGIONS.after, displayName: "線より後ろ" },
];

/** データ ラベルを線のどちら側に置くか（標準の「水平方向の位置」の「左へ移動」「右へ移動」） */
export const CATEGORY_LINE_LABEL_SIDES = { left: "left", right: "right" } as const;
export const CATEGORY_LINE_LABEL_SIDE_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_LABEL_SIDES.left, displayName: "左へ移動" },
    { value: CATEGORY_LINE_LABEL_SIDES.right, displayName: "右へ移動" },
];

/** データ ラベルを上に置くか下に置くか（標準の「縦位置」） */
export const CATEGORY_LINE_LABEL_ENDS = { top: "top", bottom: "bottom" } as const;
export const CATEGORY_LINE_LABEL_END_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_LABEL_ENDS.top, displayName: "上" },
    { value: CATEGORY_LINE_LABEL_ENDS.bottom, displayName: "下" },
];

/** データ ラベルの中身（標準の「スタイル」。標準の「双方向」は Both の訳し違えなので両方と書く） */
export const CATEGORY_LINE_LABEL_TEXTS = { value: "value", name: "name", both: "both" } as const;
export const CATEGORY_LINE_LABEL_TEXT_ITEMS: powerbi.IEnumMember[] = [
    { value: CATEGORY_LINE_LABEL_TEXTS.value, displayName: "データ値" },
    { value: CATEGORY_LINE_LABEL_TEXTS.name, displayName: "名前" },
    { value: CATEGORY_LINE_LABEL_TEXTS.both, displayName: "両方" },
];

/**
 * X 軸の定数線（標準の分析ペインの「X 軸の定数線」は連続の軸だけ。カテゴリの軸でも引けるようにした）。
 * 位置は「X 軸の定数線」の欄のメジャーが、カテゴリごとに空白でない値を返したところ。
 * まとまりと項目名は標準の定数線（線・網掛け領域・データ ラベル）に合わせた（2026-10-06、Desktop 2.158 の分析ペインで確かめた）
 */
export class CategoryLineCardSettings extends FormattingSettingsCompositeCard {
    name = "categoryLine";
    displayName = "X 軸の定数線";
    description = "「X 軸の定数線」の欄のメジャーが値を返したカテゴリに線を引く";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "X 軸の定数線",
        value: true,
    });

    topLevelSlice = this.show;

    /** 印の付いたカテゴリのどこに引くか。前は期の境目（そのカテゴリの始まり）。標準に無い項目 */
    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "線を引く場所",
        items: CATEGORY_LINE_POSITION_ITEMS,
        value: CATEGORY_LINE_POSITION_ITEMS[0],
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    transparency = new formattingSettings.NumUpDown({
        name: "transparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STROKE_STYLE_ITEMS,
        value: LINE_STROKE_STYLE_ITEMS.find((item) => item.value === LINE_STYLES.dashed)!,
    });

    dashArray = new formattingSettings.TextInput({
        name: "dashArray",
        displayName: "ダッシュ配列",
        value: "",
        placeholder: "例: 5 5 0 5",
    });

    scaleWithWidth = new formattingSettings.ToggleSwitch({
        name: "scaleWithWidth",
        displayName: "幅で拡大縮小",
        value: true,
    });

    dashCap = new formattingSettings.ItemDropdown({
        name: "dashCap",
        displayName: "ダッシュ キャップ",
        items: DASH_CAP_ITEMS,
        value: DASH_CAP_ITEMS[0],
    });

    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 1,
    });

    layer = new formattingSettings.ItemDropdown({
        name: "layer",
        displayName: "位置",
        items: CATEGORY_LINE_LAYER_ITEMS,
        value: CATEGORY_LINE_LAYER_ITEMS[1],
    });

    shadeShow = new formattingSettings.ToggleSwitch({
        name: "shadeShow",
        displayName: "網掛け領域",
        value: false,
    });

    shadeRegion = new formattingSettings.ItemDropdown({
        name: "shadeRegion",
        displayName: "位置",
        items: CATEGORY_LINE_SHADE_REGION_ITEMS,
        value: CATEGORY_LINE_SHADE_REGION_ITEMS[0],
    });

    shadeMatchLine = new formattingSettings.ToggleSwitch({
        name: "shadeMatchLine",
        displayName: "線の色を一致させる",
        value: false,
    });

    /** 網掛けの色。棒の後ろに広く塗るので、既定は淡い灰色（標準は線の色と同じ青。透過性 40% でも棒と見分けにくい） */
    shadeColor = new formattingSettings.ColorPicker({
        name: "shadeColor",
        displayName: "カラー",
        value: { value: "#E1DFDD" },
    });

    shadeTransparency = new formattingSettings.NumUpDown({
        name: "shadeTransparency",
        displayName: "透過性 (%)",
        value: 40,
    });

    labelShow = new formattingSettings.ToggleSwitch({
        name: "labelShow",
        displayName: "データ ラベル",
        value: true,
    });

    labelHorizontal = new formattingSettings.ItemDropdown({
        name: "labelHorizontal",
        displayName: "水平方向の位置",
        items: CATEGORY_LINE_LABEL_SIDE_ITEMS,
        value: CATEGORY_LINE_LABEL_SIDE_ITEMS[1],
    });

    labelVertical = new formattingSettings.ItemDropdown({
        name: "labelVertical",
        displayName: "縦位置",
        items: CATEGORY_LINE_LABEL_END_ITEMS,
        value: CATEGORY_LINE_LABEL_END_ITEMS[0],
    });

    labelText = new formattingSettings.ItemDropdown({
        name: "labelText",
        displayName: "スタイル",
        items: CATEGORY_LINE_LABEL_TEXT_ITEMS,
        value: CATEGORY_LINE_LABEL_TEXT_ITEMS[0],
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "テキストのサイズ",
        value: 9,
    });

    /** 線が近くてラベルがぶつかるとき、ずらす（縦棒は段を変え、横棒は横へ）。オフなら重ねたまま。標準に無い項目 */
    labelAvoidOverlap = new formattingSettings.ToggleSwitch({
        name: "labelAvoidOverlap",
        displayName: "ラベルの重なりを避ける",
        value: true,
    });

    lineGroup = new FormattingSettingsGroup({
        name: "categoryLineLine",
        displayName: "線",
        slices: [this.position, this.color, this.transparency, this.lineStyle, this.dashArray, this.scaleWithWidth, this.dashCap, this.width, this.layer],
    });

    shadeGroup = new FormattingSettingsGroup({
        name: "categoryLineShade",
        displayName: "網掛け領域",
        topLevelSlice: this.shadeShow,
        slices: [this.shadeRegion, this.shadeMatchLine, this.shadeColor, this.shadeTransparency],
    });

    labelGroup = new FormattingSettingsGroup({
        name: "categoryLineLabel",
        displayName: "データ ラベル",
        topLevelSlice: this.labelShow,
        slices: [this.labelHorizontal, this.labelVertical, this.labelText, this.labelColor, this.labelFontSize, this.labelAvoidOverlap],
    });

    groups = [this.lineGroup, this.shadeGroup, this.labelGroup];

    /** カスタムの線の項目は、線のスタイルがカスタムのときだけ。網掛けの色は、線の色に合わせないときだけ */
    syncVisibility(): void {
        const custom = String(this.lineStyle.value?.value) === CUSTOM_LINE_STYLE;
        this.dashArray.visible = custom;
        this.dashCap.visible = custom;
        this.shadeColor.visible = !(this.shadeMatchLine.value ?? false);
    }
}

/** Y 軸の定数線を載せる軸 */
export const VALUE_LINE_AXES = { primary: "primary", secondary: "secondary" } as const;
export const VALUE_LINE_AXIS_ITEMS: powerbi.IEnumMember[] = [
    { value: VALUE_LINE_AXES.primary, displayName: "Y 軸" },
    { value: VALUE_LINE_AXES.secondary, displayName: "第 2 Y 軸" },
];

/** Y 軸の定数線の網掛け領域を、値の小さい側に塗るか大きい側に塗るか（標準の「次の値より前:」「次の値より後:」） */
export const VALUE_LINE_SHADE_REGIONS = { below: "below", above: "above" } as const;
export const VALUE_LINE_SHADE_REGION_ITEMS: powerbi.IEnumMember[] = [
    { value: VALUE_LINE_SHADE_REGIONS.below, displayName: "値より小さい側" },
    { value: VALUE_LINE_SHADE_REGIONS.above, displayName: "値より大きい側" },
];

/**
 * Y 軸の定数線（標準の分析ペインの「Y 軸の定数線」にあたる。カスタム ビジュアルは標準の定数線を使えないので、書式ペインのカードにした）。
 * 線は、このカードの「値」に入れた数と、「Y 軸の定数線」の欄のメジャー（4 つまで）。まとまりと項目名は標準の定数線に合わせた
 */
export class ValueLineCardSettings extends FormattingSettingsCompositeCard {
    name = "valueLine";
    displayName = "Y 軸の定数線";
    description = "「値」の数と、「Y 軸の定数線」の欄のメジャーの値に線を引く";

    show = new formattingSettings.ToggleSwitch({ name: "show", displayName: "Y 軸の定数線", value: true });

    topLevelSlice = this.show;

    /** 決まった数の線（標準の「値」）。空なら引かない。桁区切りのカンマも読む */
    value = new formattingSettings.TextInput({ name: "value", displayName: "値", value: "", placeholder: "例：1000" });

    /** 「値」の線の名前（データ ラベルのスタイルが名前・両方のとき）。メジャーの線はメジャーの名前 */
    lineName = new formattingSettings.TextInput({ name: "name", displayName: "名前", value: "", placeholder: "定数線" });

    axis = new formattingSettings.ItemDropdown({ name: "axis", displayName: "軸", items: VALUE_LINE_AXIS_ITEMS, value: VALUE_LINE_AXIS_ITEMS[0] });

    /** 線が値の軸の範囲の外なら、軸を広げて線を入れる（標準は広げずに描かない。既定は標準と同じ） */
    extendAxis = new formattingSettings.ToggleSwitch({ name: "extendAxis", displayName: "軸を線に合わせて広げる", value: false });

    color = new formattingSettings.ColorPicker({ name: "color", displayName: "カラー", value: { value: "#605E5C" } });

    transparency = new formattingSettings.NumUpDown({ name: "transparency", displayName: "透過性 (%)", value: 0 });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STROKE_STYLE_ITEMS,
        value: LINE_STROKE_STYLE_ITEMS.find((item) => item.value === LINE_STYLES.dashed)!,
    });

    dashArray = new formattingSettings.TextInput({ name: "dashArray", displayName: "ダッシュ配列", value: "", placeholder: "例: 5 5 0 5" });

    scaleWithWidth = new formattingSettings.ToggleSwitch({ name: "scaleWithWidth", displayName: "幅で拡大縮小", value: true });

    dashCap = new formattingSettings.ItemDropdown({ name: "dashCap", displayName: "ダッシュ キャップ", items: DASH_CAP_ITEMS, value: DASH_CAP_ITEMS[0] });

    width = new formattingSettings.NumUpDown({ name: "width", displayName: "幅 (px)", value: 1 });

    layer = new formattingSettings.ItemDropdown({ name: "layer", displayName: "位置", items: CATEGORY_LINE_LAYER_ITEMS, value: CATEGORY_LINE_LAYER_ITEMS[1] });

    shadeShow = new formattingSettings.ToggleSwitch({ name: "shadeShow", displayName: "網掛け領域", value: false });

    shadeRegion = new formattingSettings.ItemDropdown({
        name: "shadeRegion",
        displayName: "位置",
        items: VALUE_LINE_SHADE_REGION_ITEMS,
        value: VALUE_LINE_SHADE_REGION_ITEMS[0],
    });

    shadeMatchLine = new formattingSettings.ToggleSwitch({ name: "shadeMatchLine", displayName: "線の色を一致させる", value: false });

    shadeColor = new formattingSettings.ColorPicker({ name: "shadeColor", displayName: "カラー", value: { value: "#E1DFDD" } });

    shadeTransparency = new formattingSettings.NumUpDown({ name: "shadeTransparency", displayName: "透過性 (%)", value: 40 });

    labelShow = new formattingSettings.ToggleSwitch({ name: "labelShow", displayName: "データ ラベル", value: true });

    /** 標準の Y 軸の定数線と同じく、既定は左（線の左端） */
    labelHorizontal = new formattingSettings.ItemDropdown({
        name: "labelHorizontal",
        displayName: "水平方向の位置",
        items: CATEGORY_LINE_LABEL_SIDE_ITEMS,
        value: CATEGORY_LINE_LABEL_SIDE_ITEMS[0],
    });

    labelVertical = new formattingSettings.ItemDropdown({
        name: "labelVertical",
        displayName: "縦位置",
        items: CATEGORY_LINE_LABEL_END_ITEMS,
        value: CATEGORY_LINE_LABEL_END_ITEMS[0],
    });

    labelText = new formattingSettings.ItemDropdown({
        name: "labelText",
        displayName: "スタイル",
        items: CATEGORY_LINE_LABEL_TEXT_ITEMS,
        value: CATEGORY_LINE_LABEL_TEXT_ITEMS[0],
    });

    labelColor = new formattingSettings.ColorPicker({ name: "labelColor", displayName: "カラー", value: { value: "#605E5C" } });

    labelFontSize = new formattingSettings.NumUpDown({ name: "labelFontSize", displayName: "テキストのサイズ", value: 9 });

    /** 値の表示単位は、載せた軸の表示単位に従う（データ ラベルと同じ）。小数点以下の桁数だけ選べる */
    labelPrecision = new formattingSettings.ItemDropdown({ name: "labelPrecision", displayName: "小数点以下桁数の値", items: PRECISIONS, value: PRECISIONS[0] });

    labelAvoidOverlap = new formattingSettings.ToggleSwitch({ name: "labelAvoidOverlap", displayName: "ラベルの重なりを避ける", value: true });

    lineGroup = new FormattingSettingsGroup({
        name: "valueLineLine",
        displayName: "線",
        slices: [this.value, this.lineName, this.axis, this.extendAxis, this.color, this.transparency, this.lineStyle, this.dashArray, this.scaleWithWidth, this.dashCap, this.width, this.layer],
    });

    shadeGroup = new FormattingSettingsGroup({
        name: "valueLineShade",
        displayName: "網掛け領域",
        topLevelSlice: this.shadeShow,
        slices: [this.shadeRegion, this.shadeMatchLine, this.shadeColor, this.shadeTransparency],
    });

    labelGroup = new FormattingSettingsGroup({
        name: "valueLineLabel",
        displayName: "データ ラベル",
        topLevelSlice: this.labelShow,
        slices: [this.labelHorizontal, this.labelVertical, this.labelText, this.labelColor, this.labelFontSize, this.labelPrecision, this.labelAvoidOverlap],
    });

    groups = [this.lineGroup, this.shadeGroup, this.labelGroup];

    /** カスタムの線の項目はカスタムのときだけ。網掛けの色は線の色に合わせないときだけ。軸は第 2 Y 軸があるときだけ、広げるのは Y 軸のときだけ */
    syncVisibility(hasSecondary: boolean): void {
        const custom = String(this.lineStyle.value?.value) === CUSTOM_LINE_STYLE;
        this.dashArray.visible = custom;
        this.dashCap.visible = custom;
        this.shadeColor.visible = !(this.shadeMatchLine.value ?? false);
        this.axis.visible = hasSecondary;
        this.extendAxis.visible = !hasSecondary || String(this.axis.value?.value) !== VALUE_LINE_AXES.secondary;
    }
}

/**
 * 折れ線の点の印。標準と同じく既定は出さず、「すべてのカテゴリに表示」で出す。
 * 折れ線があれば「設定の適用先」に「すべて」と線を並べ、線ごとに型・色・罫線を変えられる（標準と同じ）。
 * 標準の適用先はカテゴリごとにも選べるが、ここでは線ごとまで（カテゴリごとは見送り）
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

    /** マーカーを中心で回す角度（度、時計回り）。標準の「回転」 */
    rotation = new formattingSettings.NumUpDown({
        name: "rotation",
        displayName: "回転",
        value: 0,
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
        slices: [this.shape, this.size, this.rotation],
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

    /** マーカーの書式。own（線ごとに保存した値）があればそれを、無い項目は「すべて」の値を使う */
    values(own?: powerbi.DataViewObject): LineMarkerValues {
        const raw = (property: string): powerbi.DataViewPropertyValue | undefined => {
            const value = own?.[property];
            return value === undefined || value === null ? undefined : value;
        };
        const num = (property: string, fallback: number): number => {
            const value = raw(property);
            return typeof value === "number" ? value : fallback;
        };
        const bool = (property: string, fallback: boolean): boolean => {
            const value = raw(property);
            return typeof value === "boolean" ? value : fallback;
        };
        const fill = (property: string, fallback: string): string => {
            const color = (raw(property) as powerbi.Fill | undefined)?.solid?.color;
            return color !== undefined && color !== null ? String(color) : fallback;
        };
        const percent = (v: number) => Math.max(0, Math.min(100, v));
        const shape = raw("shape");
        return {
            show: bool("show", this.show.value ?? false),
            shape: shape !== undefined ? String(shape) : String(this.shape.value?.value ?? "circle"),
            size: Math.max(1, Math.min(20, num("size", this.size.value ?? 5))),
            rotation: (((num("rotation", Number(this.rotation.value) || 0) || 0) % 360) + 360) % 360,
            color: fill("color", this.color.value?.value ?? ""),
            transparency: percent(num("transparency", this.transparency.value ?? 0)),
            borderShow: bool("borderShow", this.borderShow.value ?? false),
            borderMatchLine: bool("borderMatchLine", this.borderMatchLine.value ?? false),
            borderColor: fill("borderFill", this.borderFill.value?.value || "#605E5C") || "#605E5C",
            borderTransparency: percent(num("borderTransparency", this.borderTransparency.value ?? 0)),
            borderWidth: Math.max(1, Math.min(10, num("borderWidth", this.borderWidth.value ?? 1))),
        };
    }

    /** 折れ線があれば、「設定の適用先」に「すべて」と線を並べる。無ければ今までどおり項目のまとまりで出す */
    applyTargets(targets: LineTarget[]): void {
        if (!targets.length) {
            this.groups = [this.optionsGroup, this.shapeGroup, this.colorGroup, this.borderGroup];
            return;
        }
        const all = new MarkerTargetItem("すべて", [
            this.show,
            this.shape,
            this.size,
            this.rotation,
            this.color,
            this.transparency,
            this.borderShow,
            this.borderMatchLine,
            this.borderFill,
            this.borderTransparency,
            this.borderWidth,
        ]);
        const items = targets.slice(0, MAX_COLUMN_TARGETS).map((target) => {
            const m = target.marker;
            const selector = target.selector;
            return new MarkerTargetItem(target.name, [
                new formattingSettings.ToggleSwitch({ name: "show", displayName: "このシリーズに表示", value: m.show, selector }),
                new formattingSettings.ItemDropdown({ name: "shape", displayName: "型", items: MARKER_SHAPE_ITEMS, value: itemOf(MARKER_SHAPE_ITEMS, m.shape), selector }),
                new formattingSettings.NumUpDown({ name: "size", displayName: "サイズ (px)", value: m.size, selector }),
                new formattingSettings.NumUpDown({ name: "rotation", displayName: "回転", value: m.rotation, selector }),
                new formattingSettings.ColorPicker({ name: "color", displayName: "カラー", value: { value: m.color }, selector }),
                new formattingSettings.NumUpDown({ name: "transparency", displayName: "透過性 (%)", value: m.transparency, selector }),
                new formattingSettings.ToggleSwitch({ name: "borderShow", displayName: "罫線", value: m.borderShow, selector }),
                new formattingSettings.ToggleSwitch({ name: "borderMatchLine", displayName: "線の色を一致させる", value: m.borderMatchLine, selector }),
                new formattingSettings.ColorPicker({ name: "borderFill", displayName: "罫線のカラー", value: { value: m.borderColor }, selector }),
                new formattingSettings.NumUpDown({ name: "borderTransparency", displayName: "罫線の透過性 (%)", value: m.borderTransparency, selector }),
                new formattingSettings.NumUpDown({ name: "borderWidth", displayName: "罫線の幅 (px)", value: m.borderWidth, selector }),
            ]);
        });
        this.groups = [
            new FormattingSettingsGroup({
                name: "markerTargets",
                slices: [],
                container: new FormattingSettingsContainer({ displayName: "設定の適用先", containerItems: [all, ...items] }),
            }),
        ];
    }
}

class MarkerTargetItem extends FormattingSettingsCard {
    name = "markerTarget";

    constructor(displayName: string, slices: FormattingSettingsSlice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
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

/** 縦のグリッド線（カテゴリの区切りの線）を引く位置 */
export const CATEGORY_GRID_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "categories", displayName: "カテゴリごと" },
    { value: "levels", displayName: "階層の区切り" },
];

export class GridlinesCardSettings extends FormattingSettingsCompositeCard {
    name = "gridlines";
    displayName = "グリッド線";

    // 横（水平・Y 軸の目盛線）と縦（垂直・X 軸の目盛線）。既定の線のスタイルは点線
    private horizontal = gridlineParts({ prefix: "horizontal", showDisplayName: "横", show: true, styleItems: LINE_STYLE_ITEMS, style: LINE_STYLE_ITEMS[0] });
    private vertical = gridlineParts({ prefix: "vertical", showDisplayName: "縦", show: false, styleItems: LINE_STYLE_ITEMS, style: LINE_STYLE_ITEMS[0] });

    horizontalShow = this.horizontal.show;
    horizontalColor = this.horizontal.color;
    horizontalTransparency = this.horizontal.transparency;
    horizontalStyle = this.horizontal.style;
    horizontalScaleWithWidth = this.horizontal.scaleWithWidth;
    horizontalDashArray = this.horizontal.dashArray;
    horizontalDashCap = this.horizontal.dashCap;
    horizontalWidth = this.horizontal.width;

    verticalShow = this.vertical.show;
    verticalColor = this.vertical.color;
    verticalTransparency = this.vertical.transparency;
    verticalStyle = this.vertical.style;
    verticalScaleWithWidth = this.vertical.scaleWithWidth;
    verticalDashArray = this.vertical.dashArray;
    verticalDashCap = this.vertical.dashCap;
    verticalWidth = this.vertical.width;

    /** 縦の線を引く位置。既定はカテゴリごと（今までと同じ）。階層の区切りなら、X 軸の階層の 1 つ上のレベルの区切りだけ */
    verticalPosition = new formattingSettings.ItemDropdown({
        name: "verticalPosition",
        displayName: "線の位置",
        description: "階層の区切りにすると、X 軸に階層を入れたとき、1 つ上のレベルの区切り（年度・四半期の境目など）にだけ引く。階層が無ければカテゴリごと",
        items: CATEGORY_GRID_POSITION_ITEMS,
        value: CATEGORY_GRID_POSITION_ITEMS[0],
    });

    horizontalGroup = new FormattingSettingsGroup({
        name: "horizontalGridlines",
        displayName: "横",
        topLevelSlice: this.horizontalShow,
        slices: this.horizontal.slices,
    });

    verticalGroup = new FormattingSettingsGroup({
        name: "verticalGridlines",
        displayName: "縦",
        topLevelSlice: this.verticalShow,
        slices: [this.verticalPosition, ...this.vertical.slices],
    });

    groups = [this.horizontalGroup, this.verticalGroup];

    /** 線のスタイルがカスタムのときだけ、ダッシュ配列とダッシュ キャップを出す */
    syncCustomDash(): void {
        this.horizontal.sync();
        this.vertical.sync();
    }
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
    compare = new CompareCardSettings();
    categoryLine = new CategoryLineCardSettings();
    valueLine = new ValueLineCardSettings();
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
        this.compare,
        this.categoryLine,
        this.valueLine,
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
        this.markers.applyTargets(options.lineTargets ?? []);
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
        this.categoryLine.displayName = horizontal ? "Y 軸の定数線" : "X 軸の定数線";
        this.categoryLine.show.displayName = this.categoryLine.displayName;
        this.valueLine.displayName = horizontal ? "X 軸の定数線" : "Y 軸の定数線";
        this.valueLine.show.displayName = this.valueLine.displayName;
        this.valueLine.axis.items = horizontal
            ? [{ value: VALUE_LINE_AXES.primary, displayName: "X 軸" }, { value: VALUE_LINE_AXES.secondary, displayName: "第 2 X 軸" }]
            : VALUE_LINE_AXIS_ITEMS;
        // 横棒では線が横に寝るので、ラベルの「水平方向の位置」は線の左端・右端、「縦位置」は線の上・下になる（標準の X 軸の定数線と同じ読み方）
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
        const columns = dataView?.metadata?.columns;
        const hasDetailField = !!columns?.some((c) => c.roles?.labelDetail);
        const hasTitleField = !!columns?.some((c) => c.roles?.labelTitle);
        const hasValueField = !!columns?.some((c) => c.roles?.labelValue);
        this.dataLabels.applyChartTypeDefaults(chartType, saved, hasDetailField, hasValueField, hasTitleField);
    }

    /**
     * グラフの種類とデータに関係ないカードを隠す（標準はその種類で使うカードだけを出す）。
     * 隠すのは書式ペインの表示だけで、保存済みの値は残る。hasLines は「折れ線の値」にフィールドがあるとき true。
     * lineOnly は「値」が空で折れ線だけを描くとき true（棒にしか効かない合計ラベル・リボン・第 2 Y 軸を隠す。
     * 「棒」カードは並び順・カテゴリ間のスペース・その他のまとめ方が線にも効くので残す）
     */
    applyCardVisibility(hasLines: boolean, lineOnly = false): void {
        const chartType = String(this.chart.chartType.value?.value ?? CHART_TYPES.clustered);
        const horizontal = String(this.chart.orientation.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
        // 折れ線の値は、どの種類・向きでも描く（横棒は既定でマーカーだけ）
        const linesDrawn = hasLines;
        this.totalLabels.visible = chartType === CHART_TYPES.stacked && !lineOnly;
        // リボンは積み上げ・100% 積み上げで出せる（縦棒・横棒とも。集合は棒が横に並ぶので帯でつながない）
        this.ribbons.visible = (chartType === CHART_TYPES.stacked || chartType === CHART_TYPES.stacked100) && !lineOnly;
        // 折れ線だけのときは線を左の軸で描くので、第 2 Y 軸は使わない
        this.valueAxis2.visible = linesDrawn && !lineOnly;
        this.lines.visible = linesDrawn;
        this.areas.visible = linesDrawn;
        this.markers.visible = linesDrawn;
        this.columns.applyChartType(chartType !== CHART_TYPES.clustered, linesDrawn);
        this.legend.applyLines(linesDrawn);
        this.gridlines.syncCustomDash();
        this.categoryLine.syncVisibility();
        this.valueLine.syncVisibility(hasLines);
    }
}
