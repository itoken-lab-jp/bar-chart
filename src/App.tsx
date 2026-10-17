"use strict";

import * as React from "react";
import powerbi from "powerbi-visuals-api";
import IViewport = powerbi.IViewport;
import ISelectionId = powerbi.visuals.ISelectionId;

import { ViewModel, DataPoint, CategoryGroup, LineSeriesInfo, LegendItemInfo, DataLabelsSettings, ParetoRank, Tick } from "./viewModel";
import { VisualFormattingSettingsModel, STEP_WIDTHS, LEVEL_ROW_PADDINGS, LINE_BLANKS } from "./settings";
import { placeLabel, labelBlock, measureTextWidth, LABEL_PADDING, FontSpec, LabelLine } from "./unitUtils";
import { continuousBand, continuousTicks, positionOf } from "./continuousAxis";
import { clusterLayout, spanOf, levelRunsOf, layerLayout, layerOpacity, categoryGridBreaks } from "./layout";
import { layoutLegend, LegendLayout, LegendItemBox, LEGEND_MARKER_GAP, LAYER_GLYPH_THICKNESS, LAYER_GLYPH_LENGTH } from "./legend";
import { linePath, areaPath, markerPath, XY } from "./linePath";
import { recommendedTickCount, fitTicks as fitTicksBy } from "./shared/ticks";
import { gridCutout, useGridMaskId } from "./shared/gridCutout";
import { CUSTOM_LINE_STYLE, gridLineStroke, lineStrokeOf } from "./shared/gridlines";

export { gridLineStroke, lineStrokeOf };
import { DARK_TEXT, autoLabelColor, blend, sameColor } from "./shared/color";
import { HighContrastColors } from "./shared/highContrast";
import { selectKeyProps, selectKeyDown } from "./shared/keyboard";
import { LegendArrowButton } from "./shared/LegendArrowButton";
import { valueEndBarPath } from "./shared/barShape";
import { placeReferenceLabels, referenceLabelSvgStyle, referenceLineSvgStyle, referenceShadeSvgStyle } from "./shared/referenceLine";
import { haloStyle } from "./shared/halo";
import { rotatedLabelLeadIn } from "./shared/rotatedLabels";
import { PT_TO_PX, truncateToWidth } from "./shared/text";
import { useScrollStart } from "./shared/scrollStart";
import { CopyImageButton } from "./shared/CopyImageButton";
import { controlColorsOf, controlVars } from "./shared/controls";
import { TEXT_ATTRIBUTE } from "./shared/copyImage";
import { fontOf, fontStyleOf } from "./shared/textStyle";

/** 長ければ末尾を省略した文字（軸のタイトル用）。省略したときは、マウスを乗せると全体が出る */
function fittedText(text: string, maxWidthPx: number, font: FontSpec): React.ReactNode {
    const shown = truncateText(text, maxWidthPx, font);
    return shown === text ? text : (
        <>
            {shown}
            <title>{text}</title>
        </>
    );
}

/**
 * 指定の幅に収まるよう、末尾を「…」で省略した文字。幅は Power BI の文字幅の計測で測る（標準と同じ省略記号 1 文字）。
 * 1.10 までは全角 1.05 文字分・「...」3 文字分と見積もっていて、標準より早く省略していた
 */
function truncateText(text: string, maxWidthPx: number, font: FontSpec): string {
    if (!text) return "";
    if (measureTextWidth(text, font) <= maxWidthPx) return text;
    const ellipsis = "…";
    // 収まるいちばん長い先頭を二分探索で探す（計測の呼び出しを少なくする）
    let lo = 0;
    let hi = text.length - 1;
    while (lo < hi) {
        const mid = Math.ceil((lo + hi) / 2);
        if (measureTextWidth(text.slice(0, mid).trimEnd() + ellipsis, font) <= maxWidthPx) lo = mid;
        else hi = mid - 1;
    }
    return (lo > 0 ? text.slice(0, lo).trimEnd() : text.charAt(0)) + ellipsis;
}


/**
 * データラベルの自動の文字色。文字のすぐ下の色で決める：背景を出すなら背景（透過性があれば下の色と混ぜた色）、
 * 出さないなら置いた場所の色（棒の中は棒の色、外はビジュアルの背景）。棒の中は白、外は灰色を優先し、
 * 読めなければ白か濃い灰色の読めるほう（shared/color.ts の autoLabelColor）。
 * ハイコントラストでは、前景色で塗った棒の中だけ背景色、それ以外は前景色
 */
export function labelAutoColor(
    inside: boolean,
    barColor: string,
    dl: Pick<DataLabelsSettings, "backgroundShow" | "backgroundColor" | "backgroundTransparency">,
    hc?: HighContrastColors | null,
    chartBackground = "#FFFFFF"
): string {
    if (hc) return inside && !dl.backgroundShow && sameColor(barColor, hc.foreground) ? hc.background : hc.foreground;
    const underneath = inside ? barColor : chartBackground;
    const base = dl.backgroundShow ? blend(dl.backgroundColor, underneath, dl.backgroundTransparency / 100) : underneath;
    return autoLabelColor(inside, base);
}

/** 斜めのカテゴリ名を、棒の下端から離す間隔 (px) */
const ROTATED_LABEL_GAP = 8;
/** 前後に点の無い折れ線の 1 点に置く点の半径の下限 (px)。線が細くても見える大きさにする */
const LONE_POINT_MIN_RADIUS = 3;

/** 階層の「帯」：隣の帯とのすき間の半分と、角の丸み (px) */
const LEVEL_BAND_OPACITY = 0.12;
/** 選んだ区切りの帯の濃さ */
const LEVEL_BAND_PICKED_OPACITY = 0.24;
const LEVEL_BOX_GAP = 2;
const LEVEL_BOX_RADIUS = 4;

/**
 * 斜めのカテゴリラベルをビジュアル左端から離しておく余白 (px)。
 * Power BI はビジュアルの枠で描画をクリップするため、左端ぎりぎりまで伸ばすと
 * 枠線・角丸・ビジュアルの内側パディングの下で文字が欠ける。その分を見込んで「...」で省略する。
 */
const CATEGORY_LABEL_LEFT_SAFE_MARGIN = 20;

/** ハイライトで該当しない部分（棒全体）の不透明度の係数 */
const HIGHLIGHT_DIM_FACTOR = 0.35;
/** 累計の切り替えボタンをオンにしたときの色（枠・文字。塗りはこれを薄くしたもの） */
const TOOLBAR_ACCENT = "#605E5C";
/** ハイコントラストで選んだ棒を囲む線の幅 (px) */
const HC_SELECTED_WIDTH = 3;


interface DataLabelLine extends LabelLine {
    kind: "title" | "value" | "detail";
}

/** データラベルの行（タイトル・値・詳細の行。標準と同じくこの順に上から）。ラベルの表示を最適化なら、最大幅を超える行を「…」で切る */
function labelLinesOf(dl: DataLabelsSettings, d: DataPoint): DataLabelLine[] {
    const lines = rawLabelLinesOf(dl, d);
    if (!dl.optimizeLabelDisplay) return lines;
    return lines.map((line) => ({ ...line, text: truncateToWidth(line.text, dl.labelMaxWidth, line.font) }));
}

function rawLabelLinesOf(dl: DataLabelsSettings, d: DataPoint): DataLabelLine[] {
    const lines: DataLabelLine[] = [];
    if (dl.titleShow && d.titleText) {
        lines.push({
            kind: "title",
            text: d.titleText,
            font: {
                family: dl.titleFontFamily,
                size: dl.titleFontSize * PT_TO_PX,
                bold: dl.titleBold,
                italic: dl.titleItalic,
                underline: dl.titleUnderline,
            },
        });
    }
    if (dl.valueShow) {
        lines.push({
            kind: "value",
            text: d.dataLabelText,
            font: { family: dl.fontFamily, size: dl.fontSize * PT_TO_PX, bold: dl.bold, italic: dl.italic, underline: dl.underline },
        });
    }
    if (dl.detailShow && d.detailText) {
        lines.push({
            kind: "detail",
            text: d.detailText,
            font: {
                family: dl.detailFontFamily,
                size: dl.detailFontSize * PT_TO_PX,
                bold: dl.detailBold,
                italic: dl.detailItalic,
                underline: dl.detailUnderline,
            },
        });
    }
    return lines;
}

/** データラベルの行の class（テストとハーネスが行を見分ける） */
function labelClassOf(line: DataLabelLine): string {
    return line.kind === "value" ? "data-label" : line.kind === "title" ? "data-label-title" : "data-label-detail";
}

/**
 * データラベルの行の文字の見た目。色が空なら autoColor（棒の中は白か濃い灰色、外は灰色。labelAutoColor）。
 * 「符号の色」は棒の外か背景を付けたラベルだけに効かせる（toneHere）。棒の中の文字を棒と別の色にすると読めなくなるので、
 * 中は今までどおり読める色（標準も棒の中のラベルは白）
 */
function labelTextStyle(dl: DataLabelsSettings, line: DataLabelLine, d: DataPoint, autoColor: string, toneHere: boolean): React.CSSProperties {
    const own = line.kind === "detail" ? { color: dl.detailColor, transparency: dl.detailTransparency } : line.kind === "title" ? { color: dl.titleColor, transparency: dl.titleTransparency } : null;
    const fill = (own ? own.color : (toneHere && d.labelToneColor) || d.labelColor) || autoColor;
    return {
        fill,
        ...(own && own.transparency > 0 ? { fillOpacity: 1 - own.transparency / 100 } : {}),
        ...(!own && dl.transparency > 0 ? { fillOpacity: 1 - dl.transparency / 100 } : {}),
        ...fontStyleOf({ ...line.font, size: line.font.size / PT_TO_PX }),
        // 文字の縁。文字と縁の色が同じなら付けない（棒の中の自動の白い文字に白い縁を付けると、文字が太るだけ）
        ...haloStyle(dl.halo, fill),
    };
}

/**
 * 値の軸の目盛りの本数の上限。標準（powerbi-visuals-utils-chartutils の getRecommendedNumberOfTicksForYAxis・ForXAxis）と同じ。
 * 縦の軸は高さ、横の軸は幅で決める
 */
/**
 * 目盛りを、数字が重ならない本数まで減らす（「目盛りの本数 (目安)」を多く入れたときなど）。
 * length は描く範囲の長さ（px）、need は隣り合う目盛りの間に要る長さ（縦の軸は文字の行の高さ、横の軸はいちばん長い数字＋余白）
 */
export function fitTicks(ticksFor: (count: number) => Tick[], target: number, length: number, need: (ticks: Tick[]) => number): Tick[] {
    return fitTicksBy(ticksFor, target, (ticks) => Math.min(...ticks.slice(1).map((tick, i) => Math.abs(tick.ratio - ticks[i].ratio))) * length, need);
}

export { recommendedTickCount };



/**
 * 積み上げで下に別の棒がある棒の始まりの側を、系列間のスペース（px）だけ縮める（標準と同じく、値の端の位置は動かさない）。
 * lo・size は棒の範囲（縦棒なら top・height）、startAtLo は始まりの側が lo の側か
 */
export function insetSpan(lo: number, size: number, startAtLo: boolean, inset: number): { lo: number; size: number } {
    const cut = Math.max(0, Math.min(inset, size));
    return startAtLo ? { lo: lo + cut, size: size - cut } : { lo, size: size - cut };
}

/**
 * 棒を描く順（後に描いた棒が手前）。凡例の順（系列を並べ替えたら並べ替えた順）。重複を反転するときだけ逆にして、凡例の最初の系列を手前にする。
 * 返す番号は系列の番号（位置ではない）
 */
function drawOrderOf(points: DataPoint[], reverse: boolean, positions: number[]): Array<[DataPoint, number]> {
    const list = points.map((d, s) => [d, s] as [DataPoint, number]).sort(([, a], [, b]) => (positions[a] ?? a) - (positions[b] ?? b));
    return reverse ? list.reverse() : list;
}

/** 棒 (または棒のハイライト部分) の path。r > 0 なら値の向きの端（上か下）だけ角を丸める */
function barPathOf(left: number, top: number, width: number, height: number, r: number, roundAtTop: boolean): string {
    return valueEndBarPath(left, top, width, height, r, roundAtTop ? "top" : "bottom");
}

/** 横棒 (または横棒のハイライト部分) の path。r > 0 なら値の向きの端（右か左）だけ角を丸める */
function hBarPathOf(left: number, top: number, width: number, height: number, r: number, roundAtRight: boolean): string {
    return valueEndBarPath(left, top, width, height, r, roundAtRight ? "right" : "left");
}

/** 定数線のラベルの置き場所（作りは shared/referenceLine.ts の placeReferenceLabels） */
export const placeCategoryLineLabels = placeReferenceLabels;

export interface AppProps {
    viewModel: ViewModel;
    viewport: IViewport;
    settings: VisualFormattingSettingsModel;
    onSelect: (id: ISelectionId, multiSelect: boolean) => void;
    onContextMenu: (id: ISelectionId, x: number, y: number) => void;
    /** このビジュアルで選ばれている棒（または系列）。空なら選択なし */
    selectedIds?: ISelectionId[];
    /** 棒以外の所をクリックした（選択の解除） */
    onClearSelection?: () => void;
    /** パレートのランクの帯を押した。そのランクの棒をまとめて選ぶ */
    onSelectMany?: (ids: ISelectionId[], multiSelect: boolean) => void;
    /** 棒にマウスが入った。座標はクライアント座標（ルート基準への変換は visual.ts） */
    onTooltipShow?: (d: DataPoint, clientX: number, clientY: number) => void;
    /** 棒の上でマウスが動いた */
    onTooltipMove?: (d: DataPoint, clientX: number, clientY: number) => void;
    /** 棒からマウスが出た（折れ線の点からも） */
    onTooltipHide?: () => void;
    /** 折れ線の点にマウスが入った。lineIndex は lines の添字、pointIndex は categoryGroups の添字 */
    onLineTooltipShow?: (lineIndex: number, pointIndex: number, clientX: number, clientY: number) => void;
    /** 折れ線の点の上でマウスが動いた */
    onLineTooltipMove?: (lineIndex: number, pointIndex: number, clientX: number, clientY: number) => void;
    /** リボンの帯にマウスが入った。fromIndex は categoryGroups の添字（帯は fromIndex と次のカテゴリのあいだ） */
    onRibbonTooltipShow?: (seriesIndex: number, fromIndex: number, clientX: number, clientY: number) => void;
    /** リボンの帯の上でマウスが動いた */
    onRibbonTooltipMove?: (seriesIndex: number, fromIndex: number, clientX: number, clientY: number) => void;
    /** 閲覧者がグラフの上の「累計」ボタンを押した。保存は visual.ts */
    onToggleCumulative?: () => void;
    /** 閲覧者がグラフの上で累計の区切りを選んだ。保存は visual.ts */
    onChangeCumulativeReset?: (reset: string) => void;
    /** 操作を受け付けるか（ダッシュボードのタイルでは false）。false なら切り替えボタンを出さない */
    interactive?: boolean;
    /** レポートを編集中か（Desktop・Web の編集）。閲覧中（Web の閲覧）は空のときの案内を出さず、標準と同じく何も描かない */
    editing?: boolean;
    /** 画像のコピーのボタンの右クリックで、ブラウザーのメニュー（「画像をコピー」）を出すか（Desktop では出ないので false） */
    browserMenu?: boolean;
    /** 凡例で選んだ比較レイヤー（compareLayers の添字）。空ならすべて描く */
    visibleLayers?: number[];
    /** 凡例の比較レイヤーを押した。描くレイヤーを決めるのは visual.ts */
    onSelectLayer?: (layer: number, multiSelect: boolean) => void;
    /** 描いた値の軸の高さ (px、縦棒)。visual.ts が次の transform に渡し、棒の外に出すデータ ラベルの余白を値の軸に取る */
    onPlotHeight?: (height: number) => void;
}

/** 比較レイヤーの棒 1 枚ぶんの描き方。offset は系列の棒の中心からのずれ */
interface LayerDraw {
    width: number;
    offset: number;
    opacity: number;
    outline: boolean;
    /** データラベルを出すか（描いている中でいちばん手前のレイヤーだけ） */
    label: boolean;
}

/** 比較レイヤーの棒（0 が手前の「値」） */
function layerPointsOf(g: CategoryGroup, layer: number): DataPoint[] {
    return layer === 0 ? g.points : g.layerPoints?.[layer - 1] ?? [];
}

/**
 * カテゴリの start〜end（表示の順。end を含む）を選ぶときの ID。カテゴリのラベルや階層の区切りを押したときに使う。
 * 棒は系列ごとの棒の ID（カテゴリ＋系列）、「その他」はまとめたカテゴリ全部。カテゴリだけの ID にすると、系列があるとき棒の ID と
 * 一致せず選んだ棒が薄くなる（棒と同じ ID で選ぶ）。折れ線の点も、そのカテゴリの点の ID で一緒に選ぶ（棒だけだと線が選ばれず、
 * 線が沈む）。折れ線だけのときは置いてある見えない棒の ID を入れない（線の点と一致せず、選ぶと線が全部薄くなる）
 */
export function categorySelectionIds(viewModel: ViewModel, start: number, end: number): ISelectionId[] {
    const groups = viewModel.categoryGroups.slice(start, end + 1);
    const lineIds = lineSelectionIdsAt(viewModel, groups);
    if (viewModel.lineOnly) return lineIds;
    return [...groups.flatMap((g) => g.points.flatMap((d) => d.selectionIds ?? [d.selectionId])), ...lineIds];
}

/** groups のカテゴリにある折れ線の点の ID（値の無い点は除く。「その他」はまとめたカテゴリ全部） */
function lineSelectionIdsAt(viewModel: ViewModel, groups: CategoryGroup[]): ISelectionId[] {
    const rows = new Set(groups.map((g) => g.rowIndex));
    return viewModel.lines.flatMap((line) =>
        line.points.filter((p) => rows.has(p.rowIndex) && p.value !== null).flatMap((p) => p.selectionIds ?? [p.selectionId])
    );
}

/** 折れ線が選ばれているか（凡例で線ごと、または線の点。「その他」の点は、まとめたカテゴリの ID のどれか） */
export function linePicked(line: LineSeriesInfo | undefined, selectedIds: ISelectionId[]): boolean {
    if (!line) return false;
    return lineWholePicked(line, selectedIds) || line.points.some((p) => linePointPicked(p, selectedIds));
}

/** 凡例で線ごと選ばれているか（線の ID） */
function lineWholePicked(line: LineSeriesInfo, selectedIds: ISelectionId[]): boolean {
    return selectedIds.some((s) => s.equals(line.selectionId));
}

/** 線の点が選ばれているか（「その他」の点は、まとめたカテゴリの ID のどれか） */
function linePointPicked(p: LineSeriesInfo["points"][number], selectedIds: ISelectionId[]): boolean {
    return (p.selectionIds ?? [p.selectionId]).some((id) => selectedIds.some((s) => s.equals(id)));
}

/** 閲覧者向けの累計の切り替えボタンの行の高さ (px)。出すときだけグラフの上に取る */
const TOOLBAR_HEIGHT = 28;

const NO_SELECTION: ISelectionId[] = [];
const NO_LAYERS: number[] = [];

export const App: React.FC<AppProps> = ({
    viewModel,
    viewport,
    onSelect,
    onContextMenu,
    selectedIds = NO_SELECTION,
    onClearSelection,
    onSelectMany,
    onTooltipShow,
    onTooltipMove,
    onTooltipHide,
    onLineTooltipShow,
    onLineTooltipMove,
    onRibbonTooltipShow,
    onRibbonTooltipMove,
    onToggleCumulative,
    onChangeCumulativeReset,
    interactive = true,
    editing = true,
    browserMenu = false,
    visibleLayers = NO_LAYERS,
    onSelectLayer,
    onPlotHeight,
}) => {
    const [hoveredKey, setHoveredKey] = React.useState<string | null>(null);
    // 凡例が入りきらないときに出しているページ。凡例の項目が変わったら最初のページに戻す
    // （描いたあとで戻すと、前のページで 1 回描いてしまうので、項目の並びと一緒に持って描くときに比べる）
    const legendKey = JSON.stringify(viewModel.legendEntries.map((e) => e.name));
    const [legendPageState, setLegendPageState] = React.useState({ key: legendKey, page: 0 });
    const legendPage = legendPageState.key === legendKey ? legendPageState.page : 0;
    const setLegendPage = (page: number) => setLegendPageState({ key: legendKey, page });
    // 縦棒の描画で決まった値の軸の高さ。描いたあとに visual.ts へ知らせる（横棒・空のときは 0）
    const plotHeightRef = React.useRef(0);
    plotHeightRef.current = 0;
    // 目盛線を棒の形で切り抜く mask の ID（gridCutout）
    const gridMaskId = useGridMaskId("bar");
    React.useEffect(() => {
        if (plotHeightRef.current > 0) onPlotHeight?.(plotHeightRef.current);
    });

    // はみ出したときの最初の位置。縦棒は横に、横棒は縦にスクロールする。末尾なら、カテゴリの数・両端が変わったら当て直す
    const horizontalBars = viewModel.orientation === "horizontal";
    const scrollGroups = viewModel.categoryGroups;
    const scrollStart = useScrollStart<HTMLDivElement>({
        axis: horizontalBars ? "y" : "x",
        target: viewModel.categoryAxis.scrollStart === "end" ? "end" : "start",
        shape: JSON.stringify([
            viewModel.orientation,
            scrollGroups.length,
            scrollGroups[0]?.levelKeys ?? null,
            scrollGroups[scrollGroups.length - 1]?.levelKeys ?? null,
        ]),
    });

    if (viewModel.isEmpty) {
        // 何も描かなくても、背景のクリック（選択を外す）と右クリック（Power BI のメニュー）は効かせる。
        // フィールドを入れる前の案内は、編集中だけ出す（閲覧中は標準と同じく何も描かない）
        if (viewModel.noData || !editing)
            return (
                <div
                    className="unit-bar-empty"
                    style={{ width: viewport.width, height: viewport.height }}
                    onClick={() => onClearSelection?.()}
                    onContextMenu={(e) => {
                        e.preventDefault();
                        onContextMenu(null as unknown as ISelectionId, e.clientX, e.clientY);
                    }}
                />
            );
        return (
            <div className="unit-bar-landing">
                <p>カテゴリと値をフィールドに配置してください</p>
            </div>
        );
    }

    const lg = viewModel.legend;
    const legendFontPx = lg.fontSize * PT_TO_PX;

    // 比較レイヤー。凡例で選んだレイヤーがあれば、それだけを手前から並べ直して描く（薄くせずに描く）
    const layerCount = viewModel.compareLayers.length;
    const layered = layerCount > 1;
    const pickedLayers = visibleLayers.filter((l) => l < layerCount);
    const layerFilterOn = layered && pickedLayers.length > 0;
    const shownLayers = layered ? (layerFilterOn ? pickedLayers : viewModel.compareLayers.map((_, l) => l)) : [0];
    /** 系列の棒 1 本ぶんの幅に、描くレイヤーを並べる。キーはレイヤーの番号 */
    const layerDrawsFor = (slotWidth: number): Map<number, LayerDraw> => {
        // メジャーごとに並べるときは、集合の系列と同じく「系列間のスペース」を空けて並べる（最初のメジャーが左。横棒は上）
        const lay = viewModel.measureGroups
            ? (({ barWidth, offsets }) => ({ width: barWidth, offsets }))(clusterLayout(slotWidth, shownLayers.length, viewModel.columns.seriesSpacing, 0))
            : layerLayout(slotWidth, shownLayers.length, viewModel.compare.overlap, viewModel.compare.rightFront);
        return new Map(
            shownLayers.map((l, k): [number, LayerDraw] => [
                l,
                {
                    width: lay.width,
                    offset: lay.offsets[k],
                    opacity: layerFilterOn || viewModel.measureGroups ? 1 : layerOpacity(l, layerCount, viewModel.compare.backOpacity),
                    outline: layered && viewModel.compare.outline,
                    // メジャーごとに並べたときは、どの棒にもデータ ラベルを出す（比較では手前の棒だけ）
                    label: k === 0 || !!viewModel.measureGroups,
                },
            ])
        );
    };

    /**
     * 奥のレイヤーを薄く描くときは、目盛線を棒の形で切り抜く（目盛線はいちばん後ろ。半透明の棒の下で透けて見えないように）。
     * mask は描いた棒のまとまり（id が `${gridMaskId}-bars`）を写す
     */
    const gridMasked = (draws: Map<number, LayerDraw>) => layered && [...draws.values()].some((draw) => draw.opacity < 1);
    const gridMaskAttr = (draws: Map<number, LayerDraw>) => (gridMasked(draws) ? `url(#${gridMaskId})` : undefined);

    /** l 枚目のレイヤーの積み上げの合計を持つカテゴリ（メジャーごとに並べたとき、合計ラベルをレイヤーごとに出す） */
    const withLayerTotals = (g: CategoryGroup, l: number): CategoryGroup => (l === 0 ? g : { ...g, totals: g.layerTotals?.[l - 1] ?? null });

    // 閲覧者向けの累計の切り替え。出すぶんだけ上を空け、凡例とグラフはその下に描く
    const toolbarOn = viewModel.cumulative.toggle && interactive;
    const toolbarHeight = toolbarOn ? TOOLBAR_HEIGHT : 0;
    const bodyHeight = Math.max(10, viewport.height - toolbarHeight);

    // 凡例（複数系列のときだけ）。グラフの外側に置き、その分だけグラフの領域を縮める
    const legend: LegendLayout | null = lg.show
        ? layoutLegend({
            // 比較レイヤーは系列・折れ線と別のまとまりにして、見出し「比較」を付ける
            entries: viewModel.legendEntries.map((e) => ({ name: e.name, color: e.color, kind: e.kind, group: e.kind === "layer" ? 1 : 0 })),
            title: lg.title,
            // 凡例のタイトルをオフにしたら、「比較」の見出しも出さない（見出しの代わりに間を広げる）
            groupTitles: lg.titleShow ? { 1: "比較" } : undefined,
            font: { family: lg.fontFamily, size: legendFontPx, bold: lg.bold, italic: lg.italic, underline: lg.underline },
            position: lg.position,
            width: viewport.width,
            height: bodyHeight,
            horizontal: viewModel.orientation === "horizontal",
            page: legendPage,
        })
        : null;

    /**
     * 自分で選んだ棒か。凡例のクリックで系列ごと選んだときは、その系列の棒をすべて含む
     * （自分で選んだ場合、Power BI は選んだビジュアル自身にはハイライトを送らない）
     */
    const isPicked = (d: DataPoint): boolean => {
        const seriesId = viewModel.series[d.seriesIndex]?.selectionId;
        const categoryId = d.categorySelectionId;
        // 「その他」は、まとめたカテゴリが全部選ばれているか、凡例でその系列が選ばれているかで見る
        // （仮の ID は、まとめた先頭のカテゴリのものなので比べない）
        if (d.selectionIds?.length) {
            return (
                d.selectionIds.every((id) => selectedIds.some((s) => s.equals(id))) ||
                (seriesId ? selectedIds.some((s) => s.equals(seriesId)) : false)
            );
        }
        return selectedIds.some(
            (s) => s.equals(d.selectionId) || (seriesId ? s.equals(seriesId) : false) || (categoryId ? s.equals(categoryId) : false)
        );
    };

    /** この線を第 2 軸で描くか。パレートでは累積比の線（選べない線）だけが第 2 軸で、ほかは左の軸を共有する */
    const onAxis2 = (line: LineSeriesInfo) =>
        viewModel.valueAxis2.show && viewModel.lines.length > 0 && (!viewModel.pareto.enabled || line.selectable === false);

    /** パレートのランクの帯の区切り（同じランクが続くところ。end は含まない）。0 以下のカテゴリ（ランク無し）は区切りを作らない */
    const rankRuns = (): Array<{ start: number; end: number; rank: ParetoRank }> => {
        const ranks = viewModel.pareto.ranks;
        const runs: Array<{ start: number; end: number; rank: ParetoRank }> = [];
        let start = 0;
        for (let i = 1; i <= ranks.length; i++) {
            if (i < ranks.length && ranks[i] === ranks[start]) continue;
            const rank = ranks[start];
            if (rank) runs.push({ start, end: i, rank });
            start = i;
        }
        return runs;
    };

    /**
     * パレートのランクの帯の 1 区切り。角の丸い帯に名前と件数を書く。
     * 押すか Enter・Space で、そのランクのカテゴリをまとめて選ぶ（Ctrl で足す）
     */
    const renderRankRun = (
        run: { start: number; end: number; rank: ParetoRank },
        box: { x: number; y: number; width: number; height: number },
        font: FontSpec,
        labelStyle: React.CSSProperties
    ) => {
        const pareto = viewModel.pareto;
        const { rank } = run;
        const count = run.end - run.start;
        const full = `${pareto.labels[rank]} ${count}件`;
        const available = box.width - 8;
        const text = measureTextWidth(full, font) <= available ? full : pareto.labels[rank];
        const showText = measureTextWidth(text, font) <= available && box.height >= font.size;
        const ids = pareto.selections[rank];
        const picked = ids.length > 0 && ids.every((id) => selectedIds.some((s) => s.equals(id)));
        const select = (multi: boolean) => onSelectMany?.(ids, multi);
        return (
            <g
                key={`rank-${run.start}`}
                className="pareto-rank"
                {...keysOf({ label: full, onSelect: select, onClear: onClearSelection, pressed: picked })}
                style={{ cursor: "pointer" }}
                onClick={(e) => {
                    e.stopPropagation();
                    select(e.ctrlKey || e.metaKey);
                }}
            >
                <rect
                    x={box.x}
                    y={box.y}
                    width={Math.max(0, box.width)}
                    height={Math.max(0, box.height)}
                    rx={4}
                    style={{ fill: pareto.colors[rank], fillOpacity: picked ? 0.45 : 0.22 }}
                />
                {showText && (
                    <text
                        x={box.x + box.width / 2}
                        y={box.y + box.height / 2 + font.size * 0.35}
                        textAnchor="middle"
                        className="pareto-rank-label"
                        style={labelStyle}
                    >
                        {text}
                    </text>
                )}
            </g>
        );
    };

    const levelIdsOf = (run: { start: number; end: number }) => categorySelectionIds(viewModel, run.start, run.end);
    const lineIdsAt = (groups: CategoryGroup[]) => lineSelectionIdsAt(viewModel, groups);
    /** 区切りの棒が全部選ばれているか（区切りを押したときも、棒を Ctrl で全部選んだときも） */
    const levelPicked = (run: { start: number; end: number }) => {
        if (selectedIds.length === 0) return false;
        const groups = viewModel.categoryGroups.slice(run.start, run.end + 1);
        if (viewModel.lineOnly) {
            const ids = lineIdsAt(groups);
            return ids.length > 0 && ids.every((id) => selectedIds.some((s) => s.equals(id)));
        }
        const points = groups.flatMap((g) => g.points);
        return points.length > 0 && points.every(isPicked);
    };

    /**
     * 階層の区切りの当たり。押すか Enter・Space で、その区切りのカテゴリをまとめて選ぶ（Ctrl で足す）。
     * 見た目は変えない（帯の濃さは選んだときに少し上げる）。名前はツールヒントに出す。
     * 読み上げの名前は上のレベルからの道筋（「FY26 H1」）。同じ名前の区切り（別の年の H1）を見分けられるように
     */
    const renderLevelHit = (
        run: { start: number; end: number; text: string },
        level: number,
        box: { x: number; y: number; width: number; height: number },
        key: string
    ) => {
        const select = (multi: boolean) => onSelectMany?.(levelIdsOf(run), multi);
        // 読み上げ・ツールヒントの名前は、軸から外したドリルの共通の親も上に付ける
        const path = [...viewModel.commonLevels, ...(viewModel.categoryGroups[run.start]?.levels.slice(0, level + 1) ?? [])].join(" ") || run.text;
        return (
            <g
                key={key}
                className="level-hit"
                {...keysOf({ label: path, onSelect: select, onClear: onClearSelection, pressed: levelPicked(run) })}
                style={{ cursor: "pointer" }}
                onClick={(e) => {
                    e.stopPropagation();
                    select(e.ctrlKey || e.metaKey);
                }}
            >
                <title>{path}</title>
                <rect x={box.x} y={box.y} width={Math.max(0, box.width)} height={Math.max(0, box.height)} fill="transparent" />
            </g>
        );
    };

    /**
     * カテゴリのラベルの当たり（標準と同じく、ラベルを押すとそのカテゴリの棒をまとめて選ぶ。Ctrl で足す、Enter・Space も）。
     * i は表示の順のカテゴリの番号、name は読み上げの名前（上のレベルからつないだ名前）。ラベルを包む g に付ける
     */
    const categoryLabelHit = (i: number, name: string) => {
        const run = { start: i, end: i };
        const select = (multi: boolean) => onSelectMany?.(levelIdsOf(run), multi);
        return {
            className: "level-hit category-label-hit",
            ...keysOf({ label: name, onSelect: select, onClear: onClearSelection, pressed: levelPicked(run) }),
            style: { cursor: "pointer" },
            onClick: (e: React.MouseEvent) => {
                e.stopPropagation();
                select(e.ctrlKey || e.metaKey);
            },
        };
    };

    /**
     * 縦棒のカテゴリのラベルを当たりで包む。ラベルを出さない（欄に入らない）ときと、カテゴリを入れていないときは包まない
     * （見えない・名前の無いボタンに Tab で止まらないように）
     */
    const withCategoryLabelHit = (i: number, name: string, label: React.ReactNode) => {
        if (label === null || label === undefined || label === false) return null;
        if (viewModel.noCategory) return label;
        return <g {...categoryLabelHit(i, name)}>{label}</g>;
    };

    /**
     * キーボードで選べる部品の属性。操作できない場所（ダッシュボードのタイルなど）では Tab で止めず、名前だけ付ける
     */
    const keysOf = (options: Parameters<typeof selectKeyProps>[0]) => (interactive ? selectKeyProps(options) : { "aria-label": options.label });

    /**
     * 透かして描いた棒の見た目の色（背景と混ざった色）。棒の中のラベルの自動の文字色を、見えている色で決める
     * （透過性や比較レイヤーの奥で薄くなった棒の上に白い文字を置かない）
     */
    const seenColorOf = (color: string, opacity: number) =>
        viewModel.highContrast || opacity >= 1 ? color : blend(color, viewModel.background ?? "#FFFFFF", 1 - opacity);

    const isLinePicked = (lineIndex: number): boolean => linePicked(viewModel.lines[lineIndex], selectedIds);

    /** 凡例の項目が選ばれているか（棒の系列なら、その系列の棒のどれかが選ばれている） */
    const isEntryPicked = (entry: LegendItemInfo | undefined): boolean =>
        !!entry &&
        (entry.kind === "layer"
            ? false
            : entry.kind === "line"
            ? isLinePicked(entry.index)
            : viewModel.dataPoints.some((d) => d.seriesIndex === entry.index && isPicked(d)));

    /**
     * 折れ線のデータ ラベル。点の上か下（横棒では右か左）に値を置く。自動は上（右）で、プロットに入らないか先に描いたラベルと重なれば反対の側。
     * boxes に、先に描いたラベル（棒のデータ ラベル・合計ラベル・前の線のラベル）の枠があれば、重なるラベルは描かない。
     * bounds は値の軸の向きのプロットの端（縦棒は上端・下端の y、横棒は左端・右端の x）
     */
    const renderLineLabels = (
        line: LineSeriesInfo,
        j: number,
        posOf: (i: number) => number,
        valueOf: (ratio: number) => number,
        horizontal: boolean,
        bounds: [number, number],
        boxes: Array<{ x0: number; x1: number; y0: number; y1: number }>
    ) => {
        const labels = line.labels;
        if (!labels) return null;
        // 線ごとに変えた設定（保存の無い項目は「すべて」の値）
        const dl = labels.settings;
        const fontPx = dl.fontSize * PT_TO_PX;
        const font: FontSpec = { family: dl.fontFamily, size: fontPx, bold: dl.bold, italic: dl.italic, underline: dl.underline };
        const fill = labels.color || line.color;
        // 点（マーカー）から離す間。マーカーが大きければその半径の外へ
        const markerRadius = line.marker?.show ? (line.marker.size ?? 5) / 2 : line.width / 2;
        const gap = Math.max(4, markerRadius + 3);
        const texts = line.points.map((p, i) => {
            const text = labels.texts[i];
            if (text === null || text === undefined || p.ratio === null) return null;
            const pos = posOf(i);
            const val = valueOf(p.ratio);
            const width = measureTextWidth(text, font);
            // 置き方の候補。上（横棒は右）なら first=true。自動は入る側を先に、重なったら反対の側を試す
            const placementOf = (first: boolean) => {
                if (horizontal) {
                    // 横棒では値の軸が横。「上」は右、「下」は左
                    const x = first ? val + gap : val - gap;
                    const box = first ? { x0: x, x1: x + width, y0: pos - fontPx / 2, y1: pos + fontPx / 2 } : { x0: x - width, x1: x, y0: pos - fontPx / 2, y1: pos + fontPx / 2 };
                    return { x, y: pos + fontPx * 0.35, anchor: (first ? "start" : "end") as "start" | "end" | "middle", box, inside: first ? box.x1 <= bounds[1] : box.x0 >= bounds[0] };
                }
                const box = { x0: pos - width / 2, x1: pos + width / 2, y0: first ? val - gap - fontPx : val + gap, y1: first ? val - gap : val + gap + fontPx };
                return { x: pos, y: first ? val - gap - fontPx * 0.2 : val + gap + fontPx * 0.8, anchor: "middle" as "start" | "end" | "middle", box, inside: first ? box.y0 >= bounds[0] : box.y1 <= bounds[1] };
            };
            const overlaps = (box: { x0: number; x1: number; y0: number; y1: number }) =>
                boxes.some((b) => b.x0 < box.x1 - 1 && box.x0 < b.x1 - 1 && b.y0 < box.y1 - 1 && box.y0 < b.y1 - 1);
            const candidates =
                labels.position === "above" ? [placementOf(true)] : labels.position === "below" ? [placementOf(false)] : [placementOf(true), placementOf(false)];
            if (labels.position === "auto" && !candidates[0].inside) candidates.reverse();
            // 自動ではプロットからはみ出す側には置かない（切れて読めなくなる）。どちらの側も入らないか重なれば描かない
            const chosen = candidates.find((c) => !overlaps(c.box) && (labels.position !== "auto" || c.inside));
            if (!chosen) return null;
            const { x, y, anchor, box } = chosen;
            boxes.push(box);
            // 背景は棒のラベルと同じく「背景」がオンのとき。棒の上で読みにくければ、線ごとに「文字の縁」を付ける
            return [
                dl.backgroundShow && (
                    <rect
                        key={`line-label-bg-${j}-${i}`}
                        className="line-data-label-bg"
                        x={box.x0 - 2}
                        y={box.y0 - 1}
                        width={box.x1 - box.x0 + 4}
                        height={box.y1 - box.y0 + 2}
                        rx={2}
                        pointerEvents="none"
                        style={{ fill: dl.backgroundColor, fillOpacity: 1 - dl.backgroundTransparency / 100 }}
                    />
                ),
                <text
                    key={`line-label-${j}-${i}`}
                    className="line-data-label"
                    x={x}
                    y={y}
                    textAnchor={anchor}
                    pointerEvents="none"
                    style={{
                        fill,
                        ...(dl.transparency > 0 ? { fillOpacity: 1 - dl.transparency / 100 } : {}),
                        ...fontStyleOf(fontOf(dl)),
                        ...haloStyle(dl.halo, fill),
                    }}
                >
                    {text}
                </text>,
            ];
        });
        return <g key={`line-labels-${j}`} className="line-data-labels">{texts}</g>;
    };

    /**
     * 折れ線 1 本。点はカテゴリの中心に置き、値の無いカテゴリでは線を切る。
     * 点の上にマウスの当たり判定の円を置き、ツールヒントと選択を受ける。
     * posOf(i) はカテゴリの軸の座標、valueOf(ratio) は値の軸の座標。横棒（horizontal）では点が上から下へ並ぶ。
     * barBand はカテゴリ 1 つの棒のまとまりの幅（集合なら系列の棒の端から端まで）。「段の幅」が「棒の幅」のときの線の長さ
     */
    const renderLineSeries = (
        line: LineSeriesInfo,
        j: number,
        part: "area" | "line",
        posOf: (i: number) => number,
        valueOf: (ratio: number) => number,
        horizontal: boolean,
        stepExtension: number,
        stepBounds: [number, number],
        barBand: number
    ) => {
        // 空白の前後をつなぐか（「線」カードの「空白の値」）。ビジュアルが計算した線（パレートの累積比）は対象外
        const connectBlanks = viewModel.lineBlanks === LINE_BLANKS.connect && line.selectable !== false;
        const pts = line.points.map((p, i) => {
            const pos = posOf(i);
            const val = p.ratio === null ? null : valueOf(p.ratio);
            const at: XY | null = val === null ? null : horizontal ? { x: val, y: pos } : { x: pos, y: val };
            return { p, i, at };
        });
        // 値の無いカテゴリで切った、途切れのない点の並び
        const runsOf = (points: Array<{ p: LineSeriesInfo["points"][number]; at: XY | null }>, joinBlanks: boolean): XY[][] => {
            const out: XY[][] = [];
            let run: XY[] = [];
            for (const pt of points) {
                // 累計の区切りで 0 に戻るところも切る（前の区切りの終わりから下がる線を引かない）
                if (pt.p.breakBefore && run.length) {
                    out.push(run);
                    run = [];
                }
                if (pt.at === null) {
                    // 前後をつなぐなら、空白を飛ばして同じ並びに続ける
                    if (joinBlanks) continue;
                    if (run.length) out.push(run);
                    run = [];
                    continue;
                }
                run.push(pt.at);
            }
            if (run.length) out.push(run);
            return out;
        };
        const runs = runsOf(pts, connectBlanks);
        const shape = {
            interpolation: line.interpolation,
            smoothing: line.smoothing,
            tension: line.tension / 100,
            stepPosition: line.stepPosition,
            stepConnect: line.stepConnect,
            stepLevelWidth: line.stepWidth === STEP_WIDTHS.bar ? barBand : 0,
        };
        const pathsOf = (rs: XY[][]) => {
            const runPaths = rs.map((r) => linePath(r, shape, horizontal, stepExtension, stepBounds));
            // 前後に点の無い 1 点は、線が長さ 0 になって見えない。マーカーを出していなくても、線の色の点を置く（標準と同じ）
            const lone = rs.filter((r, k) => r.length === 1 && !/[LCQ]/.test(runPaths[k])).map((r) => r[0]);
            return { path: runPaths.join(" "), lone };
        };
        const { path, lone: lonePoints } = pathsOf(runs);
        // ほかのビジュアルで選ばれたとき（ハイライト）は、棒と同じく線全体を沈め、該当分の線を濃く重ねる。該当しないカテゴリでは線を切る
        const highlighting = viewModel.hasHighlights && line.points.some((p) => p.highlight !== undefined);
        const highlightAt = (p: LineSeriesInfo["points"][number], i: number): XY | null => {
            if (p.highlightRatio === null || p.highlightRatio === undefined) return null;
            const pos = posOf(i);
            const val = valueOf(p.highlightRatio);
            return horizontal ? { x: val, y: pos } : { x: pos, y: val };
        };
        // 凡例で線ごと選んだときは線全体を濃く。点（カテゴリのラベル・線の点）で選んだときは、ハイライトと同じく線を沈め、
        // 選んだカテゴリの線を濃く重ねる（中で選んでも外で選ばれても同じ見え方）
        const selecting = !viewModel.hasHighlights && selectedIds.length > 0;
        const whole = selecting && lineWholePicked(line, selectedIds);
        const pickedAt = (p: LineSeriesInfo["points"][number], at: XY | null) => (linePointPicked(p, selectedIds) ? at : null);
        const highlightPaths = highlighting
            ? pathsOf(runsOf(pts.map((pt) => ({ p: pt.p, at: highlightAt(pt.p, pt.i) })), false))
            : selecting && !whole
                ? pathsOf(runsOf(pts.map((pt) => ({ p: pt.p, at: pickedAt(pt.p, pt.at) })), false))
                : null;
        const pointPicked = (p: LineSeriesInfo["points"][number]) =>
            highlighting ? p.highlight !== null && p.highlight !== undefined : !selecting || whole || linePointPicked(p, selectedIds);
        const dimmed = (selecting && !whole) || highlighting;
        const opacity = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        const stroke = lineStrokeOf(line, line.width);
        const lineOpacity = opacity * (1 - line.transparency / 100);
        const mk = line.marker ?? viewModel.markers;
        const markerFill = mk.color || line.color;
        const pointOpacity = (p: LineSeriesInfo["points"][number]) => (pointPicked(p) ? 1 : HIGHLIGHT_DIM_FACTOR);
        const markerOpacity = Math.max(0, Math.min(1, 1 - mk.transparency / 100));
        const markerStroke = mk.borderShow ? (mk.borderMatchLine ? line.color : mk.borderColor) : "none";
        const markerStrokeOpacity = Math.max(0, Math.min(1, 1 - mk.borderTransparency / 100));
        const hitR = Math.max(6, line.width * 2);
        if (part === "area") {
            // 網掛け領域は棒の上に、透かして塗る（標準と同じ）
            if (!line.areaShow) return null;
            const area = viewModel.areas;
            const baseline = valueOf(line.baselineRatio);
            return runs.map((r, k) => (
                <path
                    key={`area-${j}-${k}`}
                    d={areaPath(r, baseline, shape, horizontal, stepExtension, stepBounds)}
                    className="line-area"
                    fill={area.matchLineColor ? line.color : area.fill}
                    fillOpacity={opacity * Math.max(0, Math.min(1, 1 - area.transparency / 100))}
                    stroke="none"
                    pointerEvents="none"
                />
            ));
        }
        return (
            <g key={`line-${j}`} className="line-series">
                {/* 線は「すべての系列に表示」「このシリーズに表示」がオンのときだけ（マーカーは別に出す） */}
                {line.lineShow && (
                    <path
                        d={path}
                        className="line-path"
                        fill="none"
                        stroke={line.color}
                        strokeWidth={line.width}
                        strokeDasharray={stroke.dashArray}
                        strokeLinejoin={line.lineJoin === "miter" || line.lineJoin === "bevel" ? line.lineJoin : "round"}
                        strokeLinecap={stroke.lineCap}
                        strokeOpacity={lineOpacity}
                        pointerEvents="none"
                    />
                )}
                {line.lineShow &&
                    !mk.show &&
                    lonePoints.map((at, k) => (
                        <circle
                            key={`lone-${j}-${k}`}
                            cx={at.x}
                            cy={at.y}
                            r={Math.max(LONE_POINT_MIN_RADIUS, line.width)}
                            className="line-lone-point"
                            fill={line.color}
                            fillOpacity={lineOpacity}
                            pointerEvents="none"
                        />
                    ))}
                {line.lineShow && highlightPaths && (
                    <>
                        <path
                            d={highlightPaths.path}
                            className="line-path line-highlight"
                            fill="none"
                            stroke={line.color}
                            strokeWidth={line.width}
                            strokeDasharray={stroke.dashArray}
                            strokeLinejoin={line.lineJoin === "miter" || line.lineJoin === "bevel" ? line.lineJoin : "round"}
                            strokeLinecap={stroke.lineCap}
                            strokeOpacity={1 - line.transparency / 100}
                            pointerEvents="none"
                        />
                        {/* 濃い区間が 1 点だけ（前後が選ばれていない）なら線にならないので、マーカーの形で出す（マーカーを出していない線も） */}
                        {!mk.show &&
                            highlightPaths.lone.map((at, k) => (
                                <path
                                    key={`hl-lone-${j}-${k}`}
                                    d={markerPath(mk.shape, at.x, at.y, Math.max(mk.size, line.width + 2), horizontal)}
                                    transform={mk.rotation ? `rotate(${mk.rotation} ${at.x} ${at.y})` : undefined}
                                    className="line-picked-point"
                                    fill={line.color}
                                    fillOpacity={1 - line.transparency / 100}
                                    pointerEvents="none"
                                />
                            ))}
                    </>
                )}
                {pts.map((pt) => {
                    if (pt.at === null) return null;
                    // ハイライトの該当分があるマーカーは、該当分の高さに置く（濃く重ねた線の上）
                    const mkAt = (highlighting && highlightAt(pt.p, pt.i)) || pt.at;
                    return (
                        <g key={`pt-${j}-${pt.i}`}>
                            {mk.show && (
                                <path
                                    d={markerPath(mk.shape, mkAt.x, mkAt.y, mk.size, horizontal)}
                                    transform={mk.rotation ? `rotate(${mk.rotation} ${mkAt.x} ${mkAt.y})` : undefined}
                                    className="line-marker"
                                    style={{
                                        fill: markerFill,
                                        fillOpacity: pointOpacity(pt.p) * markerOpacity,
                                        stroke: markerStroke,
                                        strokeWidth: mk.borderShow ? mk.borderWidth : 0,
                                        strokeOpacity: pointOpacity(pt.p) * markerStrokeOpacity,
                                    }}
                                    pointerEvents="none"
                                />
                            )}
                            <circle
                                cx={pt.at.x}
                                cy={pt.at.y}
                                r={hitR}
                                className="line-hit"
                                fill="transparent"
                                onClick={(e) => {
                                    e.stopPropagation();
                                    if (pt.p.selectionIds) onSelectMany?.(pt.p.selectionIds, e.ctrlKey || e.metaKey);
                                    else onSelect(pt.p.selectionId, e.ctrlKey || e.metaKey);
                                }}
                                onContextMenu={(e) => {
                                    e.preventDefault();
                                    e.stopPropagation();
                                    // 「その他」は特定のカテゴリにしない（仮の ID を渡すと、まとめた先頭のカテゴリのメニューになる）
                                    onContextMenu(pt.p.selectionIds ? (null as unknown as ISelectionId) : pt.p.selectionId, e.clientX, e.clientY);
                                }}
                                onMouseEnter={(e) => onLineTooltipShow?.(j, pt.i, e.clientX, e.clientY)}
                                onMouseMove={(e) => onLineTooltipMove?.(j, pt.i, e.clientX, e.clientY)}
                                onMouseLeave={() => onTooltipHide?.()}
                            />
                        </g>
                    );
                })}
            </g>
        );
    };

    /**
     * リボンの帯。同じ系列の棒を、隣のカテゴリの棒と S 字の帯でつなぐ（棒の後ろに描く）。
     * 罫線は標準と同じく帯の両縁だけに引く。
     * gapOf(i) は i 番目と i+1 番目のカテゴリの棒のあいだ（カテゴリの軸の座標）、extentOf(d) は棒の両端（値の軸の座標）。
     * 縦棒ではカテゴリの軸が x、横棒では y
     */
    const renderRibbonBands = (
        groups: CategoryGroup[],
        gapOf: (i: number) => [number, number],
        extentOf: (d: DataPoint) => [number, number],
        horizontal: boolean
    ): React.ReactNode[] => {
        const rb = viewModel.ribbons;
        const at = (u: number, v: number) => (horizontal ? `${v},${u}` : `${u},${v}`);
        const nodes: React.ReactNode[] = [];
        for (let i = 0; i + 1 < groups.length; i++) {
            const [from, to] = gapOf(i);
            const pad = (to - from) * (rb.spacing / 100);
            const ua = from + pad;
            const ub = to - pad;
            if (ub <= ua) continue;
            const um = (ua + ub) / 2;
            groups[i].points.forEach((a, s) => {
                const b = groups[i + 1].points[s];
                if (!b || a.blank || b.blank) return;
                const [a0, a1] = extentOf(a);
                const [b0, b1] = extentOf(b);
                if (a1 - a0 <= 0 && b1 - b0 <= 0) return;
                const edgeStart = `M ${at(ua, a0)} C ${at(um, a0)} ${at(um, b0)} ${at(ub, b0)}`;
                const edgeEnd = `M ${at(ub, b1)} C ${at(um, b1)} ${at(um, a1)} ${at(ua, a1)}`;
                const band = `${edgeStart} L ${at(ub, b1)} C ${at(um, b1)} ${at(um, a1)} ${at(ua, a1)} Z`;
                const color = rb.matchSeriesColor ? (viewModel.seriesMode ? viewModel.series[s]?.color ?? a.color : a.color) : rb.fill;
                const dimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isPicked(a) && !isPicked(b);
                const opacity = (1 - rb.transparency / 100) * (dimmed ? HIGHLIGHT_DIM_FACTOR : 1);
                const borderColor = rb.borderMatchRibbon ? color : rb.borderFill;
                nodes.push(
                    <g key={`ribbon-${i}-${s}`} className="ribbon-group">
                        <path
                            d={band}
                            className="ribbon"
                            style={{ fill: color, fillOpacity: opacity }}
                            onClick={(e) => {
                                e.stopPropagation();
                                const id = viewModel.series[s]?.selectionId ?? a.selectionId;
                                onSelect(id, e.ctrlKey || e.metaKey);
                            }}
                            onMouseEnter={(e) => onRibbonTooltipShow?.(s, i, e.clientX, e.clientY)}
                            onMouseMove={(e) => onRibbonTooltipMove?.(s, i, e.clientX, e.clientY)}
                            onMouseLeave={() => onTooltipHide?.()}
                        />
                        {rb.borderShow &&
                            [edgeStart, edgeEnd].map((curve, k) => (
                                <path
                                    key={k}
                                    d={curve}
                                    className="ribbon-border"
                                    fill="none"
                                    stroke={borderColor}
                                    strokeWidth={rb.borderWidth}
                                    strokeOpacity={1 - rb.borderTransparency / 100}
                                    pointerEvents="none"
                                />
                            ))}
                    </g>
                );
            });
        }
        return nodes;
    };

    /**
     * ドリルダウンした位置（事業A ＞ 製品A1）。カテゴリの軸と同じ文字で、左端から maxWidth まで。
     * 入りきらなければ末尾を「…」で省き、マウスを乗せると全部出す（省かないときも出す）
     */
    const drillPathFont = (): FontSpec => ({
        family: viewModel.categoryAxis.fontFamily,
        size: viewModel.categoryAxis.fontSize * PT_TO_PX,
        bold: viewModel.categoryAxis.bold,
        italic: viewModel.categoryAxis.italic,
        underline: viewModel.categoryAxis.underline,
    });
    /** maxWidth に収まるドリルの位置の文字。1 文字と「…」も入らなければ空（出さず、行もとらない） */
    const fittedDrillPath = (maxWidth: number): string => {
        const path = viewModel.drillPath;
        if (!path || maxWidth < 1) return "";
        const font = drillPathFont();
        const shown = truncateText(path, maxWidth, font);
        return measureTextWidth(shown, font) <= maxWidth + 0.5 ? shown : "";
    };
    const renderDrillPath = (x: number, y: number, maxWidth: number) => {
        const path = viewModel.drillPath;
        const font = drillPathFont();
        const shown = fittedDrillPath(maxWidth);
        if (!shown) return null;
        return (
            <text
                x={x}
                y={y}
                className="drill-path"
                textAnchor="start"
                pointerEvents="visiblePainted"
                style={{
                    ...fontStyleOf({ ...font, size: viewModel.categoryAxis.fontSize }),
                    fill: viewModel.categoryAxis.labelColor,
                }}
            >
                {shown}
                <title>{path}</title>
            </text>
        );
    };
    /** 単位のラベルの幅（ドリルの位置を並べるときに空ける）。文字はカテゴリの軸の書体で見積もる */
    const badgeWidthOf = (text: string, fontPt: number) =>
        text ? measureTextWidth(text, { ...drillPathFont(), size: fontPt * PT_TO_PX, bold: false, italic: false, underline: false }) : 0;

    /** グラフ本体（軸・棒・ラベル）。width・height は凡例を除いた領域 */
    const renderChart = (width: number, height: number) => {
        const badgeText = viewModel.unitInfo.badgeText;
        const catAxis = viewModel.categoryAxis;
        const valAxis = viewModel.valueAxis;
        const gridlines = viewModel.gridlines;
        const columnsSettings = viewModel.columns;

        // Y軸目盛ラベルの文字数
        let maxTickChars = 2;
        if (valAxis.show) {
            // 目盛りの本数は描く範囲の高さで決まる（下の valueTicks）。いちばん細かい 8 本の数字の幅も見積もりに入れる
            for (const t of [...viewModel.ticks, ...viewModel.ticksFor(Math.max(8, valAxis.tickCount))]) {
                if (t.label.length > maxTickChars) {
                    maxTickChars = t.label.length;
                }
            }
        }
        const isTopBadge = badgeText && viewModel.unitInfo.unitPosition === "valueAxisTop";
        const badgeChars = isTopBadge ? badgeText.length : 0;

        // Y軸タイトル・目盛の精密レイアウト (Power BI 標準準拠)
        const isRightAxis = valAxis.switchPosition;
        const valTitleFontSizePx = valAxis.titleFontSize * PT_TO_PX;
        const valTickFontSizePx = valAxis.fontSize * PT_TO_PX;
        const tickWidthPx = valAxis.show ? Math.max(14, maxTickChars * (valTickFontSizePx * 0.55)) : 0;
        const badgeWidthPx = isTopBadge ? Math.max(20, badgeChars * (viewModel.unitInfo.fontSize * PT_TO_PX * 0.7)) : 0;
        const tickToAxisGap = 6;
        const gapTitleToTicks = isRightAxis ? 6 : 4;
        const leftPadding = 4;
        const rightPadding = 12;

        // 標準と同じく「値」(目盛りラベル) とタイトルは独立。値を OFF にしてもタイトルは残す
        const hasYTitle = valAxis.titleShow && Boolean(valAxis.titleText);
        const axisContentWidth = Math.max(tickWidthPx, badgeWidthPx) + tickToAxisGap;

        // 第 2 Y 軸（折れ線を右の軸で描くとき）。左の軸（Y 軸）の反対側に置く
        const axis2 = viewModel.valueAxis2;
        const axis2On = axis2.show && viewModel.lines.length > 0;
        const axis2TickFontPx = axis2.fontSize * PT_TO_PX;
        const axis2TickWidth = axis2On && axis2.valueShow
            ? Math.max(14, Math.max(2, ...[...axis2.ticks, ...axis2.ticksFor(Math.max(8, valAxis.tickCount))].map((t) => t.label.length)) * axis2TickFontPx * 0.55)
            : 0;
        const axis2TitleFontPx = axis2.titleFontSize * PT_TO_PX;
        const hasAxis2Title = axis2On && axis2.titleShow && Boolean(axis2.titleText);
        // 第 2 Y 軸の上の単位ラベル（Y 軸の単位ラベルと同じ置き方）
        const axis2BadgeText = axis2On ? axis2.badgeText : "";
        const axis2BadgeWidth = axis2BadgeText
            ? Math.max(20, axis2BadgeText.length * (axis2.unitFontSize * PT_TO_PX * 0.7))
            : 0;
        const axis2Width = axis2On
            ? Math.max(axis2TickWidth, axis2BadgeWidth) + tickToAxisGap + (hasAxis2Title ? axis2TitleFontPx + gapTitleToTicks : 0) + 4
            : 0;

        const marginLeft = !isRightAxis
            ? (hasYTitle
                ? leftPadding + valTitleFontSizePx + gapTitleToTicks + axisContentWidth
                : Math.max(16, leftPadding + axisContentWidth))
            : Math.max(24, axis2Width + leftPadding);

        const marginRight = isRightAxis
            ? (hasYTitle
                ? axisContentWidth + gapTitleToTicks + valTitleFontSizePx + rightPadding
                : Math.max(24, axisContentWidth + rightPadding))
            : Math.max(24, axis2Width + rightPadding);
        const marginTopBase = badgeText || axis2BadgeText ? 36 : 22;
        // ドリルの位置（左上）。単位のラベルと同じ行に出す（行を積まず、描く範囲を狭めない）。
        // 単位のラベルが無くても同じ高さの行をとり、文字が大きければ行を広げる。
        // 描く範囲の上にはみ出したデータラベル・合計ラベルとは、描いたラベルの枠で重なりを見て、その手前で省く（renderVerticalDrillPath）
        const drillFontPx = catAxis.fontSize * PT_TO_PX;
        const marginTopWithDrill = Math.max(36, Math.ceil(drillFontPx) + 16);
        const drillY = marginTopWithDrill - 12;
        // 左の列に単位（Y 軸の単位、Y 軸を右にしたときは第 2 Y 軸の単位）があれば、その右端（描く範囲の 6px 手前）から 10px 空ける
        // （6px だと「(千円) FY26」が 1 つの文字に見えた）。「プロット エリアの右上」の単位があれば、その手前 8px まで
        const leftColumnBadge = (isTopBadge && !isRightAxis) || (Boolean(axis2BadgeText) && isRightAxis);
        const drillX = leftColumnBadge ? marginLeft + 4 : marginLeft;
        const topRightBadgeRoom = badgeText && viewModel.unitInfo.unitPosition === "plotTopRight" ? badgeWidthOf(badgeText, viewModel.unitInfo.fontSize) + 8 : 0;
        const drillMaxWidth = width - marginRight - topRightBadgeRoom - drillX;
        // 低い図で行をとると描く範囲が潰れ、軸がビジュアルの外へ押し出される。足す高さが図の 12% を超えるなら出さない
        const drillPathOn = Boolean(fittedDrillPath(drillMaxWidth)) && marginTopWithDrill - marginTopBase <= height * 0.12;
        const marginTop = drillPathOn ? marginTopWithDrill : marginTopBase;

        // X軸カテゴリラベルの幅（実測）と回転の判定
        const catFontSizePx = catAxis.fontSize * PT_TO_PX;
        const catFont: FontSpec = { family: catAxis.fontFamily, size: catFontSizePx, bold: catAxis.bold, italic: catAxis.italic, underline: catAxis.underline };
        // 階層を段に重ねるときは、棒のすぐ下にいちばん下のレベルだけを出し、上のレベルはその下の段に出す
        const stackedLevels = catAxis.show && catAxis.levelCount > 1 && !catAxis.concatenateLabels;
        const labelOf = (g: CategoryGroup) => (stackedLevels ? g.levels[g.levels.length - 1] : g.label);

        // カテゴリ最小幅と横スクロール判定
        const viewWidth = Math.max(10, width - marginLeft - marginRight);
        const groups = viewModel.categoryGroups;
        const count = groups.length;
        const minCatWidth = Math.max(0, catAxis.minCategoryWidth || 0);
        // minCategoryWidth が設定されていて、count * minCategoryWidth が viewWidth を上回る場合に横スクロール発動
        const neededWidth = minCatWidth > 0 ? Math.ceil(count * minCatWidth) : 0;
        // X 軸の種類が「連続」なら、値の間隔どおりに置く（スクロールしない。名前の代わりに目盛りを出す）
        const continuous = viewModel.continuousAxis;
        const scrolls = !continuous && minCatWidth > 0 && neededWidth > viewWidth + 1;
        const SCROLLBAR_HEIGHT = scrolls ? 12 : 0;

        // X軸タイトル高さ (タイトル領域は高さ最大値の判定外で独立確保し、重なりを防止)
        const hasCatTitle = catAxis.titleShow && Boolean(catAxis.titleText);
        const catTitleFontSizePx = catAxis.titleFontSize * PT_TO_PX;
        const catTitleHeight = hasCatTitle ? catTitleFontSizePx + 10 : 0;
        // パレートのランクの帯。カテゴリのラベル（と階層の段）の下に 1 段取る
        const pareto = viewModel.pareto;
        const rankBandOn = pareto.enabled && pareto.showRankBand && catAxis.show;
        const rankBandHeight = rankBandOn ? catFontSizePx + 12 : 0;
        // 名前の欄（名前と階層の段）に使える高さ。上限を大きくしても（標準は 50% まで、自作は 100% まで入る）、描く範囲を
        // 押しつぶして名前がビジュアルの下へはみ出さないよう、描く範囲に高さの 4 分の 1 は残す（2026-09-24。1.29 までは 100% で棒がつぶれ、名前の頭が切れた）
        const labelRoom = Math.max(0, height * 0.75 - marginTop - catTitleHeight - rankBandHeight - SCROLLBAR_HEIGHT - 6);
        // 名前の縦の広がりの上限 (高さの最大値 % はここだけに効く)
        // maxHeight は 0〜100% (既定 25%)。標準と同じく、凡例を含むビジュアル全体の高さに対する割合。
        // 斜めのラベルでは、割合はラベルの縦の広がりに効き、棒との間隔と文字の太さの半分は別に足す
        // （Desktop で、タイトルありの標準は 8 文字、タイトルなしは 9 文字まで出した。1.13 までは 1 文字早かった）
        const labelExtentMax = Math.max(16, viewport.height * (catAxis.maxHeight / 100));

        const plotWidth = scrolls ? neededWidth : viewWidth;
        const padRatio = Math.max(0, Math.min(0.5, columnsSettings.categorySpacing / 100));
        // 外側のパディング: 最初の棒の前と最後の棒の後ろの余白。カテゴリ 1 つ分の幅
        // (step = 棒 + カテゴリ間のスペース) に対する比。「自動」はカテゴリ間のスペースの半分で、
        // 各カテゴリの枠の中央に棒を置く 1.3.1.0 までの配置と同じになる
        const outerRatio = columnsSettings.outerPadding === null
            ? padRatio / 2
            : Math.max(0, Math.min(1, columnsSettings.outerPadding / 100));
        const slots = (count || 1) - padRatio + 2 * outerRatio;
        // 斜めの名前は棒の中心から左下へ伸びるので、左の端で「…」に省かれないよう、はみ出す分だけ棒の並びの前を空ける
        // （スクロールするとき・縦に立てるときは空けない）
        const leadIn = (() => {
            if (!catAxis.show || stackedLevels || scrolls || !count || continuous) return 0;
            const step0 = plotWidth / slots;
            const widths = groups.map((g) => measureTextWidth(labelOf(g), catFont));
            if (!(Math.max(...widths) > step0 * 0.92) || step0 < catFontSizePx * 1.45) return 0;
            // 名前の長さの上限（下の maxAllowedLabelLen と同じ見積もり。名前の欄に使える高さでも頭を打つ）
            const maxLen = Math.min(labelExtentMax, Math.max(0, labelRoom - ROTATED_LABEL_GAP - catFontSizePx * 0.5)) / Math.SQRT1_2;
            // renderCategoryLabel の左端の判定（(cx − 左の余白) / sin45 までの長さなら省かない）を満たす余白（1px の余裕）
            const need = rotatedLabelLeadIn(
                widths.map((w, i) => ({ extent: Math.min(w, maxLen) * Math.SQRT1_2, ratio: (outerRatio + (1 - padRatio) / 2 + i) / slots })),
                marginLeft,
                plotWidth,
                CATEGORY_LABEL_LEFT_SAFE_MARGIN + 1
            );
            // 余白を足しても、帯がカテゴリの最小幅と「斜めのまま」の幅（文字の 1.45 倍）を下回らない所で止める
            // （下回ると最小幅を割る・縦に立つのに余白だけ残る）。極端に長い名前は、ビジュアルの幅の 30% までで省く
            const keepStep = Math.max(minCatWidth, catFontSizePx * 1.45);
            return Math.max(0, Math.min(need, plotWidth * 0.3, plotWidth - keepStep * slots));
        })();
        // 連続：いちばん狭い値の間隔を棒 1 本ぶんの帯にし、両端は帯の半分ずつ空ける
        const continuousLayout = continuous ? continuousBand(continuous, plotWidth) : null;
        const step = continuousLayout ? continuousLayout.band : (plotWidth - leadIn) / slots;
        /** i 番目のカテゴリの中心 (プロット左端からの x) */
        const centerOf = (i: number) =>
            continuous && continuousLayout
                ? continuousLayout.edge + positionOf(continuous, continuous.values[i]) * (plotWidth - 2 * continuousLayout.edge)
                : leadIn + step * (outerRatio + (1 - padRatio) / 2 + i);
        /**
         * 連続の目盛り（プロット左端からの x・文字の中心 cx と文字）。端の文字はプロットの外へはみ出さないよう内側へ寄せる（線の位置 x は変えない）。
         * 寄せたあとの文字が隣と 16px 空く本数まで減らす。2 本でも入らなければ、標準と同じく末尾を「…」で省く
         * （2026-10-12 ユーザー「細くしたときに二つ目の日付が重なる」。前は端へ寄せた文字が重なっていた）
         */
        const continuousTickXs = (() => {
            if (!continuous || !continuousLayout) return [];
            const usable = plotWidth - 2 * continuousLayout.edge;
            const at = (ratio: number) => continuousLayout.edge + ratio * usable;
            const ticksOf = (max: number) => continuousTicks(continuous, max).filter((t) => t.ratio >= -1e-9 && t.ratio <= 1 + 1e-9);
            const laid = (ticks: ReturnType<typeof ticksOf>, labelOf: (label: string) => string) =>
                ticks.map((t) => {
                    const label = labelOf(t.label);
                    const width = measureTextWidth(label, catFont);
                    const x = at(t.ratio);
                    return { x, cx: Math.max(width / 2, Math.min(plotWidth - width / 2, x)), label, width };
                });
            const fits = (items: ReturnType<typeof laid>, gap: number) => items.every((t, k) => k === 0 || t.cx - t.width / 2 >= items[k - 1].cx + items[k - 1].width / 2 + gap);
            for (let max = Math.max(2, Math.floor(plotWidth / 60)); max >= 2; max--) {
                const items = laid(ticksOf(max), (label) => label);
                if (fits(items, 16)) return items;
            }
            const ticks = ticksOf(2);
            for (let room = Math.max(...ticks.map((t) => measureTextWidth(t.label, catFont))); room > 0; room -= 2) {
                const items = laid(ticks, (label) => truncateToWidth(label, room, catFont));
                if (fits(items, 4)) return items;
            }
            return [];
        })();
        // カテゴリ 1 つぶんの帯に、系列の棒を並べる（系列 1 本なら帯いっぱいの 1 本）。
        // 積み上げは系列を 1 本の棒に積むので、帯いっぱいの 1 本ぶんの幅
        const stacked = viewModel.chartType !== "clustered";
        const cluster = clusterLayout(
            step * (1 - padRatio),
            stacked ? 1 : viewModel.series.length,
            columnsSettings.seriesSpacing,
            columnsSettings.maxBarWidth,
            columnsSettings.overlap
        );
        const barWidth = cluster.barWidth;
        const drawOrder = (points: DataPoint[]) => drawOrderOf(points, columnsSettings.overlapReverse, viewModel.seriesPositions);
        const clusterSpan = spanOf(cluster);
        // 比較レイヤー：系列の棒 1 本ぶんの幅に、レイヤーの棒を少しずつずらして重ねる。合計ラベルは手前のレイヤーの棒の上に出す
        const layerDraws = layerDrawsFor(barWidth);
        const frontTotals = !layered || shownLayers[0] === 0;
        const frontShift = layered ? layerDraws.get(shownLayers[0])!.offset : 0;
        // 積み上げのデータラベルは棒の中に置く。外側の指定（自動・外側の上）は中央に読み替える。系列ごとに変えた位置があればそれ
        const labelPositionOf = (d: DataPoint) => {
            const position = (d.labelSettings ?? viewModel.dataLabels).position;
            return stacked && (position === "auto" || position === "outsideEnd") ? "insideCenter" : position;
        };

        const categoryLabelStyle: React.CSSProperties = {
            ...fontStyleOf(fontOf(catAxis)),
            fill: catAxis.labelColor,
        };
        const boxedLevels = stackedLevels && catAxis.hierarchyStyle === "boxed";
        // 標準の段の枠は薄い点線。色はラベルの色に合わせる（ハイコントラストでもラベルと同じ色になる）
        const levelSeparatorStyle: React.CSSProperties = {
            stroke: catAxis.labelColor,
            strokeOpacity: 0.35,
            strokeWidth: 1,
            strokeDasharray: "1 2",
            // グリッド線と同じく crispEdges を掛けない（拡大率が整数でない画面で太さがそろわない）
            shapeRendering: "auto",
        };
        // 帯。ラベルの色を淡く敷くだけで、枠線は引かない（区切りは帯どうしのすき間で見せる）
        const levelBoxStyle: React.CSSProperties = {
            fill: catAxis.labelColor,
            fillOpacity: LEVEL_BAND_OPACITY,
        };
        const widestCat = !catAxis.show ? 0 : continuous ? 0 : Math.max(0, ...groups.map((g) => measureTextWidth(labelOf(g), catFont)));

        // バンド幅に収まらなければ斜め -45 度に回転
        const shouldRotateCat = catAxis.show && widestCat > step * 0.92;
        // 斜めの名前どうしの間（帯の幅 × sin45）が文字の高さより狭いと重なる（帯の幅が文字の 1.45 倍を切るとき）。
        // そのときは標準と同じく、斜めではなく縦に立てて全部の名前を出す（2026-09-24 ユーザー決定。1.29 までは斜めのまま重なった）
        const crowdedCat = shouldRotateCat && step < catFontSizePx * 1.45;
        // 階層を段に重ねるときも縦に立てる（標準と同じ。斜めだと左隣の区切りの線をまたぐ）
        const uprightCat = shouldRotateCat && (stackedLevels || crowdedCat);
        // 縦に立てても帯の幅が文字の高さより狭いと重なるので、そのときだけ間引いて出す（1 つおき、…）
        const categoryLabelEvery = uprightCat ? Math.max(1, Math.ceil((catFontSizePx * 1.1) / step)) : 1;

        // X軸ラベル領域
        const maxLabelAreaHeight = shouldRotateCat ? ROTATED_LABEL_GAP + labelExtentMax + catFontSizePx * 0.5 : labelExtentMax;

        // 間引くときは、出す名前だけで欄の高さを決める（出さない長い名前のために棒を縮めない）
        const widestShownCat =
            categoryLabelEvery > 1
                ? Math.max(0, ...groups.filter((_, i) => i % categoryLabelEvery === 0).map((g) => measureTextWidth(labelOf(g), catFont)))
                : widestCat;
        // 必要とされるラベル高さ（斜めは、棒との間隔＋ラベルの縦の広がり＋文字の太さの半分。縦はラベルの長さそのまま）
        const desiredLabelHeight = uprightCat
            ? ROTATED_LABEL_GAP + widestShownCat + catFontSizePx * 0.5
            : shouldRotateCat
                ? ROTATED_LABEL_GAP + widestShownCat * Math.SQRT1_2 + catFontSizePx * 0.5
                : catFontSizePx + 10;

        // 最低の高さ（14）も、名前の欄に使える高さに収まるときだけとる（低いビジュアルで描く範囲の 4 分の 1 を割らない）
        const labelAreaHeight = catAxis.show
            ? Math.max(Math.min(14, labelRoom), Math.min(desiredLabelHeight, maxLabelAreaHeight, labelRoom))
            : 0;

        // ラベルの長さの上限 (px)。超えたら末尾を「…」に省略。斜めは、ラベル領域の高さから棒との間隔を除いた分
        // 領域がいちばん長いラベルちょうどのときに、計算の誤差で省略しないよう 0.5px の余裕を持たせる。
        // 斜め・縦は欄の高さで決まるので下限を置かない（欄より長い名前を出すと、下のランクの帯やビジュアルの外にはみ出す）
        const maxAllowedLabelLen = uprightCat
            ? Math.max(0, labelAreaHeight - ROTATED_LABEL_GAP - catFontSizePx * 0.5 + 0.5)
            : shouldRotateCat
                ? Math.max(0, (labelAreaHeight - ROTATED_LABEL_GAP - catFontSizePx * 0.5) / Math.SQRT1_2 + 0.5)
                : Math.max(12, step * 0.92);

        // 上のレベルの段。多くてもプロットを潰さないよう、ラベルと合わせてグラフの高さの半分までにする（上のレベルから落とす）
        // 帯は上下に余白が要るので、段を少し高くする。「段の高さ」で余白を詰められる
        const levelPadding = LEVEL_ROW_PADDINGS[catAxis.levelRowHeight] ?? LEVEL_ROW_PADDINGS.normal;
        const levelRowHeight = catFontSizePx + (boxedLevels ? levelPadding.boxed : levelPadding.lines);
        const levelRows = stackedLevels
            ? Math.min(catAxis.levelCount - 1, Math.max(0, Math.floor((Math.min(height * 0.5, labelRoom) - labelAreaHeight) / levelRowHeight)))
            : 0;
        const levelAreaHeight = levelRows * levelRowHeight;

        // 全体の下部マージン = ラベル領域 + 上のレベルの段 + ランクの帯 + タイトル領域 + スクロールバー高さ + 余白
        const marginBottom = Math.max(20, labelAreaHeight + levelAreaHeight + rankBandHeight + catTitleHeight + SCROLLBAR_HEIGHT + 6);
        const plotHeight = Math.max(10, height - marginTop - marginBottom);
        plotHeightRef.current = plotHeight;
        // 目盛りの本数は、標準と同じく描く範囲の高さで決める（150px 未満は 3、300px 未満は 5、それ以上は 8 まで）
        // 「目盛りの本数 (目安)」があればその本数、空なら描く範囲の高さで決める（第 2 Y 軸も同じ上限）。
        // どちらも、数字の行が重ならない本数までにとどめる
        const tickTarget = valAxis.tickCount || recommendedTickCount(plotHeight, "vertical");
        const valueTicks = fitTicks(viewModel.ticksFor, tickTarget, plotHeight, () => valTickFontSizePx * 1.4);
        const axis2Ticks = fitTicks(axis2.ticksFor, tickTarget, plotHeight, () => axis2TickFontPx * 1.4);

        // 横グリッド線 (Y軸目盛線) の線種と透過性
        const hStroke = gridLineStroke(gridlines.horizontalStyle, gridlines.horizontalWidth, gridlines.horizontalScaleWithWidth, { dashArray: gridlines.horizontalDashArray, dashCap: gridlines.horizontalDashCap });
        const hOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));

        // 縦グリッド線 (X軸目盛線) の線種と透過性
        const vStroke = gridLineStroke(gridlines.verticalStyle, gridlines.verticalWidth, gridlines.verticalScaleWithWidth, { dashArray: gridlines.verticalDashArray, dashCap: gridlines.verticalDashCap });
        const vOpacity = Math.max(0, Math.min(1, 1 - gridlines.verticalTransparency / 100));

        // 縦グリッド線の位置: プロットの両端と、隣り合うカテゴリの中間（「階層の区切り」なら上のレベルの区切りだけ）
        const verticalGridXs = continuous ? continuousTickXs.map((t) => t.x) : [
            0,
            ...categoryGridBreaks(groups.map((g) => g.levels), catAxis.levelCount, gridlines.verticalPosition === "levels", groups.map((g) => g.levelKeys)).map(
                (i) => centerOf(i) - step / 2
            ),
            plotWidth,
        ];

        /** 軸の比率 (0〜1) の位置の Y 座標（範囲の反転を含む） */
        const yOfRatio = (ratio: number): number => {
            const clamped = Math.max(0, Math.min(1, ratio));
            const effective = valAxis.invertRange ? 1 - clamped : clamped;
            return marginTop + plotHeight * (1 - effective);
        };

        /**
         * 比率で表した 2 点のあいだの縦の範囲。棒は始まり（集合は 0、積み上げは下に積んだ棒の端）から終わりまで。
         * 0 の線は引かない（標準と同じく 0 の線は横のグリッド線の 1 本として出る）
         */
        const extentBetween = (fromRatio: number, toRatio: number): { top: number; height: number } => {
            const y0 = yOfRatio(fromRatio);
            const y1 = yOfRatio(toRatio);
            return { top: Math.min(y0, y1), height: Math.abs(y1 - y0) };
        };
        // 積み上げの系列間のスペース（px）。系列の展開のときは viewModel が置き直しているので使わない
        const stackInset = stacked && !columnsSettings.stackedExplode ? columnsSettings.stackedSpacing : 0;
        /** 棒の範囲。積み上げで下に別の棒があれば、始まりの側を系列間のスペースだけ縮める */
        const segmentExtent = (d: DataPoint, toRatio: number): { top: number; height: number } => {
            const e = extentBetween(d.startRatio, toRatio);
            if (!d.insetStart || stackInset <= 0) return e;
            const r = insetSpan(e.top, e.height, Math.abs(yOfRatio(d.startRatio) - e.top) < 1e-6, stackInset);
            return { top: r.lo, height: r.size };
        };
        // 内側の罫線を非表示は、積んだ棒が触れ合っているときだけ（すき間を空けたら、どの棒の罫線も外側の罫線）
        const outlineOnly = stacked && columnsSettings.borderOutlineOnly && columnsSettings.stackedSpacing <= 0;

        /**
         * 「内側の罫線を非表示」：積んだ棒の正・負それぞれの外側だけに罫線を引く。罫線の設定は「すべて」の値
         * （棒の色を一致させるなら、外側の端の棒の色）
         */
        const renderStackOutline = (g: CategoryGroup, cx: number) => {
            if (!columnsSettings.showBorder) return null;
            const left = cx - barWidth / 2;
            return [true, false].map((positive) => {
                const segs = g.points.filter((d) => !d.blank && (positive ? d.value > 0 : d.value < 0));
                if (!segs.length) return null;
                const extents = segs.map((d) => segmentExtent(d, d.valRatio));
                const top = Math.min(...extents.map((e) => e.top));
                const bottom = Math.max(...extents.map((e) => e.top + e.height));
                const outer = segs.find((d) => d.outermost) ?? segs[segs.length - 1];
                const r = Math.max(0, Math.min(columnsSettings.cornerRadius, barWidth / 2, bottom - top));
                const roundAtTop = valAxis.invertRange ? !positive : positive;
                return (
                    <path
                        key={`outline-${positive ? "pos" : "neg"}`}
                        d={barPathOf(left, top, barWidth, bottom - top, r, roundAtTop)}
                        className="stack-outline"
                        fill="none"
                        stroke={columnsSettings.borderMatchColumn ? outer.color : columnsSettings.borderFill}
                        strokeWidth={columnsSettings.borderWidth}
                        strokeOpacity={Math.max(0, Math.min(1, 1 - columnsSettings.borderTransparency / 100))}
                        pointerEvents="none"
                    />
                );
            });
        };

        /** 棒 1 本（系列 1 本ぶん）。cx は棒の中心の x */
        /** 描いたデータラベル・合計ラベルの枠（プロットの座標）。ドリルの位置を重ねないために集める */
        const labelBoxes: Array<{ x0: number; x1: number; y0: number; y1: number }> = [];
        // データラベルは、すべての棒を描いたあとに重ねる（棒ごとに描くと、外に出したラベルが隣の棒に隠れる。標準も棒の上）
        const labelLayer: React.ReactNode[] = [];
        const renderBar = (d: DataPoint, cx: number, key: string, layer?: LayerDraw) => {
            if (d.blank) return null;
            // 比較レイヤーの棒は、レイヤーの幅で描く
            const bw = layer?.width ?? barWidth;
            const barLeft = cx - bw / 2;

            const { top: barTop, height: barH } = segmentExtent(d, d.valRatio);

            const isHovered = hoveredKey === key;

            const isPositive = d.value >= 0;
            // 積み上げでは、角丸は外側の端の棒だけ
            const r = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, bw / 2, barH)) : 0;

            // 範囲反転時の角丸向き（通常: 正なら上、負なら下。反転時: 正なら下、負なら上）
            const roundAtTop = valAxis.invertRange ? !isPositive : isPositive;
            const barPath = barPathOf(barLeft, barTop, bw, barH, r, roundAtTop);

            // ハイライト（他のビジュアルでの選択）: 標準と同じく棒全体を薄く描き、該当分の高さを
            // 通常の濃さで重ねる。全部が該当する棒はそのまま、該当しない (null) 棒は薄いだけ
            const fullyHighlighted =
                d.highlight !== null && Math.abs(d.highlight - d.value) <= Math.abs(d.value) * 1e-9;
            const highlightDimmed = viewModel.hasHighlights && !fullyHighlighted;
            // 自分の棒（または凡例）をクリックして選んだとき: 標準と同じく、選んでいない棒を薄くする
            const selectionDimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isPicked(d);
            const dimmed = highlightDimmed || selectionDimmed;
            let highlightPath: string | null = null;
            if (highlightDimmed && d.highlight !== null && d.highlightRatio !== null) {
                const hl = segmentExtent(d, d.highlightRatio);
                if (hl.height > 0) {
                    const hlPositive = d.highlight >= 0;
                    const hlR = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, bw / 2, hl.height)) : 0;
                    highlightPath = barPathOf(
                        barLeft, hl.top, bw, hl.height, hlR,
                        valAxis.invertRange ? !hlPositive : hlPositive
                    );
                }
            }

            // 透明度は塗りだけに効かせる（opacity だと境界線まで消える）。
            // 罫線は罫線の透過性だけで決まる。ハイライト時の減光は両方に掛ける
            const dimFactor = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
            const baseFillOpacity = Math.max(0, Math.min(1, 1 - d.transparency / 100)) * (layer?.opacity ?? 1);
            const fillOpacity = dimFactor * baseFillOpacity;
            const ownBorder = d.borderShow && !outlineOnly;
            // 比較レイヤーの輪郭線は、棒の色で濃く引く（薄い棒どうしの境目を見せる）
            const stroke = layer?.outline ? (viewModel.highContrast?.foreground ?? d.color) : ownBorder ? d.borderColor : "none";
            const strokeWidth = layer?.outline ? 1 : ownBorder ? d.borderWidth : 0;
            const strokeOpacity = dimFactor * (d.borderShow ? Math.max(0, Math.min(1, 1 - d.borderTransparency / 100)) : 1);
            const seriesName = viewModel.seriesMode ? viewModel.series[d.seriesIndex]?.name : undefined;
            // ハイコントラストでは、選んだ棒を選択の色で太く囲む（薄くするだけでは見分けにくい）
            const hcPicked = !!viewModel.highContrast && selectedIds.length > 0 && isPicked(d);

            return (
                <g
                    key={key}
                    className={`bar-item ${isHovered ? "hovered" : ""}`}
                    onClick={(e) => {
                        e.stopPropagation();
                        if (d.selectionIds) onSelectMany?.(d.selectionIds, e.ctrlKey || e.metaKey);
                        else onSelect(d.selectionId, e.ctrlKey || e.metaKey);
                    }}
                    onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onContextMenu(d.selectionIds ? (null as unknown as ISelectionId) : d.selectionId, e.clientX, e.clientY);
                    }}
                    onMouseEnter={() => setHoveredKey(key)}
                    onMouseLeave={() => setHoveredKey(null)}
                    {...keysOf({
                        label: seriesName ? `${d.category} ${seriesName}: ${d.formattedValue}` : `${d.category}: ${d.formattedValue}`,
                        onSelect: (multi) => (d.selectionIds ? onSelectMany?.(d.selectionIds, multi) : onSelect(d.selectionId, multi)),
                        onClear: onClearSelection,
                        pressed: selectedIds.length > 0 && isPicked(d),
                    })}
                >
                    <path
                        d={barPath}
                        className="bar-rect"
                        style={{
                            fill: d.color,
                            fillOpacity,
                            stroke: hcPicked ? viewModel.highContrast!.selected : stroke,
                            strokeWidth: hcPicked ? HC_SELECTED_WIDTH : strokeWidth,
                            strokeOpacity,
                            strokeDasharray: hcPicked ? undefined : d.borderDash || undefined,
                        }}
                        onMouseEnter={(e) => onTooltipShow?.(d, e.clientX, e.clientY)}
                        onMouseMove={(e) => onTooltipMove?.(d, e.clientX, e.clientY)}
                        onMouseLeave={() => onTooltipHide?.()}
                    />

                    {/* ハイライトの該当分。マウスは下の棒で受ける */}
                    {highlightPath && (
                        <path
                            d={highlightPath}
                            className="bar-highlight"
                            pointerEvents="none"
                            style={{
                                // ハイコントラストでは、背景色で塗った棒でも該当分が見えるよう選択の色で塗る
                                fill: viewModel.highContrast?.selected ?? d.color,
                                fillOpacity: baseFillOpacity,
                            }}
                        />
                    )}

                    {/* データラベル。描くのは棒の後ろ（labelLayer） */}
                    {void labelLayer.push((() => {
                        // 系列ごとに変えた設定があればそれ（保存の無い項目は「すべて」の値）
                        const dl = d.labelSettings ?? viewModel.dataLabels;
                        if (!dl.show || !d.labelShow || barH <= 0 || layer?.label === false) return null;

                        const lines = labelLinesOf(dl, d);
                        if (!lines.length) return null;
                        const isVertical = dl.orientation === "vertical";

                        const barBottom = barTop + barH;
                        const placed = placeLabel({
                            position: labelPositionOf(d),
                            top: barTop,
                            bottom: barBottom,
                            barWidth: bw,
                            value: lines[0].text,
                            font: lines[0].font,
                            lines,
                            layout: { singleLine: dl.singleLine, align: dl.horizontalAlignment },
                            endAtBottom: !roundAtTop,
                            vertical: isVertical,
                            // 積み上げでは外側が上の棒なので、入りきらないラベルを外へ逃がさず出さない
                            overflow: stacked ? false : dl.overflow,
                        });

                        if (!placed.fitsVertically && !placed.outside) return null;
                        if (!placed.fitsAlongBar && !dl.overflow) return null;

                        // 縦向きのラベルは rotate(-90)：(x, y) → (y, −x)
                        const bx = placed.box;
                        const ownBox = isVertical
                            ? { x0: cx + bx.y, x1: cx + bx.y + bx.height, y0: placed.y - bx.x - bx.width, y1: placed.y - bx.x }
                            : { x0: cx + bx.x, x1: cx + bx.x + bx.width, y0: placed.y + bx.y, y1: placed.y + bx.y + bx.height };
                        // 先に描いたラベルと重なるラベルは出さない（標準と同じ。棒の幅より長いラベルを外に出すと隣と重なる）
                        if (labelBoxes.some((b) => b.x0 < ownBox.x1 - 1 && ownBox.x0 < b.x1 - 1 && b.y0 < ownBox.y1 - 1 && ownBox.y0 < b.y1 - 1)) return null;
                        labelBoxes.push(ownBox);
                        const autoColor = labelAutoColor(!placed.outside, seenColorOf(d.color, baseFillOpacity), dl, viewModel.highContrast, viewModel.background);
                        const toneHere = placed.outside || dl.backgroundShow;
                        const bgOpacity = (100 - dl.backgroundTransparency) / 100;

                        return (
                            <g
                                key={`dl-${key}`}
                                transform={`translate(${cx}, ${placed.y})${isVertical ? " rotate(-90)" : ""}`}
                                className="data-label-group"
                                pointerEvents="none"
                            >
                                {dl.backgroundShow && (
                                    <rect
                                        x={placed.box.x}
                                        y={placed.box.y}
                                        width={placed.box.width}
                                        height={placed.box.height}
                                        rx={2}
                                        fill={dl.backgroundColor}
                                        fillOpacity={bgOpacity}
                                    />
                                )}
                                {lines.map((line, k) => (
                                    <text
                                        key={line.kind}
                                        x={placed.xs[k]}
                                        y={placed.baselines[k]}
                                        className={labelClassOf(line)}
                                        textAnchor={placed.anchor}
                                        style={labelTextStyle(dl, line, d, autoColor, toneHere)}
                                    >
                                        {line.text}
                                    </text>
                                ))}
                            </g>
                        );
                    })())}
                </g>
            );
        };

        /**
         * 合計ラベル（積み上げ）。正の棒の上の端・負の棒の下の端の外側に置く。cx はカテゴリの中心。
         * 「正と負の分割」がオフなら、正と負を足した合計を 1 つ（正なら上、負なら下）
         */
        const renderTotalLabels = (g: CategoryGroup, cx: number) => {
            const tl = viewModel.totalLabels;
            const t = g.totals;
            if (!tl.show || !t) return null;
            const labels: Array<{ key: string; text: string; ratio: number; up: boolean }> = [];
            if (tl.split) {
                if (t.hasPositive) labels.push({ key: "pos", text: t.positiveText, ratio: t.positiveEndRatio, up: true });
                if (t.hasNegative) labels.push({ key: "neg", text: t.negativeText, ratio: t.negativeEndRatio, up: false });
            } else if (t.hasPositive || t.hasNegative) {
                const up = t.net >= 0 ? t.hasPositive : !t.hasNegative;
                labels.push({ key: "net", text: t.netText, ratio: up ? t.positiveEndRatio : t.negativeEndRatio, up });
            }
            const fontPx = tl.fontSize * PT_TO_PX;
            const font = { family: tl.fontFamily, size: fontPx, bold: tl.bold, italic: tl.italic, underline: tl.underline };
            return labels.map((label) => {
                // 範囲の反転では、値の大きい方が下になる
                const outwardUp = label.up !== valAxis.invertRange;
                const edgeY = yOfRatio(label.ratio);
                const baseline = outwardUp ? edgeY - 4 : edgeY + fontPx * 0.85 + 4;
                const width = measureTextWidth(label.text, font);
                const boxTop = baseline - fontPx * 0.85 - LABEL_PADDING / 2;
                labelBoxes.push({ x0: cx - width / 2 - LABEL_PADDING, x1: cx + width / 2 + LABEL_PADDING, y0: boxTop, y1: boxTop + fontPx + LABEL_PADDING });
                return (
                    <g key={`total-${label.key}`} className="total-label-group" pointerEvents="none">
                        {tl.backgroundShow && (
                            <rect
                                x={cx - width / 2 - LABEL_PADDING}
                                y={baseline - fontPx * 0.85 - LABEL_PADDING / 2}
                                width={width + LABEL_PADDING * 2}
                                height={fontPx + LABEL_PADDING}
                                rx={2}
                                fill={tl.backgroundColor}
                                fillOpacity={(100 - tl.backgroundTransparency) / 100}
                            />
                        )}
                        <text
                            x={cx}
                            y={baseline}
                            className="total-label"
                            textAnchor="middle"
                            style={{
                                fill: tl.color,
                                ...fontStyleOf(fontOf(tl)),
                            }}
                        >
                            {label.text}
                        </text>
                    </g>
                );
            });
        };

        /**
         * 階層の上のレベルの段。標準と同じく、いちばん下のレベルのラベルの下に、1 つ上のレベルから順に段を重ね、
         * 同じ親が続く区切りを点線の枠で囲む。文字は区切りの幅の中央に置き、はみ出す分は省略する
         */
        const renderLevelRows = (xOffset: number) => {
            const paths = groups.map((g) => g.levels);
            const rowsTop = marginTop + plotHeight + labelAreaHeight;
            const edgeOf = (i: number) => Math.max(0, Math.min(plotWidth, centerOf(i) - step / 2));
            const nodes: React.ReactNode[] = [];
            for (let r = 0; r < levelRows; r++) {
                // r = 0 がいちばん下のレベルのすぐ上のレベル
                const level = catAxis.levelCount - 2 - r;
                const top = rowsTop + r * levelRowHeight;
                const bottom = top + levelRowHeight;
                if (!boxedLevels) {
                    nodes.push(
                        <line key={`lv-top-${r}`} x1={xOffset} y1={top} x2={xOffset + plotWidth} y2={top} className="level-separator" style={levelSeparatorStyle} />
                    );
                }
                levelRunsOf(paths, level, groups.map((g) => g.levelKeys)).forEach((run, k) => {
                    const left = run.start === 0 ? 0 : edgeOf(run.start);
                    const right = run.end === groups.length - 1 ? plotWidth : edgeOf(run.end + 1);
                    // 区切り（帯・線の間）を押すか Enter・Space で、その区切りのカテゴリをまとめて選ぶ（Ctrl で足す。2026-09-24）
                    nodes.push(
                        renderLevelHit(run, level, { x: xOffset + left, y: top, width: right - left, height: levelRowHeight }, `lv-hit-${r}-${k}`)
                    );
                    if (boxedLevels) {
                        // 帯：区切りごとに角の丸い淡い面。隣の帯とは少し離す
                        nodes.push(
                            <rect
                                key={`lv-box-${r}-${k}`}
                                x={xOffset + left + LEVEL_BOX_GAP}
                                y={top + LEVEL_BOX_GAP}
                                width={Math.max(0, right - left - LEVEL_BOX_GAP * 2)}
                                height={Math.max(0, levelRowHeight - LEVEL_BOX_GAP * 2)}
                                rx={LEVEL_BOX_RADIUS}
                                className="level-box"
                                style={{ ...levelBoxStyle, fillOpacity: levelPicked(run) ? LEVEL_BAND_PICKED_OPACITY : LEVEL_BAND_OPACITY }}
                                pointerEvents="none"
                            />
                        );
                    }
                    // 区切りの線は棒の下から、この段の下まで
                    if (!boxedLevels) [left, right].forEach((x, e) =>
                        nodes.push(
                            <line
                                key={`lv-edge-${r}-${k}-${e}`}
                                x1={xOffset + x}
                                y1={marginTop + plotHeight}
                                x2={xOffset + x}
                                y2={bottom}
                                className="level-separator"
                                style={levelSeparatorStyle}
                            />
                        )
                    );
                    // 省略しても収まらない区切りには文字を出さない（省略は最低 1 文字＋…を残すので、隣にはみ出して重なる）
                    const available = right - left - 4 - (boxedLevels ? LEVEL_BOX_GAP * 2 : 0);
                    const shown = truncateText(run.text, available, catFont);
                    if (measureTextWidth(shown, catFont) > available) return;
                    nodes.push(
                        <text
                            key={`lv-text-${r}-${k}`}
                            x={xOffset + (left + right) / 2}
                            y={top + levelRowHeight / 2 + catFontSizePx * 0.35}
                            className="x-category-level"
                            textAnchor="middle"
                            style={categoryLabelStyle}
                            pointerEvents="none"
                        >
                            {shown}
                            <title>{run.text}</title>
                        </text>
                    );
                });
            }
            return <g className="category-levels">{nodes}</g>;
        };

        /**
         * パレートの境目。累積比の軸（右、0〜100%）の境目の高さに破線を引く。
         * 線の色は累積比の線に合わせ、棒と重なっても読めるよう薄くする
         */
        const renderThresholds = (xOffset: number) => {
            const ratioLine = viewModel.lines.find((l) => l.selectable === false);
            const color = ratioLine?.color ?? catAxis.labelColor;
            return (
                <g className="pareto-thresholds" pointerEvents="none">
                    {pareto.thresholds.map((t, k) => {
                        const y = marginTop + plotHeight * (1 - Math.max(0, Math.min(1, t)));
                        return (
                            <line
                                key={`th-${k}`}
                                x1={xOffset}
                                y1={y}
                                x2={xOffset + plotWidth}
                                y2={y}
                                className="pareto-threshold"
                                style={{ stroke: color, strokeOpacity: 0.6, strokeWidth: 1, strokeDasharray: "4 3" }}
                            />
                        );
                    })}
                </g>
            );
        };

        /**
         * パレートのランクの帯。同じランクが続くところを角の丸い帯にし、
         * 名前と件数を書く。帯を押すと、そのランクの棒をまとめて選ぶ（Ctrl で足す）
         */
        const renderRankBand = (xOffset: number) => {
            const top = marginTop + plotHeight + labelAreaHeight + levelAreaHeight + 2;
            const edgeOf = (i: number) => Math.max(0, Math.min(plotWidth, centerOf(i) - step / 2));
            const last = pareto.ranks.length;
            return (
                <g className="pareto-ranks">
                    {rankRuns().map((run) => {
                        // 最初の帯は、棒の並びの前の余白（leadIn）にはかけない
                        const left = run.start === 0 ? leadIn : edgeOf(run.start);
                        const right = run.end === last ? plotWidth : edgeOf(run.end);
                        return renderRankRun(run, { x: xOffset + left + 2, y: top, width: right - left - 4, height: rankBandHeight - 4 }, catFont, categoryLabelStyle);
                    })}
                </g>
            );
        };

        /** X軸カテゴリラベル (長さに応じて ... 省略し、左端はみ出し時も ... で省略)。cx はカテゴリの中心 */
        const renderCategoryLabel = (category: string, cx: number) => {
            // 斜めラベルは左下へ伸びるので、左端までの距離も許容長に含める
            // （スクロール時の左端はスクロール領域の端 = x 0 でクリップされる）
            const labelLeftBound = scrolls ? 0 : CATEGORY_LABEL_LEFT_SAFE_MARGIN;
            const maxLenLeft = shouldRotateCat && !uprightCat
                ? Math.max(12, (cx - labelLeftBound) / Math.sin(Math.PI / 4))
                : maxAllowedLabelLen;
            const effectiveMaxLen = Math.min(maxAllowedLabelLen, maxLenLeft);
            const displayCat = truncateText(category, effectiveMaxLen, catFont);
            // 斜め・縦の名前は、省いても 1 文字と「…」は残る（truncateText）。省いた名前が欄に入らなければ、その名前だけ出さない
            // （欄の外へはみ出さない。名前ごとに決めるので、欄に入る短い名前は出す）
            if (shouldRotateCat && measureTextWidth(displayCat, catFont) > maxAllowedLabelLen + 0.5) return null;
            const labelStyle = categoryLabelStyle;

            if (uprightCat) {
                // 縦に立てる。棒の中心の真下から下へ読み上げる向き（区切りの枠の中に収まる）
                const labelTop = marginTop + plotHeight + ROTATED_LABEL_GAP;
                return (
                    <text
                        x={cx}
                        y={labelTop}
                        transform={`rotate(-90, ${cx}, ${labelTop})`}
                        className="x-category-label"
                        textAnchor="end"
                        dominantBaseline="central"
                        style={labelStyle}
                    >
                        {displayCat}
                        <title>{category}</title>
                    </text>
                );
            }
            if (shouldRotateCat) {
                // 斜め45度回転 (標準準拠: 棒の直下から左下に伸び、Y軸ラベルの下に被さる)。
                // 文字の高さの中ほど（ベースラインから 0.35 文字）を通る線が棒の中心で終わるように、ベースラインを右下へずらす
                // （1.29 まではベースラインの端を中心より 4px 左に置き、文字の本体が左上に寄って見えた）
                const lift = catFontSizePx * 0.35 * Math.SQRT1_2;
                const labelX = cx + lift;
                const labelTop = marginTop + plotHeight + ROTATED_LABEL_GAP + lift;
                return (
                    <text
                        x={labelX}
                        y={labelTop}
                        transform={`rotate(-45, ${labelX}, ${labelTop})`}
                        className="x-category-label"
                        textAnchor="end"
                        style={labelStyle}
                    >
                        {displayCat}
                        <title>{category}</title>
                    </text>
                );
            }
            // 水平中央揃え
            return (
                <text
                    x={cx}
                    y={marginTop + plotHeight + catFontSizePx + 4}
                    className="x-category-label"
                    textAnchor="middle"
                    style={labelStyle}
                >
                    {displayCat}
                    <title>{category}</title>
                </text>
            );
        };

        /** 折れ線。点はカテゴリの中心。第 2 Y 軸には範囲の反転が無いので、そのまま下から上へ */
        const renderLine = (line: LineSeriesInfo, j: number, xOffset: number, part: "area" | "line") =>
            renderLineSeries(
                line,
                j,
                part,
                (i) => xOffset + centerOf(i),
                // 第 2 軸で描く線（パレートでは累積比の線だけ。「折れ線の値」の線は左の軸で、範囲の反転に従う）
                (ratio) => (onAxis2(line) ? marginTop + plotHeight * (1 - Math.max(0, Math.min(1, ratio))) : yOfRatio(ratio)),
                false,
                // ステップの線は、段が変わる位置と同じく、隣のカテゴリとのすき間の真ん中まで延ばす（プロットの外へは出さない）
                step / 2,
                [xOffset, xOffset + plotWidth],
                clusterSpan
            );

        /**
         * リボンの帯。同じ系列の棒を、隣のカテゴリの棒と S 字の帯でつなぐ（棒の後ろに描く）。
         * 罫線は標準と同じく帯の上と下の縁だけに引く
         */
        const renderRibbons = (xOffset: number) => {
            const half = barWidth / 2;
            return renderRibbonBands(
                groups,
                (i) => [xOffset + centerOf(i) + half, xOffset + centerOf(i + 1) - half],
                (d) => {
                    const e = segmentExtent(d, d.valRatio);
                    return [e.top, e.top + e.height];
                },
                false
            );
        };

        // プロット内容の描画。xOffset はプロット左端の x
        // （非スクロール時は marginLeft。スクロール時はプロット幅だけのスクロール領域に描くので 0）
        const renderPlot = (xOffset: number) => {
            const gridMask = gridMaskAttr(layerDraws);
            return (
                <>
                    {gridMask && gridCutout(gridMaskId, `${gridMaskId}-bars`)}
                    <g className="gridlines" mask={gridMask}>
                    {/* X軸グリッド線（垂直目盛線） */}
                    {gridlines.verticalShow && (
                        <g className="x-gridlines-group">
                            {verticalGridXs.map((x, i) => {
                                const lineX = xOffset + x;
                                return (
                                    <line
                                        key={`x-grid-${i}`}
                                        x1={lineX}
                                        y1={marginTop}
                                        x2={lineX}
                                        y2={marginTop + plotHeight}
                                        stroke={gridlines.verticalColor}
                                        strokeWidth={gridlines.verticalWidth}
                                        strokeDasharray={vStroke.dashArray}
                                        strokeLinecap={vStroke.lineCap}
                                        shapeRendering={vStroke.shapeRendering}
                                        strokeOpacity={vOpacity}
                                    />
                                );
                            })}
                        </g>
                    )}

                    {/* 横グリッド線 (Y軸目盛線 - プロット幅分) */}
                    {gridlines.horizontalShow && (
                        <g className="y-gridlines-group">
                            {valueTicks.map((tick, i) => {
                                const y = marginTop + plotHeight * (1 - tick.ratio);
                                return (
                                    <line
                                        key={`y-grid-${i}`}
                                        x1={xOffset}
                                        y1={y}
                                        x2={xOffset + plotWidth}
                                        y2={y}
                                        stroke={gridlines.horizontalColor}
                                        strokeWidth={gridlines.horizontalWidth}
                                        strokeDasharray={hStroke.dashArray}
                                        strokeLinecap={hStroke.lineCap}
                                        shapeRendering={hStroke.shapeRendering}
                                        strokeOpacity={hOpacity}
                                        className="grid-line"
                                    />
                                );
                            })}
                        </g>
                    )}
                    </g>

                    {/* リボンの帯。棒の後ろに描く */}
                    {viewModel.ribbons.show && groups.length > 1 && <g className="ribbons-group">{renderRibbons(xOffset)}</g>}

                    {/* X 軸の定数線の網掛け領域と、背面の線（棒の後ろ） */}
                    {renderCategoryLines((i) => xOffset + centerOf(i), step, [marginTop, marginTop + plotHeight], false, [xOffset, xOffset + plotWidth], "back")}
                    {/* Y 軸の定数線の網掛け領域と、背面の線 */}
                    {renderValueLines((ratio, secondary) => (secondary ? marginTop + plotHeight * (1 - ratio) : yOfRatio(ratio)), [xOffset, xOffset + plotWidth], true, [marginTop, marginTop + plotHeight], "back")}

                    {/* 棒とデータラベルとカテゴリラベル。カテゴリごとに、系列の棒を凡例の順に並べる */}
                    <g className="bars-group" id={gridMask ? `${gridMaskId}-bars` : undefined}>
                        {groups.map((g, i) => {
                            const cx = xOffset + centerOf(i);
                            return (
                                <g key={`cat-${i}`} className="category-group">
                                    {layered
                                        ? // 奥のレイヤーから描き、手前のレイヤーを上に重ねる
                                          [...shownLayers].reverse().flatMap((l) => {
                                              const draw = layerDraws.get(l)!;
                                              return drawOrder(layerPointsOf(g, l)).map(([d, s]) =>
                                                  renderBar(d, cx + (stacked ? 0 : cluster.offsets[viewModel.seriesPositions[s] ?? s] ?? 0) + draw.offset, `${i}-${s}-${l}`, draw)
                                              );
                                          })
                                        : drawOrder(g.points).map(([d, s]) => renderBar(d, cx + (stacked ? 0 : cluster.offsets[viewModel.seriesPositions[s] ?? s] ?? 0), `${i}-${s}`))}
                                    {outlineOnly && !layered && renderStackOutline(g, cx)}
                                    {viewModel.measureGroups
                                        ? shownLayers.map((l) => <g key={`totals-${l}`}>{renderTotalLabels(withLayerTotals(g, l), cx + layerDraws.get(l)!.offset)}</g>)
                                        : frontTotals && renderTotalLabels(g, cx + frontShift)}
                                    {catAxis.show && !continuous && i % categoryLabelEvery === 0 && withCategoryLabelHit(i, g.category, renderCategoryLabel(labelOf(g), cx))}
                                </g>
                            );
                        })}
                    </g>
                    <g className="data-labels-group">{labelLayer}</g>
                    {/* 連続の X 軸の目盛り。棒ごとの名前の代わりに、切りのいい日付・数値に出す */}
                    {catAxis.show && continuous && (
                        <g className="x-continuous-ticks">
                            {continuousTickXs.map((t, k) => (
                                <text
                                    key={`x-tick-${k}`}
                                    x={xOffset + t.cx}
                                    y={marginTop + plotHeight + catFontSizePx + 4}
                                    className="x-category-label"
                                    textAnchor="middle"
                                    style={categoryLabelStyle}
                                >
                                    {t.label}
                                </text>
                            ))}
                        </g>
                    )}

                    {levelRows > 0 && renderLevelRows(xOffset)}
                    {rankBandOn && renderRankBand(xOffset)}
                    {pareto.enabled && pareto.showThresholds && renderThresholds(xOffset)}

                    {/* 折れ線（複合）。棒の上に重ねる */}
                    {viewModel.lines.length > 0 && (
                        <g className="lines-group">
                            {/* 網掛け領域はすべての線の下に塗る */}
                            {viewModel.lines.map((line, j) => renderLine(line, j, xOffset, "area"))}
                            {viewModel.lines.map((line, j) => renderLine(line, j, xOffset, "line"))}
                            {/* 線のデータ ラベル。棒のラベルと重なるものは描かない */}
                            {viewModel.lines.map((line, j) =>
                                renderLineLabels(
                                    line,
                                    j,
                                    (i) => xOffset + centerOf(i),
                                    (ratio) => (onAxis2(line) ? marginTop + plotHeight * (1 - Math.max(0, Math.min(1, ratio))) : yOfRatio(ratio)),
                                    false,
                                    [marginTop, marginTop + plotHeight],
                                    labelBoxes
                                )
                            )}
                        </g>
                    )}
                    {/* X 軸の定数線。いちばん上に重ねる */}
                    {renderCategoryLines((i) => xOffset + centerOf(i), step, [marginTop, marginTop + plotHeight], false, [xOffset, xOffset + plotWidth], "front")}
                    {/* Y 軸の定数線 */}
                    {renderValueLines((ratio, secondary) => (secondary ? marginTop + plotHeight * (1 - ratio) : yOfRatio(ratio)), [xOffset, xOffset + plotWidth], true, [marginTop, marginTop + plotHeight], "front")}
                </>
            );
        };

        // 表示範囲内のプロット右端。横スクロール時の plotWidth はスクロール全幅なので使わず、
        // 見えている幅 viewWidth で決める（右側のY軸・右上の単位バッジが画面外に出ないように）
        const plotRightX = marginLeft + viewWidth;

        // Y軸（目盛り、線、タイトル、軸上バッジ）の共通描画
        const axisLineX = isRightAxis ? plotRightX : marginLeft;
        const tickLabelX = isRightAxis ? axisLineX + tickToAxisGap : axisLineX - tickToAxisGap;
        const tickTextAnchor = isRightAxis ? "start" : "end";

        const yTitleX = !isRightAxis
            ? (axisLineX - tickToAxisGap) - tickWidthPx - gapTitleToTicks
            : (axisLineX + tickToAxisGap) + tickWidthPx + gapTitleToTicks + valTitleFontSizePx * 0.5;
        const yTitleRotate = -90;

        // 「プロット エリアの右上」の単位バッジ。スクロールしても流れないよう固定側に描く
        const renderPlotTopRightBadge = () =>
            badgeText && viewModel.unitInfo.unitPosition === "plotTopRight" ? (
                <text
                    x={plotRightX}
                    y={marginTop - 12}
                    className="unit-axis-badge"
                    textAnchor="end"
                    style={{
                        fontSize: `${viewModel.unitInfo.fontSize}pt`,
                        fill: viewModel.unitInfo.color,
                    }}
                >
                    {badgeText}
                </text>
            ) : null;

        /**
         * ドリルの位置。描いたデータラベル・合計ラベル（labelBoxes）のうち、文字の行に掛かるものがあれば、いちばん左のものの 6px 手前で省く
         * （入らなければ出さない）。renderPlot のあとに呼ぶ。スクロールする図では、ドリルの位置は固定でラベルが流れるので、
         * スクロールの端まで動かしたときにラベルが来る位置で見る（右の棒のラベルも、スクロールすればドリルの位置の下を通る）
         */
        const renderVerticalDrillPath = () => {
            if (!drillPathOn) return null;
            const top = drillY - drillFontPx * 0.85;
            const bottom = drillY + drillFontPx * 0.25;
            const shift = scrolls ? marginLeft : 0;
            const maxScroll = scrolls ? Math.max(0, plotWidth - viewWidth) : 0;
            let right = drillX + drillMaxWidth;
            for (const b of labelBoxes) {
                if (b.y1 <= top || b.y0 >= bottom || b.x1 + shift <= drillX) continue;
                right = Math.min(right, b.x0 + shift - maxScroll - 6);
            }
            return renderDrillPath(drillX, drillY, right - drillX);
        };

        const renderYAxis = () => {
            return (
                <g className="y-axis-container">
                    {/* 軸上バッジ (valueAxisTop 配置の場合) */}
                    {badgeText && viewModel.unitInfo.unitPosition !== "plotTopRight" && (
                        <text
                            x={tickLabelX}
                            y={marginTop - 12}
                            className="unit-axis-badge"
                            textAnchor={tickTextAnchor}
                            style={{
                                fontSize: `${viewModel.unitInfo.fontSize}pt`,
                                fill: viewModel.unitInfo.color,
                            }}
                        >
                            {badgeText}
                        </text>
                    )}

                    {/* Y軸タイトル */}
                    {hasYTitle && (
                        <text
                            transform={`translate(${yTitleX}, ${marginTop + plotHeight / 2}) rotate(${yTitleRotate})`}
                            textAnchor="middle"
                            dominantBaseline={isRightAxis ? "central" : undefined}
                            className="y-axis-title"
                            style={{
                                ...fontStyleOf(fontOf(valAxis, "title")),
                                fill: valAxis.titleColor,
                            }}
                        >
                            {fittedText(valAxis.titleText, plotHeight, {
                                family: valAxis.titleFontFamily,
                                size: valTitleFontSizePx,
                                bold: valAxis.titleBold,
                                italic: valAxis.titleItalic,
                                underline: valAxis.titleUnderline,
                            })}
                        </text>
                    )}

                    {/* Y軸目盛りラベル。軸線は標準と同じく描かない */}
                    <g className="y-axis-group">
                        {valueTicks.map((tick, i) => {
                            const y = marginTop + plotHeight * (1 - tick.ratio);
                            return (
                                <g key={i} className="y-tick">
                                    {valAxis.show && (
                                        <text
                                            x={tickLabelX}
                                            y={y + (valTickFontSizePx * 0.35)}
                                            className="y-tick-label"
                                            textAnchor={tickTextAnchor}
                                            style={{
                                                ...fontStyleOf(fontOf(valAxis)),
                                                fill: valAxis.labelColor,
                                            }}
                                        >
                                            {tick.label}
                                        </text>
                                    )}
                                </g>
                            );
                        })}
                    </g>
                </g>
            );
        };

        /** 第 2 Y 軸の目盛りとタイトル。軸線は Y 軸と同じく描かない */
        const renderY2Axis = () => {
            if (!axis2On) return null;
            const onRight = !isRightAxis;
            const lineX = onRight ? plotRightX : marginLeft;
            const labelX = onRight ? lineX + tickToAxisGap : lineX - tickToAxisGap;
            const titleX = onRight
                ? labelX + axis2TickWidth + gapTitleToTicks + axis2TitleFontPx * 0.5
                : labelX - axis2TickWidth - gapTitleToTicks - axis2TitleFontPx * 0.5;
            return (
                <g className="y2-axis-container">
                    {hasAxis2Title && (
                        <text
                            transform={`translate(${titleX}, ${marginTop + plotHeight / 2}) rotate(-90)`}
                            textAnchor="middle"
                            dominantBaseline="central"
                            className="y2-axis-title"
                            style={{
                                ...fontStyleOf(fontOf(axis2, "title")),
                                fill: axis2.titleColor,
                            }}
                        >
                            {fittedText(axis2.titleText, plotHeight, {
                                family: axis2.titleFontFamily,
                                size: axis2TitleFontPx,
                                bold: axis2.titleBold,
                                italic: axis2.titleItalic,
                                underline: axis2.titleUnderline,
                            })}
                        </text>
                    )}
                    {axis2BadgeText && (
                        <text
                            x={labelX}
                            y={marginTop - 12}
                            className="unit-axis-badge unit-axis2-badge"
                            textAnchor={onRight ? "start" : "end"}
                            style={{
                                fontSize: `${axis2.unitFontSize}pt`,
                                fill: axis2.unitColor,
                            }}
                        >
                            {axis2BadgeText}
                        </text>
                    )}
                    {axis2.valueShow &&
                        axis2Ticks.map((tick, i) => (
                            <text
                                key={`y2-${i}`}
                                x={labelX}
                                y={marginTop + plotHeight * (1 - tick.ratio) + axis2TickFontPx * 0.35}
                                className="y2-tick-label"
                                textAnchor={onRight ? "start" : "end"}
                                style={{
                                    ...fontStyleOf(fontOf(axis2)),
                                    fill: axis2.labelColor,
                                }}
                            >
                                {tick.label}
                            </text>
                        ))}
                </g>
            );
        };

        return scrolls ? (
            <>
                {/* 1. スクロール領域: プロット領域の幅だけに置く（最底面にスクロールバー）。
                        はみ出した棒・ラベルはこの領域の overflow で切れるため、
                        Y軸の目盛り・タイトルの下に潜り込まない */}
                <div
                    className="unit-bar-scroll"
                    ref={scrollStart.ref}
                    onScroll={scrollStart.onScroll}
                    style={{
                        position: "absolute",
                        left: marginLeft,
                        top: 0,
                        width: viewWidth,
                        height,
                        overflowX: "auto",
                        overflowY: "hidden",
                    }}
                    onWheel={(e) => {
                        if (e.deltaY && !e.deltaX) {
                            e.currentTarget.scrollLeft += e.deltaY;
                        }
                    }}
                >
                    <svg
                        width={plotWidth}
                        height={height - SCROLLBAR_HEIGHT}
                        className="unit-bar-plot-svg"
                        style={{
                            width: `${plotWidth}px`,
                            minWidth: `${plotWidth}px`,
                            height: `${height - SCROLLBAR_HEIGHT}px`,
                            display: "block",
                        }}
                        role="group"
                        aria-label="単位つき棒グラフ（プロット）"
                    >
                        {renderPlot(0)}
                    </svg>
                </div>

                {/* 2. 前面固定レイヤー: Y軸（数値目盛・タイトル・軸線）と単位バッジ。スクロールしても動かない */}
                <svg
                    width={width}
                    height={height - SCROLLBAR_HEIGHT}
                    className="unit-bar-svg"
                    style={{
                        position: "absolute",
                        left: 0,
                        top: 0,
                        width: `${width}px`,
                        height: `${height - SCROLLBAR_HEIGHT}px`,
                        pointerEvents: "none",
                    }}
                    role="img"
                    aria-label="単位つき棒グラフ（Y軸）"
                >
                    {renderPlotTopRightBadge()}
                {renderVerticalDrillPath()}
                    {renderYAxis()}
                    {renderY2Axis()}
                </svg>

                {/* 3. 前面固定レイヤー: X軸タイトル (スクロールバーの直上・プロット領域の中央に固定配置) */}
                {catAxis.titleShow && catAxis.titleText && (
                    <div
                        className="unit-bar-x-title-fixed"
                        style={{
                            position: "absolute",
                            left: marginLeft,
                            bottom: SCROLLBAR_HEIGHT + 2,
                            width: viewWidth,
                            display: "flex",
                            justifyContent: "center",
                            alignItems: "center",
                            pointerEvents: "none",
                        }}
                    >
                        <span
                            className="x-axis-title"
                            {...{ [TEXT_ATTRIBUTE]: "" }}
                            style={{
                                ...fontStyleOf(fontOf(catAxis, "title")),
                                color: catAxis.titleColor,
                                whiteSpace: "nowrap",
                            }}
                            title={catAxis.titleText}
                        >
                            {truncateText(catAxis.titleText, viewWidth, {
                                family: catAxis.titleFontFamily,
                                size: catTitleFontSizePx,
                                bold: catAxis.titleBold,
                                italic: catAxis.titleItalic,
                                underline: catAxis.titleUnderline,
                            })}
                        </span>
                    </div>
                )}
            </>
        ) : (
            /* 非スクロール時: 単一SVGで全てを描画 */
            <svg
                width={width}
                height={height}
                className="unit-bar-svg"
                role="group"
                aria-label="単位つき棒グラフ"
            >
                {/* 軸上バッジ (plotTopRight 配置の場合) */}
                {renderPlotTopRightBadge()}

                {/* プロット描画 (オフセット marginLeft) */}
                {renderPlot(marginLeft)}
                {/* ドリルの位置は、描いたラベルの枠を見るのでプロットのあと */}
                {renderVerticalDrillPath()}

                {/* Y軸（目盛・線・タイトル・軸上バッジ）と第 2 Y 軸 */}
                {renderYAxis()}
                {renderY2Axis()}

                {/* X軸タイトル */}
                {catAxis.titleShow && catAxis.titleText && (
                    <text
                        x={marginLeft + leadIn + (plotWidth - leadIn) / 2}
                        y={height - 6}
                        className="x-axis-title"
                        textAnchor="middle"
                        style={{
                            ...fontStyleOf(fontOf(catAxis, "title")),
                            fill: catAxis.titleColor,
                        }}
                    >
                        {fittedText(catAxis.titleText, plotWidth, {
                            family: catAxis.titleFontFamily,
                            size: catTitleFontSizePx,
                            bold: catAxis.titleBold,
                            italic: catAxis.titleItalic,
                            underline: catAxis.titleUnderline,
                        })}
                    </text>
                )}
            </svg>
        );
    };

    /**
     * 横棒。カテゴリの軸を左、値の軸を下に置き、先頭のカテゴリを上にする（標準と同じ）。
     * 値の比率（startRatio・valRatio）と集合の並び（clusterLayout）は縦棒と同じものを使い、写し方だけ変える。
     * グリッド線は、値の線に「横」、カテゴリの区切りに「縦」の設定を使う（縦棒と同じく値かカテゴリかで対応させる）
     */
    const renderHorizontalChart = (width: number, height: number) => {
        const badgeText = viewModel.unitInfo.badgeText;
        const unitPosition = viewModel.unitInfo.unitPosition;
        const catAxis = viewModel.categoryAxis;
        const valAxis = viewModel.valueAxis;
        const gridlines = viewModel.gridlines;
        const columnsSettings = viewModel.columns;
        const groups = viewModel.categoryGroups;
        const count = groups.length;
        const stacked = viewModel.chartType !== "clustered";

        // 値の軸（下。「軸の位置を切り替える」で上）の目盛り・タイトル・単位ラベル
        const valueAtTop = valAxis.switchPosition;
        const valTickFontPx = valAxis.fontSize * PT_TO_PX;
        const valTitleFontPx = valAxis.titleFontSize * PT_TO_PX;
        const hasValTitle = valAxis.titleShow && Boolean(valAxis.titleText);
        const tickRowHeight = valAxis.show ? valTickFontPx + 8 : 0;
        const valTitleHeight = hasValTitle ? valTitleFontPx + 8 : 0;
        const badgeFontPx = viewModel.unitInfo.fontSize * PT_TO_PX;
        const badgeInAxis = Boolean(badgeText) && unitPosition !== "plotTopRight";
        const badgeAtTopRight = Boolean(badgeText) && unitPosition === "plotTopRight";
        const axisBlockHeight = tickRowHeight + valTitleHeight + (badgeInAxis ? badgeFontPx + 4 : 0);

        // 第 2 X 軸（折れ線の値を値の軸と別の尺度で描くとき）。値の軸の反対側に置く（値の軸が下なら上）
        const axis2 = viewModel.valueAxis2;
        const axis2On = axis2.show && viewModel.lines.length > 0;
        const axis2AtTop = !valueAtTop;
        const axis2TickFontPx = axis2.fontSize * PT_TO_PX;
        const axis2TickRow = axis2On && axis2.valueShow ? axis2TickFontPx + 8 : 0;
        const axis2TitleFontPx = axis2.titleFontSize * PT_TO_PX;
        const hasAxis2Title = axis2On && axis2.titleShow && Boolean(axis2.titleText);
        const axis2TitleHeight = hasAxis2Title ? axis2TitleFontPx + 8 : 0;
        const axis2BadgeText = axis2On ? axis2.badgeText : "";
        const axis2BadgeFontPx = axis2.unitFontSize * PT_TO_PX;
        const axis2BlockHeight = axis2On ? axis2TickRow + axis2TitleHeight + (axis2BadgeText ? axis2BadgeFontPx + 4 : 0) : 0;

        // カテゴリの軸（左）。ラベルは「最大幅 (%)」まで、長ければ末尾を省略
        const catFontPx = catAxis.fontSize * PT_TO_PX;
        const catTitleFontPx = catAxis.titleFontSize * PT_TO_PX;
        const hasCatTitle = catAxis.titleShow && Boolean(catAxis.titleText);
        const catFont = { family: catAxis.fontFamily, size: catFontPx, bold: catAxis.bold, italic: catAxis.italic, underline: catAxis.underline };
        // 階層を段に重ねるときは、棒のすぐ左にいちばん下のレベルだけを出し、上のレベルはさらに左の列に出す
        const stackedLevels = catAxis.show && catAxis.levelCount > 1 && !catAxis.concatenateLabels;
        const labelOf = (g: CategoryGroup) => (stackedLevels ? g.levels[g.levels.length - 1] : g.label);
        const widestLabel = catAxis.show ? Math.max(0, ...groups.map((g) => measureTextWidth(labelOf(g), catFont))) : 0;
        const catTitleWidth = hasCatTitle ? catTitleFontPx + 8 : 0;
        // パレートのランクの帯。いちばん下のレベルのラベルのすぐ左に 1 列取る
        const pareto = viewModel.pareto;
        const rankBandOn = pareto.enabled && pareto.showRankBand && catAxis.show;
        const marginRight = 16;
        // いちばん上の行：「プロット エリアの右上」の単位（右）とドリルの位置（左）。どちらか大きい文字に合わせる
        const badgeRowFontPx = badgeAtTopRight ? badgeFontPx : 0;
        const drillRowFontPx = Math.max(badgeRowFontPx, catAxis.fontSize * PT_TO_PX);
        // 幅は、縦のスクロールバー（12px）が出ても収まるように見積もる
        const drillMaxWidth = (badgeAtTopRight ? width - marginRight - 12 - badgeWidthOf(badgeText, viewModel.unitInfo.fontSize) - 8 : width - marginRight - 12) - 4;
        // 低い図で行を足すと描く範囲が潰れ、軸がビジュアルの外へ押し出される。足す高さが図の 12% を超えるなら出さない
        const drillExtra = drillRowFontPx + 6 - (badgeRowFontPx ? badgeRowFontPx + 6 : 0);
        const drillPathOn = Boolean(fittedDrillPath(drillMaxWidth)) && drillExtra <= height * 0.12;
        const topRowFontPx = drillPathOn ? drillRowFontPx : badgeRowFontPx;
        const marginTop = 10 + (valueAtTop ? axisBlockHeight : 0) + (axis2AtTop ? axis2BlockHeight : 0) + (topRowFontPx ? topRowFontPx + 6 : 0);
        const marginBottom = 6 + (valueAtTop ? axis2BlockHeight : axisBlockHeight);
        const viewHeight = Math.max(10, height - marginTop - marginBottom);

        // 「最小カテゴリの高さ」を下回るなら縦にスクロールする（高さだけで決まるので、名前の欄の幅より先に決める）
        const minCatHeight = Math.max(0, catAxis.minCategoryWidth || 0);
        const neededHeight = minCatHeight > 0 ? Math.ceil(count * minCatHeight) : 0;
        const scrolls = minCatHeight > 0 && neededHeight > viewHeight + 1;
        const SCROLLBAR_WIDTH = scrolls ? 12 : 0;

        // 名前の欄（名前・ランクの帯・上のレベルの列）に使える幅。上限を大きくしても、描く範囲に幅の 4 分の 1 は残す（縦向きと同じ）。
        // 左右の余白とスクロールバーを除く
        const labelRoom = Math.max(0, width * 0.75 - 4 - catTitleWidth - marginRight - SCROLLBAR_WIDTH);
        // 標準と同じく、凡例を含むビジュアル全体の幅に対する割合
        const nameCap = Math.max(16, viewport.width * (catAxis.maxHeight / 100));
        // ランクの帯の幅は、帯を出すランクの「名前 N件」のいちばん長いもの。ラベルの最大幅で頭を打つ（長い名前でプロットを押し出さない。入らなければ名前だけ、それも入らなければ文字を出さない）
        const rankBandWant = rankBandOn
            ? Math.min(
                nameCap,
                Math.max(0, ...(["A", "B", "C"] as ParetoRank[]).map((r) => {
                    const count = pareto.ranks.filter((x) => x === r).length;
                    return count > 0 ? measureTextWidth(`${pareto.labels[r]} ${count}件`, catFont) : 0;
                }))
            ) + 16
            : 0;
        // 名前は、名前の欄からランクの帯と名前の右の余白（8）を除いた幅まで。ただし省いても 1 文字と「…」は残る（truncateText）ので、
        // どの名前も 1 文字は出せる幅（名前ごとに、そのままか「1 文字＋…」の短いほうを実測）は、ランクの帯より先にとる
        // （とらないと、名前がランクの帯に重なる）。省いた名前が幅に入らなければ、その名前だけ出さない
        const shortestOf = (text: string) => Math.min(measureTextWidth(text, catFont), measureTextWidth(truncateText(text, 0, catFont), catFont));
        const minNameWidth = catAxis.show ? Math.max(0, ...groups.map((g) => shortestOf(labelOf(g)))) : 0;
        const maxLabelWidth = Math.min(nameCap, Math.max(0, labelRoom - 8 - rankBandWant, Math.min(minNameWidth, labelRoom - 8)));
        const nameOf = (g: CategoryGroup) => {
            const text = truncateText(labelOf(g), maxLabelWidth, catFont);
            return measureTextWidth(text, catFont) <= maxLabelWidth + 0.5 ? text : null;
        };
        const namesShown = catAxis.show && groups.some((g) => nameOf(g) !== null);
        const lowestLabelWidth = namesShown ? Math.min(widestLabel, maxLabelWidth) + 8 : 0;
        const rankBandWidth = rankBandOn ? Math.max(0, Math.min(rankBandWant, labelRoom - lowestLabelWidth)) : 0;
        // 上のレベルの列。棒に近い（下の）レベルから置き、ラベルと合わせてグラフの幅の半分（と名前の欄）までにする（上のレベルから落とす）
        const levelColumns: Array<{ level: number; width: number }> = [];
        if (stackedLevels) {
            let used = lowestLabelWidth;
            for (let level = catAxis.levelCount - 2; level >= 0; level--) {
                const widest = Math.max(0, ...groups.map((g) => measureTextWidth(g.levels[level] ?? "", catFont)));
                const columnWidth = Math.min(widest, maxLabelWidth) + 12;
                if (used + columnWidth > Math.min(width * 0.5, labelRoom - rankBandWidth)) break;
                levelColumns.push({ level, width: columnWidth });
                used += columnWidth;
            }
        }
        const labelAreaWidth = lowestLabelWidth + rankBandWidth + levelColumns.reduce((sum, c) => sum + c.width, 0);

        const marginLeft = 4 + catTitleWidth + labelAreaWidth;
        const plotHeight = scrolls ? neededHeight : viewHeight;
        const plotWidth = Math.max(10, width - marginLeft - marginRight - SCROLLBAR_WIDTH);
        // 目盛りの本数は、標準と同じく描く範囲の幅で決める（300px 未満は 3、500px 未満は 5、それ以上は 8 まで）
        // 「目盛りの本数 (目安)」があればその本数、空なら描く範囲の幅で決める（第 2 X 軸も同じ上限）。
        // どちらも、いちばん長い数字＋ 8px が隣とぶつからない本数までにとどめる
        const tickTarget = valAxis.tickCount || recommendedTickCount(plotWidth, "horizontal");
        const valTickSpec: FontSpec = { family: valAxis.fontFamily, size: valTickFontPx, bold: valAxis.bold, italic: valAxis.italic, underline: valAxis.underline };
        const axis2TickSpec: FontSpec = { family: axis2.fontFamily, size: axis2TickFontPx, bold: axis2.bold, italic: axis2.italic, underline: false };
        const widest = (spec: FontSpec) => (ticks: Tick[]) => Math.max(0, ...ticks.map((tick) => measureTextWidth(tick.label, spec))) + 8;
        const valueTicks = fitTicks(viewModel.ticksFor, tickTarget, plotWidth, widest(valTickSpec));
        const axis2Ticks = fitTicks(axis2.ticksFor, tickTarget, plotWidth, widest(axis2TickSpec));

        const padRatio = Math.max(0, Math.min(0.5, columnsSettings.categorySpacing / 100));
        const outerRatio = columnsSettings.outerPadding === null
            ? padRatio / 2
            : Math.max(0, Math.min(1, columnsSettings.outerPadding / 100));
        const step = plotHeight / ((count || 1) - padRatio + 2 * outerRatio);
        /** i 番目のカテゴリの中心 (プロット上端からの y) */
        const centerOf = (i: number) => step * (outerRatio + (1 - padRatio) / 2 + i);
        const cluster = clusterLayout(
            step * (1 - padRatio),
            stacked ? 1 : viewModel.series.length,
            columnsSettings.seriesSpacing,
            columnsSettings.maxBarWidth,
            columnsSettings.overlap
        );
        const thickness = cluster.barWidth;
        const drawOrder = (points: DataPoint[]) => drawOrderOf(points, columnsSettings.overlapReverse, viewModel.seriesPositions);
        const clusterSpan = spanOf(cluster);
        // 比較レイヤー（縦棒と同じ。手前の棒は下か上の端）
        const layerDrawsX = layerDrawsFor(thickness);
        const frontTotalsX = !layered || shownLayers[0] === 0;
        const frontShiftX = layered ? layerDrawsX.get(shownLayers[0])!.offset : 0;

        /** 値の比率 → プロット左端からの x（範囲の反転なら右から） */
        const xOfRatio = (ratio: number) => {
            const clamped = Math.max(0, Math.min(1, ratio));
            return plotWidth * (valAxis.invertRange ? 1 - clamped : clamped);
        };
        const extentX = (fromRatio: number, toRatio: number) => {
            const a = xOfRatio(fromRatio);
            const b = xOfRatio(toRatio);
            return { left: Math.min(a, b), width: Math.abs(b - a) };
        };
        const stackInsetX = stacked && !columnsSettings.stackedExplode ? columnsSettings.stackedSpacing : 0;
        /** 棒の範囲。積み上げで下に別の棒があれば、始まりの側を系列間のスペースだけ縮める */
        const segmentExtentX = (d: DataPoint, toRatio: number) => {
            const e = extentX(d.startRatio, toRatio);
            if (!d.insetStart || stackInsetX <= 0) return e;
            const r = insetSpan(e.left, e.width, Math.abs(xOfRatio(d.startRatio) - e.left) < 1e-6, stackInsetX);
            return { left: r.lo, width: r.size };
        };
        const outlineOnlyX = stacked && columnsSettings.borderOutlineOnly && columnsSettings.stackedSpacing <= 0;
        /** 「内側の罫線を非表示」（横棒）。縦棒の renderStackOutline と同じ */
        const renderHStackOutline = (g: CategoryGroup, cy: number, xOffset: number, yOffset: number) => {
            if (!columnsSettings.showBorder) return null;
            const top = yOffset + cy - thickness / 2;
            return [true, false].map((positive) => {
                const segs = g.points.filter((d) => !d.blank && (positive ? d.value > 0 : d.value < 0));
                if (!segs.length) return null;
                const extents = segs.map((d) => segmentExtentX(d, d.valRatio));
                const left = Math.min(...extents.map((e) => e.left));
                const right = Math.max(...extents.map((e) => e.left + e.width));
                const outer = segs.find((d) => d.outermost) ?? segs[segs.length - 1];
                const r = Math.max(0, Math.min(columnsSettings.cornerRadius, thickness / 2, right - left));
                return (
                    <path
                        key={`outline-${positive ? "pos" : "neg"}`}
                        d={hBarPathOf(xOffset + left, top, right - left, thickness, r, positive !== valAxis.invertRange)}
                        className="stack-outline"
                        fill="none"
                        stroke={columnsSettings.borderMatchColumn ? outer.color : columnsSettings.borderFill}
                        strokeWidth={columnsSettings.borderWidth}
                        strokeOpacity={Math.max(0, Math.min(1, 1 - columnsSettings.borderTransparency / 100))}
                        pointerEvents="none"
                    />
                );
            });
        };

        const valueStroke = gridLineStroke(gridlines.horizontalStyle, gridlines.horizontalWidth, gridlines.horizontalScaleWithWidth, { dashArray: gridlines.horizontalDashArray, dashCap: gridlines.horizontalDashCap });
        const valueOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));
        const categoryStroke = gridLineStroke(gridlines.verticalStyle, gridlines.verticalWidth, gridlines.verticalScaleWithWidth, { dashArray: gridlines.verticalDashArray, dashCap: gridlines.verticalDashCap });
        const categoryOpacity = Math.max(0, Math.min(1, 1 - gridlines.verticalTransparency / 100));

        const dl = viewModel.dataLabels;
        /** 描いた棒のデータ ラベルの枠。線のデータ ラベルはこれと重なるものを描かない */
        const hLabelBoxes: Array<{ x0: number; x1: number; y0: number; y1: number }> = [];
        // 横棒のデータラベルの位置。「自動」は集合なら外側の端、積み上げなら中央（積み上げでは外側に置かない）
        const hLabelPositionOf = (d: DataPoint) => {
            const position = (d.labelSettings ?? dl).position;
            return position === "auto" || (stacked && position === "outsideEnd") ? (stacked ? "insideCenter" : "outsideEnd") : position;
        };

        /** 横棒の折れ線。第 2 X 軸には範囲の反転が無いので、そのまま左から右へ */
        const renderHLine = (line: LineSeriesInfo, j: number, xOffset: number, yOffset: number, part: "area" | "line") =>
            renderLineSeries(
                line,
                j,
                part,
                (i) => yOffset + centerOf(i),
                (ratio) => xOffset + (onAxis2(line) ? plotWidth * Math.max(0, Math.min(1, ratio)) : xOfRatio(ratio)),
                true,
                step / 2,
                [yOffset, yOffset + plotHeight],
                clusterSpan
            );

        const renderHBar = (d: DataPoint, cy: number, key: string, xOffset: number, yOffset: number, layer?: LayerDraw) => {
            if (d.blank) return null;
            // 比較レイヤーの棒は、レイヤーの太さで描く
            const th = layer?.width ?? thickness;
            const { left: l, width: w } = segmentExtentX(d, d.valRatio);
            const left = xOffset + l;
            const top = yOffset + cy - th / 2;
            // 棒が右へ伸びるか（正の値。範囲の反転なら逆）
            const rightward = (d.value >= 0) !== valAxis.invertRange;
            const r = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, th / 2, w)) : 0;
            const barPath = hBarPathOf(left, top, w, th, r, rightward);

            const fullyHighlighted =
                d.highlight !== null && Math.abs(d.highlight - d.value) <= Math.abs(d.value) * 1e-9;
            const highlightDimmed = viewModel.hasHighlights && !fullyHighlighted;
            const selectionDimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isPicked(d);
            const dimmed = highlightDimmed || selectionDimmed;
            let highlightPath: string | null = null;
            if (highlightDimmed && d.highlight !== null && d.highlightRatio !== null) {
                const hl = segmentExtentX(d, d.highlightRatio);
                if (hl.width > 0) {
                    const hlR = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, th / 2, hl.width)) : 0;
                    highlightPath = hBarPathOf(xOffset + hl.left, top, hl.width, th, hlR, (d.highlight >= 0) !== valAxis.invertRange);
                }
            }

            const dimFactor = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
            const baseFillOpacity = Math.max(0, Math.min(1, 1 - d.transparency / 100)) * (layer?.opacity ?? 1);
            const ownBorder = d.borderShow && !outlineOnlyX;
            const stroke = layer?.outline ? (viewModel.highContrast?.foreground ?? d.color) : ownBorder ? d.borderColor : "none";
            const strokeWidth = layer?.outline ? 1 : ownBorder ? d.borderWidth : 0;
            const strokeOpacity = dimFactor * (d.borderShow ? Math.max(0, Math.min(1, 1 - d.borderTransparency / 100)) : 1);
            const seriesName = viewModel.seriesMode ? viewModel.series[d.seriesIndex]?.name : undefined;
            // ハイコントラストでは、選んだ棒を選択の色で太く囲む（縦棒と同じ）
            const hcPicked = !!viewModel.highContrast && selectedIds.length > 0 && isPicked(d);

            // データラベル（横に並べる。「縦」の方向は横棒では使わない）
            const label = (() => {
                // 系列ごとに変えた設定があればそれ（保存の無い項目は「すべて」の値）
                const dl = d.labelSettings ?? viewModel.dataLabels;
                if (!dl.show || !d.labelShow || w <= 0 || layer?.label === false) return null;
                const lines = labelLinesOf(dl, d);
                if (!lines.length) return null;
                const block = labelBlock(lines, { singleLine: dl.singleLine, align: dl.horizontalAlignment });
                const textW = block.width;
                const right = left + w;
                let x: number;
                let anchor: "start" | "middle" | "end";
                let inside = true;
                const hLabelPosition = hLabelPositionOf(d);
                if (hLabelPosition === "insideTop") {
                    // 内側の端
                    x = rightward ? right - LABEL_PADDING : left + LABEL_PADDING;
                    anchor = rightward ? "end" : "start";
                } else if (hLabelPosition === "insideCenter") {
                    x = left + w / 2;
                    anchor = "middle";
                } else if (hLabelPosition === "insideBottom") {
                    // 内側の基準（0 の側）
                    x = rightward ? left + LABEL_PADDING : right - LABEL_PADDING;
                    anchor = rightward ? "start" : "end";
                } else {
                    inside = false;
                    x = rightward ? right + LABEL_PADDING : left - LABEL_PADDING;
                    anchor = rightward ? "start" : "end";
                }
                const fitsLength = textW + LABEL_PADDING * 2 <= w;
                const blockHeight = lines.length > 1 ? block.bottomEdge - block.topEdge : lines[0].font.size;
                const fitsThickness = th >= blockHeight + 2;
                // 積み上げでは隣が別の棒なので、値の向きに入りきらないラベルは出さない
                if (inside && stacked && !fitsLength) return null;
                if (inside && !(fitsLength && fitsThickness) && !dl.overflow) return null;
                const autoColor = labelAutoColor(inside, seenColorOf(d.color, baseFillOpacity), dl, viewModel.highContrast, viewModel.background);
                const toneHere = !inside || dl.backgroundShow;
                const midY = top + th / 2;
                // 背景は 1 行なら 1.10 までと同じ高さ、複数行ならまとまりの上端から下端
                const firstPx = lines[0].font.size;
                const bgTop = lines.length > 1 ? midY + block.topEdge - 1 : midY - firstPx * 0.6 - 1;
                const bgHeight = lines.length > 1 ? blockHeight + 2 : firstPx * 1.2 + 2;
                const boxX = anchor === "start" ? x - LABEL_PADDING : anchor === "end" ? x - textW - LABEL_PADDING : x - textW / 2 - LABEL_PADDING;
                // 行はまとまりの中心から置く（左・右そろえと単一行は、行ごとに x が違う）
                const blockCenter = boxX + LABEL_PADDING + textW / 2;
                hLabelBoxes.push({ x0: boxX, x1: boxX + textW + LABEL_PADDING * 2, y0: bgTop, y1: bgTop + bgHeight });
                // 中央そろえは、どの行も棒に対する同じ位置（x・anchor）で描ける（1.35 までと同じ書き方）
                const centered = block.anchor === "middle";
                return (
                    <g className="data-label-group" pointerEvents="none">
                        {dl.backgroundShow && (
                            <rect
                                x={boxX}
                                y={bgTop}
                                width={textW + LABEL_PADDING * 2}
                                height={bgHeight}
                                rx={2}
                                fill={dl.backgroundColor}
                                fillOpacity={(100 - dl.backgroundTransparency) / 100}
                            />
                        )}
                        {lines.map((line, k) => (
                            <text
                                key={line.kind}
                                x={centered ? x : blockCenter + block.xs[k]}
                                y={midY + block.baselines[k]}
                                className={labelClassOf(line)}
                                textAnchor={centered ? anchor : block.anchor}
                                style={labelTextStyle(dl, line, d, autoColor, toneHere)}
                            >
                                {line.text}
                            </text>
                        ))}
                    </g>
                );
            })();

            return (
                <g
                    key={key}
                    className={`bar-item ${hoveredKey === key ? "hovered" : ""}`}
                    onClick={(e) => {
                        e.stopPropagation();
                        if (d.selectionIds) onSelectMany?.(d.selectionIds, e.ctrlKey || e.metaKey);
                        else onSelect(d.selectionId, e.ctrlKey || e.metaKey);
                    }}
                    onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onContextMenu(d.selectionIds ? (null as unknown as ISelectionId) : d.selectionId, e.clientX, e.clientY);
                    }}
                    onMouseEnter={() => setHoveredKey(key)}
                    onMouseLeave={() => setHoveredKey(null)}
                    {...keysOf({
                        label: seriesName ? `${d.category} ${seriesName}: ${d.formattedValue}` : `${d.category}: ${d.formattedValue}`,
                        onSelect: (multi) => (d.selectionIds ? onSelectMany?.(d.selectionIds, multi) : onSelect(d.selectionId, multi)),
                        onClear: onClearSelection,
                        pressed: selectedIds.length > 0 && isPicked(d),
                    })}
                >
                    <path
                        d={barPath}
                        className="bar-rect"
                        style={{
                            fill: d.color,
                            fillOpacity: dimFactor * baseFillOpacity,
                            stroke: hcPicked ? viewModel.highContrast!.selected : stroke,
                            strokeWidth: hcPicked ? HC_SELECTED_WIDTH : strokeWidth,
                            strokeOpacity,
                            strokeDasharray: hcPicked ? undefined : d.borderDash || undefined,
                        }}
                        onMouseEnter={(e) => onTooltipShow?.(d, e.clientX, e.clientY)}
                        onMouseMove={(e) => onTooltipMove?.(d, e.clientX, e.clientY)}
                        onMouseLeave={() => onTooltipHide?.()}
                    />
                    {highlightPath && (
                        <path
                            d={highlightPath}
                            className="bar-highlight"
                            pointerEvents="none"
                            style={{ fill: viewModel.highContrast?.selected ?? d.color, fillOpacity: baseFillOpacity }}
                        />
                    )}
                    {label}
                </g>
            );
        };

        /** 合計ラベル（積み上げ）。正の棒の右の端・負の棒の左の端の外側に置く */
        const renderHTotals = (g: CategoryGroup, cy: number, xOffset: number, yOffset: number) => {
            const tl = viewModel.totalLabels;
            const t = g.totals;
            if (!tl.show || !t) return null;
            const labels: Array<{ key: string; text: string; ratio: number; positive: boolean }> = [];
            if (tl.split) {
                if (t.hasPositive) labels.push({ key: "pos", text: t.positiveText, ratio: t.positiveEndRatio, positive: true });
                if (t.hasNegative) labels.push({ key: "neg", text: t.negativeText, ratio: t.negativeEndRatio, positive: false });
            } else if (t.hasPositive || t.hasNegative) {
                const positive = t.net >= 0 ? t.hasPositive : !t.hasNegative;
                labels.push({ key: "net", text: t.netText, ratio: positive ? t.positiveEndRatio : t.negativeEndRatio, positive });
            }
            const fontPx = tl.fontSize * PT_TO_PX;
            return labels.map((label) => {
                const outwardRight = label.positive !== valAxis.invertRange;
                const edgeX = xOffset + xOfRatio(label.ratio);
                return (
                    <text
                        key={`total-${label.key}`}
                        x={outwardRight ? edgeX + 4 : edgeX - 4}
                        y={yOffset + cy + fontPx * 0.35}
                        className="total-label"
                        textAnchor={outwardRight ? "start" : "end"}
                        pointerEvents="none"
                        style={{
                            fill: tl.color,
                            ...fontStyleOf(fontOf(tl)),
                        }}
                    >
                        {label.text}
                    </text>
                );
            });
        };

        /**
         * 階層の上のレベルの列。いちばん下のレベルのラベルの左に、1 つ上のレベルから順に列を足し、
         * 同じ親が続く区切りを点線の枠で囲む。文字は区切りの高さの中央、列の右寄せ
         */
        const renderHLevelColumns = (xOffset: number, yOffset: number) => {
            const paths = groups.map((g) => g.levels);
            const edgeOf = (i: number) => Math.max(0, Math.min(plotHeight, centerOf(i) - step / 2));
            const separator: React.CSSProperties = {
                stroke: catAxis.labelColor,
                strokeOpacity: 0.35,
                strokeWidth: 1,
                strokeDasharray: "1 2",
                shapeRendering: "auto",
            };
            const boxed = catAxis.hierarchyStyle === "boxed";
            const nodes: React.ReactNode[] = [];
            let right = xOffset - lowestLabelWidth - rankBandWidth;
            levelColumns.forEach(({ level, width: columnWidth }, c) => {
                const left = right - columnWidth;
                if (!boxed) {
                    nodes.push(
                        <line key={`lv-side-${c}`} x1={right} y1={yOffset} x2={right} y2={yOffset + plotHeight} className="level-separator" style={separator} />
                    );
                }
                levelRunsOf(paths, level, groups.map((g) => g.levelKeys)).forEach((run, k) => {
                    const top = run.start === 0 ? 0 : edgeOf(run.start);
                    const bottom = run.end === groups.length - 1 ? plotHeight : edgeOf(run.end + 1);
                    // 区切りを押すと、その区切りのカテゴリをまとめて選ぶ（縦棒と同じ）
                    nodes.push(renderLevelHit(run, level, { x: left, y: yOffset + top, width: columnWidth, height: bottom - top }, `lv-hit-${c}-${k}`));
                    if (boxed) {
                        // 帯：区切りごとに角の丸い淡い面。隣の帯とは少し離す
                        nodes.push(
                            <rect
                                key={`lv-box-${c}-${k}`}
                                x={left + LEVEL_BOX_GAP}
                                y={yOffset + top + LEVEL_BOX_GAP}
                                width={Math.max(0, columnWidth - LEVEL_BOX_GAP * 2)}
                                height={Math.max(0, bottom - top - LEVEL_BOX_GAP * 2)}
                                rx={LEVEL_BOX_RADIUS}
                                className="level-box"
                                style={{
                                    fill: catAxis.labelColor,
                                    fillOpacity: levelPicked(run) ? LEVEL_BAND_PICKED_OPACITY : LEVEL_BAND_OPACITY,
                                }}
                                pointerEvents="none"
                            />
                        );
                    }
                    // 区切りの線は棒の左から、この列の左まで
                    if (!boxed) [top, bottom].forEach((y, e) =>
                        nodes.push(
                            <line key={`lv-edge-${c}-${k}-${e}`} x1={left} y1={yOffset + y} x2={xOffset} y2={yOffset + y} className="level-separator" style={separator} />
                        )
                    );
                    // 文字の高さも無い区切りには文字を出さない（隣の区切りの文字と重なる）
                    if (bottom - top < catFontPx) return;
                    nodes.push(
                        <text
                            key={`lv-text-${c}-${k}`}
                            x={boxed ? (left + right) / 2 : right - 6}
                            y={yOffset + (top + bottom) / 2 + catFontPx * 0.35}
                            className="x-category-level"
                            textAnchor={boxed ? "middle" : "end"}
                            pointerEvents="none"
                            style={{
                                ...fontStyleOf(fontOf(catAxis)),
                                fill: catAxis.labelColor,
                            }}
                        >
                            {truncateText(run.text, columnWidth - 12, catFont)}
                            <title>{run.text}</title>
                        </text>
                    );
                });
                right = left;
            });
            return <g className="category-levels">{nodes}</g>;
        };

        /** プロット（グリッド線・棒・ラベル・カテゴリ名）。xOffset・yOffset はプロットの左上、labelRightX はカテゴリ名の右端 */
        const hGridMask = gridMaskAttr(layerDrawsX);
        const renderHPlot = (xOffset: number, yOffset: number, labelRightX: number) => (
            <>
                {hGridMask && gridCutout(gridMaskId, `${gridMaskId}-bars`)}
                <g className="gridlines" mask={hGridMask}>
                {gridlines.horizontalShow && (
                    <g className="y-gridlines-group">
                        {valueTicks.map((tick, i) => {
                            const x = xOffset + plotWidth * tick.ratio;
                            return (
                                <line
                                    key={`v-grid-${i}`}
                                    x1={x}
                                    y1={yOffset}
                                    x2={x}
                                    y2={yOffset + plotHeight}
                                    stroke={gridlines.horizontalColor}
                                    strokeWidth={gridlines.horizontalWidth}
                                    strokeDasharray={valueStroke.dashArray}
                                    strokeLinecap={valueStroke.lineCap}
                                    shapeRendering={valueStroke.shapeRendering}
                                    strokeOpacity={valueOpacity}
                                    className="grid-line"
                                />
                            );
                        })}
                    </g>
                )}
                {gridlines.verticalShow && (
                    <g className="x-gridlines-group">
                        {[
                            0,
                            ...categoryGridBreaks(groups.map((g) => g.levels), catAxis.levelCount, gridlines.verticalPosition === "levels", groups.map((g) => g.levelKeys)).map(
                                (i) => centerOf(i) - step / 2
                            ),
                            plotHeight,
                        ].map((y, i) => (
                            <line
                                key={`h-grid-${i}`}
                                x1={xOffset}
                                y1={yOffset + y}
                                x2={xOffset + plotWidth}
                                y2={yOffset + y}
                                stroke={gridlines.verticalColor}
                                strokeWidth={gridlines.verticalWidth}
                                strokeDasharray={categoryStroke.dashArray}
                                strokeLinecap={categoryStroke.lineCap}
                                shapeRendering={categoryStroke.shapeRendering}
                                strokeOpacity={categoryOpacity}
                            />
                        ))}
                    </g>
                )}
                </g>
                {viewModel.ribbons.show && groups.length > 1 && (
                    <g className="ribbons-group">
                        {renderRibbonBands(
                            groups,
                            (i) => [yOffset + centerOf(i) + thickness / 2, yOffset + centerOf(i + 1) - thickness / 2],
                            (d) => {
                                const e = segmentExtentX(d, d.valRatio);
                                return [xOffset + e.left, xOffset + e.left + e.width];
                            },
                            true
                        )}
                    </g>
                )}
                {/* Y 軸（カテゴリの軸）の定数線の網掛け領域と、背面の線（棒の後ろ） */}
                {renderCategoryLines((i) => yOffset + centerOf(i), step, [xOffset, xOffset + plotWidth], true, [yOffset, yOffset + plotHeight], "back")}
                {renderValueLines((ratio, secondary) => xOffset + (secondary ? plotWidth * ratio : xOfRatio(ratio)), [yOffset, yOffset + plotHeight], false, [xOffset, xOffset + plotWidth], "back")}
                <g className="bars-group" id={hGridMask ? `${gridMaskId}-bars` : undefined}>
                    {groups.map((g, i) => {
                        const cy = centerOf(i);
                        return (
                            <g key={`cat-${i}`} className="category-group">
                                {layered
                                    ? [...shownLayers].reverse().flatMap((l) => {
                                          const draw = layerDrawsX.get(l)!;
                                          return drawOrder(layerPointsOf(g, l)).map(([d, s]) =>
                                              renderHBar(d, cy + (stacked ? 0 : cluster.offsets[viewModel.seriesPositions[s] ?? s] ?? 0) + draw.offset, `${i}-${s}-${l}`, xOffset, yOffset, draw)
                                          );
                                      })
                                    : drawOrder(g.points).map(([d, s]) =>
                                          renderHBar(d, cy + (stacked ? 0 : cluster.offsets[viewModel.seriesPositions[s] ?? s] ?? 0), `${i}-${s}`, xOffset, yOffset)
                                      )}
                                {outlineOnlyX && !layered && renderHStackOutline(g, cy, xOffset, yOffset)}
                                {viewModel.measureGroups
                                    ? shownLayers.map((l) => <g key={`totals-${l}`}>{renderHTotals(withLayerTotals(g, l), cy + layerDrawsX.get(l)!.offset, xOffset, yOffset)}</g>)
                                    : frontTotalsX && renderHTotals(g, cy + frontShiftX, xOffset, yOffset)}
                                {namesShown && nameOf(g) !== null && !viewModel.noCategory && (
                                    <g {...categoryLabelHit(i, g.category)}>
                                        <text
                                            x={labelRightX}
                                            y={yOffset + cy + catFontPx * 0.35}
                                            className="x-category-label"
                                            textAnchor="end"
                                            style={{
                                                ...fontStyleOf(fontOf(catAxis)),
                                                fill: catAxis.labelColor,
                                            }}
                                        >
                                            {nameOf(g)}
                                            <title>{g.category}</title>
                                        </text>
                                    </g>
                                )}
                            </g>
                        );
                    })}
                </g>
                {levelColumns.length > 0 && renderHLevelColumns(xOffset, yOffset)}
                {rankBandOn && (
                    <g className="pareto-ranks">
                        {rankRuns().map((run) => {
                            const edgeOf = (i: number) => Math.max(0, Math.min(plotHeight, centerOf(i) - step / 2));
                            const top = run.start === 0 ? 0 : edgeOf(run.start);
                            const bottom = run.end === pareto.ranks.length ? plotHeight : edgeOf(run.end);
                            const left = xOffset - lowestLabelWidth - rankBandWidth;
                            return renderRankRun(
                                run,
                                { x: left + 2, y: yOffset + top + 2, width: rankBandWidth - 4, height: bottom - top - 4 },
                                catFont,
                                {
                                    ...fontStyleOf(fontOf(catAxis)),
                                    fill: catAxis.labelColor,
                                }
                            );
                        })}
                    </g>
                )}
                {pareto.enabled && pareto.showThresholds && (
                    // 境目は累積比の軸（上、0〜100%）の位置に縦の破線
                    <g className="pareto-thresholds" pointerEvents="none">
                        {pareto.thresholds.map((t, k) => {
                            const x = xOffset + plotWidth * Math.max(0, Math.min(1, t));
                            const ratioLine = viewModel.lines.find((l) => l.selectable === false);
                            return (
                                <line
                                    key={`th-${k}`}
                                    x1={x}
                                    y1={yOffset}
                                    x2={x}
                                    y2={yOffset + plotHeight}
                                    className="pareto-threshold"
                                    style={{ stroke: ratioLine?.color ?? catAxis.labelColor, strokeOpacity: 0.6, strokeWidth: 1, strokeDasharray: "4 3" }}
                                />
                            );
                        })}
                    </g>
                )}
                {/* 折れ線の値。点はカテゴリの中心で、上から下へ並ぶ。既定では線を出さずマーカーだけ */}
                {viewModel.lines.length > 0 && (
                    <g className="lines-group">
                        {viewModel.lines.map((line, j) => renderHLine(line, j, xOffset, yOffset, "area"))}
                        {viewModel.lines.map((line, j) => renderHLine(line, j, xOffset, yOffset, "line"))}
                        {(() => {
                            const boxes = hLabelBoxes;
                            return viewModel.lines.map((line, j) =>
                                renderLineLabels(
                                    line,
                                    j,
                                    (i) => yOffset + centerOf(i),
                                    (ratio) => xOffset + (onAxis2(line) ? plotWidth * Math.max(0, Math.min(1, ratio)) : xOfRatio(ratio)),
                                    true,
                                    [xOffset, xOffset + plotWidth],
                                    boxes
                                )
                            );
                        })()}
                    </g>
                )}
                {/* Y 軸（カテゴリの軸）の定数線 */}
                {renderCategoryLines((i) => yOffset + centerOf(i), step, [xOffset, xOffset + plotWidth], true, [yOffset, yOffset + plotHeight], "front")}
                {/* X 軸（値の軸）の定数線 */}
                {renderValueLines((ratio, secondary) => xOffset + (secondary ? plotWidth * ratio : xOfRatio(ratio)), [yOffset, yOffset + plotHeight], false, [xOffset, xOffset + plotWidth], "front")}
            </>
        );

        /** 値の軸（目盛り・タイトル・単位ラベル）とカテゴリの軸のタイトル。スクロールしても動かない */
        const plotRightX = marginLeft + plotWidth;

        /** 第 2 X 軸の目盛り・タイトル・単位ラベル。値の軸の反対側（上か下）に、プロットから外へ向かって並べる */
        const renderH2Axis = () => {
            if (!axis2On) return null;
            const tickY = axis2AtTop ? marginTop - 4 : marginTop + viewHeight + axis2TickFontPx + 2;
            const titleY = axis2AtTop ? marginTop - axis2TickRow - 4 : marginTop + viewHeight + axis2TickRow + axis2TitleFontPx + 2;
            const badgeY = axis2AtTop
                ? marginTop - axis2TickRow - axis2TitleHeight - 4
                : marginTop + viewHeight + axis2TickRow + axis2TitleHeight + axis2BadgeFontPx + 2;
            return (
                <g className="y2-axis-container">
                    {axis2.valueShow &&
                        axis2Ticks.map((tick, i) => (
                            <text
                                key={`x2-${i}`}
                                x={marginLeft + plotWidth * tick.ratio}
                                y={tickY}
                                className="y2-tick-label"
                                textAnchor="middle"
                                style={{
                                    ...fontStyleOf(fontOf(axis2)),
                                    fill: axis2.labelColor,
                                }}
                            >
                                {tick.label}
                            </text>
                        ))}
                    {hasAxis2Title && (
                        <text
                            x={marginLeft + plotWidth / 2}
                            y={titleY}
                            className="y2-axis-title"
                            textAnchor="middle"
                            style={{
                                ...fontStyleOf(fontOf(axis2, "title")),
                                fill: axis2.titleColor,
                            }}
                        >
                            {fittedText(axis2.titleText, plotWidth, {
                                family: axis2.titleFontFamily,
                                size: axis2TitleFontPx,
                                bold: axis2.titleBold,
                                italic: axis2.titleItalic,
                                underline: axis2.titleUnderline,
                            })}
                        </text>
                    )}
                    {axis2BadgeText && (
                        <text
                            x={plotRightX}
                            y={badgeY}
                            className="unit-axis-badge unit-axis2-badge"
                            textAnchor="end"
                            style={{ fontSize: `${axis2.unitFontSize}pt`, fill: axis2.unitColor }}
                        >
                            {axis2BadgeText}
                        </text>
                    )}
                </g>
            );
        };
        const renderHAxes = () => {
            const tickY = valueAtTop ? marginTop - 4 : marginTop + viewHeight + valTickFontPx + 2;
            const titleY = valueAtTop ? marginTop - tickRowHeight - 4 : marginTop + viewHeight + tickRowHeight + valTitleFontPx + 2;
            const badgeY = valueAtTop
                ? marginTop - tickRowHeight - valTitleHeight - 4
                : marginTop + viewHeight + tickRowHeight + valTitleHeight + badgeFontPx + 2;
            const valueTextStyle = {
                ...fontStyleOf(fontOf(valAxis)),
                fill: valAxis.labelColor,
            };
            return (
                <g className="y-axis-container">
                    {valAxis.show && (
                        <g className="y-axis-group">
                            {valueTicks.map((tick, i) => (
                                <text
                                    key={i}
                                    x={marginLeft + plotWidth * tick.ratio}
                                    y={tickY}
                                    className="y-tick-label"
                                    textAnchor="middle"
                                    style={valueTextStyle}
                                >
                                    {tick.label}
                                </text>
                            ))}
                        </g>
                    )}
                    {hasValTitle && (
                        <text
                            x={marginLeft + plotWidth / 2}
                            y={titleY}
                            className="y-axis-title"
                            textAnchor="middle"
                            style={{
                                ...fontStyleOf(fontOf(valAxis, "title")),
                                fill: valAxis.titleColor,
                            }}
                        >
                            {fittedText(valAxis.titleText, plotWidth, {
                                family: valAxis.titleFontFamily,
                                size: valTitleFontPx,
                                bold: valAxis.titleBold,
                                italic: valAxis.titleItalic,
                                underline: valAxis.titleUnderline,
                            })}
                        </text>
                    )}
                    {badgeText && (
                        <text
                            x={plotRightX}
                            y={badgeAtTopRight ? 10 + topRowFontPx * 0.85 : badgeY}
                            className="unit-axis-badge"
                            textAnchor="end"
                            style={{ fontSize: `${viewModel.unitInfo.fontSize}pt`, fill: viewModel.unitInfo.color }}
                        >
                            {badgeText}
                        </text>
                    )}
                    {drillPathOn &&
                        renderDrillPath(4, 10 + topRowFontPx * 0.85, drillMaxWidth)}
                    {renderH2Axis()}
                    {hasCatTitle && (
                        <text
                            transform={`translate(${4 + catTitleFontPx * 0.5}, ${marginTop + viewHeight / 2}) rotate(-90)`}
                            textAnchor="middle"
                            dominantBaseline="central"
                            className="x-axis-title"
                            style={{
                                ...fontStyleOf(fontOf(catAxis, "title")),
                                fill: catAxis.titleColor,
                            }}
                        >
                            {fittedText(catAxis.titleText, viewHeight, {
                                family: catAxis.titleFontFamily,
                                size: catTitleFontPx,
                                bold: catAxis.titleBold,
                                italic: catAxis.titleItalic,
                                underline: catAxis.titleUnderline,
                            })}
                        </text>
                    )}
                </g>
            );
        };

        if (scrolls) {
            return (
                <>
                    {/* 縦にスクロールする領域。カテゴリ名も棒と一緒に流れる */}
                    <div
                        className="unit-bar-vscroll"
                        ref={scrollStart.ref}
                        onScroll={scrollStart.onScroll}
                        style={{
                            position: "absolute",
                            left: marginLeft - labelAreaWidth,
                            top: marginTop,
                            width: labelAreaWidth + plotWidth + SCROLLBAR_WIDTH,
                            height: viewHeight,
                            overflowX: "hidden",
                            overflowY: "auto",
                        }}
                    >
                        <svg
                            width={labelAreaWidth + plotWidth}
                            height={plotHeight}
                            className="unit-bar-plot-svg"
                            style={{ display: "block" }}
                            role="group"
                            aria-label="単位つき棒グラフ（プロット）"
                        >
                            {renderHPlot(labelAreaWidth, 0, labelAreaWidth - 6)}
                        </svg>
                    </div>
                    <svg
                        width={width}
                        height={height}
                        className="unit-bar-svg"
                        style={{ position: "absolute", left: 0, top: 0, width: `${width}px`, height: `${height}px`, pointerEvents: "none" }}
                        role="img"
                        aria-label="単位つき棒グラフ（軸）"
                    >
                        {renderHAxes()}
                    </svg>
                </>
            );
        }
        return (
            <svg width={width} height={height} className="unit-bar-svg" role="group" aria-label="単位つき棒グラフ">
                {renderHPlot(marginLeft, marginTop, marginLeft - 6)}
                {renderHAxes()}
            </svg>
        );
    };

    /** 向きで描き分ける */
    const renderBody = (width: number, height: number) =>
        viewModel.orientation === "horizontal" ? renderHorizontalChart(width, height) : renderChart(width, height);

    /**
     * 凡例。項目をクリックすると系列ごと選ぶ（Ctrl で追加）。
     * 選択があるときは、選んだ系列（または選んだ棒のある系列）以外の印を薄くする
     */
    /**
     * 凡例の棒の印。棒と同じ色・透過性・罫線の四角にし、棒に角丸があれば値の向きの端を丸める（縦棒は上、横棒は右）
     */
    const renderLegendBarGlyph = (entry: LegendItemInfo | undefined, item: LegendItemBox, dimmed: boolean, r: number) => {
        const side = r * 2;
        const style = entry?.barStyle;
        const dim = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        const corner = Math.min(viewModel.columns.cornerRadius, side / 3);
        const left = item.x;
        const top = item.y - r;
        const d = viewModel.orientation === "horizontal"
            ? hBarPathOf(left, top, side, side, corner, true)
            : barPathOf(left, top, side, side, corner, true);
        return (
            <path
                d={d}
                className="legend-marker"
                style={{
                    fill: item.color,
                    fillOpacity: dim * Math.max(0, Math.min(1, 1 - (style?.transparency ?? 0) / 100)) * (entry?.layerOpacity ?? 1),
                    stroke: style?.borderShow ? style.borderColor : "none",
                    strokeWidth: style?.borderShow ? Math.min(2, style.borderWidth) : 0,
                    strokeOpacity: dim,
                    strokeDasharray: style?.borderDash || undefined,
                }}
            />
        );
    };

    /**
     * 凡例の折れ線の印。線と同じ色・線のスタイルの短い線を引き、マーカーを出していれば真ん中に同じ形で重ねる。
     * 横棒では線が上から下へ流れるので、縦の短い線にする（ユーザーの提案）。
     * 線を出していなければマーカーだけ（線もマーカーも出していなければ、印が消えないよう線を引く）
     */
    const renderLegendLineGlyph = (entry: LegendItemInfo, item: LegendItemBox, dimmed: boolean, r: number) => {
        const line = viewModel.lines[entry.index];
        const mk = line?.marker ?? viewModel.markers;
        const dim = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        const color = line?.color ?? item.color;
        const width = Math.max(1, Math.min(3, line?.width ?? 2));
        // スタイル：線とマーカー（既定・1.35 まで）・線・マーカー・マーカー（円）
        const style = viewModel.legend.markerStyle;
        const markerOnly = style === "markerOnly" || style === "markerCircleDefault";
        const drawLine = style === "lineOnly" || (!markerOnly && (!line || line.lineShow || !mk.show));
        const drawMarker = markerOnly || (style !== "lineOnly" && mk.show);
        const markerShape = style === "markerCircleDefault" ? "circle" : mk.shape;
        const markerFill = viewModel.legend.matchLineColor ? color : mk.color || color;
        const custom = line?.lineStyle === CUSTOM_LINE_STYLE ? lineStrokeOf(line, width) : null;
        const dash = custom
            ? custom.dashArray
            : line?.lineStyle === "dashed" ? `${width * 2} ${width * 1.5}` : line?.lineStyle === "dotted" ? `${width * 0.1} ${width * 2}` : undefined;
        const cx = item.x + item.glyphWidth / 2;
        const horizontal = viewModel.orientation === "horizontal";
        // 横棒の縦の線は、行の高さに収まる長さ
        const half = horizontal ? r * 1.4 : item.glyphWidth / 2;
        // マーカーは行に収まる大きさまで（半径はサイズの 0.75 倍）
        const size = Math.min(mk.size, (r * 1.3) / 0.75);
        return (
            <>
                {drawLine && (
                    <line
                        x1={horizontal ? cx : cx - half}
                        y1={horizontal ? item.y - half : item.y}
                        x2={horizontal ? cx : cx + half}
                        y2={horizontal ? item.y + half : item.y}
                        className="legend-line"
                        stroke={color}
                        strokeWidth={width}
                        strokeDasharray={dash}
                        strokeLinecap={custom ? custom.lineCap : line?.lineStyle === "dotted" ? "round" : "butt"}
                        strokeOpacity={dim}
                    />
                )}
                {drawMarker && (
                    <path
                        d={markerPath(markerShape, cx, item.y, markerOnly ? Math.min(Math.max(size, r / 0.75), (r * 1.3) / 0.75) : size, horizontal)}
                        transform={mk.rotation && markerShape !== "circle" ? `rotate(${mk.rotation} ${cx} ${item.y})` : undefined}
                        className="legend-marker"
                        style={{
                            fill: markerFill,
                            fillOpacity: dim * Math.max(0, Math.min(1, 1 - mk.transparency / 100)),
                            stroke: mk.borderShow ? (mk.borderMatchLine ? color : mk.borderColor) : "none",
                            strokeWidth: mk.borderShow ? Math.min(2, mk.borderWidth) : 0,
                            strokeOpacity: dim,
                        }}
                    />
                )}
            </>
        );
    };

    /**
     * 凡例の比較レイヤーの印。棒の形（縦棒なら縦長、横棒なら横長）を、そのレイヤーの濃さで塗る。
     * 棒の系列の四角と形で見分けられ、レイヤーどうしは濃さで見分ける（2026-10-06 ユーザーと決めた案 C）
     */
    const renderLegendLayerGlyph = (entry: LegendItemInfo, item: LegendItemBox, dimmed: boolean, r: number) => {
        const horizontal = viewModel.orientation === "horizontal";
        const width = r * (horizontal ? LAYER_GLYPH_LENGTH : LAYER_GLYPH_THICKNESS);
        const height = r * (horizontal ? LAYER_GLYPH_THICKNESS : LAYER_GLYPH_LENGTH);
        const dim = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        return (
            <rect
                className="legend-layer-marker"
                x={item.x}
                y={item.y - height / 2}
                width={width}
                height={height}
                style={{ fill: item.color, fillOpacity: dim * (entry.layerOpacity ?? 1) }}
            />
        );
    };

    /**
     * 定数線（X 軸・Y 軸）を描く。items の at は線の位置（線に直交する方向の座標）。
     * lineHorizontal なら線は横（span は x の範囲、clamp は y の範囲）、そうでなければ縦（span は y、clamp は x）。
     * part が back なら網掛け領域と背面の線（棒より先に描く）、front なら前面の線とデータ ラベル（棒のあとに描く）。
     * shadeOf は網掛け領域の範囲（clamp の方向の [始め, 終わり]）。ラベルの置き場所は placeCategoryLineLabels
     */
    const renderReferenceLines = (
        kind: string,
        items: Array<{ key: string; label: string; at: number }>,
        style: Omit<ViewModel["categoryLine"], "position" | "shadeRegion">,
        lineHorizontal: boolean,
        span: [number, number],
        clamp: [number, number],
        part: "back" | "front",
        shadeOf: (at: number) => [number, number]
    ) => {
        if (!style.show || !items.length) return null;
        const lineStyle = referenceLineSvgStyle(style);
        const drawLines = (style.layer === "back") === (part === "back");
        const lineOf = (key: string, at: number) => (
            <line
                key={key}
                x1={lineHorizontal ? span[0] : at}
                y1={lineHorizontal ? at : span[0]}
                x2={lineHorizontal ? span[1] : at}
                y2={lineHorizontal ? at : span[1]}
                className={kind}
                style={lineStyle}
            />
        );
        if (part === "back") {
            return (
                <g className={`${kind}s-back`} pointerEvents="none">
                    {style.shadeShow &&
                        items.map(({ key, at }) => {
                            const [a, b] = shadeOf(at);
                            const [from, to] = [Math.min(a, b), Math.max(a, b)];
                            return (
                                <rect
                                    key={`${key}-shade`}
                                    className={`${kind}-shade`}
                                    x={lineHorizontal ? span[0] : from}
                                    y={lineHorizontal ? from : span[0]}
                                    width={lineHorizontal ? span[1] - span[0] : to - from}
                                    height={lineHorizontal ? to - from : span[1] - span[0]}
                                    style={referenceShadeSvgStyle(style)}
                                />
                            );
                        })}
                    {drawLines && items.map(({ key, at }) => lineOf(key, at))}
                </g>
            );
        }
        const fontPx = style.labelFontSize * PT_TO_PX;
        const font: FontSpec = { family: viewModel.categoryAxis.fontFamily, size: fontPx, bold: false, italic: false, underline: false };
        const placed = placeReferenceLabels(
            items,
            lineHorizontal,
            clamp,
            span,
            fontPx,
            (text) => measureTextWidth(text, font),
            style.labelAvoidOverlap,
            style.labelHorizontal,
            style.labelVertical
        );
        return (
            <g className={`${kind}s`} pointerEvents="none">
                {drawLines && items.map(({ key, at }) => lineOf(key, at))}
                {style.labelShow &&
                    items.map(({ key, label }, k) =>
                        label ? (
                            <text
                                key={`${key}-label`}
                                x={placed[k].x}
                                y={placed[k].y}
                                textAnchor={placed[k].anchor}
                                className={`${kind}-label`}
                                // 棒や隣の線の上でも読めるよう、文字のまわりに白い縁を付ける
                                style={referenceLabelSvgStyle(style, viewModel.categoryAxis.fontFamily, viewModel.highContrast?.background ?? "#FFFFFF")}
                            >
                                {label}
                            </text>
                        ) : null
                    )}
            </g>
        );
    };

    /**
     * X 軸の定数線。印の付いたカテゴリの前（境目）・中央・後ろに、プロットを横切る線を引く。
     * centerAt は i 番目のカテゴリの中心。縦棒では線は縦、横棒では横。網掛けは線より前（左・上）か後ろ（右・下）
     */
    const renderCategoryLines = (
        centerAt: (i: number) => number,
        step: number,
        span: [number, number],
        horizontal: boolean,
        clamp: [number, number],
        part: "back" | "front"
    ) => {
        const cl = viewModel.categoryLine;
        const offset = cl.position === "before" ? -step / 2 : cl.position === "after" ? step / 2 : 0;
        const items = viewModel.categoryLines.flatMap((line, j) =>
            line.marks.map((mark) => ({ key: `cl-${j}-${mark.index}`, label: mark.label, at: Math.max(clamp[0], Math.min(clamp[1], centerAt(mark.index) + offset)) }))
        );
        return renderReferenceLines("category-line", items, cl, horizontal, span, clamp, part, (at) => (cl.shadeRegion === "before" ? [clamp[0], at] : [at, clamp[1]]));
    };

    /**
     * Y 軸の定数線。値の軸の位置に、プロットを横切る線を引く。縦棒では線は横、横棒では縦。
     * coordOf は値の比率（0〜1）を値の方向の座標にする（第 2 軸の線は secondary）。網掛けは値の小さい側か大きい側
     */
    const renderValueLines = (
        coordOf: (ratio: number, secondary: boolean) => number,
        span: [number, number],
        lineHorizontal: boolean,
        clamp: [number, number],
        part: "back" | "front"
    ) => {
        const vl = viewModel.valueLine;
        const items = viewModel.valueLines.map((line, j) => ({ key: `vl-${j}`, label: line.label, at: coordOf(line.ratio, line.secondary), secondary: line.secondary }));
        return renderReferenceLines("value-line", items, vl, lineHorizontal, span, clamp, part, (at) => {
            const line = items.find((item) => item.at === at);
            const end = coordOf(vl.shadeRegion === "below" ? 0 : 1, line?.secondary ?? false);
            return [at, end];
        });
    };

    const renderLegend = (layout: LegendLayout) => {
        const r = layout.markerRadius;
        const anyPicked = selectedIds.length > 0 && !viewModel.hasHighlights;
        // ページ送りの矢印の色（Fluent 2 の枠なしのボタン。テーマの地から作る。ハイコントラストでは前景色と背景色だけ）
        const legendArrowColors = viewModel.highContrast
            ? controlColorsOf(viewModel.highContrast.foreground, viewModel.highContrast.background, true)
            : controlColorsOf(DARK_TEXT, viewModel.background ?? "#FFFFFF");
        // 太字は項目ごと（凡例の題は太字、項目は書式の太字）
        const textStyle = { ...fontStyleOf(fontOf(lg)), fill: lg.color };
        return (
            <svg
                className="unit-bar-legend"
                width={layout.box.width}
                height={layout.box.height}
                style={{ position: "absolute", left: layout.box.x, top: layout.box.y, overflow: "hidden" }}
                role="group"
                aria-label="凡例"
            >
                {layout.title && (
                    <text
                        x={layout.title.x}
                        y={layout.title.y + legendFontPx * 0.35}
                        className="legend-title"
                        style={{ ...textStyle, fontWeight: "bold" }}
                    >
                        {layout.title.text}
                    </text>
                )}
                {layout.groupTitles.map((t) => (
                    <text key={`legend-group-${t.text}`} x={t.x} y={t.y + legendFontPx * 0.35} className="legend-title" style={{ ...textStyle, fontWeight: "bold" }}>
                        {t.text}
                    </text>
                ))}
                {layout.arrows &&
                    [layout.arrows.prev, layout.arrows.next].map((arrow, k) => (
                        <LegendArrowButton
                            key={`legend-arrow-${k}`}
                            arrow={arrow}
                            color={lg.color}
                            colors={legendArrowColors}
                            label={k === 0 ? "凡例の前のページ" : "凡例の次のページ"}
                            onPress={() => setLegendPage(Math.max(0, Math.min(layout.pageCount - 1, layout.page + (k === 0 ? -1 : 1))))}
                        />
                    ))}
                {/* 一覧の選択肢は項目だけ（矢印・タイトルは一覧の外） */}
                <g role="listbox" aria-multiselectable={true} aria-label="凡例の項目">
                {layout.items.map((item) => {
                    const entry = viewModel.legendEntries[item.index];
                    // 比較レイヤーは、凡例で選んだレイヤー以外を薄くする（通常の選択では薄くしない）
                    const dimmed = entry?.kind === "layer" ? layerFilterOn && !pickedLayers.includes(entry.index) : anyPicked && !isEntryPicked(entry);
                    const selectEntry = (multi: boolean) => {
                        if (entry?.kind === "layer") onSelectLayer?.(entry.index, multi);
                        else if (entry?.selectionId) onSelect(entry.selectionId, multi);
                    };
                    return (
                        <g
                            key={`legend-${item.index}`}
                            className="legend-item"
                            onClick={(e) => {
                                e.stopPropagation();
                                selectEntry(e.ctrlKey || e.metaKey);
                            }}
                            // 押せる項目は Tab で止まる。凡例の一覧の項目のまま（読み上げは「一覧の項目」）
                            // 凡例は一覧の選択肢。選べる項目だけ Tab で止まり、選ばれているかを読み上げる
                            {...(entry && (entry.kind === "layer" || entry.selectionId) && interactive
                                ? {
                                      tabIndex: 0,
                                      onKeyDown: selectKeyDown(selectEntry, onClearSelection),
                                      "aria-selected": entry.kind === "layer" ? !layerFilterOn || pickedLayers.includes(entry.index) : anyPicked && isEntryPicked(entry),
                                  }
                                : { "aria-disabled": true })}
                            role="option"
                            aria-label={item.name}
                        >
                            {entry?.kind === "layer"
                                ? renderLegendLayerGlyph(entry, item, dimmed, r)
                                : entry?.kind === "line"
                                  ? renderLegendLineGlyph(entry, item, dimmed, r)
                                  : renderLegendBarGlyph(entry, item, dimmed, r)}
                            <text
                                x={item.x + item.glyphWidth + LEGEND_MARKER_GAP}
                                y={item.y + legendFontPx * 0.35}
                                className="legend-label"
                                style={{ ...textStyle, fontWeight: lg.bold ? "bold" : "normal" }}
                            >
                                {item.text}
                                {item.text !== item.name && <title>{item.name}</title>}
                            </text>
                        </g>
                    );
                })}
                </g>
            </svg>
        );
    };

    const { width } = viewport;
    const height = bodyHeight;
    const chartWidth = legend ? Math.max(10, width - legend.reserve.left - legend.reserve.right) : width;
    const chartHeight = legend ? Math.max(10, height - legend.reserve.top - legend.reserve.bottom) : height;

    /**
     * 累計の切り替えボタンと区切りの選択（閲覧者が押せる所）。押した状態は visual.ts がレポートに保存する。
     * 棒のクリックと同じく、ここでの操作は選択の解除に伝えない
     */
    const renderToolbar = () => {
        const cumulative = viewModel.cumulative;
        const on = cumulative.enabled;
        // ハイコントラストでは 1 本目の系列が背景色なので、前景色にする
        // オンのときの色は、棒の色に引きずられない固定の灰色（ハイコントラストでは前景色）
        const accent = viewModel.highContrast?.foreground ?? TOOLBAR_ACCENT;
        const textColor = viewModel.categoryAxis.labelColor;
        const font: React.CSSProperties = { fontFamily: viewModel.categoryAxis.fontFamily, fontSize: "9pt" };
        // ボタンの枠・押した地の灰は、標準の字（#252423）と白から作る（形と動きは visual.less）。ハイコントラストでは前景色と背景色だけ
        const hc = viewModel.highContrast;
        const controls = controlVars(hc ? controlColorsOf(hc.foreground, hc.background, true) : controlColorsOf(DARK_TEXT, "#FFFFFF"));
        const stop = (e: React.SyntheticEvent) => e.stopPropagation();
        const resetItems = [{ value: "none", displayName: "区切らない" }, ...cumulative.levels];
        return (
            <div
                className="cumulative-toolbar"
                style={{ position: "absolute", top: 0, right: 0, height: TOOLBAR_HEIGHT, display: "flex", alignItems: "center", gap: 6, paddingRight: 4, whiteSpace: "nowrap", ...controls } as React.CSSProperties}
                onClick={stop}
                onContextMenu={stop}
            >
                {on && cumulative.levels.length > 0 && (
                    <select
                        className="cumulative-reset"
                        aria-label="累計の区切り"
                        value={cumulative.reset}
                        onChange={(e) => onChangeCumulativeReset?.(e.target.value)}
                        style={{ ...font, color: textColor }}
                    >
                        {resetItems.map((item) => (
                            <option key={String(item.value)} value={String(item.value)}>
                                {`区切り: ${String(item.displayName)}`}
                            </option>
                        ))}
                    </select>
                )}
                <button
                    type="button"
                    className="cumulative-toggle"
                    aria-pressed={on}
                    onClick={() => onToggleCumulative?.()}
                    style={{ ...font, color: on ? accent : textColor, "--toggle-accent": accent } as React.CSSProperties}
                >
                    <span aria-hidden="true" className="cumulative-dot" />
                    累計
                </button>
            </div>
        );
    };

    /** 凡例とグラフ。凡例があれば、凡例を除いた領域にグラフを描く（中の絶対配置はこの領域が基準になる） */
    const renderLegendAndBody = () =>
        legend ? (
            <>
                <div
                    className="unit-bar-chart-area"
                    style={{
                        position: "absolute",
                        left: legend.reserve.left,
                        top: legend.reserve.top,
                        width: chartWidth,
                        height: chartHeight,
                    }}
                >
                    {renderBody(chartWidth, chartHeight)}
                </div>
                {renderLegend(legend)}
            </>
        ) : (
            renderBody(width, height)
        );

    return (
        <div
            className="unit-bar-container"
            // キーで移ったときの枠の色（ハイコントラストでは選択の色）。style/visual.less の --focus-color
            style={{ width, height: viewport.height, position: "relative", ["--focus-color" as string]: viewModel.highContrast?.selected ?? "#252423" }}
            // 棒のクリック・右クリックは棒側で止めるので、ここに来るのは棒以外の所だけ。
            // （以前は target === currentTarget で判定していたが、SVG が全面を覆うため常に偽だった）
            onClick={() => onClearSelection?.()}
            onContextMenu={(e) => {
                e.preventDefault();
                onContextMenu(null as unknown as ISelectionId, e.clientX, e.clientY);
            }}
        >
            {toolbarOn && renderToolbar()}
            {toolbarOn ? (
                // ボタンの行の下を、凡例とグラフの領域にする（中の絶対配置はこの領域が基準になる）
                <div className="unit-bar-body" style={{ position: "absolute", left: 0, top: toolbarHeight, width, height }}>
                    {renderLegendAndBody()}
                </div>
            ) : (
                renderLegendAndBody()
            )}
            {viewModel.notice && (
                <div
                    className="unit-bar-notice"
                    style={{
                        position: "absolute",
                        left: 4,
                        bottom: 2,
                        padding: "0 4px",
                        fontFamily: viewModel.categoryAxis.fontFamily,
                        fontSize: "8pt",
                        color: viewModel.categoryAxis.labelColor,
                        background: viewModel.highContrast?.background ?? "rgba(255, 255, 255, 0.85)",
                        pointerEvents: "none",
                    }}
                >
                    {viewModel.notice}
                </div>
            )}
            {interactive && viewModel.copyButton && (
                <CopyImageButton
                    alt="グラフの画像"
                    stamp={[viewModel, viewport.width, viewport.height, selectedIds, legendPage]}
                    {...(viewModel.highContrast
                        ? { background: viewModel.highContrast.background, colors: { foreground: viewModel.highContrast.foreground, background: viewModel.highContrast.background } }
                        : {})}
                    browserMenu={browserMenu}
                />
            )}
        </div>
    );
};

