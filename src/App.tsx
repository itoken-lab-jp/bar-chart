"use strict";

import * as React from "react";
import powerbi from "powerbi-visuals-api";
import IViewport = powerbi.IViewport;
import ISelectionId = powerbi.visuals.ISelectionId;

import { ViewModel, DataPoint, CategoryGroup, LineSeriesInfo, LegendItemInfo, DataLabelsSettings, ParetoRank } from "./viewModel";
import { VisualFormattingSettingsModel, STEP_WIDTHS } from "./settings";
import { contrastingText, placeLabel, labelBlock, measureTextWidth, LABEL_PADDING, FontSpec, LabelLine } from "./unitUtils";
import { clusterLayout, spanOf, levelRunsOf } from "./layout";
import { layoutLegend, LegendLayout, LegendItemBox, LEGEND_MARKER_GAP } from "./legend";
import { linePath, areaPath, markerPath, XY } from "./linePath";

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

/** 斜めのカテゴリ名を、棒の下端から離す間隔 (px) */
const ROTATED_LABEL_GAP = 8;
/** 階層の「囲み」の枠：隣の枠とのすき間の半分と、角の丸み (px) */
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

/** pt -> px 換算比率 (1pt = 4/3 px = 1.3333...px) */
const PT_TO_PX = 4 / 3;

interface DataLabelLine extends LabelLine {
    kind: "value" | "detail";
}

/** データラベルの行（値の行と詳細の行。標準と同じく値が上） */
function labelLinesOf(dl: DataLabelsSettings, d: DataPoint): DataLabelLine[] {
    const lines: DataLabelLine[] = [];
    if (dl.valueShow) {
        lines.push({
            kind: "value",
            text: d.dataLabelText,
            font: { family: dl.fontFamily, size: dl.fontSize * PT_TO_PX, bold: dl.bold, italic: dl.italic, underline: false },
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

/** データラベルの行の文字の見た目。色が空なら autoColor（棒の色に合わせた白か黒） */
function labelTextStyle(dl: DataLabelsSettings, line: DataLabelLine, d: DataPoint, autoColor: string): React.CSSProperties {
    const detail = line.kind === "detail";
    return {
        fill: (detail ? dl.detailColor : d.labelColor) || autoColor,
        ...(detail && dl.detailTransparency > 0 ? { fillOpacity: 1 - dl.detailTransparency / 100 } : {}),
        fontSize: `${line.font.size / PT_TO_PX}pt`,
        fontFamily: line.font.family,
        fontWeight: line.font.bold ? "bold" : "normal",
        fontStyle: line.font.italic ? "italic" : "normal",
        ...(line.font.underline ? { textDecoration: "underline" } : {}),
    };
}

/**
 * グリッド線の線種。標準に合わせ、点線は細かい点（長さ 1 の線に丸い端）、破線は 4px 刻み。
 * 点線は端を丸めるので crispEdges を掛けない（点がつぶれる）
 */
export function gridLineStroke(style: string, width = 1, scaleWithWidth = false): {
    dashArray?: string;
    lineCap?: "round";
    shapeRendering: "crispEdges" | "auto";
} {
    // 「幅で拡大縮小」がオンなら、点線・破線の模様を線の幅に比例させる（標準と同じく、細い線では模様が細かくなる）。
    // オフなら、幅によらず 1.10 までと同じ模様
    const w = Math.max(1, width);
    if (style === "dotted") {
        return { dashArray: scaleWithWidth ? `${w} ${2 * w}` : "1 3", lineCap: "round", shapeRendering: "auto" };
    }
    if (style === "dashed") return { dashArray: scaleWithWidth ? `${3 * w} ${3 * w}` : "4 4", shapeRendering: "crispEdges" };
    return { shapeRendering: "crispEdges" };
}

/** 棒 (または棒のハイライト部分) の path。r > 0 なら値の向きの端だけ角を丸める */
function barPathOf(left: number, top: number, width: number, height: number, r: number, roundAtTop: boolean): string {
    if (r <= 0) {
        return `M ${left},${top} h ${width} v ${height} h ${-width} Z`;
    }
    if (roundAtTop) {
        return `M ${left},${top + height} ` +
            `L ${left},${top + r} ` +
            `Q ${left},${top} ${left + r},${top} ` +
            `L ${left + width - r},${top} ` +
            `Q ${left + width},${top} ${left + width},${top + r} ` +
            `L ${left + width},${top + height} Z`;
    }
    return `M ${left},${top} ` +
        `L ${left + width},${top} ` +
        `L ${left + width},${top + height - r} ` +
        `Q ${left + width},${top + height} ${left + width - r},${top + height} ` +
        `L ${left + r},${top + height} ` +
        `Q ${left},${top + height} ${left},${top + height - r} Z`;
}

/** 横棒 (または横棒のハイライト部分) の path。r > 0 なら値の向きの端（右か左）だけ角を丸める */
function hBarPathOf(left: number, top: number, width: number, height: number, r: number, roundAtRight: boolean): string {
    if (r <= 0) {
        return `M ${left},${top} h ${width} v ${height} h ${-width} Z`;
    }
    if (roundAtRight) {
        return `M ${left},${top} ` +
            `L ${left + width - r},${top} ` +
            `Q ${left + width},${top} ${left + width},${top + r} ` +
            `L ${left + width},${top + height - r} ` +
            `Q ${left + width},${top + height} ${left + width - r},${top + height} ` +
            `L ${left},${top + height} Z`;
    }
    return `M ${left + width},${top} ` +
        `L ${left + r},${top} ` +
        `Q ${left},${top} ${left},${top + r} ` +
        `L ${left},${top + height - r} ` +
        `Q ${left},${top + height} ${left + r},${top + height} ` +
        `L ${left + width},${top + height} Z`;
}

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
}

/** 閲覧者向けの累計の切り替えボタンの行の高さ (px)。出すときだけグラフの上に取る */
const TOOLBAR_HEIGHT = 28;

const NO_SELECTION: ISelectionId[] = [];

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
}) => {
    const [hoveredKey, setHoveredKey] = React.useState<string | null>(null);

    if (viewModel.isEmpty) {
        return (
            <div className="unit-bar-landing">
                <p>カテゴリと値をフィールドに配置してください</p>
            </div>
        );
    }

    const lg = viewModel.legend;
    const legendFontPx = lg.fontSize * PT_TO_PX;

    // 閲覧者向けの累計の切り替え。出すぶんだけ上を空け、凡例とグラフはその下に描く
    const toolbarOn = viewModel.cumulative.toggle && interactive;
    const toolbarHeight = toolbarOn ? TOOLBAR_HEIGHT : 0;
    const bodyHeight = Math.max(10, viewport.height - toolbarHeight);

    // 凡例（複数系列のときだけ）。グラフの外側に置き、その分だけグラフの領域を縮める
    const legend: LegendLayout | null = lg.show
        ? layoutLegend({
            entries: viewModel.legendEntries.map((e) => ({ name: e.name, color: e.color, kind: e.kind })),
            title: lg.title,
            font: { family: lg.fontFamily, size: legendFontPx, bold: lg.bold, italic: lg.italic, underline: lg.underline },
            position: lg.position,
            width: viewport.width,
            height: bodyHeight,
            horizontal: viewModel.orientation === "horizontal",
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
                role="button"
                tabIndex={0}
                aria-label={full}
                aria-pressed={picked}
                style={{ cursor: "pointer" }}
                onClick={(e) => {
                    e.stopPropagation();
                    select(e.ctrlKey || e.metaKey);
                }}
                onKeyDown={(e) => {
                    if (e.key !== "Enter" && e.key !== " ") return;
                    e.preventDefault();
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

    /** 折れ線が選ばれているか（凡例で線ごと、または線の点） */
    const isLinePicked = (lineIndex: number): boolean => {
        const line = viewModel.lines[lineIndex];
        return !!line && selectedIds.some((s) => s.equals(line.selectionId) || line.points.some((p) => s.equals(p.selectionId)));
    };

    /** 凡例の項目が選ばれているか（棒の系列なら、その系列の棒のどれかが選ばれている） */
    const isEntryPicked = (entry: LegendItemInfo | undefined): boolean =>
        !!entry &&
        (entry.kind === "line"
            ? isLinePicked(entry.index)
            : viewModel.dataPoints.some((d) => d.seriesIndex === entry.index && isPicked(d)));

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
        const pts = line.points.map((p, i) => {
            const pos = posOf(i);
            const val = p.ratio === null ? null : valueOf(p.ratio);
            const at: XY | null = val === null ? null : horizontal ? { x: val, y: pos } : { x: pos, y: val };
            return { p, i, at };
        });
        // 値の無いカテゴリで切った、途切れのない点の並び
        const runs: XY[][] = [];
        let run: XY[] = [];
        for (const pt of pts) {
            // 累計の区切りで 0 に戻るところも切る（前の区切りの終わりから下がる線を引かない）
            if (pt.p.breakBefore && run.length) {
                runs.push(run);
                run = [];
            }
            if (pt.at === null) {
                if (run.length) runs.push(run);
                run = [];
                continue;
            }
            run.push(pt.at);
        }
        if (run.length) runs.push(run);
        const shape = {
            interpolation: line.interpolation,
            smoothing: line.smoothing,
            tension: line.tension / 100,
            stepPosition: line.stepPosition,
            stepConnect: line.stepConnect,
            stepLevelWidth: line.stepWidth === STEP_WIDTHS.bar ? barBand : 0,
        };
        const path = runs.map((r) => linePath(r, shape, horizontal, stepExtension, stepBounds)).join(" ");
        const dimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isLinePicked(j);
        const opacity = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        const dash =
            line.lineStyle === "dashed"
                ? `${line.width * 3} ${line.width * 2}`
                : line.lineStyle === "dotted"
                    ? `${line.width * 0.1} ${line.width * 2}`
                    : undefined;
        const mk = viewModel.markers;
        const markerFill = mk.color || line.color;
        const markerOpacity = opacity * Math.max(0, Math.min(1, 1 - mk.transparency / 100));
        const markerStroke = mk.borderShow ? (mk.borderMatchLine ? line.color : mk.borderColor) : "none";
        const markerStrokeOpacity = opacity * Math.max(0, Math.min(1, 1 - mk.borderTransparency / 100));
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
                        strokeDasharray={dash}
                        strokeLinejoin={line.lineJoin === "miter" || line.lineJoin === "bevel" ? line.lineJoin : "round"}
                        strokeLinecap={line.lineStyle === "dotted" ? "round" : "butt"}
                        strokeOpacity={opacity}
                        pointerEvents="none"
                    />
                )}
                {pts.map((pt) =>
                    pt.at === null ? null : (
                        <g key={`pt-${j}-${pt.i}`}>
                            {mk.show && (
                                <path
                                    d={markerPath(mk.shape, pt.at.x, pt.at.y, mk.size, horizontal)}
                                    className="line-marker"
                                    style={{
                                        fill: markerFill,
                                        fillOpacity: markerOpacity,
                                        stroke: markerStroke,
                                        strokeWidth: mk.borderShow ? mk.borderWidth : 0,
                                        strokeOpacity: markerStrokeOpacity,
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
                    )
                )}
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
            for (const t of viewModel.ticks) {
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
            ? Math.max(14, Math.max(2, ...axis2.ticks.map((t) => t.label.length)) * axis2TickFontPx * 0.55)
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
        const marginTop = badgeText || axis2BadgeText ? 36 : 22;

        // カテゴリ最小幅と横スクロール判定
        const viewWidth = Math.max(10, width - marginLeft - marginRight);
        const groups = viewModel.categoryGroups;
        const count = groups.length;
        const minCatWidth = Math.max(0, catAxis.minCategoryWidth || 0);
        // minCategoryWidth が設定されていて、count * minCategoryWidth が viewWidth を上回る場合に横スクロール発動
        const neededWidth = minCatWidth > 0 ? Math.ceil(count * minCatWidth) : 0;
        const scrolls = minCatWidth > 0 && neededWidth > viewWidth + 1;
        const SCROLLBAR_HEIGHT = scrolls ? 12 : 0;

        const plotWidth = scrolls ? neededWidth : viewWidth;
        const padRatio = Math.max(0, Math.min(0.5, columnsSettings.categorySpacing / 100));
        // 外側のパディング: 最初の棒の前と最後の棒の後ろの余白。カテゴリ 1 つ分の幅
        // (step = 棒 + カテゴリ間のスペース) に対する比。「自動」はカテゴリ間のスペースの半分で、
        // 各カテゴリの枠の中央に棒を置く 1.3.1.0 までの配置と同じになる
        const outerRatio = columnsSettings.outerPadding === null
            ? padRatio / 2
            : Math.max(0, Math.min(1, columnsSettings.outerPadding / 100));
        const step = plotWidth / ((count || 1) - padRatio + 2 * outerRatio);
        /** i 番目のカテゴリの中心 (プロット左端からの x) */
        const centerOf = (i: number) => step * (outerRatio + (1 - padRatio) / 2 + i);
        // カテゴリ 1 つぶんの帯に、系列の棒を並べる（系列 1 本なら帯いっぱいの 1 本）。
        // 積み上げは系列を 1 本の棒に積むので、帯いっぱいの 1 本ぶんの幅
        const stacked = viewModel.chartType !== "clustered";
        const cluster = clusterLayout(
            step * (1 - padRatio),
            stacked ? 1 : viewModel.series.length,
            columnsSettings.seriesSpacing,
            columnsSettings.maxBarWidth
        );
        const barWidth = cluster.barWidth;
        const clusterSpan = spanOf(cluster);
        // 積み上げのデータラベルは棒の中に置く。外側の指定（自動・外側の上）は中央に読み替える
        const labelPosition =
            stacked && (viewModel.dataLabels.position === "auto" || viewModel.dataLabels.position === "outsideEnd")
                ? "insideCenter"
                : viewModel.dataLabels.position;

        // X軸カテゴリラベルの幅（実測）と回転の判定
        const catFontSizePx = catAxis.fontSize * PT_TO_PX;
        const catFont: FontSpec = { family: catAxis.fontFamily, size: catFontSizePx, bold: catAxis.bold, italic: catAxis.italic, underline: catAxis.underline };
        // 階層を段に重ねるときは、棒のすぐ下にいちばん下のレベルだけを出し、上のレベルはその下の段に出す
        const stackedLevels = catAxis.show && catAxis.levelCount > 1 && !catAxis.concatenateLabels;
        const labelOf = (g: CategoryGroup) => (stackedLevels ? g.levels[g.levels.length - 1] : g.category);
        const categoryLabelStyle: React.CSSProperties = {
            fontSize: `${catAxis.fontSize}pt`,
            fontFamily: catAxis.fontFamily,
            fontWeight: catAxis.bold ? "bold" : "normal",
            fontStyle: catAxis.italic ? "italic" : "normal",
            textDecoration: catAxis.underline ? "underline" : undefined,
            fill: catAxis.labelColor,
        };
        const boxedLevels = stackedLevels && catAxis.hierarchyStyle === "boxed";
        // 標準の段の枠は薄い点線。色はラベルの色に合わせる（ハイコントラストでもラベルと同じ色になる）
        const levelSeparatorStyle: React.CSSProperties = {
            stroke: catAxis.labelColor,
            strokeOpacity: 0.35,
            strokeWidth: 1,
            strokeDasharray: "1 2",
            shapeRendering: "crispEdges",
        };
        // 囲みの枠。ラベルの色をごく淡く敷き、同じ色の細い線で縁取る
        const levelBoxStyle: React.CSSProperties = {
            fill: catAxis.labelColor,
            fillOpacity: 0.06,
            stroke: catAxis.labelColor,
            strokeOpacity: 0.3,
            strokeWidth: 1,
        };
        const widestCat = catAxis.show ? Math.max(0, ...groups.map((g) => measureTextWidth(labelOf(g), catFont))) : 0;

        // バンド幅に収まらなければ斜め -45 度に回転
        const shouldRotateCat = catAxis.show && widestCat > step * 0.92;
        // 階層を段に重ねるときは、斜めではなく縦に立てる（標準と同じ。斜めだと左隣の区切りの線をまたぐ）
        const uprightCat = shouldRotateCat && stackedLevels;

        // X軸タイトル高さ (タイトル領域は高さ最大値の判定外で独立確保し、重なりを防止)
        const hasCatTitle = catAxis.titleShow && Boolean(catAxis.titleText);
        const catTitleFontSizePx = catAxis.titleFontSize * PT_TO_PX;
        const catTitleHeight = hasCatTitle ? catTitleFontSizePx + 10 : 0;

        // X軸ラベル領域 (高さの最大値 % はここだけに効く)
        // maxHeight は 0〜100% (既定 25%)。標準と同じく、凡例を含むビジュアル全体の高さに対する割合。
        // 斜めのラベルでは、割合はラベルの縦の広がりに効き、棒との間隔と文字の太さの半分は別に足す
        // （Desktop で、タイトルありの標準は 8 文字、タイトルなしは 9 文字まで出した。1.13 までは 1 文字早かった）
        const labelExtentMax = Math.max(16, viewport.height * (catAxis.maxHeight / 100));
        const maxLabelAreaHeight = shouldRotateCat ? ROTATED_LABEL_GAP + labelExtentMax + catFontSizePx * 0.5 : labelExtentMax;

        // 必要とされるラベル高さ（斜めは、棒との間隔＋ラベルの縦の広がり＋文字の太さの半分。縦はラベルの長さそのまま）
        const desiredLabelHeight = uprightCat
            ? ROTATED_LABEL_GAP + widestCat + catFontSizePx * 0.5
            : shouldRotateCat
                ? ROTATED_LABEL_GAP + widestCat * Math.SQRT1_2 + catFontSizePx * 0.5
                : catFontSizePx + 10;

        const labelAreaHeight = catAxis.show
            ? Math.max(14, Math.min(desiredLabelHeight, maxLabelAreaHeight))
            : 0;

        // ラベルの長さの上限 (px)。超えたら末尾を「…」に省略。斜めは、ラベル領域の高さから棒との間隔を除いた分
        // 領域がいちばん長いラベルちょうどのときに、計算の誤差で省略しないよう 0.5px の余裕を持たせる
        const maxAllowedLabelLen = uprightCat
            ? Math.max(12, labelAreaHeight - ROTATED_LABEL_GAP - catFontSizePx * 0.5 + 0.5)
            : shouldRotateCat
                ? Math.max(12, (labelAreaHeight - ROTATED_LABEL_GAP - catFontSizePx * 0.5) / Math.SQRT1_2 + 0.5)
                : Math.max(12, step * 0.92);

        // 上のレベルの段。多くてもプロットを潰さないよう、ラベルと合わせてグラフの高さの半分までにする（上のレベルから落とす）
        // 囲みは枠の上下に余白が要るので、段を少し高くする
        const levelRowHeight = catFontSizePx + (boxedLevels ? 12 : 8);
        const levelRows = stackedLevels
            ? Math.min(catAxis.levelCount - 1, Math.max(0, Math.floor((height * 0.5 - labelAreaHeight) / levelRowHeight)))
            : 0;
        const levelAreaHeight = levelRows * levelRowHeight;

        // パレートのランクの帯。カテゴリのラベル（と階層の段）の下に 1 段取る
        const pareto = viewModel.pareto;
        const rankBandOn = pareto.enabled && pareto.showRankBand && catAxis.show;
        const rankBandHeight = rankBandOn ? catFontSizePx + 12 : 0;

        // 全体の下部マージン = ラベル領域 + 上のレベルの段 + ランクの帯 + タイトル領域 + スクロールバー高さ + 余白
        const marginBottom = Math.max(20, labelAreaHeight + levelAreaHeight + rankBandHeight + catTitleHeight + SCROLLBAR_HEIGHT + 6);
        const plotHeight = Math.max(10, height - marginTop - marginBottom);

        // 横グリッド線 (Y軸目盛線) の線種と透過性
        const hStroke = gridLineStroke(gridlines.horizontalStyle, gridlines.horizontalWidth, gridlines.horizontalScaleWithWidth);
        const hOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));

        // 縦グリッド線 (X軸目盛線) の線種と透過性
        const vStroke = gridLineStroke(gridlines.verticalStyle, gridlines.verticalWidth, gridlines.verticalScaleWithWidth);
        const vOpacity = Math.max(0, Math.min(1, 1 - gridlines.verticalTransparency / 100));

        // 縦グリッド線の位置: プロットの両端と、隣り合うカテゴリの中間
        const verticalGridXs = [
            0,
            ...groups.slice(1).map((_, k) => centerOf(k + 1) - step / 2),
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

        /** 棒 1 本（系列 1 本ぶん）。cx は棒の中心の x */
        const renderBar = (d: DataPoint, cx: number, key: string) => {
            if (d.blank) return null;
            const barLeft = cx - barWidth / 2;

            const { top: barTop, height: barH } = extentBetween(d.startRatio, d.valRatio);

            const isHovered = hoveredKey === key;

            const isPositive = d.value >= 0;
            // 積み上げでは、角丸は外側の端の棒だけ
            const r = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, barWidth / 2, barH)) : 0;

            // 範囲反転時の角丸向き（通常: 正なら上、負なら下。反転時: 正なら下、負なら上）
            const roundAtTop = valAxis.invertRange ? !isPositive : isPositive;
            const barPath = barPathOf(barLeft, barTop, barWidth, barH, r, roundAtTop);

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
                const hl = extentBetween(d.startRatio, d.highlightRatio);
                if (hl.height > 0) {
                    const hlPositive = d.highlight >= 0;
                    const hlR = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, barWidth / 2, hl.height)) : 0;
                    highlightPath = barPathOf(
                        barLeft, hl.top, barWidth, hl.height, hlR,
                        valAxis.invertRange ? !hlPositive : hlPositive
                    );
                }
            }

            // 透明度は塗りだけに効かせる（opacity だと境界線まで消える）。
            // 罫線は罫線の透過性だけで決まる。ハイライト時の減光は両方に掛ける
            const dimFactor = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
            const baseFillOpacity = Math.max(0, Math.min(1, 1 - d.transparency / 100));
            const fillOpacity = dimFactor * baseFillOpacity;
            const stroke = d.borderShow ? d.borderColor : "none";
            const strokeWidth = d.borderShow ? d.borderWidth : 0;
            const strokeOpacity = dimFactor * (d.borderShow ? Math.max(0, Math.min(1, 1 - d.borderTransparency / 100)) : 1);
            const seriesName = viewModel.seriesMode ? viewModel.series[d.seriesIndex]?.name : undefined;

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
                    role="graphics-symbol"
                    aria-label={seriesName ? `${d.category} ${seriesName}: ${d.formattedValue}` : `${d.category}: ${d.formattedValue}`}
                >
                    <path
                        d={barPath}
                        className="bar-rect"
                        style={{
                            fill: d.color,
                            fillOpacity,
                            stroke,
                            strokeWidth,
                            strokeOpacity,
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
                                fill: d.color,
                                fillOpacity: baseFillOpacity,
                            }}
                        />
                    )}

                    {/* データラベル */}
                    {(() => {
                        const dl = viewModel.dataLabels;
                        if (!dl.show || !d.labelShow || barH <= 0) return null;

                        const lines = labelLinesOf(dl, d);
                        if (!lines.length) return null;
                        const isVertical = dl.orientation === "vertical";

                        const barBottom = barTop + barH;
                        const placed = placeLabel({
                            position: labelPosition,
                            top: barTop,
                            bottom: barBottom,
                            barWidth,
                            value: lines[0].text,
                            font: lines[0].font,
                            lines,
                            vertical: isVertical,
                            // 積み上げでは外側が上の棒なので、入りきらないラベルを外へ逃がさず出さない
                            overflow: stacked ? false : dl.overflow,
                        });

                        if (!placed.fitsVertically && !placed.outside) return null;
                        if (!placed.fitsAlongBar && !dl.overflow) return null;

                        const autoColor = placed.outside || dl.backgroundShow ? "#252423" : contrastingText(d.color);
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
                                        x={0}
                                        y={placed.baselines[k]}
                                        className={line.kind === "value" ? "data-label" : "data-label-detail"}
                                        textAnchor="middle"
                                        style={labelTextStyle(dl, line, d, autoColor)}
                                    >
                                        {line.text}
                                    </text>
                                ))}
                            </g>
                        );
                    })()}
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
                                fontSize: `${tl.fontSize}pt`,
                                fontFamily: tl.fontFamily,
                                fontWeight: tl.bold ? "bold" : "normal",
                                fontStyle: tl.italic ? "italic" : "normal",
                                textDecoration: tl.underline ? "underline" : undefined,
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
                    if (boxedLevels) {
                        // 囲み：区切りごとに角の丸い淡い枠。隣の枠とは少し離す
                        nodes.push(
                            <rect
                                key={`lv-box-${r}-${k}`}
                                x={xOffset + left + LEVEL_BOX_GAP}
                                y={top + LEVEL_BOX_GAP}
                                width={Math.max(0, right - left - LEVEL_BOX_GAP * 2)}
                                height={Math.max(0, levelRowHeight - LEVEL_BOX_GAP * 2)}
                                rx={LEVEL_BOX_RADIUS}
                                className="level-box"
                                style={levelBoxStyle}
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
                        const left = run.start === 0 ? 0 : edgeOf(run.start);
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
                // 斜め45度回転 (標準準拠: 棒の直下から左下に伸び、Y軸ラベルの下に被さる)
                const labelTop = marginTop + plotHeight + ROTATED_LABEL_GAP;
                return (
                    <text
                        x={cx - 4}
                        y={labelTop}
                        transform={`rotate(-45, ${cx - 4}, ${labelTop})`}
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
                    const e = extentBetween(d.startRatio, d.valRatio);
                    return [e.top, e.top + e.height];
                },
                false
            );
        };

        // プロット内容の描画。xOffset はプロット左端の x
        // （非スクロール時は marginLeft。スクロール時はプロット幅だけのスクロール領域に描くので 0）
        const renderPlot = (xOffset: number) => {
            return (
                <>
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
                            {viewModel.ticks.map((tick, i) => {
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

                    {/* リボンの帯。棒の後ろに描く */}
                    {viewModel.ribbons.show && groups.length > 1 && <g className="ribbons-group">{renderRibbons(xOffset)}</g>}

                    {/* 棒とデータラベルとカテゴリラベル。カテゴリごとに、系列の棒を凡例の順に並べる */}
                    <g className="bars-group">
                        {groups.map((g, i) => {
                            const cx = xOffset + centerOf(i);
                            return (
                                <g key={`cat-${i}`} className="category-group">
                                    {g.points.map((d, s) => renderBar(d, cx + (stacked ? 0 : cluster.offsets[s] ?? 0), `${i}-${s}`))}
                                    {renderTotalLabels(g, cx)}
                                    {catAxis.show && renderCategoryLabel(labelOf(g), cx)}
                                </g>
                            );
                        })}
                    </g>

                    {levelRows > 0 && renderLevelRows(xOffset)}
                    {rankBandOn && renderRankBand(xOffset)}
                    {pareto.enabled && pareto.showThresholds && renderThresholds(xOffset)}

                    {/* 折れ線（複合）。棒の上に重ねる */}
                    {viewModel.lines.length > 0 && (
                        <g className="lines-group">
                            {/* 網掛け領域はすべての線の下に塗る */}
                            {viewModel.lines.map((line, j) => renderLine(line, j, xOffset, "area"))}
                            {viewModel.lines.map((line, j) => renderLine(line, j, xOffset, "line"))}
                        </g>
                    )}
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
                                fontSize: `${valAxis.titleFontSize}pt`,
                                fontFamily: valAxis.titleFontFamily,
                                fontWeight: valAxis.titleBold ? "bold" : "normal",
                                fontStyle: valAxis.titleItalic ? "italic" : "normal",
                                textDecoration: valAxis.titleUnderline ? "underline" : undefined,
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
                        {viewModel.ticks.map((tick, i) => {
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
                                                fontSize: `${valAxis.fontSize}pt`,
                                                fontFamily: valAxis.fontFamily,
                                                fontWeight: valAxis.bold ? "bold" : "normal",
                                                fontStyle: valAxis.italic ? "italic" : "normal",
                                                textDecoration: valAxis.underline ? "underline" : undefined,
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
                                fontSize: `${axis2.titleFontSize}pt`,
                                fontFamily: axis2.titleFontFamily,
                                fontWeight: axis2.titleBold ? "bold" : "normal",
                                fontStyle: axis2.titleItalic ? "italic" : "normal",
                                textDecoration: axis2.titleUnderline ? "underline" : undefined,
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
                        axis2.ticks.map((tick, i) => (
                            <text
                                key={`y2-${i}`}
                                x={labelX}
                                y={marginTop + plotHeight * (1 - tick.ratio) + axis2TickFontPx * 0.35}
                                className="y2-tick-label"
                                textAnchor={onRight ? "start" : "end"}
                                style={{
                                    fontSize: `${axis2.fontSize}pt`,
                                    fontFamily: axis2.fontFamily,
                                    fontWeight: axis2.bold ? "bold" : "normal",
                                    fontStyle: axis2.italic ? "italic" : "normal",
                                    textDecoration: axis2.underline ? "underline" : undefined,
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
                        role="img"
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
                            style={{
                                fontSize: `${catAxis.titleFontSize}pt`,
                                fontFamily: catAxis.titleFontFamily,
                                fontWeight: catAxis.titleBold ? "bold" : "normal",
                                fontStyle: catAxis.titleItalic ? "italic" : "normal",
                                textDecoration: catAxis.titleUnderline ? "underline" : undefined,
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
                role="img"
                aria-label="単位つき棒グラフ"
            >
                {/* 軸上バッジ (plotTopRight 配置の場合) */}
                {renderPlotTopRightBadge()}

                {/* プロット描画 (オフセット marginLeft) */}
                {renderPlot(marginLeft)}

                {/* Y軸（目盛・線・タイトル・軸上バッジ）と第 2 Y 軸 */}
                {renderYAxis()}
                {renderY2Axis()}

                {/* X軸タイトル */}
                {catAxis.titleShow && catAxis.titleText && (
                    <text
                        x={marginLeft + plotWidth / 2}
                        y={height - 6}
                        className="x-axis-title"
                        textAnchor="middle"
                        style={{
                            fontSize: `${catAxis.titleFontSize}pt`,
                            fontFamily: catAxis.titleFontFamily,
                            fontWeight: catAxis.titleBold ? "bold" : "normal",
                            fontStyle: catAxis.titleItalic ? "italic" : "normal",
                            textDecoration: catAxis.titleUnderline ? "underline" : undefined,
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
        const labelOf = (g: CategoryGroup) => (stackedLevels ? g.levels[g.levels.length - 1] : g.category);
        const widestLabel = catAxis.show ? Math.max(0, ...groups.map((g) => measureTextWidth(labelOf(g), catFont))) : 0;
        // 標準と同じく、凡例を含むビジュアル全体の幅に対する割合
        const maxLabelWidth = Math.max(16, viewport.width * (catAxis.maxHeight / 100));
        const lowestLabelWidth = catAxis.show ? Math.min(widestLabel, maxLabelWidth) + 8 : 0;
        // 上のレベルの列。棒に近い（下の）レベルから置き、ラベルと合わせてグラフの幅の半分までにする（上のレベルから落とす）
        const levelColumns: Array<{ level: number; width: number }> = [];
        if (stackedLevels) {
            let used = lowestLabelWidth;
            for (let level = catAxis.levelCount - 2; level >= 0; level--) {
                const widest = Math.max(0, ...groups.map((g) => measureTextWidth(g.levels[level] ?? "", catFont)));
                const columnWidth = Math.min(widest, maxLabelWidth) + 12;
                if (used + columnWidth > width * 0.5) break;
                levelColumns.push({ level, width: columnWidth });
                used += columnWidth;
            }
        }
        // パレートのランクの帯。いちばん下のレベルのラベルのすぐ左に 1 列取る
        const pareto = viewModel.pareto;
        const rankBandOn = pareto.enabled && pareto.showRankBand && catAxis.show;
        // 幅は、帯を出すランクの「名前 N件」のいちばん長いもの。ラベルの最大幅で頭を打つ（長い名前でプロットを押し出さない。入らなければ名前だけ、それも入らなければ文字を出さない）
        const rankBandWidth = rankBandOn
            ? Math.min(
                maxLabelWidth,
                Math.max(0, ...(["A", "B", "C"] as ParetoRank[]).map((r) => {
                    const count = pareto.ranks.filter((x) => x === r).length;
                    return count > 0 ? measureTextWidth(`${pareto.labels[r]} ${count}件`, catFont) : 0;
                }))
            ) + 16
            : 0;
        const labelAreaWidth = lowestLabelWidth + rankBandWidth + levelColumns.reduce((sum, c) => sum + c.width, 0);
        const catTitleWidth = hasCatTitle ? catTitleFontPx + 8 : 0;

        const marginLeft = 4 + catTitleWidth + labelAreaWidth;
        const marginRight = 16;
        const marginTop = 10 + (valueAtTop ? axisBlockHeight : 0) + (axis2AtTop ? axis2BlockHeight : 0) + (badgeAtTopRight ? badgeFontPx + 6 : 0);
        const marginBottom = 6 + (valueAtTop ? axis2BlockHeight : axisBlockHeight);
        const viewHeight = Math.max(10, height - marginTop - marginBottom);

        // 「最小カテゴリの高さ」を下回るなら縦にスクロールする
        const minCatHeight = Math.max(0, catAxis.minCategoryWidth || 0);
        const neededHeight = minCatHeight > 0 ? Math.ceil(count * minCatHeight) : 0;
        const scrolls = minCatHeight > 0 && neededHeight > viewHeight + 1;
        const SCROLLBAR_WIDTH = scrolls ? 12 : 0;
        const plotHeight = scrolls ? neededHeight : viewHeight;
        const plotWidth = Math.max(10, width - marginLeft - marginRight - SCROLLBAR_WIDTH);

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
            columnsSettings.maxBarWidth
        );
        const thickness = cluster.barWidth;
        const clusterSpan = spanOf(cluster);

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

        const valueStroke = gridLineStroke(gridlines.horizontalStyle, gridlines.horizontalWidth, gridlines.horizontalScaleWithWidth);
        const valueOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));
        const categoryStroke = gridLineStroke(gridlines.verticalStyle, gridlines.verticalWidth, gridlines.verticalScaleWithWidth);
        const categoryOpacity = Math.max(0, Math.min(1, 1 - gridlines.verticalTransparency / 100));

        const dl = viewModel.dataLabels;
        // 横棒のデータラベルの位置。「自動」は集合なら外側の端、積み上げなら中央（積み上げでは外側に置かない）
        const hLabelPosition =
            dl.position === "auto" || (stacked && dl.position === "outsideEnd")
                ? (stacked ? "insideCenter" : "outsideEnd")
                : dl.position;

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

        const renderHBar = (d: DataPoint, cy: number, key: string, xOffset: number, yOffset: number) => {
            if (d.blank) return null;
            const { left: l, width: w } = extentX(d.startRatio, d.valRatio);
            const left = xOffset + l;
            const top = yOffset + cy - thickness / 2;
            // 棒が右へ伸びるか（正の値。範囲の反転なら逆）
            const rightward = (d.value >= 0) !== valAxis.invertRange;
            const r = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, thickness / 2, w)) : 0;
            const barPath = hBarPathOf(left, top, w, thickness, r, rightward);

            const fullyHighlighted =
                d.highlight !== null && Math.abs(d.highlight - d.value) <= Math.abs(d.value) * 1e-9;
            const highlightDimmed = viewModel.hasHighlights && !fullyHighlighted;
            const selectionDimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isPicked(d);
            const dimmed = highlightDimmed || selectionDimmed;
            let highlightPath: string | null = null;
            if (highlightDimmed && d.highlight !== null && d.highlightRatio !== null) {
                const hl = extentX(d.startRatio, d.highlightRatio);
                if (hl.width > 0) {
                    const hlR = d.outermost ? Math.max(0, Math.min(columnsSettings.cornerRadius, thickness / 2, hl.width)) : 0;
                    highlightPath = hBarPathOf(xOffset + hl.left, top, hl.width, thickness, hlR, (d.highlight >= 0) !== valAxis.invertRange);
                }
            }

            const dimFactor = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
            const baseFillOpacity = Math.max(0, Math.min(1, 1 - d.transparency / 100));
            const stroke = d.borderShow ? d.borderColor : "none";
            const strokeWidth = d.borderShow ? d.borderWidth : 0;
            const strokeOpacity = dimFactor * (d.borderShow ? Math.max(0, Math.min(1, 1 - d.borderTransparency / 100)) : 1);
            const seriesName = viewModel.seriesMode ? viewModel.series[d.seriesIndex]?.name : undefined;

            // データラベル（横に並べる。「縦」の方向は横棒では使わない）
            const label = (() => {
                if (!dl.show || !d.labelShow || w <= 0) return null;
                const lines = labelLinesOf(dl, d);
                if (!lines.length) return null;
                const block = labelBlock(lines);
                const textW = block.width;
                const right = left + w;
                let x: number;
                let anchor: "start" | "middle" | "end";
                let inside = true;
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
                const fitsThickness = thickness >= blockHeight + 2;
                // 積み上げでは隣が別の棒なので、値の向きに入りきらないラベルは出さない
                if (inside && stacked && !fitsLength) return null;
                if (inside && !(fitsLength && fitsThickness) && !dl.overflow) return null;
                const autoColor = !inside || dl.backgroundShow ? "#252423" : contrastingText(d.color);
                const midY = top + thickness / 2;
                // 背景は 1 行なら 1.10 までと同じ高さ、複数行ならまとまりの上端から下端
                const firstPx = lines[0].font.size;
                const bgTop = lines.length > 1 ? midY + block.topEdge - 1 : midY - firstPx * 0.6 - 1;
                const bgHeight = lines.length > 1 ? blockHeight + 2 : firstPx * 1.2 + 2;
                const boxX = anchor === "start" ? x - LABEL_PADDING : anchor === "end" ? x - textW - LABEL_PADDING : x - textW / 2 - LABEL_PADDING;
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
                                x={x}
                                y={midY + block.baselines[k]}
                                className={line.kind === "value" ? "data-label" : "data-label-detail"}
                                textAnchor={anchor}
                                style={labelTextStyle(dl, line, d, autoColor)}
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
                    role="graphics-symbol"
                    aria-label={seriesName ? `${d.category} ${seriesName}: ${d.formattedValue}` : `${d.category}: ${d.formattedValue}`}
                >
                    <path
                        d={barPath}
                        className="bar-rect"
                        style={{ fill: d.color, fillOpacity: dimFactor * baseFillOpacity, stroke, strokeWidth, strokeOpacity }}
                        onMouseEnter={(e) => onTooltipShow?.(d, e.clientX, e.clientY)}
                        onMouseMove={(e) => onTooltipMove?.(d, e.clientX, e.clientY)}
                        onMouseLeave={() => onTooltipHide?.()}
                    />
                    {highlightPath && (
                        <path d={highlightPath} className="bar-highlight" pointerEvents="none" style={{ fill: d.color, fillOpacity: baseFillOpacity }} />
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
                            fontSize: `${tl.fontSize}pt`,
                            fontFamily: tl.fontFamily,
                            fontWeight: tl.bold ? "bold" : "normal",
                            fontStyle: tl.italic ? "italic" : "normal",
                            textDecoration: tl.underline ? "underline" : undefined,
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
                shapeRendering: "crispEdges",
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
                    if (boxed) {
                        // 囲み：区切りごとに角の丸い淡い枠。隣の枠とは少し離す
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
                                    fillOpacity: 0.06,
                                    stroke: catAxis.labelColor,
                                    strokeOpacity: 0.3,
                                    strokeWidth: 1,
                                }}
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
                            style={{
                                fontSize: `${catAxis.fontSize}pt`,
                                fontFamily: catAxis.fontFamily,
                                fontWeight: catAxis.bold ? "bold" : "normal",
                                fontStyle: catAxis.italic ? "italic" : "normal",
                                textDecoration: catAxis.underline ? "underline" : undefined,
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
        const renderHPlot = (xOffset: number, yOffset: number, labelRightX: number) => (
            <>
                {gridlines.horizontalShow && (
                    <g className="y-gridlines-group">
                        {viewModel.ticks.map((tick, i) => {
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
                        {[0, ...groups.slice(1).map((_, k) => centerOf(k + 1) - step / 2), plotHeight].map((y, i) => (
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
                {viewModel.ribbons.show && groups.length > 1 && (
                    <g className="ribbons-group">
                        {renderRibbonBands(
                            groups,
                            (i) => [yOffset + centerOf(i) + thickness / 2, yOffset + centerOf(i + 1) - thickness / 2],
                            (d) => {
                                const e = extentX(d.startRatio, d.valRatio);
                                return [xOffset + e.left, xOffset + e.left + e.width];
                            },
                            true
                        )}
                    </g>
                )}
                <g className="bars-group">
                    {groups.map((g, i) => {
                        const cy = centerOf(i);
                        return (
                            <g key={`cat-${i}`} className="category-group">
                                {g.points.map((d, s) =>
                                    renderHBar(d, cy + (stacked ? 0 : cluster.offsets[s] ?? 0), `${i}-${s}`, xOffset, yOffset)
                                )}
                                {renderHTotals(g, cy, xOffset, yOffset)}
                                {catAxis.show && (
                                    <text
                                        x={labelRightX}
                                        y={yOffset + cy + catFontPx * 0.35}
                                        className="x-category-label"
                                        textAnchor="end"
                                        style={{
                                            fontSize: `${catAxis.fontSize}pt`,
                                            fontFamily: catAxis.fontFamily,
                                            fontWeight: catAxis.bold ? "bold" : "normal",
                                            fontStyle: catAxis.italic ? "italic" : "normal",
                                            textDecoration: catAxis.underline ? "underline" : undefined,
                                            fill: catAxis.labelColor,
                                        }}
                                    >
                                        {truncateText(labelOf(g), maxLabelWidth, catFont)}
                                        <title>{g.category}</title>
                                    </text>
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
                                    fontSize: `${catAxis.fontSize}pt`,
                                    fontFamily: catAxis.fontFamily,
                                    fontWeight: catAxis.bold ? "bold" : "normal",
                                    fontStyle: catAxis.italic ? "italic" : "normal",
                                    textDecoration: catAxis.underline ? "underline" : undefined,
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
                    </g>
                )}
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
                        axis2.ticks.map((tick, i) => (
                            <text
                                key={`x2-${i}`}
                                x={marginLeft + plotWidth * tick.ratio}
                                y={tickY}
                                className="y2-tick-label"
                                textAnchor="middle"
                                style={{
                                    fontSize: `${axis2.fontSize}pt`,
                                    fontFamily: axis2.fontFamily,
                                    fontWeight: axis2.bold ? "bold" : "normal",
                                    fontStyle: axis2.italic ? "italic" : "normal",
                                    textDecoration: axis2.underline ? "underline" : undefined,
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
                                fontSize: `${axis2.titleFontSize}pt`,
                                fontFamily: axis2.titleFontFamily,
                                fontWeight: axis2.titleBold ? "bold" : "normal",
                                fontStyle: axis2.titleItalic ? "italic" : "normal",
                                textDecoration: axis2.titleUnderline ? "underline" : undefined,
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
                fontSize: `${valAxis.fontSize}pt`,
                fontFamily: valAxis.fontFamily,
                fontWeight: valAxis.bold ? "bold" : "normal",
                fontStyle: valAxis.italic ? "italic" : "normal",
                textDecoration: valAxis.underline ? "underline" : undefined,
                fill: valAxis.labelColor,
            };
            return (
                <g className="y-axis-container">
                    {valAxis.show && (
                        <g className="y-axis-group">
                            {viewModel.ticks.map((tick, i) => (
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
                                fontSize: `${valAxis.titleFontSize}pt`,
                                fontFamily: valAxis.titleFontFamily,
                                fontWeight: valAxis.titleBold ? "bold" : "normal",
                                fontStyle: valAxis.titleItalic ? "italic" : "normal",
                                textDecoration: valAxis.titleUnderline ? "underline" : undefined,
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
                            y={badgeAtTopRight ? 10 + badgeFontPx * 0.85 : badgeY}
                            className="unit-axis-badge"
                            textAnchor="end"
                            style={{ fontSize: `${viewModel.unitInfo.fontSize}pt`, fill: viewModel.unitInfo.color }}
                        >
                            {badgeText}
                        </text>
                    )}
                    {renderH2Axis()}
                    {hasCatTitle && (
                        <text
                            transform={`translate(${4 + catTitleFontPx * 0.5}, ${marginTop + viewHeight / 2}) rotate(-90)`}
                            textAnchor="middle"
                            dominantBaseline="central"
                            className="x-axis-title"
                            style={{
                                fontSize: `${catAxis.titleFontSize}pt`,
                                fontFamily: catAxis.titleFontFamily,
                                fontWeight: catAxis.titleBold ? "bold" : "normal",
                                fontStyle: catAxis.titleItalic ? "italic" : "normal",
                                textDecoration: catAxis.titleUnderline ? "underline" : undefined,
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
                            role="img"
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
            <svg width={width} height={height} className="unit-bar-svg" role="img" aria-label="単位つき棒グラフ">
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
                    fillOpacity: dim * Math.max(0, Math.min(1, 1 - (style?.transparency ?? 0) / 100)),
                    stroke: style?.borderShow ? style.borderColor : "none",
                    strokeWidth: style?.borderShow ? Math.min(2, style.borderWidth) : 0,
                    strokeOpacity: dim,
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
        const mk = viewModel.markers;
        const dim = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
        const color = line?.color ?? item.color;
        const width = Math.max(1, Math.min(3, line?.width ?? 2));
        const drawLine = !line || line.lineShow || !mk.show;
        const dash =
            line?.lineStyle === "dashed" ? `${width * 2} ${width * 1.5}` : line?.lineStyle === "dotted" ? `${width * 0.1} ${width * 2}` : undefined;
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
                        strokeLinecap={line?.lineStyle === "dotted" ? "round" : "butt"}
                        strokeOpacity={dim}
                    />
                )}
                {mk.show && (
                    <path
                        d={markerPath(mk.shape, cx, item.y, size, horizontal)}
                        className="legend-marker"
                        style={{
                            fill: mk.color || color,
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

    const renderLegend = (layout: LegendLayout) => {
        const r = layout.markerRadius;
        const anyPicked = selectedIds.length > 0 && !viewModel.hasHighlights;
        const textStyle = {
            fontSize: `${lg.fontSize}pt`,
            fontFamily: lg.fontFamily,
            fontStyle: lg.italic ? "italic" : "normal",
            textDecoration: lg.underline ? "underline" : undefined,
            fill: lg.color,
        };
        return (
            <svg
                className="unit-bar-legend"
                width={layout.box.width}
                height={layout.box.height}
                style={{ position: "absolute", left: layout.box.x, top: layout.box.y, overflow: "hidden" }}
                role="list"
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
                {layout.items.map((item) => {
                    const entry = viewModel.legendEntries[item.index];
                    const dimmed = anyPicked && !isEntryPicked(entry);
                    return (
                        <g
                            key={`legend-${item.index}`}
                            className="legend-item"
                            role="listitem"
                            aria-label={item.name}
                            onClick={(e) => {
                                e.stopPropagation();
                                if (entry?.selectionId) onSelect(entry.selectionId, e.ctrlKey || e.metaKey);
                            }}
                        >
                            {entry?.kind === "line" ? renderLegendLineGlyph(entry, item, dimmed, r) : renderLegendBarGlyph(entry, item, dimmed, r)}
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
        const accent = viewModel.series[0]?.color ?? viewModel.categoryAxis.labelColor;
        const textColor = viewModel.categoryAxis.labelColor;
        const font: React.CSSProperties = { fontFamily: viewModel.categoryAxis.fontFamily, fontSize: "9pt" };
        const stop = (e: React.SyntheticEvent) => e.stopPropagation();
        const resetItems = [{ value: "none", displayName: "区切らない" }, ...cumulative.levels];
        return (
            <div
                className="cumulative-toolbar"
                style={{ position: "absolute", top: 0, right: 0, height: TOOLBAR_HEIGHT, display: "flex", alignItems: "center", gap: 6, paddingRight: 4, whiteSpace: "nowrap" }}
                onClick={stop}
                onContextMenu={stop}
            >
                {on && cumulative.levels.length > 0 && (
                    <select
                        className="cumulative-reset"
                        aria-label="累計の区切り"
                        value={cumulative.reset}
                        onChange={(e) => onChangeCumulativeReset?.(e.target.value)}
                        style={{ ...font, width: "auto", maxWidth: 180, flexShrink: 0, color: textColor, height: 22, border: `1px solid ${withAlpha(textColor, 0.4)}`, borderRadius: 11, padding: "0 6px", background: "transparent" }}
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
                    style={{
                        ...font,
                        height: 22,
                        padding: "0 10px 0 8px",
                        borderRadius: 11,
                        border: `1px solid ${on ? accent : withAlpha(textColor, 0.4)}`,
                        background: on ? withAlpha(accent, 0.16) : "transparent",
                        color: on ? accent : textColor,
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                        flexShrink: 0,
                    }}
                >
                    <span
                        aria-hidden="true"
                        style={{ width: 9, height: 9, borderRadius: "50%", background: on ? accent : "transparent", border: on ? "none" : `1px solid ${withAlpha(textColor, 0.5)}` }}
                    />
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
            style={{ width, height: viewport.height, position: "relative" }}
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
        </div>
    );
};

/** 色に透明度を付ける（#RRGGBB のときだけ。ほかの書き方ならそのまま） */
function withAlpha(color: string, alpha: number): string {
    if (!/^#[0-9a-fA-F]{6}$/.test(color)) return color;
    const a = Math.round(Math.max(0, Math.min(1, alpha)) * 255).toString(16).padStart(2, "0");
    return `${color}${a}`;
}
