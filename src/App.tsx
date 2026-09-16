"use strict";

import * as React from "react";
import powerbi from "powerbi-visuals-api";
import IViewport = powerbi.IViewport;
import ISelectionId = powerbi.visuals.ISelectionId;

import { ViewModel, DataPoint, CategoryGroup, LineSeriesInfo, LegendItemInfo } from "./viewModel";
import { VisualFormattingSettingsModel } from "./settings";
import { contrastingText, placeLabel, measureTextWidth, LABEL_PADDING } from "./unitUtils";
import { clusterLayout } from "./layout";
import { layoutLegend, LegendLayout, LEGEND_MARKER_GAP } from "./legend";

/**
 * ラベル文字列を指定ピクセル幅に収まるよう末尾「...」で省略
 */
/** 長ければ末尾を省略した文字（軸のタイトル用）。省略したときは、マウスを乗せると全体が出る */
function fittedText(text: string, maxWidthPx: number, fontSizePx: number): React.ReactNode {
    const shown = truncateText(text, maxWidthPx, fontSizePx);
    return shown === text ? text : (
        <>
            {shown}
            <title>{text}</title>
        </>
    );
}

function truncateText(text: string, maxWidthPx: number, fontSizePx: number): string {
    if (!text) return "";
    const getWidth = (s: string) => {
        let w = 0;
        for (let i = 0; i < s.length; i++) {
            w += s.charCodeAt(i) > 255 ? fontSizePx * 1.05 : fontSizePx * 0.6;
        }
        return w;
    };

    if (getWidth(text) <= maxWidthPx) {
        return text;
    }

    const ellipsis = "...";
    const ellipsisW = getWidth(ellipsis);
    if (maxWidthPx <= ellipsisW) {
        return text.charAt(0) + ellipsis;
    }

    let result = "";
    for (let i = 0; i < text.length; i++) {
        const next = text.slice(0, i + 1);
        if (getWidth(next) + ellipsisW > maxWidthPx) {
            break;
        }
        result = next;
    }
    return (result || text.charAt(0)) + ellipsis;
}

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

/**
 * グリッド線の線種。標準に合わせ、点線は細かい点（長さ 1 の線に丸い端）、破線は 4px 刻み。
 * 点線は端を丸めるので crispEdges を掛けない（点がつぶれる）
 */
export function gridLineStroke(style: string): {
    dashArray?: string;
    lineCap?: "round";
    shapeRendering: "crispEdges" | "auto";
} {
    if (style === "dotted") return { dashArray: "1 3", lineCap: "round", shapeRendering: "auto" };
    if (style === "dashed") return { dashArray: "4 4", shapeRendering: "crispEdges" };
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
}

const NO_SELECTION: ISelectionId[] = [];

export const App: React.FC<AppProps> = ({
    viewModel,
    viewport,
    onSelect,
    onContextMenu,
    selectedIds = NO_SELECTION,
    onClearSelection,
    onTooltipShow,
    onTooltipMove,
    onTooltipHide,
    onLineTooltipShow,
    onLineTooltipMove,
    onRibbonTooltipShow,
    onRibbonTooltipMove,
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

    // 凡例（複数系列のときだけ）。グラフの外側に置き、その分だけグラフの領域を縮める
    const legend: LegendLayout | null = lg.show
        ? layoutLegend({
            entries: viewModel.legendEntries.map((e) => ({ name: e.name, color: e.color })),
            title: lg.title,
            font: { family: lg.fontFamily, size: legendFontPx, bold: lg.bold, italic: lg.italic, underline: lg.underline },
            position: lg.position,
            width: viewport.width,
            height: viewport.height,
        })
        : null;

    /**
     * 自分で選んだ棒か。凡例のクリックで系列ごと選んだときは、その系列の棒をすべて含む
     * （自分で選んだ場合、Power BI は選んだビジュアル自身にはハイライトを送らない）
     */
    const isPicked = (d: DataPoint): boolean => {
        const seriesId = viewModel.series[d.seriesIndex]?.selectionId;
        return selectedIds.some((s) => s.equals(d.selectionId) || (seriesId ? s.equals(seriesId) : false));
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
        // 積み上げのデータラベルは棒の中に置く。外側の指定（自動・外側の上）は中央に読み替える
        const labelPosition =
            stacked && (viewModel.dataLabels.position === "auto" || viewModel.dataLabels.position === "outsideEnd")
                ? "insideCenter"
                : viewModel.dataLabels.position;

        // X軸カテゴリラベルの文字幅・回転判定
        let maxCatChars = 0;
        for (const g of groups) {
            if (g.category.length > maxCatChars) {
                maxCatChars = g.category.length;
            }
        }
        const catFontSizePx = catAxis.fontSize * PT_TO_PX;
        const approxCatWidth = maxCatChars * (catFontSizePx * 0.85);

        // バンド幅に対してラベルが収まらなければ斜め-45度回転
        const shouldRotateCat = catAxis.show && (approxCatWidth > step * 0.85);

        // X軸タイトル高さ (タイトル領域は高さ最大値の判定外で独立確保し、重なりを防止)
        const hasCatTitle = catAxis.titleShow && Boolean(catAxis.titleText);
        const catTitleFontSizePx = catAxis.titleFontSize * PT_TO_PX;
        const catTitleHeight = hasCatTitle ? catTitleFontSizePx + 10 : 0;

        // X軸ラベル領域 (高さの最大値 % はここだけに効く)
        // maxHeight は 0〜100% (既定 25%)。範囲は viewModel でクランプ済み
        const maxLabelAreaHeight = Math.max(16, height * (catAxis.maxHeight / 100));

        // 必要とされるラベル高さ
        const desiredLabelHeight = shouldRotateCat
            ? approxCatWidth * Math.sin(Math.PI / 4) + 12
            : catFontSizePx + 10;

        const labelAreaHeight = catAxis.show
            ? Math.max(14, Math.min(desiredLabelHeight, maxLabelAreaHeight))
            : 0;

        // ラベル表示の最大許容長 (px) - これを超える場合は末尾を「...」に省略
        const maxAllowedLabelLen = shouldRotateCat
            ? Math.max(12, (labelAreaHeight - 8) / Math.sin(Math.PI / 4))
            : Math.max(12, step * 0.92);

        // 全体の下部マージン = ラベル領域 + タイトル領域 + スクロールバー高さ + 余白
        const marginBottom = Math.max(20, labelAreaHeight + catTitleHeight + SCROLLBAR_HEIGHT + 6);
        const plotHeight = Math.max(10, height - marginTop - marginBottom);

        // 横グリッド線 (Y軸目盛線) の線種と透過性
        const hStroke = gridLineStroke(gridlines.horizontalStyle);
        const hOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));

        // 縦グリッド線 (X軸目盛線) の線種と透過性
        const vStroke = gridLineStroke(gridlines.verticalStyle);
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
                        onSelect(d.selectionId, e.ctrlKey || e.metaKey);
                    }}
                    onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onContextMenu(d.selectionId, e.clientX, e.clientY);
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

                        const text = d.dataLabelText;
                        const isVertical = dl.orientation === "vertical";
                        const fontSpec = {
                            family: dl.fontFamily,
                            size: dl.fontSize * PT_TO_PX,
                            bold: dl.bold,
                            italic: dl.italic,
                            underline: false,
                        };

                        const barBottom = barTop + barH;
                        const placed = placeLabel({
                            position: labelPosition,
                            top: barTop,
                            bottom: barBottom,
                            barWidth,
                            value: text,
                            font: fontSpec,
                            vertical: isVertical,
                            // 積み上げでは外側が上の棒なので、入りきらないラベルを外へ逃がさず出さない
                            overflow: stacked ? false : dl.overflow,
                        });

                        if (!placed.fitsVertically && !placed.outside) return null;
                        if (!placed.fitsAlongBar && !dl.overflow) return null;

                        let textColor = d.labelColor;
                        if (!textColor) {
                            if (placed.outside || dl.backgroundShow) {
                                textColor = "#252423";
                            } else {
                                textColor = contrastingText(d.color);
                            }
                        }

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
                                <text
                                    x={0}
                                    y={dl.fontSize * PT_TO_PX * 0.35}
                                    className="data-label"
                                    textAnchor="middle"
                                    style={{
                                        fill: textColor,
                                        fontSize: `${dl.fontSize}pt`,
                                        fontFamily: dl.fontFamily,
                                        fontWeight: dl.bold ? "bold" : "normal",
                                        fontStyle: dl.italic ? "italic" : "normal",
                                    }}
                                >
                                    {text}
                                </text>
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

        /** X軸カテゴリラベル (長さに応じて ... 省略し、左端はみ出し時も ... で省略)。cx はカテゴリの中心 */
        const renderCategoryLabel = (category: string, cx: number) => {
            // 斜めラベルは左下へ伸びるので、左端までの距離も許容長に含める
            // （スクロール時の左端はスクロール領域の端 = x 0 でクリップされる）
            const labelLeftBound = scrolls ? 0 : CATEGORY_LABEL_LEFT_SAFE_MARGIN;
            const maxLenLeft = shouldRotateCat
                ? Math.max(12, (cx - labelLeftBound) / Math.sin(Math.PI / 4))
                : maxAllowedLabelLen;
            const effectiveMaxLen = Math.min(maxAllowedLabelLen, maxLenLeft);
            const displayCat = truncateText(category, effectiveMaxLen, catFontSizePx);
            const labelStyle = {
                fontSize: `${catAxis.fontSize}pt`,
                fontFamily: catAxis.fontFamily,
                fontWeight: catAxis.bold ? "bold" : "normal",
                fontStyle: catAxis.italic ? "italic" : "normal",
                textDecoration: catAxis.underline ? "underline" : undefined,
                fill: catAxis.labelColor,
            };

            if (shouldRotateCat) {
                // 斜め45度回転 (標準準拠: 棒の直下から左下に伸び、Y軸ラベルの下に被さる)
                const labelTop = marginTop + plotHeight + 10;
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

        /**
         * 折れ線 1 本。点はカテゴリの中心に置き、値の無いカテゴリでは線を切る。
         * 点の上にマウスの当たり判定の円を置き、ツールヒントと選択を受ける
         */
        const renderLine = (line: LineSeriesInfo, j: number, xOffset: number) => {
            // 第 2 Y 軸には範囲の反転が無いので、そのまま下から上へ
            const yOf = (ratio: number) =>
                axis2On ? marginTop + plotHeight * (1 - Math.max(0, Math.min(1, ratio))) : yOfRatio(ratio);
            const pts = line.points.map((p, i) => ({ p, i, x: xOffset + centerOf(i), y: p.ratio === null ? null : yOf(p.ratio) }));
            let path = "";
            let pen = false;
            for (const pt of pts) {
                if (pt.y === null) {
                    pen = false;
                    continue;
                }
                path += `${pen ? "L" : "M"} ${pt.x},${pt.y} `;
                pen = true;
            }
            const dimmed = !viewModel.hasHighlights && selectedIds.length > 0 && !isLinePicked(j);
            const opacity = dimmed ? HIGHLIGHT_DIM_FACTOR : 1;
            const dash =
                line.lineStyle === "dashed"
                    ? `${line.width * 3} ${line.width * 2}`
                    : line.lineStyle === "dotted"
                        ? `${line.width * 0.1} ${line.width * 2}`
                        : undefined;
            const markerR = Math.max(1, viewModel.markers.size / 2 + 0.5);
            const hitR = Math.max(6, line.width * 2);
            return (
                <g key={`line-${j}`} className="line-series">
                    <path
                        d={path.trim()}
                        className="line-path"
                        fill="none"
                        stroke={line.color}
                        strokeWidth={line.width}
                        strokeDasharray={dash}
                        strokeLinejoin="round"
                        strokeLinecap={line.lineStyle === "dotted" ? "round" : "butt"}
                        strokeOpacity={opacity}
                        pointerEvents="none"
                    />
                    {pts.map((pt) =>
                        pt.y === null ? null : (
                            <g key={`pt-${j}-${pt.i}`}>
                                {viewModel.markers.show && (
                                    <circle
                                        cx={pt.x}
                                        cy={pt.y}
                                        r={markerR}
                                        className="line-marker"
                                        style={{ fill: line.color, fillOpacity: opacity }}
                                        pointerEvents="none"
                                    />
                                )}
                                <circle
                                    cx={pt.x}
                                    cy={pt.y}
                                    r={hitR}
                                    className="line-hit"
                                    fill="transparent"
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        onSelect(pt.p.selectionId, e.ctrlKey || e.metaKey);
                                    }}
                                    onContextMenu={(e) => {
                                        e.preventDefault();
                                        e.stopPropagation();
                                        onContextMenu(pt.p.selectionId, e.clientX, e.clientY);
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
         * リボンの帯（#76）。同じ系列の棒を、隣のカテゴリの棒と S 字の帯でつなぐ（棒の後ろに描く）。
         * 罫線は標準と同じく帯の上と下の縁だけに引く
         */
        const renderRibbons = (xOffset: number) => {
            const rb = viewModel.ribbons;
            const half = barWidth / 2;
            const nodes: React.ReactNode[] = [];
            for (let i = 0; i + 1 < groups.length; i++) {
                const left = xOffset + centerOf(i) + half;
                const right = xOffset + centerOf(i + 1) - half;
                const pad = (right - left) * (rb.spacing / 100);
                const xa = left + pad;
                const xb = right - pad;
                if (xb <= xa) continue;
                const xm = (xa + xb) / 2;
                groups[i].points.forEach((a, s) => {
                    const b = groups[i + 1].points[s];
                    if (!b || a.blank || b.blank) return;
                    const ea = extentBetween(a.startRatio, a.valRatio);
                    const eb = extentBetween(b.startRatio, b.valRatio);
                    if (ea.height <= 0 && eb.height <= 0) return;
                    const aTop = ea.top;
                    const aBottom = ea.top + ea.height;
                    const bTop = eb.top;
                    const bBottom = eb.top + eb.height;
                    const topCurve = `M ${xa},${aTop} C ${xm},${aTop} ${xm},${bTop} ${xb},${bTop}`;
                    const bottomCurve = `M ${xb},${bBottom} C ${xm},${bBottom} ${xm},${aBottom} ${xa},${aBottom}`;
                    const band = `${topCurve} L ${xb},${bBottom} C ${xm},${bBottom} ${xm},${aBottom} ${xa},${aBottom} Z`;
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
                                [topCurve, bottomCurve].map((curve, k) => (
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
                                    {catAxis.show && renderCategoryLabel(g.category, cx)}
                                </g>
                            );
                        })}
                    </g>

                    {/* 折れ線（複合）。棒の上に重ねる */}
                    {viewModel.lines.length > 0 && (
                        <g className="lines-group">{viewModel.lines.map((line, j) => renderLine(line, j, xOffset))}</g>
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
                            {fittedText(valAxis.titleText, plotHeight, valTitleFontSizePx)}
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
                            {fittedText(axis2.titleText, plotHeight, axis2TitleFontPx)}
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
                            {truncateText(catAxis.titleText, viewWidth, catTitleFontSizePx)}
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
                        {fittedText(catAxis.titleText, plotWidth, catTitleFontSizePx)}
                    </text>
                )}
            </svg>
        );
    };

    /**
     * 横棒（#75）。カテゴリの軸を左、値の軸を下に置き、先頭のカテゴリを上にする（標準と同じ）。
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

        // カテゴリの軸（左）。ラベルは「最大幅 (%)」まで、長ければ末尾を省略
        const catFontPx = catAxis.fontSize * PT_TO_PX;
        const catTitleFontPx = catAxis.titleFontSize * PT_TO_PX;
        const hasCatTitle = catAxis.titleShow && Boolean(catAxis.titleText);
        const catFont = { family: catAxis.fontFamily, size: catFontPx, bold: catAxis.bold, italic: catAxis.italic, underline: catAxis.underline };
        const widestLabel = catAxis.show ? Math.max(0, ...groups.map((g) => measureTextWidth(g.category, catFont))) : 0;
        const maxLabelWidth = Math.max(16, width * (catAxis.maxHeight / 100));
        const labelAreaWidth = catAxis.show ? Math.min(widestLabel, maxLabelWidth) + 8 : 0;
        const catTitleWidth = hasCatTitle ? catTitleFontPx + 8 : 0;

        const marginLeft = 4 + catTitleWidth + labelAreaWidth;
        const marginRight = 16;
        const marginTop = 10 + (valueAtTop ? axisBlockHeight : 0) + (badgeAtTopRight ? badgeFontPx + 6 : 0);
        const marginBottom = 6 + (valueAtTop ? 0 : axisBlockHeight);
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

        const valueStroke = gridLineStroke(gridlines.horizontalStyle);
        const valueOpacity = Math.max(0, Math.min(1, 1 - gridlines.horizontalTransparency / 100));
        const categoryStroke = gridLineStroke(gridlines.verticalStyle);
        const categoryOpacity = Math.max(0, Math.min(1, 1 - gridlines.verticalTransparency / 100));

        const dl = viewModel.dataLabels;
        const dlFontPx = dl.fontSize * PT_TO_PX;
        const dlFont = { family: dl.fontFamily, size: dlFontPx, bold: dl.bold, italic: dl.italic, underline: false };
        // 横棒のデータラベルの位置。「自動」は集合なら外側の端、積み上げなら中央（積み上げでは外側に置かない）
        const hLabelPosition =
            dl.position === "auto" || (stacked && dl.position === "outsideEnd")
                ? (stacked ? "insideCenter" : "outsideEnd")
                : dl.position;

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
                const text = d.dataLabelText;
                const textW = measureTextWidth(text, dlFont);
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
                const fitsThickness = thickness >= dlFontPx + 2;
                // 積み上げでは隣が別の棒なので、値の向きに入りきらないラベルは出さない
                if (inside && stacked && !fitsLength) return null;
                if (inside && !(fitsLength && fitsThickness) && !dl.overflow) return null;
                let textColor = d.labelColor;
                if (!textColor) textColor = !inside || dl.backgroundShow ? "#252423" : contrastingText(d.color);
                const midY = top + thickness / 2;
                const boxX = anchor === "start" ? x - LABEL_PADDING : anchor === "end" ? x - textW - LABEL_PADDING : x - textW / 2 - LABEL_PADDING;
                return (
                    <g className="data-label-group" pointerEvents="none">
                        {dl.backgroundShow && (
                            <rect
                                x={boxX}
                                y={midY - dlFontPx * 0.6 - 1}
                                width={textW + LABEL_PADDING * 2}
                                height={dlFontPx * 1.2 + 2}
                                rx={2}
                                fill={dl.backgroundColor}
                                fillOpacity={(100 - dl.backgroundTransparency) / 100}
                            />
                        )}
                        <text
                            x={x}
                            y={midY + dlFontPx * 0.35}
                            className="data-label"
                            textAnchor={anchor}
                            style={{
                                fill: textColor,
                                fontSize: `${dl.fontSize}pt`,
                                fontFamily: dl.fontFamily,
                                fontWeight: dl.bold ? "bold" : "normal",
                                fontStyle: dl.italic ? "italic" : "normal",
                            }}
                        >
                            {text}
                        </text>
                    </g>
                );
            })();

            return (
                <g
                    key={key}
                    className={`bar-item ${hoveredKey === key ? "hovered" : ""}`}
                    onClick={(e) => {
                        e.stopPropagation();
                        onSelect(d.selectionId, e.ctrlKey || e.metaKey);
                    }}
                    onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        onContextMenu(d.selectionId, e.clientX, e.clientY);
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
                                        {truncateText(g.category, maxLabelWidth, catFontPx)}
                                        <title>{g.category}</title>
                                    </text>
                                )}
                            </g>
                        );
                    })}
                </g>
            </>
        );

        /** 値の軸（目盛り・タイトル・単位ラベル）とカテゴリの軸のタイトル。スクロールしても動かない */
        const plotRightX = marginLeft + plotWidth;
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
                            {fittedText(valAxis.titleText, plotWidth, valTitleFontPx)}
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
                            {fittedText(catAxis.titleText, viewHeight, catTitleFontPx)}
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
                            <circle
                                cx={item.x + r}
                                cy={item.y}
                                r={r}
                                className="legend-marker"
                                style={{ fill: item.color, fillOpacity: dimmed ? HIGHLIGHT_DIM_FACTOR : 1 }}
                            />
                            <text
                                x={item.x + r * 2 + LEGEND_MARKER_GAP}
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

    const { width, height } = viewport;
    const chartWidth = legend ? Math.max(10, width - legend.reserve.left - legend.reserve.right) : width;
    const chartHeight = legend ? Math.max(10, height - legend.reserve.top - legend.reserve.bottom) : height;

    return (
        <div
            className="unit-bar-container"
            style={{ width, height, position: "relative" }}
            // 棒のクリック・右クリックは棒側で止めるので、ここに来るのは棒以外の所だけ。
            // （以前は target === currentTarget で判定していたが、SVG が全面を覆うため常に偽だった）
            onClick={() => onClearSelection?.()}
            onContextMenu={(e) => {
                e.preventDefault();
                onContextMenu(null as unknown as ISelectionId, e.clientX, e.clientY);
            }}
        >
            {legend ? (
                <>
                    {/* 凡例を除いた領域にグラフを描く。中の絶対配置はこの領域が基準になる */}
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
            )}
        </div>
    );
};
