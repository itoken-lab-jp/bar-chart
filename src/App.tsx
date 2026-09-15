"use strict";

import * as React from "react";
import powerbi from "powerbi-visuals-api";
import IViewport = powerbi.IViewport;
import ISelectionId = powerbi.visuals.ISelectionId;

import { ViewModel, DataPoint } from "./viewModel";
import { VisualFormattingSettingsModel } from "./settings";
import { contrastingText, placeLabel } from "./unitUtils";

/**
 * ラベル文字列を指定ピクセル幅に収まるよう末尾「...」で省略
 */
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

export interface AppProps {
    viewModel: ViewModel;
    viewport: IViewport;
    settings: VisualFormattingSettingsModel;
    onSelect: (id: ISelectionId, multiSelect: boolean) => void;
    onContextMenu: (id: ISelectionId, x: number, y: number) => void;
    /** このビジュアルで選ばれている棒。空なら選択なし */
    selectedIds?: ISelectionId[];
    /** 棒以外の所をクリックした（選択の解除） */
    onClearSelection?: () => void;
    /** 棒にマウスが入った。座標はクライアント座標（ルート基準への変換は visual.ts） */
    onTooltipShow?: (d: DataPoint, clientX: number, clientY: number) => void;
    /** 棒の上でマウスが動いた */
    onTooltipMove?: (d: DataPoint, clientX: number, clientY: number) => void;
    /** 棒からマウスが出た */
    onTooltipHide?: () => void;
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
}) => {
    const [hoveredIndex, setHoveredIndex] = React.useState<number | null>(null);

    if (viewModel.isEmpty) {
        return (
            <div className="unit-bar-landing">
                <p>カテゴリと値をフィールドに配置してください</p>
            </div>
        );
    }

    const { width, height } = viewport;
    const badgeText = viewModel.unitInfo.badgeText;
    const catAxis = viewModel.categoryAxis;
    const valAxis = viewModel.valueAxis;
    const gridlines = viewModel.gridlines;
    const columnsSettings = viewModel.columns;

    // pt -> px 換算比率 (1pt = 4/3 px = 1.3333...px)
    const PT_TO_PX = 4 / 3;

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

    const marginLeft = !isRightAxis
        ? (hasYTitle
            ? leftPadding + valTitleFontSizePx + gapTitleToTicks + axisContentWidth
            : Math.max(16, leftPadding + axisContentWidth))
        : 24;

    const marginRight = isRightAxis
        ? (hasYTitle
            ? axisContentWidth + gapTitleToTicks + valTitleFontSizePx + rightPadding
            : Math.max(24, axisContentWidth + rightPadding))
        : 24;
    const marginTop = badgeText ? 36 : 22;

    // カテゴリ最小幅と横スクロール判定
    const viewWidth = Math.max(10, width - marginLeft - marginRight);
    const count = viewModel.dataPoints.length;
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
    let barWidth = Math.max(2, step * (1 - padRatio));
    if (columnsSettings.maxBarWidth > 0) {
        barWidth = Math.min(barWidth, columnsSettings.maxBarWidth);
    }

    // X軸カテゴリラベルの文字幅・回転判定
    let maxCatChars = 0;
    for (const d of viewModel.dataPoints) {
        if (d.category.length > maxCatChars) {
            maxCatChars = d.category.length;
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
        ...viewModel.dataPoints.slice(1).map((_, k) => centerOf(k + 1) - step / 2),
        plotWidth,
    ];

    // 棒の基準 (0) の Y 座標。線は引かない（標準と同じく 0 の線は横のグリッド線の 1 本として出る）
    const clampedZeroRatio = Math.max(0, Math.min(1, viewModel.zeroRatio));
    const effectiveZeroRatio = valAxis.invertRange ? 1 - clampedZeroRatio : clampedZeroRatio;
    const clampedZeroY = marginTop + plotHeight * (1 - effectiveZeroRatio);

    /** 比率 (0〜1) で表した値から、基準 (0) までの縦の範囲 */
    const verticalExtent = (ratio: number): { top: number; height: number } => {
        const clamped = Math.max(0, Math.min(1, ratio));
        const effective = valAxis.invertRange ? 1 - clamped : clamped;
        const y = marginTop + plotHeight * (1 - effective);
        return { top: Math.min(y, clampedZeroY), height: Math.max(0, Math.abs(y - clampedZeroY)) };
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

                {/* 棒とデータラベルとカテゴリラベル */}
                <g className="bars-group">
                    {viewModel.dataPoints.map((d: DataPoint, i: number) => {
                        const cx = xOffset + centerOf(i);
                        const barLeft = cx - barWidth / 2;

                        const { top: barTop, height: barH } = verticalExtent(d.valRatio);

                        const isHovered = hoveredIndex === i;

                        const isPositive = d.value >= 0;
                        const r = Math.max(0, Math.min(columnsSettings.cornerRadius, barWidth / 2, barH));

                        // 範囲反転時の角丸向き（通常: 正なら上、負なら下。反転時: 正なら下、負なら上）
                        const roundAtTop = valAxis.invertRange ? !isPositive : isPositive;
                        const barPath = barPathOf(barLeft, barTop, barWidth, barH, r, roundAtTop);

                        // ハイライト（他のビジュアルでの選択）: 標準と同じく棒全体を薄く描き、該当分の高さを
                        // 通常の濃さで重ねる。全部が該当する棒はそのまま、該当しない (null) 棒は薄いだけ
                        const fullyHighlighted =
                            d.highlight !== null && Math.abs(d.highlight - d.value) <= Math.abs(d.value) * 1e-9;
                        const highlightDimmed = viewModel.hasHighlights && !fullyHighlighted;
                        // 自分の棒をクリックして選んだとき: 標準と同じく、選んでいない棒を薄くする
                        // （自分で選んだ場合、Power BI は選んだビジュアル自身にはハイライトを送らない）
                        const selectionDimmed =
                            !viewModel.hasHighlights &&
                            selectedIds.length > 0 &&
                            !selectedIds.some((s) => s.equals(d.selectionId));
                        const dimmed = highlightDimmed || selectionDimmed;
                        let highlightPath: string | null = null;
                        if (highlightDimmed && d.highlight !== null && d.highlightRatio !== null) {
                            const hl = verticalExtent(d.highlightRatio);
                            if (hl.height > 0) {
                                const hlPositive = d.highlight >= 0;
                                const hlR = Math.max(0, Math.min(columnsSettings.cornerRadius, barWidth / 2, hl.height));
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

                        return (
                            <g
                                key={d.category + i}
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
                                onMouseEnter={() => setHoveredIndex(i)}
                                onMouseLeave={() => setHoveredIndex(null)}
                                role="graphics-symbol"
                                aria-label={`${d.category}: ${d.formattedValue}`}
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
                                    if (!dl.show || barH <= 0) return null;

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
                                        position: dl.position,
                                        top: barTop,
                                        bottom: barBottom,
                                        barWidth,
                                        value: text,
                                        font: fontSpec,
                                        vertical: isVertical,
                                        overflow: dl.overflow,
                                    });

                                    if (!placed.fitsVertically && !placed.outside) return null;
                                    if (!placed.fitsAlongBar && !dl.overflow) return null;

                                    let textColor = dl.color;
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
                                            key={`dl-${i}`}
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

                                {/* X軸カテゴリラベル (長さに応じて ... 省略し、左端はみ出し時も ... で省略) */}
                                {catAxis.show && (() => {
                                    // 斜めラベルは左下へ伸びるので、左端までの距離も許容長に含める
                                    // （スクロール時の左端はスクロール領域の端 = x 0 でクリップされる）
                                    const labelLeftBound = scrolls ? 0 : CATEGORY_LABEL_LEFT_SAFE_MARGIN;
                                    const maxLenLeft = shouldRotateCat
                                        ? Math.max(12, (cx - labelLeftBound) / Math.sin(Math.PI / 4))
                                        : maxAllowedLabelLen;
                                    const effectiveMaxLen = Math.min(maxAllowedLabelLen, maxLenLeft);
                                    const displayCat = truncateText(d.category, effectiveMaxLen, catFontSizePx);

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
                                                style={{
                                                    fontSize: `${catAxis.fontSize}pt`,
                                                    fontFamily: catAxis.fontFamily,
                                                    fontWeight: catAxis.bold ? "bold" : "normal",
                                                    fontStyle: catAxis.italic ? "italic" : "normal",
                                                    textDecoration: catAxis.underline ? "underline" : undefined,
                                                    fill: catAxis.labelColor,
                                                }}
                                            >
                                                {displayCat}
                                                <title>{d.category}</title>
                                            </text>
                                        );
                                    } else {
                                        // 水平中央揃え
                                        return (
                                            <text
                                                x={cx}
                                                y={marginTop + plotHeight + catFontSizePx + 4}
                                                className="x-category-label"
                                                textAnchor="middle"
                                                style={{
                                                    fontSize: `${catAxis.fontSize}pt`,
                                                    fontFamily: catAxis.fontFamily,
                                                    fontWeight: catAxis.bold ? "bold" : "normal",
                                                    fontStyle: catAxis.italic ? "italic" : "normal",
                                                    textDecoration: catAxis.underline ? "underline" : undefined,
                                                    fill: catAxis.labelColor,
                                                }}
                                            >
                                                {displayCat}
                                                <title>{d.category}</title>
                                            </text>
                                        );
                                    }
                                })()}
                            </g>
                        );
                    })}
                </g>
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
                        {valAxis.titleText}
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
            {scrolls ? (
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
                            >
                                {catAxis.titleText}
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

                    {/* Y軸（目盛・線・タイトル・軸上バッジ） */}
                    {renderYAxis()}

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
                            {catAxis.titleText}
                        </text>
                    )}
                </svg>
            )}
        </div>
    );
};
