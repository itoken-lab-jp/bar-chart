"use strict";

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import PrimitiveValue = powerbi.PrimitiveValue;
import DataViewValueColumn = powerbi.DataViewValueColumn;

import { BLANK_TEXT, toRootCoordinates } from "./shared/tooltip";

/** ツールヒントに出す 1 列。values は DataView の配列をそのまま持ち、行番号で引く */
export interface TooltipColumn {
    displayName: string;
    /** モデルの書式文字列。無ければ列の型から決まる既定 */
    format: string | undefined;
    values: PrimitiveValue[];
}

/** 複数系列のときの、系列 1 本ぶんのツールヒントの中身 */
export interface TooltipSeries {
    /** 凡例のフィールド名。凡例を使わない（値が複数の）ときは null */
    legendName: string | null;
    /** 系列の名前（凡例の値） */
    seriesName: string;
    measure: TooltipColumn;
    /** 累計のときの、累計の前の値 */
    before?: TooltipColumn;
    extras: TooltipColumn[];
}

export interface TooltipSource {
    /** カテゴリのフィールド名 */
    categoryName: string;
    /**
     * 階層。X 軸のフィールドが複数あって展開しているときだけ持つ。
     * names はレベルの名前、texts は行ごとのレベルの表示（上のレベルから）
     */
    categoryLevels?: { names: string[]; texts: string[][] };
    /** 「値」のフィールド */
    measure: TooltipColumn;
    /** 累計のときの、累計の前の値。measure の名前には（累計）が付く */
    before?: TooltipColumn;
    /** 「ツールヒント」に追加されたフィールド */
    extras: TooltipColumn[];
    /** 複数系列のときの系列ごとの中身。無ければ measure と extras を使う（系列 1 本） */
    series?: TooltipSeries[];
}

export function tooltipColumnOf(column: DataViewValueColumn): TooltipColumn {
    return {
        displayName: column.source?.displayName ?? "",
        format: column.source ? valueFormatter.getFormatStringByColumn(column.source) : undefined,
        values: column.values ?? [],
    };
}

export { BLANK_TEXT };

/**
 * 標準の集合縦棒グラフと同じく、表示単位（億・百万など）で丸めず、
 * モデルの書式文字列どおりの全桁で出す。軸・データラベルの単位付き表記とは別物
 */
export function formatTooltipValue(value: PrimitiveValue | undefined, format: string | undefined): string {
    if (value === null || value === undefined) return BLANK_TEXT;
    return valueFormatter.format(value, format);
}

/** カテゴリの行。階層なら標準と同じくレベルごとに 1 行ずつ（上のレベルから） */
export function categoryTooltipRows(source: TooltipSource, rowIndex: number, category: string): VisualTooltipDataItem[] {
    const levels = source.categoryLevels;
    if (levels && levels.names.length > 1) {
        return levels.names.map((name, k) => ({ displayName: name, value: levels.texts[rowIndex]?.[k] ?? "" }));
    }
    return [{ displayName: source.categoryName, value: category }];
}

/** 積み上げのときにツールヒントへ足すもの */
export interface TooltipStack {
    /** 積み上げ: カテゴリの合計（正と負を足した値）。「合計」の行を足す */
    total?: number | null;
    /** 100% 積み上げ: カテゴリの中の割合（-1〜1）。値の後ろにかっこで付ける */
    share?: number | null;
}

/** 100% 積み上げの割合。標準と同じく小数 2 桁（900 (69.23%)） */
export function formatShare(share: number): string {
    return `${(share * 100).toFixed(2)}%`;
}

/** 積み上げの「合計」の行の名前 */
export const TOTAL_TEXT = "合計";

/**
 * 棒 1 本ぶんのツールヒント。標準と同じ並び: カテゴリ → 凡例（あれば）→ 値 → 合計（積み上げ）→ 追加したフィールド。
 * 表示のたびに作る（全カテゴリぶんを update ごとに書式化しない）
 */
export function tooltipItemsOf(
    source: TooltipSource,
    rowIndex: number,
    category: string,
    seriesIndex = 0,
    stack: TooltipStack = {}
): VisualTooltipDataItem[] {
    const series = source.series?.[seriesIndex];
    const measure = series?.measure ?? source.measure;
    const before = series ? series.before : source.before;
    const extras = series?.extras ?? source.extras;
    const measureText = formatTooltipValue(measure.values[rowIndex], measure.format);
    const hasShare = stack.share !== null && stack.share !== undefined && measure.values[rowIndex] !== null && measure.values[rowIndex] !== undefined;
    return [
        ...categoryTooltipRows(source, rowIndex, category),
        ...(series?.legendName ? [{ displayName: series.legendName, value: series.seriesName }] : []),
        { displayName: measure.displayName, value: hasShare ? `${measureText} (${formatShare(stack.share!)})` : measureText },
        ...(before ? [{ displayName: before.displayName, value: formatTooltipValue(before.values[rowIndex], before.format) }] : []),
        ...(stack.total !== null && stack.total !== undefined
            ? [{ displayName: TOTAL_TEXT, value: formatTooltipValue(stack.total, measure.format) }]
            : []),
        ...extras.map((column) => ({
            displayName: column.displayName,
            value: formatTooltipValue(column.values[rowIndex], column.format),
        })),
    ];
}

export { toRootCoordinates };
