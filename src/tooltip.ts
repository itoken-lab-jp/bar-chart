"use strict";

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import PrimitiveValue = powerbi.PrimitiveValue;
import DataViewValueColumn = powerbi.DataViewValueColumn;

/** ツールヒントに出す 1 列。values は DataView の配列をそのまま持ち、行番号で引く */
export interface TooltipColumn {
    displayName: string;
    /** モデルの書式文字列。無ければ列の型から決まる既定 */
    format: string | undefined;
    values: PrimitiveValue[];
}

export interface TooltipSource {
    /** カテゴリのフィールド名 */
    categoryName: string;
    /** 「値」のフィールド */
    measure: TooltipColumn;
    /** 「ツールヒント」に追加されたフィールド */
    extras: TooltipColumn[];
}

export function tooltipColumnOf(column: DataViewValueColumn): TooltipColumn {
    return {
        displayName: column.source?.displayName ?? "",
        format: column.source ? valueFormatter.getFormatStringByColumn(column.source) : undefined,
        values: column.values ?? [],
    };
}

/**
 * 空白の表記。valueFormatter の既定は英語の "(Blank)" なので、日本語の Power BI の表記に合わせる
 */
export const BLANK_TEXT = "(空白)";

/**
 * 標準の集合縦棒グラフと同じく、表示単位（億・百万など）で丸めず、
 * モデルの書式文字列どおりの全桁で出す。軸・データラベルの単位付き表記とは別物
 */
export function formatTooltipValue(value: PrimitiveValue | undefined, format: string | undefined): string {
    if (value === null || value === undefined) return BLANK_TEXT;
    return valueFormatter.format(value, format);
}

/**
 * 棒 1 本ぶんのツールヒント。標準と同じ並び: カテゴリ → 値 → 追加したフィールド。
 * 表示のたびに作る（全カテゴリぶんを update ごとに書式化しない）
 */
export function tooltipItemsOf(source: TooltipSource, rowIndex: number, category: string): VisualTooltipDataItem[] {
    return [
        { displayName: source.categoryName, value: category },
        ...[source.measure, ...source.extras].map((column) => ({
            displayName: column.displayName,
            value: formatTooltipValue(column.values[rowIndex], column.format),
        })),
    ];
}

/**
 * クライアント座標を、ビジュアルのルート要素の内側の座標に直す。
 * tooltipService はルート基準の座標を受け取る（powerbi-visuals-utils-tooltiputils と同じ計算）
 */
export function toRootCoordinates(
    clientX: number,
    clientY: number,
    root: Pick<HTMLElement, "getBoundingClientRect" | "clientLeft" | "clientTop">
): [number, number] {
    const rect = root.getBoundingClientRect();
    return [clientX - rect.left - root.clientLeft, clientY - rect.top - root.clientTop];
}
