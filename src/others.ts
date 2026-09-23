"use strict";

/**
 * 上位 N 件＋「その他」。
 *
 * DataView を読む前に、値（「値」のロールの列の合計）の大きい順に上位 N 件のカテゴリを残し、残りを 1 行の「その他」に畳む。
 * 畳んだ DataView をそのまま transform に渡すので、描画・積み上げ・累計・パレートは今の仕組みで動く。
 * 「値」の列は足し、ほかの列（折れ線の値・ツールヒント・ラベルの詳細）は足しようが無いので空白にする。
 * X 軸の階層を展開しているとき（カテゴリの列が 2 つ以上）は、どのレベルで上位を数えるか決まらないので畳まない。
 */
import powerbi from "powerbi-visuals-api";

import DataView = powerbi.DataView;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import PrimitiveValue = powerbi.PrimitiveValue;

export interface CollapsedOthers {
    dataView: DataView | undefined;
    /** 畳んだあとの DataView での「その他」の行番号。畳まなければ -1 */
    otherRow: number;
    /** 「その他」にまとめた元の行番号（元の DataView の並び） */
    mergedRows: number[];
    /**
     * まとめたカテゴリの正の値の合計。パレートの累積比では「その他」をこの値で数える
     * （正と負をそのまま足すと相殺して分母が変わり、「0 以下のカテゴリは数えない」の決まりが崩れる）
     */
    otherPositiveTotal: number;
}

const isMeasure = (column: DataViewValueColumn) => !!column.source?.roles?.measure;

const toNumber = (raw: PrimitiveValue | undefined): number | null => {
    if (raw === null || raw === undefined) return null;
    const n = typeof raw === "number" ? raw : Number(raw);
    return Number.isFinite(n) ? n : null;
};

/** 数の合計。すべて空白なら空白 */
const sumOf = (values: Array<PrimitiveValue | undefined>): number | null => {
    let total: number | null = null;
    for (const raw of values) {
        const n = toNumber(raw);
        if (n !== null) total = (total ?? 0) + n;
    }
    return total;
};

/**
 * 「折れ線の値」の列ごとに、凡例に左右されないメジャーか（どの行でも系列ごとの値が同じ）を、畳む前の DataView で判定する。
 * 畳むと行が減って判定が変わり、残したカテゴリの線の値まで変わるため（transform は畳んだ DataView で判定する）
 */
export function lineIndependenceOf(dataView: DataView | undefined): Map<string, boolean> {
    const result = new Map<string, boolean>();
    const values = dataView?.categorical?.values;
    const rowCount = dataView?.categorical?.categories?.[0]?.values.length ?? 0;
    if (!values) return result;
    const groups: DataViewValueColumnGroup[] = values.grouped?.() ?? [{ values } as unknown as DataViewValueColumnGroup];
    const byKey = new Map<string, DataViewValueColumn[]>();
    for (const group of groups) {
        for (const column of group.values) {
            if (!column.source?.roles?.lineMeasure) continue;
            const key = column.source.queryName ?? column.source.displayName;
            byKey.set(key, [...(byKey.get(key) ?? []), column]);
        }
    }
    byKey.forEach((columns, key) => {
        let independent = true;
        for (let i = 0; i < rowCount && independent; i++) {
            const numbers = columns.map((c) => toNumber(c.values[i])).filter((n): n is number => n !== null);
            independent = numbers.every((n) => n === numbers[0]);
        }
        result.set(key, independent);
    });
    return result;
}

export function collapseOthers(dataView: DataView | undefined, topCount: number, label: string): CollapsedOthers {
    const none: CollapsedOthers = { dataView, otherRow: -1, mergedRows: [], otherPositiveTotal: 0 };
    const categorical = dataView?.categorical;
    const categories = categorical?.categories;
    const values = categorical?.values;
    if (!categorical || !categories || categories.length !== 1 || !values) return none;
    const n = Math.floor(Number.isFinite(topCount) ? topCount : 0);
    const rowCount = categories[0].values.length;
    // 1 行だけを「その他」にしても意味が無いので、2 行以上まとめるときだけ畳む
    if (n <= 0 || rowCount <= n + 1) return none;

    const measures = values.filter(isMeasure);
    const totalAt = (i: number) => sumOf(measures.map((column) => column.values[i])) ?? 0;
    // 値の大きい順に上位 N 件。同じ値なら元の順
    const ranked = Array.from({ length: rowCount }, (_, i) => i).sort((a, b) => totalAt(b) - totalAt(a));
    const keepSet = new Set(ranked.slice(0, n));
    // 残す行は元の並びのまま（並べ替えは transform の値順・パレートに任せる）
    const kept = Array.from({ length: rowCount }, (_, i) => i).filter((i) => keepSet.has(i));
    const mergedRows = Array.from({ length: rowCount }, (_, i) => i).filter((i) => !keepSet.has(i));
    const pick = <T>(list: T[] | undefined, other: T): T[] | undefined => (list ? [...kept.map((i) => list[i]), other] : undefined);

    const category = categories[0];
    const newCategory = {
        ...category,
        values: pick(category.values, label as PrimitiveValue)!,
        // 「その他」の ID は、まとめた最初の行の ID を仮に使う（選択は mergedRows から作り直す。書式の保存先にはしない）
        identity: pick(category.identity, category.identity?.[mergedRows[0]]),
        objects: category.objects ? pick(category.objects, undefined) : undefined,
    };

    const columnMap = new Map<DataViewValueColumn, DataViewValueColumn>();
    const collapse = (column: DataViewValueColumn): DataViewValueColumn => {
        const found = columnMap.get(column);
        if (found) return found;
        const measure = isMeasure(column);
        const next: DataViewValueColumn = {
            ...column,
            values: pick(column.values, measure ? sumOf(mergedRows.map((i) => column.values[i])) : null)!,
            ...(column.highlights
                ? { highlights: pick(column.highlights, measure ? sumOf(mergedRows.map((i) => column.highlights![i])) : null) }
                : {}),
        };
        columnMap.set(column, next);
        return next;
    };
    const newColumns = values.map(collapse);
    const groups: DataViewValueColumnGroup[] = values.grouped?.() ?? [];
    const newGroups = groups.map((group) => ({ ...group, values: group.values.map(collapse) }));
    const newValues = Object.assign(newColumns, {
        grouped: () => newGroups,
        ...(values.source ? { source: values.source } : {}),
        ...(values.identityFields ? { identityFields: values.identityFields } : {}),
    }) as powerbi.DataViewValueColumns;

    return {
        dataView: { ...dataView!, categorical: { ...categorical, categories: [newCategory], values: newValues } },
        otherRow: kept.length,
        mergedRows,
        otherPositiveTotal: mergedRows.reduce((sum, i) => sum + Math.max(0, totalAt(i)), 0),
    };
}
