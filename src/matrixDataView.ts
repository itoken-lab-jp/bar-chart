"use strict";

/**
 * matrix で受けたデータを、categorical と同じ形に詰め替える。
 *
 * 凡例で分けると、categorical では値が凡例の値ごとにしか届かず、率のメジャー（利益 ÷ 売上）を折れ線に入れると
 * 凡例ごとの率を足した値になる。matrix で受けて列の小計を頼むと、Power BI がカテゴリ全体で計算し直した値が届く。
 * 描画・並べ替え・累計などは categorical の形のまま読むので、ここで同じ形に詰め替え、折れ線の値だけ小計の値にする。
 *
 * - 行（カテゴリ）：行の木の葉を 1 行にする。階層は上の段から 1 列ずつ。選択 ID の元になる行の節点は、カテゴリの identity の位置に入れる
 *   （「その他」にまとめたときも、行と一緒に並べ替わる）
 * - 列（凡例）：凡例の値ごとに 1 グループ。グループには列の節点を持たせる
 * - 「折れ線の値」：凡例があれば、列の小計の値を 1 列にする（凡例ごとの値は使わない）。小計が届かなければ凡例ごとの値のまま
 */
import powerbi from "powerbi-visuals-api";

import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewHierarchyLevel = powerbi.DataViewHierarchyLevel;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import DataViewValueColumns = powerbi.DataViewValueColumns;
import CustomVisualOpaqueIdentity = powerbi.visuals.CustomVisualOpaqueIdentity;
import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
import PrimitiveValue = powerbi.PrimitiveValue;

/** 詰め替えた categorical に付ける、matrix の段（選択 ID を作るのに使う） */
const MATRIX_LEVELS = Symbol("matrixLevels");
/** 詰め替えたグループに付ける、列の節点 */
const MATRIX_NODE = Symbol("matrixNode");
/**
 * 行の節点から、上の段からの道。行の葉はその段の値（四半期 = Q2）しか持たず、葉だけで選択 ID を作ると
 * 別の年の Q2 まで選ばれる（2026-10-03、Desktop）。上の段から順に withMatrixNode を重ねる
 */
const ROW_PATHS = new WeakMap<DataViewMatrixNode, DataViewMatrixNode[]>();

interface MatrixLevels {
    rows: DataViewHierarchyLevel[];
    columns: DataViewHierarchyLevel[];
}
type MatrixCategorical = DataViewCategorical & { [MATRIX_LEVELS]?: MatrixLevels };
type MatrixGroup = DataViewValueColumnGroup & { [MATRIX_NODE]?: DataViewMatrixNode };

/** 凡例をまとめた全体の値（列の小計）で受ける役割 */
const WHOLE_ROLES = ["lineMeasure", "categoryLine", "valueLine"];

/** 列の木の葉。凡例の値（無ければ null）、小計か、どのメジャーか */
interface ColumnLeaf {
    series: DataViewMatrixNode | null;
    subtotal: boolean;
    measure: number;
}

function columnLeavesOf(root: DataViewMatrixNode | undefined, hasSeries: boolean): ColumnLeaf[] {
    const leaves: ColumnLeaf[] = [];
    for (const top of root?.children ?? []) {
        if (!hasSeries) {
            leaves.push({ series: null, subtotal: false, measure: top.levelSourceIndex ?? 0 });
            continue;
        }
        const measures = top.children?.length ? top.children : [top];
        for (const measure of measures) {
            leaves.push({ series: top.isSubtotal ? null : top, subtotal: !!top.isSubtotal, measure: measure.levelSourceIndex ?? 0 });
        }
    }
    return leaves;
}

/** 行の木の葉までの道（上の段から順）。小計の行は飛ばす */
function rowPathsOf(root: DataViewMatrixNode | undefined, depth: number): DataViewMatrixNode[][] {
    const paths: DataViewMatrixNode[][] = [];
    const walk = (node: DataViewMatrixNode, path: DataViewMatrixNode[]) => {
        for (const child of node.children ?? []) {
            if (child.isSubtotal) continue;
            const next = [...path, child];
            if (next.length >= depth || !child.children?.length) paths.push(next);
            else walk(child, next);
        }
    };
    if (root) walk(root, []);
    return paths;
}

const isSeriesLevel = (level: DataViewHierarchyLevel | undefined) => !!level?.sources?.some((source) => source.roles?.series);

/** capabilities の行（カテゴリ）と列（凡例）の上限。届いた数がこれに達したら、切られたとみなす */
export const ROW_LIMIT = 30000;
export const SERIES_LIMIT = 2000;

/**
 * matrix で届いた行（カテゴリ。階層なら葉、比較の列があればカテゴリ × 比較の列の組）と、凡例の値の数。
 * 行を window（続きを読む受け方）にすると、凡例は 60 で切られ、列の小計も届かなくなった（2026-10-07、Desktop）。
 * そこで行も top で受け、上限に達したかをこの数で見る（Power BI は続きがあることを知らせない）
 */
export function receivedCounts(dataView: DataView | undefined): { rows: number; series: number } {
    const matrix = dataView?.matrix;
    if (!matrix) return { rows: dataView?.categorical?.categories?.[0]?.values?.length ?? 0, series: dataView?.categorical?.values?.grouped?.().length ?? 0 };
    const rows = rowPathsOf(matrix.rows?.root, matrix.rows?.levels?.length ?? 0).length;
    const hasSeries = isSeriesLevel(matrix.columns?.levels?.[0]);
    const series = hasSeries ? (matrix.columns?.root?.children ?? []).filter((node) => !node.isSubtotal).length : 0;
    return { rows, series };
}

/**
 * matrix を categorical の形に詰め替えた DataView。categorical で届いたときは、そのまま返す
 */
export function categoricalOf(dataView: DataView | undefined): DataView | undefined {
    const matrix = dataView?.matrix;
    if (!dataView || dataView.categorical || !matrix) return dataView;
    const rowLevels = matrix.rows?.levels ?? [];
    const columnLevels = matrix.columns?.levels ?? [];
    const valueSources = matrix.valueSources ?? [];
    const hasSeries = isSeriesLevel(columnLevels[0]);
    const legendSource = hasSeries ? columnLevels[0].sources[0] : undefined;
    const paths = rowPathsOf(matrix.rows?.root, rowLevels.length);
    for (const path of paths) path.forEach((node, depth) => ROW_PATHS.set(node, path.slice(0, depth + 1)));
    // どのメジャーの列かは、値のセルの valueSourceIndex で決める（無ければ 0）。同じフィールドを 2 つの欄に入れると、
    // 列の葉の levelSourceIndex は重なった列の後ろの番号になり（1・1・2 など）、葉からは決められない（2026-10-03、Desktop）
    // 行によってはセルが欠けるので、どれかの行で最初に見つかったセルの番号を使う（セルが 1 つも無ければ葉の番号）。
    // 番号が値の列の外なら、その列は使わない
    const cellAt = (k: number) => {
        for (const path of paths) {
            const cell = path[path.length - 1].values?.[k];
            if (cell) return cell;
        }
        return undefined;
    };
    const leaves = columnLeavesOf(matrix.columns?.root, hasSeries).map((leaf, k) => {
        const cell = cellAt(k);
        const measure = cell ? cell.valueSourceIndex ?? 0 : leaf.measure;
        return { ...leaf, measure: measure >= 0 && measure < valueSources.length ? measure : -1 };
    });
    if (!rowLevels.length || !paths.length) {
        const empty = Object.assign([] as DataViewValueColumn[], { grouped: (): DataViewValueColumnGroup[] => [] }) as DataViewValueColumns;
        return { ...dataView, categorical: { categories: [], values: empty } };
    }

    const valuesAt = (leafIndex: number) => paths.map((path) => path[path.length - 1].values?.[leafIndex]);
    const columnOf = (leafIndex: number, source: DataViewMetadataColumn): DataViewValueColumn => {
        const cells = valuesAt(leafIndex);
        const hasHighlights = cells.some((cell) => cell?.highlight !== undefined);
        return {
            source,
            values: cells.map((cell) => (cell?.value ?? null) as PrimitiveValue),
            ...(hasHighlights ? { highlights: cells.map((cell) => (cell?.highlight ?? null) as PrimitiveValue) } : {}),
        };
    };

    // 同じフィールドを 2 つの欄に入れると（「値」と「折れ線の率の分母」など）、matrix では欄ごとに値の列が届く。
    // categorical と同じく、役割を合わせた 1 列にする（2 列のまま足すと、分母や分子が 2 倍になる）
    const keyOf = (source: DataViewMetadataColumn) => source.queryName ?? source.displayName;
    const merged = new Map<string, { source: DataViewMetadataColumn; first: number }>();
    valueSources.forEach((source, m) => {
        const found = merged.get(keyOf(source));
        if (found) found.source = { ...found.source, roles: { ...found.source.roles, ...source.roles } };
        else merged.set(keyOf(source), { source, first: m });
    });
    const isFirst = (measure: number) => !!valueSources[measure] && merged.get(keyOf(valueSources[measure]))?.first === measure;
    const sourceOf = (measure: number) => merged.get(keyOf(valueSources[measure]))!.source;
    // 凡例をまとめたカテゴリ全体の値（列の小計）で受ける役割：「折れ線の値」「X 軸の定数線」「Y 軸の定数線」（凡例ごとに評価すると、先頭の凡例が空白のとき線が消える）
    const wholeRolesOf = (measure: number) => WHOLE_ROLES.filter((role) => sourceOf(measure).roles?.[role]);
    const isLine = (measure: number) => wholeRolesOf(measure).length > 0;
    const leafList = leaves.map((leaf, k) => ({ leaf, k })).filter(({ leaf }) => isFirst(leaf.measure));
    const subtotalLines = hasSeries ? leafList.filter(({ leaf }) => leaf.subtotal && isLine(leaf.measure)) : [];
    const useSubtotal = subtotalLines.length > 0;
    /** 凡例ごとの列の役割。小計を使うなら「折れ線の値」「X 軸の定数線」の役割を外す（ほかの役割が無ければ列にしない） */
    const seriesSourceOf = (measure: number): DataViewMetadataColumn | null => {
        const source = sourceOf(measure);
        if (!useSubtotal || !isLine(measure)) return source;
        const roles = Object.fromEntries(Object.entries(source.roles ?? {}).filter(([role]) => !WHOLE_ROLES.includes(role)));
        return Object.keys(roles).length ? { ...source, roles } : null;
    };
    const groups: MatrixGroup[] = [];
    if (hasSeries) {
        const seriesNodes = (matrix.columns?.root?.children ?? []).filter((node) => !node.isSubtotal);
        for (const node of seriesNodes) {
            const values = leafList
                .filter(({ leaf }) => leaf.series === node && seriesSourceOf(leaf.measure))
                .map(({ leaf, k }) => columnOf(k, { ...seriesSourceOf(leaf.measure)!, groupName: node.value as PrimitiveValue }));
            groups.push({
                name: node.value as PrimitiveValue,
                identity: node.identity as CustomVisualOpaqueIdentity,
                ...(node.objects ? { objects: node.objects } : {}),
                values,
                [MATRIX_NODE]: node,
            });
        }
        // 「折れ線の値」「X 軸の定数線」はカテゴリ全体の値（列の小計）を 1 列にして、最初のグループに置く（凡例ごとの値が同じときと同じ扱いになる）
        if (groups.length) {
            groups[0].values.push(...subtotalLines.map(({ leaf, k }) => columnOf(k, { ...sourceOf(leaf.measure), roles: Object.fromEntries(wholeRolesOf(leaf.measure).map((role) => [role, true])) })));
        }
    } else {
        groups.push({ values: leafList.map(({ leaf, k }) => columnOf(k, sourceOf(leaf.measure))) } as MatrixGroup);
    }

    const flat = groups.flatMap((group) => group.values);
    const values = Object.assign(flat, {
        grouped: () => groups,
        ...(legendSource ? { source: legendSource } : {}),
    }) as DataViewValueColumns;

    const categories: DataViewCategoryColumn[] = rowLevels.map((level, depth) => {
        const nodes = paths.map((path) => path[Math.min(depth, path.length - 1)]);
        const lowest = depth === rowLevels.length - 1;
        const objects = lowest && nodes.some((node) => node.objects) ? nodes.map((node) => node.objects) : undefined;
        return {
            source: level.sources[0],
            values: nodes.map((node) => node.value as PrimitiveValue),
            // 選択 ID は withMatrixNode で作るので、identity の位置に行の節点を入れて持ち回る
            identity: nodes as unknown as CustomVisualOpaqueIdentity[],
            ...(objects ? { objects: objects as powerbi.DataViewObjects[] } : {}),
        };
    });

    const categorical: MatrixCategorical = { categories, values, [MATRIX_LEVELS]: { rows: rowLevels, columns: columnLevels } };
    return { ...dataView, categorical };
}

/** カテゴリ（行）の i 番目を選択 ID に足す。matrix から詰め替えたなら行の節点で、そうでなければ今までどおり */
export function withCategoryOf(builder: ISelectionIdBuilder, categorical: DataViewCategorical, column: DataViewCategoryColumn, i: number): ISelectionIdBuilder {
    const levels = (categorical as MatrixCategorical)[MATRIX_LEVELS];
    if (!levels) return builder.withCategory(column, i);
    const node = column.identity![i] as unknown as DataViewMatrixNode;
    return (ROW_PATHS.get(node) ?? [node]).reduce((next, step) => next.withMatrixNode(step, levels.rows), builder);
}

/** 系列（凡例の値）を選択 ID に足す。matrix から詰め替えたなら列の節点で、そうでなければ今までどおり */
export function withSeriesOf(builder: ISelectionIdBuilder, categorical: DataViewCategorical, group: DataViewValueColumnGroup): ISelectionIdBuilder {
    const levels = (categorical as MatrixCategorical)[MATRIX_LEVELS];
    const node = (group as MatrixGroup)[MATRIX_NODE];
    if (!levels || !node) return builder.withSeries(categorical.values!, group);
    return builder.withMatrixNode(node, levels.columns);
}
