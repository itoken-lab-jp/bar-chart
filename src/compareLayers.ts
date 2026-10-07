"use strict";

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import { COMPARE_DIRECTIONS, COMPARE_ORDERS, MAX_COMPARE_LAYERS, VisualFormattingSettingsModel } from "./settings";
import { categoricalOf } from "./matrixDataView";
import { CalculationRuntime, CompareSettings, DataPoint, LegendItemInfo, ValueExtent, ViewModel, transform, tooltipStackOf } from "./viewModel";
import { tooltipItemsOf } from "./tooltip";
import { layerOpacity } from "./layout";

/** 比較レイヤー 1 枚ぶんの元データ。view はこのレイヤーを「値」として読む DataView */
interface LayerSource {
    name: string;
    view: DataView;
}

/** 凡例（系列）があるときの、比較レイヤーの凡例の印の色 */
export const LAYER_LEGEND_GRAY = "#8A8886";

/** 比較のレイヤーが上限を超えたときの警告の見出し */
export const COMPARE_LIMIT_TITLE = "比較のレイヤーが多すぎます";
/** 折れ線だけ（値が空）で、比較の欄にフィールドを入れたときの警告の見出し */
export const COMPARE_LINE_ONLY_TITLE = "折れ線だけでは比較で重ねません";

/**
 * 比較レイヤー。手前のレイヤーに、奥のレイヤーを同じ幅で少しずつずらして重ねる。レイヤーの作り方は 2 通り：
 * - 横持ち：「値」を手前に、「比較値」に入れたメジャーを入れた順に奥へ
 * - 縦持ち：「比較の列」（予定・見通し・実績が 1 列に並んだもの）の値ごとに 1 枚。matrix の行のいちばん下の段で届く
 *   （行の木を値ごとに分ける。docs: dev/probes/compareRowsProbe/README.md）。どれを手前にするかは「比較」カードで選ぶ
 *
 * レイヤーごとに transform を走らせ（奥のレイヤーは、そのレイヤーの値を「値」として読む）、手前のレイヤーに重ねる。
 * 積み上げ・累計・ツールヒントは、レイヤーごとに今までと同じ規則で計算される。
 * どのレイヤーも同じ軸で描くよう、軸はすべてのレイヤーの範囲で決める。カテゴリの並びと「その他」にまとめる行は、
 * 手前のレイヤーで決めたものを奥のレイヤーにも使う（値順でも、奥の棒が別のカテゴリの位置に入れ替わらない）
 */
export function transformWithLayers(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    runtime: CalculationRuntime = {}
): ViewModel {
    const compare = compareSettingsOf(settings);
    const split = splitByCompareBy(dataView, settings);
    const view = categoricalOf(split ? split.layers[0].view : dataView);
    const sources = split ? [] : compareSourcesOf(view);
    const front = transform(view, host, settings, runtime);
    const show = settings.compare.show.value ?? true;
    const layers: LayerSource[] = split
        ? split.layers
        : [{ name: frontNameOf(view), view: view! }, ...sources.map((source) => ({ name: source.displayName ?? "", view: layerViewOf(view!, source) }))];
    // 折れ線だけ（値が空）のときは、重ねる棒が無いので重ねない。線は手前の値で描き、そのことを知らせる
    if (front.lineOnly && layers.length > 1 && show) {
        const compareLineOnlyWarning = split
            ? `「比較の列」は棒を前後に重ねる欄で、折れ線だけのときは重ねません。線は${split.name}が「${layers[0].name}」の値で描いています。`
            : "「比較値」は棒を前後に重ねる欄で、折れ線だけのときは描きません。";
        return { ...front, compare, compareLineOnlyWarning };
    }
    const layered = layers.length > 1 && show && !front.isEmpty && !front.pareto.enabled;
    // 「比較の列」で受けて重ねないとき（オフ・パレート・値が 1 つ）は、手前の値だけを描く
    if (!layered) return { ...front, compare, ...(split?.warning && show ? { compareWarning: split.warning } : {}) };

    const backViews = layers.slice(1, MAX_COMPARE_LAYERS).map((layer) => categoricalOf(layer.view)!);
    const order = {
        keepRows: front.basis.keepRows ?? undefined,
        displayOrder: front.basis.displayOrder,
    };
    // 1 回目でレイヤーごとの範囲を集め、2 回目ですべてのレイヤーの範囲の軸で描き直す
    const firstPass = backViews.map((v) => transform(v, host, settings, runtime, order));
    const extent = unionOf([front.basis.extent, ...firstPass.map((vm) => vm.basis.extent)]);
    const merged = transform(view, host, settings, runtime, { extent });
    const backs = backViews.map((v) => transform(v, host, settings, runtime, { ...order, extent }));

    const backsByRow = backs.map((vm) => new Map(vm.categoryGroups.map((g) => [g.rowIndex, g.points])));
    for (const group of merged.categoryGroups) {
        group.points.forEach((d) => (d.layerIndex = 0));
        group.layerPoints = backsByRow.map((byRow, k) => {
            const points: DataPoint[] = byRow.get(group.rowIndex) ?? [];
            points.forEach((d) => {
                d.layerIndex = k + 1;
                // 色は手前の同じ系列の棒にそろえる（奥のレイヤーは別のメジャーなので、テーマの色の割り当てが変わる）。
                // レイヤーの違いは濃さで見せる
                const front = group.points.find((p) => p.seriesIndex === d.seriesIndex);
                if (front) {
                    d.color = front.color;
                    d.borderColor = front.borderColor;
                }
            });
            return points;
        });
    }
    const compareLayers = layers.slice(0, MAX_COMPARE_LAYERS).map((layer) => ({ name: layer.name }));
    const legendEntries = legendEntriesOf(merged, compareLayers.map((l) => l.name), compare, settings.legend.reverseOrder.value ?? false);
    return {
        ...merged,
        compare,
        compareLayers,
        layerTooltips: [merged.tooltip, ...backs.map((vm) => vm.tooltip)],
        legendEntries,
        legend: { ...merged.legend, show: (settings.legend.show.value ?? true) && (merged.seriesMode || legendEntries.length > 1) },
        // リボンは系列の棒の端をつなぐので、幅を割ったレイヤーの棒とは合わない。比較のあいだは描かない
        ribbons: { ...merged.ribbons, show: false },
        ...(split?.warning ? { compareWarning: split.warning } : {}),
        ...(split ? { compareByName: split.name } : {}),
    };
}

/**
 * 棒のツールヒント。比較レイヤーの奥の棒は、そのレイヤーの値で出す。
 * 「比較の列」で重ねたときは、どの値（予定・見通し・実績）の棒かをカテゴリの行の次に出す（「比較値」はメジャーの名前で分かる）
 */
export function barTooltipItems(viewModel: ViewModel, d: DataPoint): powerbi.extensibility.VisualTooltipDataItem[] {
    const source = viewModel.layerTooltips?.[d.layerIndex ?? 0] ?? viewModel.tooltip;
    if (!source) return [];
    const layer = d.layerIndex !== undefined ? viewModel.compareLayers[d.layerIndex] : undefined;
    const layerRow = viewModel.compareByName && layer ? { displayName: viewModel.compareByName, value: layer.name } : undefined;
    return tooltipItemsOf(source, d.rowIndex, d.category, d.seriesIndex, tooltipStackOf(viewModel, d), layerRow);
}

/**
 * 「比較の列」で受けた matrix を、列の値ごとの DataView に分ける（手前にする値が先頭）。「比較の列」が無ければ null。
 * 比較の列は行のいちばん下の段で届く。その段を外し、カテゴリのいちばん下の節点に、その値の行の値を移す。
 * どのレイヤーも行の木の形（カテゴリの並び）は同じにする（値の無いカテゴリは値の無い節点）。行の番号がそろい、手前と重ねられる
 */
export function splitByCompareBy(
    dataView: DataView | undefined,
    settings: VisualFormattingSettingsModel
): { layers: LayerSource[]; name: string; warning?: string } | null {
    const matrix = dataView?.matrix;
    const levels = matrix?.rows?.levels ?? [];
    const depth = levels.findIndex((level) => level.sources.some((source) => source.roles?.compareBy));
    // カテゴリが無い（比較の列だけ）ときは分けない（カテゴリが無ければ棒グラフは描かない）
    if (!dataView || !matrix?.rows?.root || depth <= 0) return null;
    const keyOf = (value: unknown) => (value instanceof Date ? `date:${value.getTime()}` : `${typeof value}:${String(value)}`);
    const nameOf = (value: unknown) => (value === null || value === undefined || value === "" ? "(空白)" : value instanceof Date ? value.toLocaleDateString() : String(value));

    // 列の値（届いた順。Power BI は列の並べ替えの順で返す）
    const values = new Map<string, unknown>();
    const collect = (node: DataViewMatrixNode, level: number) => {
        for (const child of node.children ?? []) {
            if (child.isSubtotal) continue;
            if (level === depth) {
                if (!values.has(keyOf(child.value))) values.set(keyOf(child.value), child.value);
            } else collect(child, level + 1);
        }
    };
    collect(matrix.rows.root, 0);
    if (!values.size) return null;

    // 手前の値：名前で指定があればその値、無ければ並びの最後（既定）か最初。奥へは並びを手前から離れる向きにたどる
    const card = settings.compare;
    let ordered = [...values.keys()];
    if (String(card.order.value?.value ?? COMPARE_ORDERS.lastFront) !== COMPARE_ORDERS.firstFront) ordered.reverse();
    const frontName = (card.frontValue.value ?? "").trim();
    const named = frontName ? ordered.find((key) => nameOf(values.get(key)) === frontName) : undefined;
    if (named) ordered = [named, ...ordered.filter((key) => key !== named)];
    const warning =
        ordered.length > MAX_COMPARE_LAYERS
            ? `重ねられるのは ${MAX_COMPARE_LAYERS} 枚までです。「比較の列」の ${ordered.slice(MAX_COMPARE_LAYERS).map((key) => nameOf(values.get(key))).join("・")} は描いていません。`
            : undefined;

    const layerOf = (key: string): LayerSource => {
        const strip = (node: DataViewMatrixNode, level: number): DataViewMatrixNode => {
            // カテゴリのいちばん下の節点：その値の子の値を自分の値にして、比較の列の段を外す
            if (level === depth - 1) {
                const match = (node.children ?? []).find((child) => !child.isSubtotal && keyOf(child.value) === key);
                const { children: _children, ...rest } = node;
                return { ...rest, values: match?.values ?? {} };
            }
            return { ...node, children: (node.children ?? []).filter((child) => !child.isSubtotal).map((child) => strip(child, level + 1)) };
        };
        const root = strip(matrix.rows.root, -1);
        return {
            name: nameOf(values.get(key)),
            view: {
                ...dataView,
                matrix: { ...matrix, rows: { ...matrix.rows, root, levels: levels.filter((_, k) => k !== depth) } },
            },
        };
    };
    return { layers: ordered.map(layerOf), name: levels[depth].sources[0]?.displayName ?? "", ...(warning ? { warning } : {}) };
}

function compareSettingsOf(settings: VisualFormattingSettingsModel): CompareSettings {
    const card = settings.compare;
    const percent = (value: number | undefined, fallback: number) =>
        Math.max(0, Math.min(100, Number.isFinite(value) ? (value as number) : fallback)) / 100;
    return {
        rightFront: String(card.direction.value?.value ?? COMPARE_DIRECTIONS.rightFront) !== COMPARE_DIRECTIONS.leftFront,
        overlap: percent(card.overlap.value, 60),
        backOpacity: percent(card.backOpacity.value, 30),
        outline: card.outline.value ?? false,
        showInLegend: card.showInLegend.value ?? true,
    };
}

const keyOf = (source: DataViewMetadataColumn) => source.queryName ?? source.displayName;

/** 「比較値」に入れたメジャー（入れた順）。凡例があると凡例の値ごとに複製されて届くので、1 つにまとめる */
function compareSourcesOf(view: DataView | undefined): DataViewMetadataColumn[] {
    const found = new Map<string, DataViewMetadataColumn>();
    for (const column of view?.categorical?.values ?? []) {
        const source = column.source;
        if (source?.roles?.compare && !found.has(keyOf(source))) found.set(keyOf(source), source);
    }
    // 欄の中の順は rolesIndex（型定義に無いが Power BI は付けて渡す）。無ければ届いた順のまま
    const position = (source: DataViewMetadataColumn) =>
        (source as DataViewMetadataColumn & { rolesIndex?: Record<string, number[]> }).rolesIndex?.compare?.[0] ?? 0;
    return [...found.values()].sort((a, b) => position(a) - position(b));
}

/** 手前のレイヤーの名前（「値」に入れたメジャーの名前） */
function frontNameOf(view: DataView | undefined): string {
    const column = (view?.categorical?.values ?? []).find((c) => c.source?.roles?.measure);
    return column?.source.displayName ?? "";
}

/**
 * 奥のレイヤーを読む DataView。「値」の役割を外し、この「比較値」の列を「値」として読む。
 * ほかの役割（折れ線・ツールヒント・棒の色など）はそのまま残す。matrix から詰め替えた印（シンボルのキー）もそのまま持ち回る
 */
function layerViewOf(view: DataView, source: DataViewMetadataColumn): DataView {
    const categorical = view.categorical!;
    const values = categorical.values!;
    const key = keyOf(source);
    const columnMap = new Map<DataViewValueColumn, DataViewValueColumn | null>();
    const swap = (column: DataViewValueColumn): DataViewValueColumn | null => {
        if (columnMap.has(column)) return columnMap.get(column)!;
        const roles = { ...(column.source?.roles ?? {}) };
        const isLayer = !!roles.compare && keyOf(column.source) === key;
        if (!isLayer && !roles.measure) {
            columnMap.set(column, column);
            return column;
        }
        delete roles.measure;
        if (isLayer) {
            delete roles.compare;
            roles.measure = true;
        }
        const next = Object.keys(roles).length ? { ...column, source: { ...column.source, roles } } : null;
        columnMap.set(column, next);
        return next;
    };
    const keep = (column: DataViewValueColumn | null): column is DataViewValueColumn => column !== null;
    const groups: DataViewValueColumnGroup[] = values.grouped?.() ?? [];
    const newGroups = groups.map((group) => ({ ...group, values: group.values.map(swap).filter(keep) }));
    const newValues = Object.assign(values.map(swap).filter(keep), {
        grouped: () => newGroups,
        ...(values.source ? { source: values.source } : {}),
        ...(values.identityFields ? { identityFields: values.identityFields } : {}),
    }) as powerbi.DataViewValueColumns;
    return { ...view, categorical: { ...categorical, values: newValues } };
}

function unionOf(extents: ValueExtent[]): ValueExtent {
    return extents.reduce((a, b) => ({
        min: Math.min(a.min, b.min),
        max: Math.max(a.max, b.max),
        valueMaxAbs: Math.max(a.valueMaxAbs, b.valueMaxAbs),
        magMin: Math.min(a.magMin, b.magMin),
        magMax: Math.max(a.magMax, b.magMax),
    }));
}

/**
 * 凡例。系列（凡例があるとき）と折れ線の後ろに、比較レイヤーを並べる（凡例では見出し「比較」の別のまとまり）。
 * レイヤーはグラフの棒と同じ向きに並べる：手前の棒が右（横棒は下）なら奥から手前へ、左（上）なら手前から奥へ。
 * 凡例の「表示順を反転」なら、レイヤーの並びも逆にする
 * 系列が 1 本なら、棒の項目（「値」の名前）は手前のレイヤーの項目と重なるので出さない。
 * レイヤーの印は、系列が 1 本なら棒の色、凡例（系列）があれば灰色で、レイヤーの濃さに塗る（系列の色と混ざらないように）
 */
function legendEntriesOf(vm: ViewModel, names: string[], compare: CompareSettings, reverse: boolean): LegendItemInfo[] {
    if (!compare.showInLegend) return vm.legendEntries;
    const color = vm.seriesMode ? LAYER_LEGEND_GRAY : vm.series[0]?.color ?? vm.columns.fill;
    const layers = names.map((name, index): LegendItemInfo => ({
        kind: "layer",
        index,
        name,
        color,
        selectionId: null,
        layerOpacity: layerOpacity(index, names.length, compare.backOpacity),
    }));
    if (compare.rightFront !== reverse) layers.reverse();
    const bars = vm.legendEntries.filter((e) => e.kind === "bar" && vm.seriesMode);
    const lines = vm.legendEntries.filter((e) => e.kind === "line");
    // レイヤーは別のまとまり（見出し「比較」）として最後に置く。系列と折れ線は今までの並びのまま
    return [...bars, ...lines, ...layers];
}

/**
 * 凡例でレイヤーを押したあとに描くレイヤー。押したレイヤーだけにする。もう一度押すとすべてに戻す。
 * Ctrl を押しながらなら足したり外したりする（すべて外れたらすべてに戻る）
 */
export function nextVisibleLayers(current: number[], layer: number, multiSelect: boolean): number[] {
    if (!multiSelect) return current.length === 1 && current[0] === layer ? [] : [layer];
    const next = current.includes(layer) ? current.filter((l) => l !== layer) : [...current, layer];
    return next.sort((a, b) => a - b);
}
