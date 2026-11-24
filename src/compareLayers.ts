"use strict";

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import { CALCULATION_MODES, COMPARE_DIRECTIONS, COMPARE_ORDERS, MAX_COMPARE_LAYERS, STACK_OWN, VisualFormattingSettingsModel } from "./settings";
import { categoricalOf, wholeMeasureValuesOf } from "./matrixDataView";
import {
    CalculationRuntime,
    CompareSettings,
    DataPoint,
    LegendItemInfo,
    LabelRoom,
    STACK_WHOLE_ROLE,
    StackWholeColumn,
    ValueExtent,
    ViewModel,
    readTransform,
    transform,
    tooltipStackOf,
    unionLabelRoom,
} from "./viewModel";
import { tooltipItemsOf } from "./tooltip";
import { labelLayersOf, layerOpacity } from "./layout";
import { formatOfColumn, nameTextOf } from "./shared/dataValue";

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
 * 行を読むのはレイヤーごとに 1 回。読み終えたところで止め、すべてのレイヤーの範囲をそろえてから続きを組む。
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
    const split = splitByCompareBy(dataView, settings);
    const whole = categoricalOf(split ? split.layers[0].view : dataView);
    // 「積む柱」を選んだメジャーがあれば、柱ごとに 1 枚のレイヤーにして、重ねずに隣り合わせに並べる
    const pillars = split ? null : pillarPlanOf(whole, settings);
    // 凡例があり「値」にメジャーが 2 つ以上なら、メジャーごとに 1 枚のレイヤーにして、重ねずに隣り合わせに並べる
    const measures = split || pillars ? [] : groupedMeasuresOf(whole);
    const grouped = measures.length > 1 || !!pillars;
    const pillarLayers: LayerSource[] | null = pillars
        ? pillars.roots.map((root) => ({ name: root.displayName ?? "", view: pillarViewOf(whole!, pillars, root) }))
        : null;
    const compare = grouped ? GROUPED_LAYOUT : compareSettingsOf(settings);
    const view = pillarLayers ? pillarLayers[0].view : grouped ? measureViewOf(whole!, keyOf(measures[0])) : whole;
    const sources = split || grouped ? [] : compareSourcesOf(view);
    const show = grouped || (settings.compare.show.value ?? true);
    const layers: LayerSource[] = split
        ? split.layers
        : pillarLayers
            ? pillarLayers
            : grouped
                ? measures.map((source) => ({ name: source.displayName ?? "", view: measureViewOf(whole!, keyOf(source)) }))
                : [{ name: frontNameOf(view), view: view! }, ...sources.map((source) => ({ name: source.displayName ?? "", view: layerViewOf(view!, source) }))];
    const pillarBasis = pillars ? { pillar: true } : {};
    // 手前のレイヤーの行を読む。重ねないなら、このレイヤーの範囲のまま続きを組んで返す
    const frontRead = readTransform(view, host, settings, runtime, pillarBasis);
    const alone = (): ViewModel => {
        const pending = frontRead.resume({});
        return pending.finish(pending.labelRoom);
    };
    const notLayered = (front: ViewModel): ViewModel => {
        // 折れ線だけ（値が空）のときは、重ねる棒が無いので重ねない。線は手前の値で描き、そのことを知らせる
        if (front.lineOnly && layers.length > 1 && show) {
            const compareLineOnlyWarning = split
                ? `「比較の列」は棒を前後に重ねる欄で、折れ線だけのときは重ねません。線は${split.name}が「${layers[0].name}」の値で描いています。`
                : "「比較値」は棒を前後に重ねる欄で、折れ線だけのときは描きません。";
            return { ...front, compare, compareLineOnlyWarning };
        }
        // 「比較の列」で受けて重ねないとき（オフ・パレート・値が 1 つ）は、手前の値だけを描く
        return { ...front, compare, ...(split?.warning && show ? { compareWarning: split.warning } : {}) };
    };
    // 重ねうるのは、レイヤーが 2 枚以上で「比較」がオンで、折れ線だけ・パレートでないとき（データが無ければ読む前に組み終えて重ねない）
    if (!(layers.length > 1 && show && frontRead.canLayer)) return notLayered(alone());

    const frontBasis = frontRead.basis;
    const backViews = layers.slice(1, MAX_COMPARE_LAYERS).map((layer) => categoricalOf(layer.view)!);
    const order = {
        keepRows: frontBasis.keepRows ?? undefined,
        displayOrder: frontBasis.displayOrder,
        seriesOrder: frontBasis.seriesOrder,
    };
    // 奥のレイヤーの行を読んで範囲を集め、すべてのレイヤーの範囲の軸でどのレイヤーも続きを組む（行は読み直さない）
    const backReads = backViews.map((v) => readTransform(v, host, settings, runtime, { ...order, ...pillarBasis }));
    const extent = unionOf([frontBasis.extent, ...backReads.map((read) => read.basis.extent)]);
    const titleNames = pillars ? pillars.measures.map((m) => m.displayName ?? "") : layers.map((l) => l.name);
    const pendingFront = frontRead.resume({ extent, ...(grouped ? { valueTitleNames: titleNames } : {}) });
    const pendingBacks = backReads.map((read) => read.resume({ extent }));
    // 値の軸の余白は、データ ラベルを描くすべてのレイヤー（メジャーを並べるなら描くすべて、重ねるなら描くうちのいちばん手前。
    // 凡例で隠したレイヤーは除く）のラベルで取り、どのレイヤーも同じ範囲で組む
    const layerLabelRooms = [pendingFront, ...pendingBacks].map((pending) => pending.labelRoom);
    const room = unionLabelRoom(labelLayersOf(layerLabelRooms.length, runtime.visibleLayers ?? [], grouped).map((l) => layerLabelRooms[l]));
    const merged = pendingFront.finish(room);
    // 棒が 1 本も無かった（値がすべて空白など）なら重ねず、手前の値だけを組み直して描く（棒が無いので軽い）
    if (merged.isEmpty) return notLayered(transform(view, host, settings, runtime));
    const backs = pendingBacks.map((pending) => pending.finish(room));

    // 行の番号で引く（行ごとに探すと、行の数の 2 乗になる）
    const backsByRow = backs.map((vm) => new Map(vm.categoryGroups.map((g) => [g.rowIndex, g])));
    for (const group of merged.categoryGroups) {
        group.points.forEach((d) => (d.layerIndex = 0));
        const frontBySeries = new Map<number, DataPoint>();
        // 同じ系列が 2 つあれば先のもの（find と同じ）
        for (const p of group.points) if (!frontBySeries.has(p.seriesIndex)) frontBySeries.set(p.seriesIndex, p);
        group.layerPoints = backsByRow.map((byRow, k) => {
            const points: DataPoint[] = byRow.get(group.rowIndex)?.points ?? [];
            points.forEach((d) => {
                d.layerIndex = k + 1;
                // 色は手前の同じ系列の棒にそろえる（奥のレイヤーは別のメジャーなので、テーマの色の割り当てが変わる）。
                // レイヤーの違いは濃さで見せる
                const front = frontBySeries.get(d.seriesIndex);
                if (front) {
                    d.color = front.color;
                    d.borderColor = front.borderColor;
                }
            });
            return points;
        });
        if (grouped) group.layerTotals = backsByRow.map((byRow) => byRow.get(group.rowIndex)?.totals ?? null);
    }
    const compareLayers = layers.slice(0, MAX_COMPARE_LAYERS).map((layer) => ({ name: layer.name }));
    const legendEntries = legendEntriesOf(merged, compareLayers.map((l) => l.name), compare, settings.legend.reverseOrder.value ?? false);
    return {
        ...merged,
        compare,
        compareLayers,
        ...(grouped ? { measureGroups: true } : {}),
        layerLabelRooms,
        layerTooltips: [merged.tooltip, ...backs.map((vm) => vm.tooltip)],
        legendEntries,
        legend: { ...merged.legend, show: (settings.legend.show.value ?? true) && legendEntries.length > 0 && (merged.seriesMode || merged.legend.show || legendEntries.length > 1) },
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
    // カテゴリが無い（比較の列だけ）ときは分けない（比較の列をカテゴリとして描く）
    if (!dataView || !matrix?.rows?.root || depth <= 0) return null;
    const keyOf = (value: unknown) => (value instanceof Date ? `date:${value.getTime()}` : `${typeof value}:${String(value)}`);
    // 値の名前は比較の列の書式で出す（カテゴリと同じ書き方。日付は書式どおりで、実行環境の言語に左右されない。
    // 空白は「(空白)」、真偽は True・False、NaN・±Infinity は空白にせず字で出す）
    const nameOf = nameTextOf(formatOfColumn(levels[depth].sources.find((source) => source.roles?.compareBy)));
    // 1.36.10 までの名前（書式を通さない。日付は実行環境の toLocaleDateString）。その版で保存した「手前にする値の名前」を読むため
    const legacyNameOf = (value: unknown) =>
        value === null || value === undefined || value === "" ? "(空白)" : value instanceof Date ? value.toLocaleDateString() : String(value);

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
    // まず今の名前で探し、無ければ前の版の名前で探す
    const named = frontName
        ? (ordered.find((key) => nameOf(values.get(key)) === frontName) ?? ordered.find((key) => legacyNameOf(values.get(key)) === frontName))
        : undefined;
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

/**
 * メジャーごとに棒を並べるときのレイヤーの置き方。最初のメジャーを左（横棒は上）にして、重ねずに並べる。
 * どの棒も薄くせず、凡例にはレイヤーを出さない（どのメジャーかは位置と軸のタイトル・ツールヒントで読む）
 */
const GROUPED_LAYOUT: CompareSettings = { rightFront: false, overlap: 0, backOpacity: 1, outline: false, showInLegend: false };

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

/**
 * 凡例があるときに「値」に入れたメジャー（入れた順）。凡例の値ごとに複製されて届くので、1 つにまとめる。
 * 凡例が無ければ空（メジャーごとの系列として、今までどおり 1 本の棒の中に並べる・積む）
 */
function groupedMeasuresOf(view: DataView | undefined): DataViewMetadataColumn[] {
    if (!view?.categorical?.values?.source) return [];
    const found = new Map<string, DataViewMetadataColumn>();
    for (const column of view.categorical.values) {
        const source = column.source;
        if (source?.roles?.measure && !found.has(keyOf(source))) found.set(keyOf(source), source);
    }
    const position = (source: DataViewMetadataColumn) =>
        (source as DataViewMetadataColumn & { rolesIndex?: Record<string, number[]> }).rolesIndex?.measure?.[0] ?? 0;
    return [...found.values()].sort((a, b) => position(a) - position(b));
}

/** 「積む柱」で柱を並べるときの組み方。柱の根のメジャー（入れた順）と、ほかの柱に積むメジャー */
export interface PillarPlan {
    /** 「値」のメジャー（入れた順） */
    measures: DataViewMetadataColumn[];
    /** 柱ごとのいちばん下のメジャー（入れた順）。最初の柱が左（横棒は上） */
    roots: DataViewMetadataColumn[];
    /** ほかの柱に積むメジャー（入れた順）と、積む先の柱の根の key */
    stacked: Array<{ source: DataViewMetadataColumn; root: string }>;
    /** 積むメジャーの負の値を最初の柱の上に積む（オフなら積む柱の 0 の下へ） */
    lossToFirst: boolean;
}

/** 「値」に入れたメジャー（入れた順）。凡例があると凡例の値ごとに複製されて届くので、1 つにまとめる */
function measuresOf(view: DataView | undefined): DataViewMetadataColumn[] {
    const found = new Map<string, DataViewMetadataColumn>();
    for (const column of view?.categorical?.values ?? []) {
        const source = column.source;
        if (source?.roles?.measure && !found.has(keyOf(source))) found.set(keyOf(source), source);
    }
    const position = (source: DataViewMetadataColumn) =>
        (source as DataViewMetadataColumn & { rolesIndex?: Record<string, number[]> }).rolesIndex?.measure?.[0] ?? 0;
    return [...found.values()].sort((a, b) => position(a) - position(b));
}

/** メジャーに保存した「積む柱」（積む先のメジャーの queryName）。自分の柱なら空 */
function stackOntoOf(source: DataViewMetadataColumn): string {
    const value = source.objects?.columns?.stackOnto;
    return typeof value === "string" && value !== STACK_OWN ? value : "";
}

/**
 * 「積む柱」の組み方。どのメジャーもほかの柱に積まないなら null（今までどおり）。
 * 積む先は、自分の柱に立つメジャー（それ自身がほかの柱に積まないもの）だけ。積む先が見つからない・それ自身もほかに積むなら自分の柱。
 * パレートは柱を並べない（累積比はカテゴリの合計で引くので、柱を分けても意味が変わらない）
 */
export function pillarPlanOf(view: DataView | undefined, settings: VisualFormattingSettingsModel): PillarPlan | null {
    const mode = String(settings.calculation.mode.value?.value ?? CALCULATION_MODES.none);
    if (mode === CALCULATION_MODES.pareto) return null;
    const measures = measuresOf(view);
    if (measures.length < 2) return null;
    const byKey = new Map(measures.map((m) => [keyOf(m), m]));
    const targetOf = (m: DataViewMetadataColumn): string => {
        const target = stackOntoOf(m);
        if (!target || target === keyOf(m) || !byKey.has(target)) return "";
        return stackOntoOf(byKey.get(target)!) && byKey.has(stackOntoOf(byKey.get(target)!)) ? "" : target;
    };
    const stacked = measures.filter((m) => targetOf(m)).map((source) => ({ source, root: targetOf(source) }));
    if (!stacked.length) return null;
    return { measures, roots: measures.filter((m) => !targetOf(m)), stacked, lossToFirst: settings.columns.stackLossToFirst.value ?? true };
}

/** 数として読める値（空白・数でない値は null） */
const numberOf = (value: unknown): number | null => (typeof value === "number" && Number.isFinite(value) ? value : null);

/**
 * 柱 1 本ぶんを読む DataView。凡例があれば、柱の根のメジャーを凡例ごとに積み、ほかの柱に積むメジャーは凡例で分けずに
 * 全体の値（列の小計）で 1 つの系列として積む。凡例が無ければ、柱の根のメジャーを系列にし、ほかの柱の根は値の無い系列として残す
 * （どの柱も系列の並びをそろえ、色・凡例・書式の対象を柱をまたいで同じにする）。
 * ほかの柱に積むメジャーは、どの柱にも 1 つの系列として置き、積む柱にだけ正の値を入れる。負の値（赤字）は、
 * 最初の柱の上に絶対値で積み、ラベルとツールヒントは元の負の値で出す（PL の比例縮尺図で、損失を収益の側に置くのと同じ）
 */
function pillarViewOf(view: DataView, plan: PillarPlan, root: DataViewMetadataColumn): DataView {
    const categorical = view.categorical!;
    const values = categorical.values!;
    const legend = !!values.source;
    const rootKey = keyOf(root);
    const firstKey = keyOf(plan.roots[0]);
    const rootKeys = new Set(plan.roots.map(keyOf));
    const stackedKeys = new Set(plan.stacked.map((s) => keyOf(s.source)));
    const columnMap = new Map<DataViewValueColumn, DataViewValueColumn | null>();
    const swap = (column: DataViewValueColumn): DataViewValueColumn | null => {
        if (columnMap.has(column)) return columnMap.get(column)!;
        const roles = { ...(column.source?.roles ?? {}) };
        const key = column.source ? keyOf(column.source) : "";
        let next: DataViewValueColumn | null = column;
        if (roles.measure && key !== rootKey) {
            if (!legend && rootKeys.has(key)) {
                // 凡例なし：ほかの柱の根は、値の無い系列として残す（ハイライトも空）
                const { highlights: _highlights, ...rest } = column;
                next = { ...rest, values: column.values.map((): null => null) };
            } else {
                delete roles.measure;
                next = Object.keys(roles).length ? { ...column, source: { ...column.source, roles } } : null;
            }
        }
        columnMap.set(column, next);
        return next;
    };
    const keep = (column: DataViewValueColumn | null): column is DataViewValueColumn => column !== null;
    const groups: DataViewValueColumnGroup[] = values.grouped?.() ?? [];
    const newGroups = groups.map((group) => ({ ...group, values: group.values.map(swap).filter(keep) }));

    const rowCount = categorical.categories?.[0]?.values.length ?? values[0]?.values.length ?? 0;
    // 積むメジャーの全体の値。凡例があれば列の小計、届かなければ凡例ごとの値を足す
    const wholeOf = (source: DataViewMetadataColumn): Array<number | null> => {
        const key = keyOf(source);
        const subtotal = legend ? wholeMeasureValuesOf(categorical, key) : undefined;
        if (subtotal) return subtotal.map(numberOf);
        const columns = values.filter((c) => c.source && keyOf(c.source) === key && c.source.roles?.measure);
        if (!legend) return (columns[0]?.values ?? []).map(numberOf);
        return Array.from({ length: rowCount }, (_, i) => {
            const nums = columns.map((c) => numberOf(c.values[i])).filter((v): v is number => v !== null);
            return nums.length ? nums.reduce((a, b) => a + b, 0) : null;
        });
    };
    const wholeColumns = plan.stacked.filter((s) => stackedKeys.has(keyOf(s.source))).map(({ source, root: target }): StackWholeColumn => {
        const amounts = wholeOf(source);
        const own = target === rootKey;
        // 最初の柱に積むメジャーは、負の値もそのまま（ほかの積み上げと同じく 0 の下へ）
        const lossHere = plan.lossToFirst && rootKey === firstKey && target !== firstKey;
        const roles = { [STACK_WHOLE_ROLE]: true };
        const cells = amounts.map((v) => (v === null ? null : own && (v > 0 || target === firstKey || !plan.lossToFirst) ? v : lossHere && v < 0 ? -v : null));
        return {
            source: { ...source, roles, groupName: undefined } as DataViewMetadataColumn,
            values: cells,
            ...(lossHere ? { signed: { source: { ...source, roles }, values: amounts } } : {}),
        };
    });
    if (newGroups.length) newGroups[0] = { ...newGroups[0], values: [...newGroups[0].values, ...wholeColumns] };
    else newGroups.push({ values: wholeColumns } as DataViewValueColumnGroup);
    const newValues = Object.assign(newGroups.flatMap((group) => group.values), {
        grouped: () => newGroups,
        ...(values.source ? { source: values.source } : {}),
        ...(values.identityFields ? { identityFields: values.identityFields } : {}),
    }) as powerbi.DataViewValueColumns;
    return { ...view, categorical: { ...categorical, values: newValues } };
}

/**
 * 1 つのメジャーだけを「値」として読む DataView。ほかのメジャーは「値」の役割を外す（ほかの役割も無ければ列ごと外す）。
 * 凡例のグループ・折れ線・ツールヒントなどはそのまま残す
 */
function measureViewOf(view: DataView, key: string): DataView {
    const categorical = view.categorical!;
    const values = categorical.values!;
    const columnMap = new Map<DataViewValueColumn, DataViewValueColumn | null>();
    const swap = (column: DataViewValueColumn): DataViewValueColumn | null => {
        if (columnMap.has(column)) return columnMap.get(column)!;
        const roles = { ...(column.source?.roles ?? {}) };
        if (!roles.measure || keyOf(column.source) === key) {
            columnMap.set(column, column);
            return column;
        }
        delete roles.measure;
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
    const color = vm.highContrast ? vm.highContrast.foreground : vm.seriesMode ? LAYER_LEGEND_GRAY : vm.series[0]?.color ?? vm.columns.fill;
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
 * 凡例で描くレイヤーを before から after に変えると、データ ラベルを描くレイヤーが変わって値の軸の余白が変わるか。
 * 変わるなら組み直す（変わらなければ、組んだ viewModel のまま描き直す）
 */
export function labelRoomChanges(viewModel: ViewModel, before: readonly number[], after: readonly number[]): boolean {
    const rooms: LabelRoom[] = viewModel.layerLabelRooms ?? [];
    if (!rooms.some((room) => room.top.length || room.bottom.length)) return false;
    const roomOf = (visible: readonly number[]) =>
        JSON.stringify(unionLabelRoom(labelLayersOf(rooms.length, visible, !!viewModel.measureGroups).map((l) => rooms[l])));
    return roomOf(before) !== roomOf(after);
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
