/**
 * 値の軸の長さ（縦棒は描く範囲の高さ、横棒は幅）で組み直す決まり。値の軸の長さは描いてから分かるので、入力から見積もった長さで組み、
 * 描いて知らせた長さ（App の onValueAxisLength）で組み直す（visual.ts とハーネスが使う）
 */
import powerbi from "powerbi-visuals-api";

import { recommendedTickCount, tickTargetOf } from "./shared/ticks";
import { ORIENTATIONS, VisualFormattingSettingsModel } from "./settings";
import { ViewModel } from "./viewModel";

/** 1 回の組み立て（update・閲覧者の操作）の中で、値の軸の長さで組み直す回数の上限（ふつうは 1 回で落ち着く。長さが行き来しても止める） */
export const MAX_AXIS_REBUILDS = 2;

/**
 * ビューの大きさから引いて値の軸の長さを見積もる、値の軸の外（名前の欄・凡例・軸の文字・余白）の長さ (px)。ハーネスの 32 の見本では
 * 40〜150px で、真ん中が 70〜90px
 */
const OUTSIDE_AXIS_ESTIMATE = 80;

/**
 * 最初の組み立てに使う値の軸の長さ。入力（ビューの大きさと棒の向き）だけで見積もり、前に描いた長さは使わない（同じデータ・大きさ・書式なら、
 * それまでの大きさの順によらず同じ絵になる）。見積もりがずれても、描いて知らせた長さで組み直す（needsAxisRebuild）
 */
export function initialValueAxisLength(viewport: powerbi.IViewport, settings: VisualFormattingSettingsModel): number {
    const horizontal = String(settings.chart.orientation?.value?.value ?? ORIENTATIONS.vertical) === ORIENTATIONS.horizontal;
    return Math.max(1, (horizontal ? viewport.width : viewport.height) - OUTSIDE_AXIS_ESTIMATE);
}

type AxisBasis = Pick<ViewModel, "valueAxisRoom" | "roundsByLength" | "roundCount" | "builtAxisLength" | "orientation">;

/**
 * 描いて知らせた値の軸の長さ length で組み直すか。値の軸にラベルの余白を取るとき（valueAxisRoom）は、組み立てに使った長さから 1px 以上
 * 変わったとき。範囲を長さで丸めるとき（roundsByLength）は、length で決まる目盛りの本数の目安が、丸めに使った本数と違うとき
 * （同じなら、組み直しても範囲も目盛りも変わらない）
 */
export function needsAxisRebuild(viewModel: AxisBasis, length: number): boolean {
    if (!(length > 0)) return false;
    if (viewModel.valueAxisRoom && Math.abs(length - (viewModel.builtAxisLength ?? 0)) >= 1) return true;
    return !!viewModel.roundsByLength && recommendedTickCount(length, axisOf(viewModel)) !== viewModel.roundCount;
}

/**
 * 値の軸の目盛りの本数の目安。範囲を描く範囲の長さで丸めたとき（roundsByLength）は、丸めに使った本数（roundCount）。ふつうは描いた
 * 長さで決まる本数と同じで、組み直しの上限で止まって違っても、範囲と目盛りを同じ本数から出す。ほかは tickTargetOf（「目盛りの本数 (目安)」か描いた長さ）
 */
export function valueTickTarget(viewModel: AxisBasis, tickCount: number, length: number): number {
    return viewModel.roundsByLength && viewModel.roundCount ? viewModel.roundCount : tickTargetOf(tickCount, length, axisOf(viewModel));
}

const axisOf = (viewModel: AxisBasis): "vertical" | "horizontal" => (viewModel.orientation === ORIENTATIONS.horizontal ? "horizontal" : "vertical");
