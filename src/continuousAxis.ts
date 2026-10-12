"use strict";

/**
 * X 軸の種類「連続」。日付・数値のカテゴリを、値の間隔どおりの位置に置く。
 * 位置は 0〜1 の比（いちばん小さい値が 0、いちばん大きい値が 1）で持ち、描く側がプロットの幅に写す。
 * 目盛りは切りのいい日付（年・月・日）か数値に置く
 */
import { scaleLinear, scaleTime } from "d3-scale";

export interface ContinuousAxis {
    /** カテゴリの値（並びは categoryGroups と同じ。日付はミリ秒） */
    values: number[];
    isDate: boolean;
    /** 値が整数だけ（整数の列か、届いた値がどれも整数）。目盛りも整数だけにする */
    integer?: boolean;
    min: number;
    max: number;
    /** 数値の目盛りの書き方（モデルの書式）。日付は年・月・日で決める */
    formatNumber: (value: number) => string;
}

export interface ContinuousTick {
    /** 位置（0〜1。min が 0、max が 1） */
    ratio: number;
    label: string;
}

/** 値が日付・数値として読めれば数（日付はミリ秒）、読めなければ null */
export function axisValueOf(raw: unknown, isDate: boolean): number | null {
    if (raw === null || raw === undefined || raw === "") return null;
    if (isDate) {
        const time = raw instanceof Date ? raw.getTime() : new Date(raw as string | number).getTime();
        return Number.isFinite(time) ? time : null;
    }
    const value = typeof raw === "number" ? raw : Number(raw);
    return Number.isFinite(value) ? value : null;
}

/** 値の位置（0〜1）。値が 1 つしかない（min = max）ときは真ん中 */
export function positionOf(axis: Pick<ContinuousAxis, "min" | "max">, value: number): number {
    return axis.max > axis.min ? (value - axis.min) / (axis.max - axis.min) : 0.5;
}

/**
 * 棒 1 本ぶんの帯の幅と、両端の余白（px）。隣り合う値のいちばん狭い間隔を帯にし、両端は帯の半分ずつ空ける
 * （いちばん端の棒がプロットからはみ出さない）。width はプロットの幅
 */
export function continuousBand(axis: ContinuousAxis, width: number): { band: number; edge: number } {
    const sorted = [...axis.values].sort((a, b) => a - b);
    let gap = Infinity;
    for (let k = 1; k < sorted.length; k++) {
        const d = sorted[k] - sorted[k - 1];
        if (d > 0 && d < gap) gap = d;
    }
    const span = axis.max - axis.min;
    // 値が 1 つ（か全部同じ）なら、プロットの半分を帯にする
    if (!(span > 0) || !Number.isFinite(gap)) return { band: width / 2, edge: width / 4 };
    // 帯 b のとき、値を置ける幅は width - b。いちばん狭い間隔 g（比）がちょうど帯になるのは b = g (width - b)
    const g = gap / span;
    const band = (g * width) / (1 + g);
    return { band, edge: band / 2 };
}

const DAY = 24 * 60 * 60 * 1000;
const pad2 = (n: number) => String(n).padStart(2, "0");

/**
 * 目盛り。本数は max まで。日付は年・月・日の切りのいい所に置き、間隔に合わせて「2026」「2026/04」「04/15」と書く
 * （年をまたぐ日の目盛りは年も付ける）
 */
export function continuousTicks(axis: ContinuousAxis, max: number): ContinuousTick[] {
    const count = Math.max(2, Math.floor(max));
    const single = (): ContinuousTick[] => [{ ratio: 0.5, label: axis.isDate ? dateLabel(new Date(axis.min), "day", true) : axis.formatNumber(axis.min) }];
    if (!(axis.max > axis.min)) return single();
    if (!axis.isDate) {
        // 整数の軸に 0.5 などの目盛りを置かない
        const values = scaleLinear()
            .domain([axis.min, axis.max])
            .ticks(count)
            .filter((v) => !axis.integer || Number.isInteger(v));
        return dedupe(values.map((v) => ({ ratio: positionOf(axis, v), label: axis.formatNumber(v) })), single);
    }
    // 日付は日より細かい目盛り（12 時間ごとなど）を置かない（日付までしか書かないので同じ文字が並ぶ）
    let dates = scaleTime()
        .domain([new Date(axis.min), new Date(axis.max)])
        .ticks(count)
        .filter((d) => d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0 && d.getMilliseconds() === 0);
    if (dates.length > count) dates = dates.filter((_, k) => k % Math.ceil(dates.length / count) === 0);
    // 目盛りの細かさ：全部が 1 月 1 日なら年、全部が 1 日なら月、それ以外は日
    const unit = dates.every((d) => d.getMonth() === 0 && d.getDate() === 1 && d.getHours() === 0)
        ? "year"
        : dates.every((d) => d.getDate() === 1 && d.getHours() === 0)
          ? "month"
          : "day";
    const multiYear = axis.max - axis.min > 365 * DAY || new Date(axis.min).getFullYear() !== new Date(axis.max).getFullYear();
    return dedupe(dates.map((d) => ({ ratio: positionOf(axis, d.getTime()), label: dateLabel(d, unit, multiYear) })), single);
}

/** 書いた文字が隣と同じ目盛りを落とす（書式で丸めると 1 と 1.5 が同じ「2」になる、など）。1 本も残らなければ fallback */
function dedupe(ticks: ContinuousTick[], fallback: () => ContinuousTick[]): ContinuousTick[] {
    const kept = ticks.filter((t, k) => k === 0 || t.label !== ticks[k - 1].label);
    return kept.length ? kept : fallback();
}

function dateLabel(d: Date, unit: "year" | "month" | "day", withYear: boolean): string {
    if (unit === "year") return String(d.getFullYear());
    if (unit === "month") return `${d.getFullYear()}/${pad2(d.getMonth() + 1)}`;
    return withYear ? `${d.getFullYear()}/${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}` : `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}`;
}
