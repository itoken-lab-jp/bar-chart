import { scaleLinear } from "d3-scale";

/**
 * 目盛りの値。本数が max を超えず、数字が同じ表示にならない（丸めで 0.005 と 0.01 がどちらも 0.01 になる、など）ように、
 * d3 の ticks の目安を max から下げていく（ticks の引数は目安で、0〜4 に 3 を渡すと 5 本返る）
 */
export function ticksUpTo(domain: [number, number], max: number, labelsOf: (values: number[]) => string[]): number[] {
    let values: number[] = [];
    for (let count = Math.max(1, max); count >= 1; count--) {
        values = scaleLinear().domain(domain).ticks(count);
        const labels = labelsOf(values);
        if (values.length <= max && new Set(labels).size === labels.length) return values;
    }
    return values;
}

/** 「目盛りの本数 (目安)」の入力を本数にする。空・数でない・1 以下は 0（自動）。多すぎる指定は 20 本まで */
export function tickCountOf(text: string | undefined): number {
    // 全角の数字も読む（「５」など）
    const ascii = String(text ?? "").trim().replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0));
    const n = Math.round(Number(ascii));
    return Number.isFinite(n) && n >= 2 ? Math.min(20, n) : 0;
}

/**
 * 値の軸の目盛りの本数の上限。標準（powerbi-visuals-utils-chartutils の getRecommendedNumberOfTicksForYAxis・ForXAxis）と同じ。
 * 縦の軸は高さ、横の軸は幅で決める
 */
export function recommendedTickCount(length: number, axis: "vertical" | "horizontal"): number {
    const [small, medium] = axis === "vertical" ? [150, 300] : [300, 500];
    return length < small ? 3 : length < medium ? 5 : 8;
}

/** 書式の最小値・最大値の文字を数に直す。桁区切りのカンマ（1,000）も読む。空・数でなければ null（自動） */
export function boundOf(text: string | undefined | null): number | null {
    const trimmed = (text ?? "").trim().replace(/,/g, "");
    if (!trimmed) return null;
    const value = Number(trimmed);
    return Number.isFinite(value) ? value : null;
}
