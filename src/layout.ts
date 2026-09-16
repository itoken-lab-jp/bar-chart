"use strict";

/**
 * 棒の並びの計算。座標は「カテゴリの方向」で持ち、画面の x・y には描画側で写す
 * （縦棒ではカテゴリの方向が x。横棒に切り替えるときはここを変えずに写し方だけ変える）。
 */

export interface ClusterLayout {
    /** 棒 1 本の幅（カテゴリの方向） */
    barWidth: number;
    /** カテゴリの中心から見た、各系列の棒の中心の位置。凡例の順に左（小さい方）から */
    offsets: number[];
}

/**
 * 集合の棒の並び。カテゴリ 1 つぶんの帯（棒を置ける幅）に、系列の棒を凡例の順に並べる。
 *
 * - 系列が 1 本なら 1.4 までと同じ（帯いっぱい。最小 2px、最大幅があればそこで止める）
 * - 系列間のスペース (%) は、系列 1 本ぶんの枠に対するすき間。0 なら棒どうしが接する（標準の既定）
 * - 最大幅で棒が細くなったときは、すき間の比率を保ったまま帯の中央に寄せる
 */
export function clusterLayout(bandWidth: number, count: number, seriesSpacing: number, maxBarWidth: number): ClusterLayout {
    if (count <= 1) {
        let width = Math.max(2, bandWidth);
        if (maxBarWidth > 0) width = Math.min(width, maxBarWidth);
        return { barWidth: width, offsets: [0] };
    }
    const gap = Math.max(0, Math.min(0.9, seriesSpacing / 100));
    let slot = bandWidth / count;
    let width = Math.max(1, slot * (1 - gap));
    if (maxBarWidth > 0 && width > maxBarWidth) {
        width = maxBarWidth;
        slot = width / (1 - gap);
    }
    return {
        barWidth: width,
        offsets: Array.from({ length: count }, (_, k) => (k - (count - 1) / 2) * slot),
    };
}
