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

/** カテゴリ 1 つの棒のまとまりの幅（いちばん左の棒の左端から、いちばん右の棒の右端まで）。系列 1 本なら棒の幅 */
export function spanOf(cluster: ClusterLayout): number {
    const { offsets, barWidth } = cluster;
    return offsets.length > 1 ? offsets[offsets.length - 1] - offsets[0] + barWidth : barWidth;
}

/** 階層の上のレベルの 1 区切り。start〜end は並び（categoryGroups）の添字で、end を含む */
export interface LevelRun {
    start: number;
    end: number;
    text: string;
}

/**
 * 階層の上のレベル（level 番目）で、同じ親が続くところを 1 つにまとめる。
 * 親は上のレベルから level 番目までの並びで見る。上が違えば、同じ名前（別の年の Q1 など）でも別の区切り。
 * keys（元の値のキー）があればそれで見分ける。表示が同じでも値が違えば別の区切り
 */
export function levelRunsOf(levels: string[][], level: number, keys?: string[][]): LevelRun[] {
    const runs: LevelRun[] = [];
    let key: string | null = null;
    levels.forEach((path, i) => {
        const next = JSON.stringify((keys?.[i] ?? path).slice(0, level + 1));
        if (key === next) {
            runs[runs.length - 1].end = i;
        } else {
            runs.push({ start: i, end: i, text: path[level] ?? "" });
            key = next;
        }
    });
    return runs;
}
