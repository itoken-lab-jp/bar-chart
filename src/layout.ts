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
 * - 重複（overlap）のときは、系列間のスペースを重ねる割合（%）として読む（標準と同じ）。0 なら並べたまま、100 なら全部が
 *   帯いっぱいの幅で重なる。棒の幅は「帯 ÷ 系列の数」と帯の幅のあいだを割合で決め、残りを等しくずらす
 */
export function clusterLayout(bandWidth: number, count: number, seriesSpacing: number, maxBarWidth: number, overlap = false): ClusterLayout {
    if (overlap && count > 1) {
        const o = Math.max(0, Math.min(1, seriesSpacing / 100));
        const share = 1 / count + (1 - 1 / count) * o;
        let span = bandWidth;
        let width = Math.max(1, span * share);
        if (maxBarWidth > 0 && width > maxBarWidth) {
            width = maxBarWidth;
            span = width / share;
        }
        const step = (span - width) / (count - 1);
        return { barWidth: width, offsets: Array.from({ length: count }, (_, k) => (k - (count - 1) / 2) * step) };
    }
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

/** 比較レイヤーの並び。layers の順（手前から奥）に、棒の幅と、系列の棒の中心から見た位置を持つ */
export interface LayerLayout {
    width: number;
    offsets: number[];
}

/**
 * 比較レイヤーの棒の並び。系列の棒 1 本ぶんの幅（slotWidth）に、count 枚の同じ幅の棒を少しずつずらして収める。
 * 重なり overlap（0〜1）が 0 なら隣り合わせ（集合と同じ並び）、1 なら全部が同じ位置に重なる。
 * 手前の棒（添字 0）は rightFront なら右（横棒は下）の端、そうでなければ左（上）の端に置き、奥ほど反対の側へずれる
 */
export function layerLayout(slotWidth: number, count: number, overlap: number, rightFront: boolean): LayerLayout {
    if (count <= 1) return { width: slotWidth, offsets: [0] };
    const shift = 1 - Math.max(0, Math.min(1, overlap));
    const width = slotWidth / (1 + (count - 1) * shift);
    const offsets = Array.from({ length: count }, (_, k) => {
        const position = rightFront ? count - 1 - k : k;
        return -slotWidth / 2 + width / 2 + position * width * shift;
    });
    return { width, offsets };
}

/**
 * いちばん奥の不透明度 back（0〜1）まで、手前から同じ割合で薄くする。k は手前から何枚目か（0 が手前）
 */
export function layerOpacity(k: number, count: number, back: number): number {
    if (k <= 0 || count <= 1) return 1;
    if (k >= count - 1) return Math.max(0, Math.min(1, back));
    return 1 - (1 - Math.max(0, Math.min(1, back))) * Math.min(1, k / (count - 1));
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
