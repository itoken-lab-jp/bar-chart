/**
 * 斜め（-45°）の項目名の前の余白。斜めの名前は棒の中心から左下へ伸びるので、項目の並びの前を空けないと、
 * 先頭の名前の頭がビジュアルの左端の外へ出て切れる。
 */

/** 余白を求める項目名 1 つ。extent は名前が中心から左へ伸びる長さ（px）、ratio は中心の位置（並べる長さに対する比 0〜1） */
export interface RotatedLabel {
    extent: number;
    ratio: number;
}

/**
 * すべての名前の左端が、ビジュアルの左端から margin 以上内側に来るよう、項目の並びの前に空ける長さ（px）。
 * 項目は start（ビジュアルの左端から並べ始めるまで）から length の中に並べ、前を空けると残りの長さで並べ直すので、
 * 中心 = start + leadIn + (length − leadIn) × ratio として解く。上限はかけないので、呼ぶ側で抑える
 */
export function rotatedLabelLeadIn(labels: RotatedLabel[], start: number, length: number, margin: number): number {
    let need = 0;
    for (const { extent, ratio } of labels) {
        const deficit = margin + extent - start - length * ratio;
        if (deficit > 0 && ratio < 1) need = Math.max(need, deficit / (1 - ratio));
    }
    return need;
}
