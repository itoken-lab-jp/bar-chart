/**
 * 棒の形（角丸）。どの端（上・下・左・右）を丸めるかと半径を渡し、棒の太さと長さに収まる半径で丸める。
 * どのグラフでも同じ形にする
 */

/** 長方形（左上と大きさ） */
export interface Rect {
    x: number;
    y: number;
    width: number;
    height: number;
}

/** 棒の端（丸める側） */
export type CornerSide = "top" | "bottom" | "left" | "right";

/** 角を丸める端と半径 */
export interface Corner {
    sides: CornerSide[];
    radius: number;
}

/** 実際に描く角の半径。棒の太さの半分と、長さ（両端を丸めるなら長さの半分）まで。丸めなければ 0 */
export function cornerRadiusOf(r: Rect, corner: Corner | null): number {
    if (!corner || !corner.sides.length) return 0;
    const has = (side: CornerSide) => corner.sides.includes(side);
    const alongHeight = (has("top") ? 1 : 0) + (has("bottom") ? 1 : 0);
    const alongWidth = (has("left") ? 1 : 0) + (has("right") ? 1 : 0);
    return Math.max(
        0,
        Math.min(
            corner.radius,
            alongHeight ? Math.min(r.width / 2, r.height / alongHeight) : Infinity,
            alongWidth ? Math.min(r.height / 2, r.width / alongWidth) : Infinity
        )
    );
}

/**
 * 棒（区画）の path。corner があれば、その端（sides）の角だけを丸める（二次曲線）。
 * corner が無い・半径が 0 なら長方形
 */
export function barPath(r: Rect, corner: Corner | null): string {
    const { x, y, width: w, height: h } = r;
    const has = (side: CornerSide) => corner?.sides.includes(side) ?? false;
    const k = cornerRadiusOf(r, corner);
    if (!corner || k <= 0) return `M ${x},${y} h ${w} v ${h} h ${-w} Z`;
    const x2 = x + w;
    const y2 = y + h;
    // 角ごとの半径（左上から時計回り）。丸めない角は 0（二次曲線が点に縮む）
    const tl = has("top") || has("left") ? k : 0;
    const tr = has("top") || has("right") ? k : 0;
    const br = has("bottom") || has("right") ? k : 0;
    const bl = has("bottom") || has("left") ? k : 0;
    return (
        `M ${x + tl},${y} L ${x2 - tr},${y} Q ${x2},${y} ${x2},${y + tr} ` +
        `L ${x2},${y2 - br} Q ${x2},${y2} ${x2 - br},${y2} ` +
        `L ${x + bl},${y2} Q ${x},${y2} ${x},${y2 - bl} ` +
        `L ${x},${y + tl} Q ${x},${y} ${x + tl},${y} Z`
    );
}

/**
 * 値の向きの端だけを丸めた棒の path（縦棒は上か下、横棒は右か左）。半径は棒の太さの半分と長さまで。
 * r が 0 以下なら長方形
 */
export function valueEndBarPath(left: number, top: number, width: number, height: number, r: number, end: CornerSide): string {
    return barPath({ x: left, y: top, width, height }, r > 0 ? { sides: [end], radius: r } : null);
}
