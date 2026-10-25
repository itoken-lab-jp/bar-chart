/**
 * 配列の最小・最大。Math.min(...xs)・Math.max(...xs) は配列を引数に広げるので、要素が多い（Node 22 で約 12.5 万を超える）と
 * 引数の数の上限で RangeError を投げる。こちらは要素の数に依らない。値は Math.min(...xs)・Math.max(...xs) と同じ
 * （空なら Infinity・-Infinity、NaN があれば NaN、0 と -0 の扱いも同じ）
 */

export function minOf(xs: ArrayLike<number>): number {
    let m = Infinity;
    for (let i = 0; i < xs.length; i++) m = Math.min(m, xs[i]);
    return m;
}

export function maxOf(xs: ArrayLike<number>): number {
    let m = -Infinity;
    for (let i = 0; i < xs.length; i++) m = Math.max(m, xs[i]);
    return m;
}
