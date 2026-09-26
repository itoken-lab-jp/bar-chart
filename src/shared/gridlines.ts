/**
 * グリッド線の線種の模様（stroke-dasharray）。標準に合わせ、点線は細かい点（丸い端と組み合わせる）、破線は 4px 刻み。
 * 「幅で拡大縮小」がオンなら、点線・破線の模様を線の幅に比例させる（標準と同じく、細い線では模様が細かくなる）。
 * 実線は null
 */
export function gridDashOf(style: string, width: number, scaleWithWidth: boolean): string | null {
    const w = Math.max(1, width);
    if (style === "dotted") return scaleWithWidth ? `${w} ${2 * w}` : "1 3";
    if (style === "dashed") return scaleWithWidth ? `${3 * w} ${3 * w}` : "4 4";
    return null;
}
