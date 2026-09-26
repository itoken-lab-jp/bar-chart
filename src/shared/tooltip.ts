/**
 * 空白の表記。valueFormatter の既定は英語の "(Blank)" なので、日本語の Power BI の表記に合わせる
 */
export const BLANK_TEXT = "(空白)";

/**
 * クライアント座標を、ビジュアルのルート要素の内側の座標に直す。
 * tooltipService はルート基準の座標を受け取る（powerbi-visuals-utils-tooltiputils と同じ計算）
 */
export function toRootCoordinates(
    clientX: number,
    clientY: number,
    root: Pick<HTMLElement, "getBoundingClientRect" | "clientLeft" | "clientTop">
): [number, number] {
    const rect = root.getBoundingClientRect();
    return [clientX - rect.left - root.clientLeft, clientY - rect.top - root.clientTop];
}
