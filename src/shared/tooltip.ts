import powerbi from "powerbi-visuals-api";

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

/**
 * ツールヒントを出す（move なら動かす）。ツールヒントが切られているときは何もしない。
 * identities は、ドリルスルーやレポート ページのツールヒントが対象の行を知るための ID
 */
export function showTooltipAt(
    service: powerbi.extensibility.ITooltipService,
    root: HTMLElement,
    clientX: number,
    clientY: number,
    move: boolean,
    dataItems: powerbi.extensibility.VisualTooltipDataItem[],
    identities: powerbi.visuals.ISelectionId[]
): void {
    if (!service.enabled()) return;
    const options = { coordinates: toRootCoordinates(clientX, clientY, root), isTouchEvent: false, dataItems, identities };
    if (move) service.move(options);
    else service.show(options);
}
