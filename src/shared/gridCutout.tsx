import * as React from "react";

/** 目盛線を棒の形で切り抜く mask の ID の通し番号（1 つの文書に複数のグラフを描いても重ならないように） */
let gridMaskSeq = 0;

/** 目盛線を切り抜く mask の ID。描き直しても変えない。写す棒のまとまりには `${id}-bars` を付ける */
export function useGridMaskId(prefix: string): string {
    const ref = React.useRef("");
    if (!ref.current) ref.current = `${prefix}-grid-mask-${++gridMaskSeq}`;
    return ref.current;
}

/**
 * 目盛線を棒の形で切り抜く mask。半透明に描いた棒の下では、先に描いた目盛線が透けて棒の前にあるように見える。
 * 棒のまとまり（barsId）を写し、色を黒・不透明に変えて mask から抜く。
 * 背景を塗らずに隠すので、ビジュアルの背景が透明でも壊れない
 */
export function gridCutout(maskId: string, barsId: string): React.ReactNode {
    return (
        <defs>
            <filter id={`${maskId}-solid`} x="-1%" y="-1%" width="102%" height="102%" colorInterpolationFilters="sRGB">
                {/* 色を黒に、少しでも塗った所は不透明にする */}
                <feColorMatrix type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 255 0" />
            </filter>
            <mask id={maskId} maskUnits="userSpaceOnUse" x={-1e5} y={-1e5} width={2e5} height={2e5}>
                <rect x={-1e5} y={-1e5} width={2e5} height={2e5} fill="white" />
                <use href={`#${barsId}`} filter={`url(#${maskId}-solid)`} />
            </mask>
        </defs>
    );
}
