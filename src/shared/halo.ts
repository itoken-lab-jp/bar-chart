/**
 * 文字の縁。SVG の文字のまわりに縁の色の線を引き、棒や線の上でも読めるようにする（縁を先に塗って、文字を上に描く）。
 * 書式ペインの部品は formatCards.ts の labelHaloParts
 */
import { sameColor } from "./color";

/** 縁の設定。幅 0 なら付けない */
export interface Halo {
    color: string;
    width: number;
}

/** 縁を付けない設定（既定） */
export const NO_HALO: Halo = { color: "#FFFFFF", width: 0 };

/** 書式の値（表示・カラー・幅）から縁の設定を作る。幅は 0〜10 px */
export function haloOf(show: boolean | undefined, color: string | undefined, width: number | undefined): Halo {
    return { color: color || "#FFFFFF", width: show ? Math.max(0, Math.min(10, width ?? 3)) : 0 };
}

/**
 * 文字に縁を付けるスタイル。文字と縁が同じ色なら付けない（棒の中の白い文字に白い縁を付けても、文字が太るだけ）。
 * fill は文字の色
 */
export function haloStyle(halo: Halo, fill: string): { stroke?: string; strokeWidth?: number; strokeLinejoin?: "round"; paintOrder?: string } {
    if (!(halo.width > 0) || sameColor(fill, halo.color)) return {};
    return { stroke: halo.color, strokeWidth: halo.width, strokeLinejoin: "round", paintOrder: "stroke" };
}
