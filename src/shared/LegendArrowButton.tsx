import * as React from "react";
import { CONTROL_RADIUS, CONTROL_TRANSITION, ControlColors } from "./controls";

/** 凡例のページを送る矢印 1 つ。x・y は矢印の中心、size は一辺（押せる範囲） */
export interface LegendArrow {
    x: number;
    y: number;
    size: number;
    /** 向き（上下に置いた凡例は左右、左右に置いた凡例は上下） */
    direction: "left" | "right" | "up" | "down";
    /** 押せるか（最初のページの「前」・最後のページの「次」は押せない） */
    enabled: boolean;
}

/** 入りきらない凡例のページ（標準と同じく矢印で送る）。入りきれば arrows は null */
export interface LegendPaging {
    pageCount: number;
    page: number;
    arrows: { prev: LegendArrow; next: LegendArrow } | null;
}

/** 幅（間を含む）の並びを、room に収まるページに分ける。1 つで room を超える物はそのページに 1 つだけ置く */
export function pagesByWidth(widths: number[], room: number, gap: number): number[][] {
    const pages: number[][] = [];
    let current: number[] = [];
    let used = 0;
    widths.forEach((w, i) => {
        const need = (current.length ? gap : 0) + w;
        if (current.length && used + need > room) {
            pages.push(current);
            current = [];
            used = 0;
        }
        used += (current.length ? gap : 0) + w;
        current.push(i);
    });
    if (current.length) pages.push(current);
    return pages;
}
/** 動きを減らす設定・強制カラー（Windows のハイコントラスト）か。読めない環境では false */
const mediaMatches = (query: string): boolean => {
    try {
        return typeof window !== "undefined" && !!window.matchMedia?.(query).matches;
    } catch {
        return false;
    }
};

/**
 * 凡例のページを送る矢印（三角）。形は Fluent 2 の枠なしのボタン：押せる範囲に角丸の地を持ち、マウスを乗せると hover の地と
 * 字の色の三角、押している間は pressed の地、押せないときは三角を disabled の色にする（Tab の並びからは外さない）。
 * キーボードで移ったときは字の色の 2px の枠（強制カラーでは Highlight）。色は colors（テーマの字と地から controlColorsOf で作ったもの）。
 * 押しても選択の解除（背景のクリック）には伝えない
 */
export function LegendArrowButton({
    arrow,
    color,
    colors,
    label,
    onPress,
}: {
    arrow: LegendArrow;
    /** 三角のふだんの色（凡例の文字と同じ） */
    color: string;
    colors: ControlColors;
    label: string;
    onPress: () => void;
}): React.JSX.Element {
    const [hover, setHover] = React.useState(false);
    const [pressed, setPressed] = React.useState(false);
    const [focusVisible, setFocusVisible] = React.useState(false);
    const h = arrow.size / 2;
    const t = h * 0.75;
    const points =
        arrow.direction === "left"
            ? `${arrow.x - t / 2},${arrow.y} ${arrow.x + t / 2},${arrow.y - t} ${arrow.x + t / 2},${arrow.y + t}`
            : arrow.direction === "right"
              ? `${arrow.x + t / 2},${arrow.y} ${arrow.x - t / 2},${arrow.y - t} ${arrow.x - t / 2},${arrow.y + t}`
              : arrow.direction === "up"
                ? `${arrow.x},${arrow.y - t / 2} ${arrow.x - t},${arrow.y + t / 2} ${arrow.x + t},${arrow.y + t / 2}`
                : `${arrow.x},${arrow.y + t / 2} ${arrow.x - t},${arrow.y - t / 2} ${arrow.x + t},${arrow.y - t / 2}`;
    const press = (e: React.SyntheticEvent) => {
        e.stopPropagation();
        if (arrow.enabled) onPress();
    };
    const transition = mediaMatches("(prefers-reduced-motion: reduce)") ? undefined : `fill ${CONTROL_TRANSITION}`;
    const active = arrow.enabled && (hover || pressed);
    const radius = parseFloat(CONTROL_RADIUS);
    return (
        <g
            className={`legend-arrow${arrow.enabled ? "" : " disabled"}`}
            role="button"
            aria-label={label}
            aria-disabled={!arrow.enabled}
            // 押せなくなっても Tab の並びから外さない（押した矢印が最後のページで消えて、フォーカスの行き先が無くならないように）
            tabIndex={0}
            style={{ cursor: arrow.enabled ? "pointer" : "default", outline: "none" }}
            onClick={press}
            onMouseEnter={() => setHover(true)}
            onMouseLeave={() => {
                setHover(false);
                setPressed(false);
            }}
            onMouseDown={() => setPressed(true)}
            onMouseUp={() => setPressed(false)}
            // キーで移ったときだけ枠を出す（マウスで押したときは出さない）
            onFocus={(e) => setFocusVisible(focusVisibleOf(e.currentTarget))}
            onBlur={() => setFocusVisible(false)}
            onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    // 押し続けで最後のページまで飛ばない
                    if (!e.repeat) press(e);
                }
            }}
        >
            {/* 押せる範囲（角丸の地）。ふだんは透明、マウスを乗せると hover、押している間は pressed */}
            <rect
                x={arrow.x - h}
                y={arrow.y - h}
                width={arrow.size}
                height={arrow.size}
                rx={radius}
                style={{ fill: active ? (pressed ? colors.pressed : colors.hover) : "transparent", transition }}
            />
            {focusVisible && (
                <rect
                    x={arrow.x - h + 1}
                    y={arrow.y - h + 1}
                    width={Math.max(0, arrow.size - 2)}
                    height={Math.max(0, arrow.size - 2)}
                    rx={radius}
                    pointerEvents="none"
                    style={{ fill: "none", stroke: mediaMatches("(forced-colors: active)") ? "Highlight" : colors.focus, strokeWidth: 2 }}
                />
            )}
            <polygon points={points} style={{ fill: !arrow.enabled ? colors.disabled : active ? colors.text : color, transition }} />
        </g>
    );
}

/** キーボードで移ったか（:focus-visible）。読めない環境ではキーボードとみなす */
function focusVisibleOf(element: Element): boolean {
    try {
        return element.matches(":focus-visible");
    } catch {
        return true;
    }
}
