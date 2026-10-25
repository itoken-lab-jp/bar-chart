/**
 * スクロールで隠れた所も含めて、グラフ・表の全体を 1 枚の画像にする（「画像としてコピー」の capture。CopyImageButton）。
 *
 * - 大きさの上限 `fitsWholeImage`：画像は画面の IMAGE_SCALE 倍の画素で作る。canvas の 1 辺の上限（Chromium・Edge は 32767 画素）と、
 *   作る canvas の大きさ（1 画素 4 バイト。64M 画素で 256MB）を超えるときは、全体をあきらめて見えている範囲を写す
 * - 描き直して撮る `wholeChartImage`（項目の部分だけがスクロールするグラフ）：軸・タイトル・凡例はスクロールの外に置いてあるので、スクロールの箱を
 *   広げるだけでは軸の線・タイトル・凡例が元の幅のまま残る。同じグラフを、スクロールしない大きさで入れ物の中の見えない所に描き直し、
 *   それを chartImage で撮る（画面のグラフは動かさない）。描いたあとにまだスクロールするなら、隠れた分だけ伸ばして描き直す
 * HTML で描いた表は軸を外に持たないので、描き直さずに写しのスクロールの箱を広げて撮ればよい（大きさの上限だけをここから使う）
 */
import * as React from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";

import { ChartImage, IMAGE_SCALE, chartImage } from "./copyImage";

/** 画像の 1 辺の上限（画素） */
export const WHOLE_IMAGE_MAX_SIDE = 32767;
/** 画像の面積の上限（画素） */
export const WHOLE_IMAGE_MAX_AREA = 64 * 1024 * 1024;
/** 描き直して、まだスクロールするときに伸ばし直す回数の上限 */
const MAX_ROUNDS = 4;

/** 画面の大きさ（px）の全体を、画像にできるか */
export function fitsWholeImage(width: number, height: number, scale = IMAGE_SCALE): boolean {
    const w = Math.ceil(width * scale);
    const h = Math.ceil(height * scale);
    return w > 0 && h > 0 && w <= WHOLE_IMAGE_MAX_SIDE && h <= WHOLE_IMAGE_MAX_SIDE && w * h <= WHOLE_IMAGE_MAX_AREA;
}

export interface Size {
    width: number;
    height: number;
}

/**
 * スクロールの箱が隠している分（px。拡大する前）。隠していなければ 0。スクロールしない向き（overflow が auto・scroll でない）は数えない
 * （横にスクロールする箱は、横のスクロールバーの分だけ中身が箱の高さを越える）
 */
export function hiddenOf(scroller: Element | null): Size {
    const el = scroller as HTMLElement | null;
    if (!el) return { width: 0, height: 0 };
    const style = getComputedStyle(el);
    const scrolls = (overflow: string) => overflow === "auto" || overflow === "scroll";
    const width = scrolls(style.overflowX) ? el.scrollWidth - el.clientWidth : 0;
    const height = scrolls(style.overflowY) ? el.scrollHeight - el.clientHeight : 0;
    return { width: width > 1 ? width : 0, height: height > 1 ? height : 0 };
}

/**
 * スクロールしない大きさを、今の大きさと隠れた分から見積もる。隠れていなければ null。
 * 伸ばした大きさは整数に切り上げる（端数の大きさで描くと、伸ばし直しが 1px ずつ続くことがある）
 */
export function grownSize(viewport: Size, hidden: Size): Size | null {
    if (hidden.width <= 0 && hidden.height <= 0) return null;
    return { width: Math.ceil(viewport.width + hidden.width), height: Math.ceil(viewport.height + hidden.height) };
}

const nextFrame = () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

/**
 * 入れ物（root）の中の見えない所に、render(大きさ) を描いて画像にする。描いたあと、scroller で描いたものの中のスクロールの箱を探し、
 * 隠れた分があれば伸ばして描き直す。画面のスクロールの箱が隠していなければ（start が null）、大きすぎれば、伸ばし直しても
 * スクロールが消えなければ undefined（見えている範囲を写す）
 */
export async function wholeChartImage(
    root: HTMLElement,
    background: string,
    start: Size | null,
    render: (viewport: Size) => React.ReactElement,
    scroller: (host: HTMLElement) => Element | null
): Promise<ChartImage | null | undefined> {
    if (!start || !fitsWholeImage(start.width, start.height)) return undefined;
    const host = document.createElement("div");
    host.setAttribute("aria-hidden", "true");
    // 入れ物の左上に重ね、見えないように切る（位置と大きさは測れる。chartImage は clip-path を見ない）
    host.style.cssText = "position:absolute;left:0;top:0;clip-path:inset(50%);pointer-events:none";
    root.appendChild(host);
    const reactRoot = createRoot(host);
    try {
        let viewport = start;
        for (let round = 0; ; round++) {
            flushSync(() => reactRoot.render(render(viewport)));
            // 描いたあとに測って直す配置（useLayoutEffect・そこからの描き直し）を待つ
            await nextFrame();
            await nextFrame();
            const next = grownSize(viewport, hiddenOf(scroller(host)));
            if (!next) break;
            if (round + 1 >= MAX_ROUNDS || !fitsWholeImage(next.width, next.height)) return undefined;
            viewport = next;
        }
        const chart = host.firstElementChild;
        return chart instanceof HTMLElement ? await chartImage(chart, background) : null;
    } finally {
        reactRoot.unmount();
        host.remove();
    }
}
