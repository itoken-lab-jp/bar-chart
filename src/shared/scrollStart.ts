/**
 * スクロールの最初の位置。中身がはみ出してスクロールするとき、開いたときにどこから見せるか。
 *
 * - 見る人が自分でスクロールするまでは、描き直すたびに当て直す。Power BI は開いた直後に大きさを変えながら
 *   何度も update を呼ぶので、最初の 1 回だけ当てると、そのあとの大きさの変化でずれる
 * - 見る人がスクロールしたあとは、書式や大きさが変わっても戻さない
 * - 当て直すのは、スクロールする箱が作り直されたとき、設定を変えたとき、はみ出さなくなったとき（見る人の位置は
 *   そこで消える）と、末尾のときに中身の形（カテゴリ・列の並び）が変わったとき。先頭のときは形が変わっても
 *   当て直さない（ブラウザーがそのまま残す位置を変えない）
 */
import powerbi from "powerbi-visuals-api";
import * as React from "react";

export const SCROLL_STARTS = {
    start: "start",
    end: "end",
} as const;

/** 書式ペインの選択肢（先頭・末尾）。末尾は最後のカテゴリの側（縦に並べれば右端、横に並べれば下端） */
export const SCROLL_START_ITEMS: powerbi.IEnumMember[] = [
    { value: SCROLL_STARTS.start, displayName: "先頭" },
    { value: SCROLL_STARTS.end, displayName: "末尾" },
];

export type ScrollTarget = "start" | "end";

/** スクロールする量（px）。中身が窓に収まれば 0 */
export function scrollOffsetOf(target: ScrollTarget, content: number, view: number): number {
    return target === "end" ? Math.max(0, content - view) : 0;
}

/** 見る人のスクロールとみなす差（px）。当てた位置はブラウザーが丸めるので、読み戻した値と比べる */
const MOVED_PX = 1;

/** 前の描画で覚えたこと */
export interface ScrollMemory {
    /** スクロールする箱（作り直されたら別のもの） */
    box: unknown;
    axis: "x" | "y";
    target: ScrollTarget;
    shape: string;
    /** 見る人がスクロールした */
    moved: boolean;
    /** 最後に当てた位置（読み戻した値） */
    placed: number;
}

/** 今の描画で見えるもの。position は当てる前の位置 */
export interface ScrollNow {
    box: unknown;
    axis: "x" | "y";
    target: ScrollTarget;
    shape: string;
    position: number;
    content: number;
    view: number;
}

/**
 * 当てた位置から見る人が動かしたか。中身が縮んでブラウザーが端へ詰めた（当てた位置より手前の、いちばん奥）ときは、
 * 見る人の操作ではない
 */
export function movedFrom(placed: number, position: number, content: number, view: number): boolean {
    if (Math.abs(position - placed) <= MOVED_PX) return false;
    const max = Math.max(0, content - view);
    return !(position < placed && position >= max - MOVED_PX);
}

/** 描き直したときの判断。offset が数なら、そこへ当てる（当てたら placed を読み戻した値にする） */
export function nextScroll(memory: ScrollMemory | null, now: ScrollNow): { memory: ScrollMemory; offset: number | null } {
    const fresh =
        memory === null ||
        memory.box !== now.box ||
        memory.axis !== now.axis ||
        memory.target !== now.target ||
        (now.target === SCROLL_STARTS.end && memory.shape !== now.shape);
    // はみ出さなければ、見る人の位置はもう無い。次にはみ出したときは選んだ位置から
    const fits = now.content - now.view < MOVED_PX;
    // scroll イベントより先に描き直しが来ても、当てる前の位置で見る人のスクロールを拾う
    const moved = !fresh && !fits && (memory.moved || movedFrom(memory.placed, now.position, now.content, now.view));
    const base = { box: now.box, axis: now.axis, target: now.target, shape: now.shape };
    if (moved) return { memory: { ...base, moved: true, placed: memory?.placed ?? 0 }, offset: null };
    const offset = scrollOffsetOf(now.target, now.content, now.view);
    return { memory: { ...base, moved: false, placed: offset }, offset };
}

export interface ScrollStartOptions {
    axis: "x" | "y";
    target: ScrollTarget;
    /** 中身の形（カテゴリ・列の並び）。末尾のとき、変わったら見る人のスクロールを捨てて当て直す */
    shape: string;
}

/**
 * スクロールする箱に ref と onScroll を付けて使う。箱が無い描画（はみ出さないとき）でも毎回呼んでよい
 */
export function useScrollStart<T extends HTMLElement>(options: ScrollStartOptions): { ref: React.RefObject<T>; onScroll: () => void } {
    const ref = React.useRef<T>(null);
    const memory = React.useRef<ScrollMemory | null>(null);

    React.useLayoutEffect(() => {
        const box = ref.current;
        if (!box) {
            memory.current = null;
            return;
        }
        const x = options.axis === "x";
        const next = nextScroll(memory.current, {
            box,
            axis: options.axis,
            target: options.target,
            shape: options.shape,
            position: x ? box.scrollLeft : box.scrollTop,
            content: x ? box.scrollWidth : box.scrollHeight,
            view: x ? box.clientWidth : box.clientHeight,
        });
        if (next.offset !== null) {
            if (x) box.scrollLeft = next.offset;
            else box.scrollTop = next.offset;
            next.memory.placed = x ? box.scrollLeft : box.scrollTop;
        }
        memory.current = next.memory;
    });

    const onScroll = React.useCallback(() => {
        const box = ref.current;
        const m = memory.current;
        if (!box || !m || box !== m.box || m.moved) return;
        const x = m.axis === "x";
        if (movedFrom(m.placed, x ? box.scrollLeft : box.scrollTop, x ? box.scrollWidth : box.scrollHeight, x ? box.clientWidth : box.clientHeight)) {
            m.moved = true;
        }
    }, []);

    return { ref, onScroll };
}
