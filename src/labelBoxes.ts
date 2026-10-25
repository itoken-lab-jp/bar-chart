/**
 * ラベルの重なり。棒のデータ ラベル・合計ラベル・線のデータ ラベルは、先に置いたラベルと重なるものを描かない（標準と同じ）。
 * 重なりは描くものの枠で見る：背景を描くラベルは背景の四角、描かないラベルは文字の四角に 1px 足したもの。文字の縁を付けるなら縁の分も入れる。
 * 置いた枠はマスに分けて持ち、新しい枠は掛かるマスの枠とだけ比べる（ラベルが何万あっても、1 つを見るのに全部と比べない）
 */
import { Halo, haloStyle } from "./shared/halo";

/** ラベルの枠（プロットの座標） */
export interface LabelBox {
    x0: number;
    x1: number;
    y0: number;
    y1: number;
    /**
     * 合計ラベルの積み上げ（カテゴリとレイヤー）。同じ積み上げの棒の中のデータ ラベルとは、合計ラベルの背景がデータ ラベルの文字に掛かるときだけ重なるとみなす
     * （データ ラベルの文字は棒に入り、合計ラベルは背景ごと積み上げの端の外に置くので、ふつうは掛からない。
     *  背景の余白が触れるだけで省くと、既定の 9pt でも一番外の棒のラベルが消える。背景を描かない合計ラベルとは比べない）
     */
    stack?: string;
    /** 合計ラベルが背景を描く（枠は背景の四角） */
    background?: boolean;
}

/** 積み上げの棒のデータ ラベル。stack はその積み上げ（合計ラベルの stack と同じ書き方）、text は文字の四角 */
export interface StackedLabel {
    stack: string;
    text: LabelBox;
}

/** 背景を描かないラベルの、文字の四角に足す余白（ほかの物とくっつきすぎないように） */
export const HIT_PAD = 1;

/**
 * 重なりを見る枠。背景を描くなら背景の四角、描かないなら文字の四角に HIT_PAD を足したもの。
 * halo は文字の縁が文字の四角の外へ出る幅（haloOutsetOf）。縁も描くものなので枠に入れる（背景からはみ出す分も）
 */
export function hitBoxOf(text: LabelBox, background: LabelBox | null, halo = 0): LabelBox {
    const around = (pad: number) => ({ x0: text.x0 - pad, x1: text.x1 + pad, y0: text.y0 - pad, y1: text.y1 + pad });
    if (!background) return around(HIT_PAD + halo);
    if (!(halo > 0)) return background;
    const inked = around(halo);
    return {
        x0: Math.min(background.x0, inked.x0),
        x1: Math.max(background.x1, inked.x1),
        y0: Math.min(background.y0, inked.y0),
        y1: Math.max(background.y1, inked.y1),
    };
}

/** 文字の縁が文字の四角の外へ出る幅（縁の太さの半分）。fills はラベルの行の文字の色で、縁を描く行（文字と縁の色が違う行）が無ければ 0 */
export function haloOutsetOf(halo: Halo, fills: readonly string[]): number {
    return fills.some((fill) => haloStyle(halo, fill).strokeWidth) ? halo.width / 2 : 0;
}

/** 枠どうしが重なるか。1px までの重なりは見ない（背景の余白が触れるだけのもの） */
export function boxesOverlap(a: LabelBox, b: LabelBox): boolean {
    return a.x0 < b.x1 - 1 && b.x0 < a.x1 - 1 && a.y0 < b.y1 - 1 && b.y0 < a.y1 - 1;
}

/** マスの一辺 (px) */
const CELL = 64;
/** これより多くのマスに掛かる枠と、座標が数でない枠は、マスに分けずにいつも比べる */
const MAX_CELLS = 1024;

/** 置いたラベルの枠 */
export interface LabelBoxes {
    /** 置いた枠（置いた順）。ドリルの位置のように、全部を見るところで使う */
    readonly boxes: readonly LabelBox[];
    /**
     * 置いた枠のどれかと重なるか。stacked（積み上げの棒のデータ ラベル）を渡すと、その積み上げの合計ラベルとは、
     * 合計ラベルの背景がデータ ラベルの文字に掛かるかだけを見る
     */
    overlaps(box: LabelBox, stacked?: StackedLabel): boolean;
    /** 重なりを見ずに置く */
    add(box: LabelBox): void;
    /** 重ならなければ置いて true、重なれば置かずに false（stacked は overlaps と同じ） */
    place(box: LabelBox, stacked?: StackedLabel): boolean;
    /** これまでに比べた枠の数（試験で、全部と比べていないことを見る） */
    readonly comparisons: number;
}

/** 枠が掛かるマスの範囲。マスに分けない枠は null */
function cellsOf(box: LabelBox): { cx0: number; cx1: number; cy0: number; cy1: number } | null {
    const cx0 = Math.floor(Math.min(box.x0, box.x1) / CELL);
    const cx1 = Math.floor(Math.max(box.x0, box.x1) / CELL);
    const cy0 = Math.floor(Math.min(box.y0, box.y1) / CELL);
    const cy1 = Math.floor(Math.max(box.y0, box.y1) / CELL);
    if (![cx0, cx1, cy0, cy1].every(Number.isFinite)) return null;
    return (cx1 - cx0 + 1) * (cy1 - cy0 + 1) > MAX_CELLS ? null : { cx0, cx1, cy0, cy1 };
}

export function createLabelBoxes(): LabelBoxes {
    const boxes: LabelBox[] = [];
    const cells = new Map<string, LabelBox[]>();
    /** マスに分けない枠 */
    const loose: LabelBox[] = [];
    let comparisons = 0;
    const hits = (list: readonly LabelBox[] | undefined, box: LabelBox, stacked: StackedLabel | undefined): boolean => {
        for (const b of list ?? []) {
            comparisons++;
            if (stacked && b.stack === stacked.stack) {
                // 同じ積み上げの合計ラベル。文字は text の内側なので、背景が文字に掛かれば text を含む box のマスにもある
                if (b.background && boxesOverlap(b, stacked.text)) return true;
                continue;
            }
            if (boxesOverlap(b, box)) return true;
        }
        return false;
    };
    const overlaps = (box: LabelBox, stacked?: StackedLabel): boolean => {
        const range = cellsOf(box);
        if (!range) return hits(boxes, box, stacked);
        if (hits(loose, box, stacked)) return true;
        for (let cx = range.cx0; cx <= range.cx1; cx++) {
            for (let cy = range.cy0; cy <= range.cy1; cy++) {
                if (hits(cells.get(`${cx},${cy}`), box, stacked)) return true;
            }
        }
        return false;
    };
    const add = (box: LabelBox): void => {
        boxes.push(box);
        const range = cellsOf(box);
        if (!range) {
            loose.push(box);
            return;
        }
        for (let cx = range.cx0; cx <= range.cx1; cx++) {
            for (let cy = range.cy0; cy <= range.cy1; cy++) {
                const key = `${cx},${cy}`;
                const list = cells.get(key);
                if (list) list.push(box);
                else cells.set(key, [box]);
            }
        }
    };
    return {
        boxes,
        overlaps,
        add,
        place: (box, stacked) => {
            if (overlaps(box, stacked)) return false;
            add(box);
            return true;
        },
        get comparisons() {
            return comparisons;
        },
    };
}

/**
 * ラベルの置き方 1 つ。box は重なりを見る枠、draw は残したときだけ呼ぶ。
 * stacked は積み上げの棒のデータ ラベルで、その積み上げの合計ラベルとは、背景が文字に掛かるかだけを見る
 */
export interface LabelPlacement<T> {
    box: LabelBox;
    stacked?: StackedLabel;
    draw: () => T;
}

/**
 * 重なりを見て描くかを決めるラベル。order の小さいものから置く（同じなら出した順）。
 * fallback は、この置き方が重なるときに、同じ順番のうちに（次のラベルを置く前に）試す代わりの置き方を作る
 * （集合の棒の「自動」で、外に置くと重なるラベルを棒の中に置く。標準と同じ）。null なら代わりは無い。それも重なれば出さない。
 * 重なったときだけ呼ぶ（重ならないラベルの分の置き方を測らない）
 */
export interface LabelCandidate<T> extends LabelPlacement<T> {
    order: readonly number[];
    fallback?: () => LabelPlacement<T> | null;
}

/** order の比べ方（先の数から順に比べる） */
function compareOrder(a: readonly number[], b: readonly number[]): number {
    for (let k = 0; k < Math.max(a.length, b.length); k++) {
        const d = (a[k] ?? 0) - (b[k] ?? 0);
        if (d !== 0) return d;
    }
    return 0;
}

/**
 * 候補を order の順に置き、置いた枠（placed。先に置いた合計ラベルなど）と重なるものを省く。重なる候補に代わりの置き方（fallback）があれば、
 * 次の候補の前にそれを試し、重ならなければそちらを置く。残したものの枠は placed に足す。
 * 返すのは残したものの描画で、並びは候補を出した順（描く順と重なりの前後は変えない）
 */
export function keepUnoverlapped<T>(candidates: readonly LabelCandidate<T>[], placed: LabelBoxes): T[] {
    const kept = new Array<LabelPlacement<T> | null>(candidates.length).fill(null);
    const ranked = candidates.map((c, k) => ({ c, k })).sort((a, b) => compareOrder(a.c.order, b.c.order) || a.k - b.k);
    for (const { c, k } of ranked) {
        if (placed.place(c.box, c.stacked)) {
            kept[k] = c;
            continue;
        }
        const alternative = c.fallback?.();
        if (alternative && placed.place(alternative.box, alternative.stacked)) kept[k] = alternative;
    }
    return kept.flatMap((option) => (option ? [option.draw()] : []));
}

/**
 * 端で入らないラベルを、カテゴリの向き（縦棒は左右、横棒は上下）に描く範囲 [lo, hi] へ収まるだけ内側へ寄せる量。
 * from・to は描くもの（背景か文字。文字の縁も）の端。入っていれば 0。範囲より長いラベルは範囲の真ん中に置く（両側にはみ出す）。
 * 標準は合計ラベルを中心のまま描くが、ここは値の軸を延ばして合計ラベルを描く範囲の中に置くので、中心のままだと軸の目盛りの数字に重なる。
 * 標準はビジュアルの端まで使うとみられるが、第 2 Y 軸・右に移した Y 軸の目盛りに掛けないよう、描く範囲で止める
 */
export function shiftIntoSpan(from: number, to: number, lo: number, hi: number): number {
    if (!(hi > lo)) return 0;
    if (to - from > hi - lo) return (lo + hi) / 2 - (from + to) / 2;
    if (from < lo) return lo - from;
    if (to > hi) return hi - to;
    return 0;
}

/** 描くものの枠（寄せる量を測る枠）。背景を描くなら背景（と、はみ出す文字の縁）、描かないなら文字の四角と縁。hitBoxOf の 1px の余白は入れない */
export function inkBoxOf(text: LabelBox, background: LabelBox | null, halo = 0): LabelBox {
    if (background) return hitBoxOf(text, background, halo);
    return { x0: text.x0 - halo, x1: text.x1 + halo, y0: text.y0 - halo, y1: text.y1 + halo };
}
