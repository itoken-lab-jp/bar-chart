import { scaleLinear } from "d3-scale";

import { parseNumberText } from "./dataValue";
import { maxOf } from "./minMax";

/** 目盛りの本数の下限。1 本では大きさが読めない（標準も、2 本を切ると目安を 1 つ上げて取り直す） */
const MIN_TICKS = 2;

/**
 * 目盛りの値。本数が max を超えず、数字が同じ表示にならない（丸めで 0.005 と 0.01 がどちらも 0.01 になる、など）ように、
 * d3 の ticks の目安を max から下げていく（ticks の引数は目安で、0〜4 に 3 を渡すと 5 本返る）。
 * 下げて max 本以下になった目盛りが 2 本を切るときは、それより前の目安で取れた 2 本以上で数字の重ならない目盛りのうち、一番少ないものにする
 * （max 本を超えてもよい。-400〜200 に 3 本の目安なら、0 の 1 本にせず、200 おきの 4 本）。
 * どの目安でも 2 本以上にならない範囲（幅が 0 など）は、下げきった目盛りのまま
 */
export function ticksUpTo(domain: [number, number], max: number, labelsOf: (values: number[]) => string[]): number[] {
    let values: number[] = [];
    let fewest: number[] | null = null;
    for (let count = Math.max(1, max); count >= 1; count--) {
        values = scaleLinear().domain(domain).ticks(count);
        const labels = labelsOf(values);
        const distinct = new Set(labels).size === labels.length;
        if (values.length <= max && distinct) return values.length < MIN_TICKS && fewest ? fewest : values;
        if (distinct && values.length >= MIN_TICKS && !(fewest && fewest.length <= values.length)) fewest = values;
    }
    return values;
}

/** 「目盛りの本数 (目安)」の入力を本数にする。全角（「５」など）も読む。空・数でない・1 以下は 0（自動）。多すぎる指定は 20 本まで */
export function tickCountOf(text: string | undefined): number {
    const n = Math.round(parseNumberText(text) ?? NaN);
    return Number.isFinite(n) && n >= 2 ? Math.min(20, n) : 0;
}

/**
 * 値の軸の目盛りの本数の上限。標準（powerbi-visuals-utils-chartutils の getRecommendedNumberOfTicksForYAxis・ForXAxis）と同じ。
 * 縦の軸は高さ、横の軸は幅で決める
 */
export function recommendedTickCount(length: number, axis: "vertical" | "horizontal"): number {
    const [small, medium] = axis === "vertical" ? [150, 300] : [300, 500];
    return length < small ? 3 : length < medium ? 5 : 8;
}

/**
 * 値の軸の目盛りの本数の目安。書式の「目盛りの本数 (目安)」の本数（tickCountOf。0 は自動）があればその本数、無ければ描く範囲の長さで
 * 決める（recommendedTickCount）。範囲を丸める本数（niceDomain）にも同じ本数を使う。描く範囲の長さで決まる本数で丸めるのは、標準と同じ
 * （Desktop で測った）。「目盛りの本数 (目安)」は標準に無い欄で、その本数で範囲を丸めることには標準の裏付けが無い（目盛りの本数にそろえただけ）
 */
export function tickTargetOf(userCount: number, length: number, axis: "vertical" | "horizontal"): number {
    return userCount || recommendedTickCount(length, axis);
}

/**
 * 範囲 lower〜upper を切りのいい値に丸める（d3 の nice(count) と同じ）。count は目盛りと同じ本数の目安（tickTargetOf）。
 * 描く範囲の長さで決まる本数で丸めるのは標準と同じ（高さ 150px 未満の縦の軸なら 3 本の目安で、-520〜900 は -1,000〜1,000。
 * 300px 以上なら 8 本の目安で -600〜1,000）。「目盛りの本数 (目安)」の本数で丸めることには、標準の裏付けが無い。
 * count が無ければ d3 の既定（10 本の目安）
 */
export function niceDomain(lower: number, upper: number, count?: number): [number, number] {
    return scaleLinear().domain([lower, upper]).nice(count).domain() as [number, number];
}

/**
 * 書式の最小値・最大値の文字を数に直す。桁区切りのカンマ（1,000）と、全角の数字・マイナス・カンマ・小数点（「－５０」「１，０００」）も読む。
 * 空・数でなければ null（自動）
 */
export function boundOf(text: string | undefined | null): number | null {
    return parseNumberText(text);
}

/**
 * 横の値の軸の目盛りを、数字が重ならない本数まで減らして取り直す（どの目盛りも切りのいい数のまま間隔が広がる）。
 * target から 1 本ずつ減らし、隣の目盛りとの間隔（spacing）が数字に要る幅（need）以上になる本数で止める。2 本まで減らしても重なれば 2 本。
 * 減らして 2 本を切る（範囲に切りのいい値が 1 つしか入らない）ときは、その前の本数の目盛りにする（数字は重なるが、1 本では大きさが読めない）。
 * 縦の値の軸は rowTicks（間隔で止めず、数字の行が入る数で本数の目安を下げる）
 */
export function fitTicks<T>(ticksFor: (count: number) => T[], target: number, spacing: (ticks: T[]) => number, need: (ticks: T[]) => number): T[] {
    let previous: T[] | null = null;
    for (let count = target; count > 2; count--) {
        const ticks = ticksFor(count);
        if (ticks.length < MIN_TICKS) return previous ?? ticks;
        if (spacing(ticks) >= need(ticks)) return ticks;
        previous = ticks;
    }
    const last = ticksFor(2);
    return last.length < MIN_TICKS && previous ? previous : last;
}

/**
 * 縦の値の軸の目盛り。本数の目安 target を、目盛りを引く範囲の長さ length（px）に数字の行（高さ rowHeight。目盛りの文字の大きさの px）が
 * 入る数（length ÷ rowHeight の切り捨て）までに下げ、その目安で取る（ticksFor。目安以下の本数の切りのいい目盛り）。
 * 取った目盛りは、隣との間隔が数字の行より狭くても、それ以上は減らさない（標準の集合縦棒を Desktop で測った形と同じ。下の数）。
 * 目安は 2 より下げない。2 の目安で 2 本を切る（範囲に切りのいい値が 1 つしか入らない）ときは、目安を上げて 2 本以上取れる所で取る（target まで）。
 * 標準で測った形（目盛りの数字の行が重なりそうな形）：
 * - -400〜800（5 の目安で丸めた範囲）・描く範囲の高さ 177px・目盛り 24pt（32px）：入る行は 5。5 本以下の目盛りは 0・500 の 2 本
 * - -300〜800（8 の目安）・36pt（48px）：目盛りを引く範囲は約 290〜300px（測りの縮尺の読み方で違う）で、入る行は 6。-200〜800 の 200 おきの
 *   6 本（間隔 約 53〜54.5px で、文字の 1.4 倍より狭い）。描く範囲は 300px を切るので、描く範囲の高さで決める本数の目安では 8 にならない。
 *   標準がこの形を 8 の目安にするわけは分かっていない
 * - -1,000〜1,000（5 の目安）・36pt：-1,000・0・1,000（入る行が 4 で、500 おきの 5 本は取らない。12・16pt は 500 おき）
 * 前の作りは、隣の間隔が文字の 1.4 倍以上になるまで本数を減らした（fitTicks と同じ決まり）。それだと 2 つ目の形が 0・500 の 2 本になった
 */
export function rowTicks<T>(ticksFor: (count: number) => T[], target: number, length: number, rowHeight: number): T[] {
    const rows = rowHeight > 0 && length >= 0 ? Math.floor(length / rowHeight + 1e-9) : target;
    const count = Math.max(MIN_TICKS, Math.min(target, rows));
    const ticks = ticksFor(count);
    if (ticks.length >= MIN_TICKS) return ticks;
    for (let up = count + 1; up <= target; up++) {
        const more = ticksFor(up);
        if (more.length >= MIN_TICKS) return more;
    }
    return ticks;
}

/**
 * 目盛りの列の幅を、描く目盛りの数字の幅にそろえて組み直す回数の上限（見積もった幅で組んだあとの回数。落ち着かないときに、選んだ幅で
 * 組む最後の 1 回は数えない）。ふつうは 1 回で落ち着く
 */
export const COLUMN_PASSES = 2;
/** 列の幅を比べるときの許し (px) */
const COLUMN_EPS = 1e-6;

/** 目盛りの列の幅を決める組み直しの途中の状態（nextColumnTry） */
export interface ColumnTry {
    /** この回に使う目盛りの列の幅（軸ごと） */
    widths: number[];
    /** これまでに組んだ回の、使った幅と、その回で描く目盛りの数字の幅（軸ごと） */
    tried: Array<{ widths: number[]; drawn: number[] }>;
    /** 決めた幅で組む最後の回か（もう組み直さない） */
    final: boolean;
}

/**
 * 値の軸の目盛りの列の幅を、描く目盛りの数字の幅（一番広い数字）にそろえる。列の幅で描く範囲の幅が変わり、名前の回し方・下の帯・
 * 描く範囲の長さ・目盛りの本数が変わるので、見積もった幅で組んでから、描く目盛りの数字の幅で組み直す。この関数は、組んだ回の結果から
 * 次に組む幅を返す（null なら、その回の組み方のまま描く）。current はその回の状態（はじめの回は undefined）、used はその回に使った列の幅、
 * drawn はその回で描く目盛りの数字の幅（軸ごと。数字を描かない軸は 0）。
 * - drawn が used と同じ（どの軸も）なら落ち着いた：null
 * - 違えば drawn の幅で組み直す（COLUMN_PASSES 回まで）
 * - 上限まで組み直しても落ち着かない（幅を変えると描く範囲の長さが変わって、描く目盛りが行き来する）ときは、組んだ回のうち、描く数字が
 *   列に入る（どの軸も drawn ≤ used）回で列の幅の合計が一番狭い回の幅で、もう一度組んで描く（描く数字は列に入り、列は描く数字より広いことがある）。
 *   入る回が無ければ、組んだ回の描く数字の幅の、軸ごとの一番広い幅で組んで描く
 */
export function nextColumnTry(current: ColumnTry | undefined, used: number[], drawn: number[]): ColumnTry | null {
    if (current?.final) return null;
    const same = (a: number[], b: number[]) => a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) <= COLUMN_EPS);
    if (same(used, drawn)) return null;
    const tried = [...(current?.tried ?? []), { widths: used, drawn }];
    if (tried.length <= COLUMN_PASSES) return { widths: drawn, tried, final: false };
    const fitting = tried.filter((t) => t.drawn.every((d, i) => d <= t.widths[i] + COLUMN_EPS));
    const sum = (values: number[]) => values.reduce((a, b) => a + b, 0);
    const pick = fitting.length
        ? fitting.reduce((a, b) => (sum(b.widths) < sum(a.widths) - COLUMN_EPS ? b : a)).widths
        : used.map((_, i) => maxOf(tried.map((t) => t.drawn[i])));
    return same(pick, used) ? null : { widths: pick, tried, final: true };
}

/** endMarginOf で、要る余白へ決め直す回数の上限（目盛りが変わらなければ 1 回で落ち着く） */
const END_MARGIN_PASSES = 4;
/** endMarginOf で、落ち着かないときに数字の入る一番小さい余白を詰める二分法の回数（幅 40px の間なら 0.04px まで） */
const END_MARGIN_SEARCH = 10;
/** 余白を比べるときの許し (px) */
const END_MARGIN_EPS = 1e-6;

/**
 * 値の軸の端の側にとる描く範囲の外の余白 (px) を、実際に描く目盛りの数字がビューに入るように決める（横向きの値の軸の右の余白）。
 * needAt(margin) は、余白を margin にしたとき（描く範囲の長さが決まり、その長さで描く目盛りも決まる）に、その目盛りの数字を
 * ビューに入れるのに要る余白。目盛りが同じなら、その余白でちょうど入る値を返す。
 * base から始めて要る余白へ決め直し、目盛りが変われば、変わったあとの目盛りでもう一度見る（END_MARGIN_PASSES 回まで）。
 * 落ち着かない（広げると目盛りが減って要る余白が下がり、戻すと増える）ときは、試した余白のうち数字が入る一番小さいものと、
 * それより小さく入らなかったものの間を二分法で詰め、入る側を採る（描く目盛りのために要る分だけ広げ、描かない目盛りのためには広げない）。
 * 余白は base〜max（max は描く範囲に残す長さの下限から出す）。max でも入らなければ max（数字の端は切れる）
 */
export function endMarginOf(base: number, max: number, needAt: (margin: number) => number): number {
    const top = Math.max(base, max);
    const clamp = (v: number) => Math.min(top, Math.max(base, Number.isFinite(v) ? v : base));
    /** 余白 margin で、描く目盛りの数字が入るか（上限に掛かれば、入らなくても入った扱い） */
    const fitsAt = (margin: number) => clamp(needAt(margin)) <= margin + END_MARGIN_EPS;
    let fit = top;
    const misses: number[] = [];
    let margin = base;
    for (let pass = 0; pass < END_MARGIN_PASSES; pass++) {
        const need = clamp(needAt(margin));
        if (Math.abs(need - margin) <= END_MARGIN_EPS) return margin;
        if (need < margin) fit = Math.min(fit, margin);
        else misses.push(margin);
        margin = need;
    }
    let low = maxOf(misses.filter((m) => m < fit));
    if (!Number.isFinite(low)) return fit;
    let high = fit;
    for (let k = 0; k < END_MARGIN_SEARCH; k++) {
        const mid = (low + high) / 2;
        if (fitsAt(mid)) high = mid;
        else low = mid;
    }
    return high;
}
