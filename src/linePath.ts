/**
 * 折れ線の形（補間）とマーカーの形の SVG パス。標準の「線」カードの補間の種類と、
 * 「マーカー」カードの型に合わせる（docs/barchart-standard-gaps.md の「標準の動き」）。
 * 曲線の作り方は d3-shape の curveMonotoneX・curveCardinal・curveStep と同じ（依存は足さない）。
 */

export interface XY {
    x: number;
    y: number;
}

export interface LineShape {
    /** "linear" | "smooth" | "step" */
    interpolation: string;
    /** スムーズのときの種類。"monotone" | "cardinal" */
    smoothing: string;
    /** カーディナルのテンション（0〜1）。1 で直線、0 で最もなめらか */
    tension: number;
    /** ステップのときの段の位置。"before" | "center" | "after" */
    stepPosition: string;
    /** ステップの段と段をつなぐ線を出すか。省略するとつなぐ */
    stepConnect?: boolean;
    /**
     * 段のつなぎを出さないステップで、値ごとの線をこの長さ（カテゴリの軸の px）で点を中心に引く（「段の幅」が「棒の幅」のとき）。
     * 省略か 0 なら、これまでどおりカテゴリの間隔いっぱいに引く
     */
    stepLevelWidth?: number;
}

const round = (v: number) => Math.round(v * 100) / 100;
type Fmt = (x: number, y: number) => string;
const pointXY: Fmt = (x, y) => `${round(x)},${round(y)}`;
/** 横棒では、カテゴリの軸（y）を曲線の独立変数にするので、x と y を入れ替えて作り、書き出すときに戻す */
const pointYX: Fmt = (x, y) => `${round(y)},${round(x)}`;
const swapped = (points: XY[]) => points.map((p) => ({ x: p.y, y: p.x }));

/**
 * ステップのときに、並びの最初の点の手前と最後の点の先へ延ばす範囲（カテゴリの軸の座標）。
 * カテゴリの間隔（step）の半分を渡すと、段が変わる位置（隣のカテゴリとのすき間の真ん中）と同じところまで延びる。
 * bounds（プロットの端）の外へは延ばさない。ステップ以外では延ばさない
 */
function stepEndsOf(shape: LineShape, first: XY, last: XY, extension: number, bounds?: [number, number]): [number, number] | null {
    if (shape.interpolation !== "step" || !(extension > 0)) return null;
    const [min, max] = bounds ?? [-Infinity, Infinity];
    return [Math.min(first.x, Math.max(min, first.x - extension)), Math.max(last.x, Math.min(max, last.x + extension))];
}

/**
 * 途切れのない点の並び 1 つぶんの線（先頭の M を含む）。
 * horizontal なら、点は上から下へ並び（横棒のカテゴリ）、補間もその向きで行う。
 * extension はステップのときに両端を延ばす長さ（点が 1 つでも、その長さの段を引く）、bounds は延ばせる範囲
 */
export function linePath(points: XY[], shape: LineShape, horizontal = false, extension = 0, bounds?: [number, number]): string {
    if (!points.length) return "";
    const pts = horizontal ? swapped(points) : points;
    const pt = horizontal ? pointYX : pointXY;
    const first = pts[0];
    const last = pts[pts.length - 1];
    const ends = stepEndsOf(shape, first, last, extension, bounds);
    if (shape.interpolation === "step" && shape.stepConnect === false) {
        const barWidth = shape.stepLevelWidth ?? 0;
        if (barWidth > 0) {
            // 棒の幅：点（カテゴリの中心）から左右に棒の幅の半分ずつ。ステップの位置と端の延長は効かない
            return pts.map((p) => `M ${pt(p.x - barWidth / 2, p.y)} L ${pt(p.x + barWidth / 2, p.y)}`).join(" ");
        }
        return stepLevels(pts, shape.stepPosition, ends ?? [first.x, last.x])
            .map(([from, to, y]) => `M ${pt(from, y)} L ${pt(to, y)}`)
            .join(" ");
    }
    if (ends) {
        const middle = pts.length > 1 ? ` ${segmentsOf(pts, shape, pt)}` : "";
        return `M ${pt(ends[0], first.y)} L ${pt(first.x, first.y)}${middle} L ${pt(ends[1], last.y)}`;
    }
    const head = `M ${pt(first.x, first.y)}`;
    if (pts.length === 1) return head;
    return `${head} ${segmentsOf(pts, shape, pt)}`;
}

/**
 * 線と基準線（値 0 の位置）のあいだを塗る領域（網掛け領域）。線と同じ補間で辺を作り、
 * 基準線に下ろして閉じる。baseline は値の軸の座標（縦棒なら y、横棒なら x）。
 * extension・bounds はステップのときに両端を延ばす長さと範囲（線と同じ範囲を塗る）
 */
export function areaPath(
    points: XY[],
    baseline: number,
    shape: LineShape,
    horizontal = false,
    extension = 0,
    bounds?: [number, number]
): string {
    if (!points.length) return "";
    const pts = horizontal ? swapped(points) : points;
    const pt = horizontal ? pointYX : pointXY;
    const first = pts[0];
    const last = pts[pts.length - 1];
    const ends = stepEndsOf(shape, first, last, extension, bounds);
    if (!ends && pts.length < 2) return "";
    const middle = pts.length > 1 ? ` ${segmentsOf(pts, shape, pt)}` : "";
    if (ends) {
        return `M ${pt(ends[0], baseline)} L ${pt(ends[0], first.y)} L ${pt(first.x, first.y)}${middle} L ${pt(ends[1], last.y)} L ${pt(ends[1], baseline)} Z`;
    }
    return `M ${pt(first.x, baseline)} L ${pt(first.x, first.y)}${middle} L ${pt(last.x, baseline)} Z`;
}

function segmentsOf(points: XY[], shape: LineShape, pt: Fmt): string {
    if (shape.interpolation === "step") return stepSegments(points, shape.stepPosition, pt);
    if (shape.interpolation === "smooth" && points.length > 2) {
        return shape.smoothing === "cardinal" ? cardinalSegments(points, shape.tension, pt) : monotoneSegments(points, pt);
    }
    return points.slice(1).map((p) => `L ${pt(p.x, p.y)}`).join(" ");
}

/**
 * 段のつなぎを出さないステップの、値ごとの水準の線（[始まり, 終わり, 高さ]）。
 * つなぐときと同じ位置で段が変わる：中央は隣の点との真ん中、「次の値より前」は前の点、「次の値より後」は次の点。
 * ends は並びの両端（延ばした先）
 */
function stepLevels(points: XY[], position: string, ends: [number, number]): Array<[number, number, number]> {
    const n = points.length;
    return points.map((p, i) => {
        let from: number;
        let to: number;
        if (position === "before") {
            from = i > 0 ? points[i - 1].x : ends[0];
            to = i < n - 1 ? p.x : ends[1];
        } else if (position === "after") {
            from = i > 0 ? p.x : ends[0];
            to = i < n - 1 ? points[i + 1].x : ends[1];
        } else {
            from = i > 0 ? (points[i - 1].x + p.x) / 2 : ends[0];
            to = i < n - 1 ? (p.x + points[i + 1].x) / 2 : ends[1];
        }
        return [from, to, p.y];
    });
}

function stepSegments(points: XY[], position: string, pt: Fmt): string {
    const parts: string[] = [];
    for (let i = 1; i < points.length; i++) {
        const a = points[i - 1];
        const b = points[i];
        if (position === "before") {
            parts.push(`L ${pt(a.x, b.y)} L ${pt(b.x, b.y)}`);
        } else if (position === "after") {
            parts.push(`L ${pt(b.x, a.y)} L ${pt(b.x, b.y)}`);
        } else {
            const mid = (a.x + b.x) / 2;
            parts.push(`L ${pt(mid, a.y)} L ${pt(mid, b.y)} L ${pt(b.x, b.y)}`);
        }
    }
    return parts.join(" ");
}

const sign = (v: number) => (v < 0 ? -1 : 1);

/** 単調な 3 次補間（Steffen の方法。d3 の curveMonotoneX と同じ）。山と谷を行き過ぎない */
function monotoneSegments(points: XY[], pt: Fmt): string {
    const n = points.length;
    const slopes: number[] = new Array(n).fill(0);
    for (let i = 1; i < n - 1; i++) {
        const h0 = points[i].x - points[i - 1].x;
        const h1 = points[i + 1].x - points[i].x;
        const s0 = h0 ? (points[i].y - points[i - 1].y) / h0 : 0;
        const s1 = h1 ? (points[i + 1].y - points[i].y) / h1 : 0;
        const p = h0 + h1 ? (s0 * h1 + s1 * h0) / (h0 + h1) : 0;
        slopes[i] = (sign(s0) + sign(s1)) * Math.min(Math.abs(s0), Math.abs(s1), 0.5 * Math.abs(p)) || 0;
    }
    const endSlope = (a: XY, b: XY, t: number) => {
        const h = b.x - a.x;
        return h ? (3 * (b.y - a.y) / h - t) / 2 : t;
    };
    slopes[0] = endSlope(points[0], points[1], slopes[1]);
    slopes[n - 1] = endSlope(points[n - 2], points[n - 1], slopes[n - 2]);
    const parts: string[] = [];
    for (let i = 0; i < n - 1; i++) {
        const a = points[i];
        const b = points[i + 1];
        const dx = (b.x - a.x) / 3;
        parts.push(`C ${pt(a.x + dx, a.y + dx * slopes[i])} ${pt(b.x - dx, b.y - dx * slopes[i + 1])} ${pt(b.x, b.y)}`);
    }
    return parts.join(" ");
}

/** カーディナル スプライン（d3 の curveCardinal と同じ。両端の接線は 0） */
function cardinalSegments(points: XY[], tension: number, pt: Fmt): string {
    const n = points.length;
    const k = (1 - Math.max(0, Math.min(1, tension))) / 6;
    const parts: string[] = [];
    for (let i = 0; i < n - 1; i++) {
        const a = points[i];
        const b = points[i + 1];
        const before = i > 0 ? points[i - 1] : b;
        const after = i + 2 < n ? points[i + 2] : a;
        const c1 = { x: a.x + k * (b.x - before.x), y: a.y + k * (b.y - before.y) };
        const c2 = { x: b.x + k * (a.x - after.x), y: b.y + k * (a.y - after.y) };
        parts.push(`C ${pt(c1.x, c1.y)} ${pt(c2.x, c2.y)} ${pt(b.x, b.y)}`);
    }
    return parts.join(" ");
}

/**
 * マーカー 1 つの形。size はサイズ (px)。塗りの形として返す（× と ＋ とダッシュも太さのある形）。
 * 半径はサイズの 0.75 倍（Desktop で標準のサイズ 20 の円が直径 30px ほどだった。1.13 まではサイズの半分 + 0.5）。
 * ダッシュは棒を横切る向きに引く（縦棒では横、横棒 horizontal では縦。棒に目標の印を打つ使い方）
 */
export function markerPath(shape: string, cx: number, cy: number, size: number, horizontal = false): string {
    const r = Math.max(1, size * 0.75);
    const pt = pointXY;
    const poly = (coords: Array<[number, number]>) =>
        `M ${coords.map(([dx, dy]) => pt(cx + dx, cy + dy)).join(" L ")} Z`;
    const across = (coords: Array<[number, number]>) => poly(horizontal ? coords.map(([dx, dy]) => [dy, dx] as [number, number]) : coords);
    switch (shape) {
        case "square":
            return poly([[-r, -r], [r, -r], [r, r], [-r, r]]);
        case "diamond": {
            const d = r * 1.3;
            return poly([[0, -d], [d, 0], [0, d], [-d, 0]]);
        }
        case "triangle": {
            const d = r * 1.3;
            return poly([[0, -d], [d * 0.87, d * 0.5], [-d * 0.87, d * 0.5]]);
        }
        case "shortDash":
            return across([[-r, -r * 0.3], [r, -r * 0.3], [r, r * 0.3], [-r, r * 0.3]]);
        case "longDash":
            return across([[-r * 2, -r * 0.3], [r * 2, -r * 0.3], [r * 2, r * 0.3], [-r * 2, r * 0.3]]);
        case "plus": {
            const t = r * 0.3;
            return poly([
                [-t, -r], [t, -r], [t, -t], [r, -t], [r, t], [t, t],
                [t, r], [-t, r], [-t, t], [-r, t], [-r, -t], [-t, -t],
            ]);
        }
        case "cross": {
            // ＋ を 45 度回した形
            const t = r * 0.3;
            const c = Math.SQRT1_2;
            const rotate = ([x, y]: [number, number]): [number, number] => [(x - y) * c, (x + y) * c];
            return poly(
                ([
                    [-t, -r], [t, -r], [t, -t], [r, -t], [r, t], [t, t],
                    [t, r], [-t, r], [-t, t], [-r, t], [-r, -t], [-t, -t],
                ] as Array<[number, number]>).map(rotate)
            );
        }
        default:
            // 円（既定）
            return `M ${pt(cx - r, cy)} A ${round(r)},${round(r)} 0 1 0 ${pt(cx + r, cy)} A ${round(r)},${round(r)} 0 1 0 ${pt(cx - r, cy)} Z`;
    }
}
