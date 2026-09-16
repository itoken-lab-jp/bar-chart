/**
 * 折れ線の形（補間）とマーカーの形の SVG パス。標準の「線」カードの補間の種類と、
 * 「マーカー」カードの型に合わせる（docs/barchart-standard-gaps.md の「標準の動き」、#92）。
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
}

const round = (v: number) => Math.round(v * 100) / 100;
const pt = (x: number, y: number) => `${round(x)},${round(y)}`;

/** 途切れのない点の並び 1 つぶんの線（先頭の M を含む） */
export function linePath(points: XY[], shape: LineShape): string {
    if (!points.length) return "";
    const head = `M ${pt(points[0].x, points[0].y)}`;
    if (points.length === 1) return head;
    return `${head} ${segmentsOf(points, shape)}`;
}

/**
 * 線と基準線（値 0 の高さ）のあいだを塗る領域（網掛け領域）。線と同じ補間で上の辺を作り、
 * 基準線に下ろして閉じる
 */
export function areaPath(points: XY[], baselineY: number, shape: LineShape): string {
    if (points.length < 2) return "";
    const first = points[0];
    const last = points[points.length - 1];
    return `M ${pt(first.x, baselineY)} L ${pt(first.x, first.y)} ${segmentsOf(points, shape)} L ${pt(last.x, baselineY)} Z`;
}

function segmentsOf(points: XY[], shape: LineShape): string {
    if (shape.interpolation === "step") return stepSegments(points, shape.stepPosition);
    if (shape.interpolation === "smooth" && points.length > 2) {
        return shape.smoothing === "cardinal" ? cardinalSegments(points, shape.tension) : monotoneSegments(points);
    }
    return points.slice(1).map((p) => `L ${pt(p.x, p.y)}`).join(" ");
}

function stepSegments(points: XY[], position: string): string {
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
function monotoneSegments(points: XY[]): string {
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
function cardinalSegments(points: XY[], tension: number): string {
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
 * 半径はサイズの 0.75 倍（Desktop で標準のサイズ 20 の円が直径 30px ほどだった。1.13 まではサイズの半分 + 0.5）
 */
export function markerPath(shape: string, cx: number, cy: number, size: number): string {
    const r = Math.max(1, size * 0.75);
    const poly = (coords: Array<[number, number]>) =>
        `M ${coords.map(([dx, dy]) => pt(cx + dx, cy + dy)).join(" L ")} Z`;
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
            return poly([[-r, -r * 0.3], [r, -r * 0.3], [r, r * 0.3], [-r, r * 0.3]]);
        case "longDash":
            return poly([[-r * 2, -r * 0.3], [r * 2, -r * 0.3], [r * 2, r * 0.3], [-r * 2, r * 0.3]]);
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
