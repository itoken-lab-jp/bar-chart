import powerbi from "powerbi-visuals-api";

/** 凡例を描く位置（辺と寄せ）。保存値は下の LEGEND_POSITION_ITEMS（標準と同じ値）で、ここへ読み替える */
export const LEGEND_POSITIONS = {
    topLeft: "topLeft",
    topCenter: "topCenter",
    topRight: "topRight",
    bottomLeft: "bottomLeft",
    bottomCenter: "bottomCenter",
    bottomRight: "bottomRight",
    leftTop: "leftTop",
    leftCenter: "leftCenter",
    leftBottom: "leftBottom",
    rightTop: "rightTop",
    rightCenter: "rightCenter",
    rightBottom: "rightBottom",
} as const;
export type LegendPosition = (typeof LEGEND_POSITIONS)[keyof typeof LEGEND_POSITIONS];

/**
 * 凡例の位置の保存値は、標準のビジュアルとレポートテーマと同じ値（Top・Bottom など）にする。
 * 基本テーマ（Fluent 2 は "Bottom"）やカスタムテーマの値がそのまま届く。左下・右下は標準に無いので独自の値
 */
const LEGEND_POSITION_PLACEMENTS: Record<string, LegendPosition> = {
    Top: LEGEND_POSITIONS.topLeft,
    TopCenter: LEGEND_POSITIONS.topCenter,
    TopRight: LEGEND_POSITIONS.topRight,
    Bottom: LEGEND_POSITIONS.bottomLeft,
    BottomCenter: LEGEND_POSITIONS.bottomCenter,
    BottomRight: LEGEND_POSITIONS.bottomRight,
    Left: LEGEND_POSITIONS.leftTop,
    LeftCenter: LEGEND_POSITIONS.leftCenter,
    LeftBottom: LEGEND_POSITIONS.leftBottom,
    Right: LEGEND_POSITIONS.rightTop,
    RightCenter: LEGEND_POSITIONS.rightCenter,
    RightBottom: LEGEND_POSITIONS.rightBottom,
};

/** 12 通り。既定は上詰め (左) */
export const LEGEND_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "Top", displayName: "上詰め (左)" },
    { value: "TopCenter", displayName: "上詰め (中央)" },
    { value: "TopRight", displayName: "上詰め (右)" },
    { value: "Bottom", displayName: "下詰め (左)" },
    { value: "BottomCenter", displayName: "下詰め (中央)" },
    { value: "BottomRight", displayName: "下詰め (右)" },
    { value: "Left", displayName: "左上" },
    { value: "LeftCenter", displayName: "左中央" },
    { value: "LeftBottom", displayName: "左下" },
    { value: "Right", displayName: "右上" },
    { value: "RightCenter", displayName: "右中央" },
    { value: "RightBottom", displayName: "右下" },
];

/** 保存値を標準の値にする。前の版の保存値（topLeft など）も読み替える。知らない値は undefined */
export function standardLegendPosition(value: unknown): string | undefined {
    if (typeof value !== "string") return undefined;
    if (value in LEGEND_POSITION_PLACEMENTS) return value;
    return Object.keys(LEGEND_POSITION_PLACEMENTS).find((k) => LEGEND_POSITION_PLACEMENTS[k] === value);
}

/** 保存値を、描画で使う位置（topLeft など）にする。知らない値は上詰め (左) */
export function legendPlacementValue(value: unknown): LegendPosition {
    return LEGEND_POSITION_PLACEMENTS[standardLegendPosition(value) ?? "Top"];
}

/**
 * 上下に置いた凡例の行数。保存値は数の形にしない（数の形の値は Desktop が数で保存して届けるため）。
 * 行数を足すときは、ここと LEGEND_ROW_ITEMS に足す
 */
export const LEGEND_ROWS = {
    one: "one",
    two: "two",
} as const;
const LEGEND_ROW_COUNTS: Record<string, number> = { [LEGEND_ROWS.one]: 1, [LEGEND_ROWS.two]: 2 };

/** 既定は 1 行 */
export const LEGEND_ROW_ITEMS: powerbi.IEnumMember[] = [
    { value: LEGEND_ROWS.one, displayName: "1 行" },
    { value: LEGEND_ROWS.two, displayName: "2 行" },
];

/** 保存値を行数にする。知らない値は 1 行 */
export function legendRowsOf(value: unknown): number {
    return (typeof value === "string" && LEGEND_ROW_COUNTS[value]) || 1;
}

export type LegendSide = "top" | "bottom" | "left" | "right";
export type LegendAlign = "start" | "center" | "end";

/** "topLeft" のような位置の値を、置く辺と寄せ方に分ける。知らない値は上詰め (左) */
export function legendPlacement(position: string): { side: LegendSide; align: LegendAlign } {
    const m = /^(top|bottom|left|right)(Left|Center|Right|Top|Bottom)$/.exec(position);
    if (!m) return { side: "top", align: "start" };
    const align: LegendAlign = m[2] === "Center" ? "center" : m[2] === "Left" || m[2] === "Top" ? "start" : "end";
    return { side: m[1] as LegendSide, align };
}

/** 凡例の行数を効かせるか（上下に置いたときだけ。左右は 1 列に並べる） */
export function legendRowsApply(position: unknown): boolean {
    const side = legendPlacement(legendPlacementValue(position)).side;
    return side === "top" || side === "bottom";
}

/**
 * 凡例の印の半径（px）。印の外形（四角の一辺・丸の直径）はこの 2 倍で、標準の集合縦棒の丸の印の直径に合わせる：
 * 文字の半分、小さくても 9px（13.5pt までは 9px、14pt で約 9.3px、16pt で約 10.7px、24pt で 16px、40pt で約 26.7px）。
 * 標準の丸は 14pt まで 9〜10px、16pt 11・24pt 16・32pt 21・40pt 28px だった（2026-10-20 の Desktop）
 */
export function legendMarkerRadius(fontPx: number): number {
    return Math.max(4.5, fontPx * 0.25);
}

/*
 * 凡例の行と箱の決まり。組み方（legendLayout.ts の layoutLegend）と、作品の描き方（字のベースライン・印と字のあいだ）が使う。
 * 標準の集合縦棒を、凡例を上・右・下に置いて 8〜40pt で測った値（2026-10-20 の Desktop）に合わせてある。標準は、行の上下の空きが
 * 文字の大きさによらずほぼ一定で、文字が大きくなっても行は文字の大きさの分だけ高くなる。標準を測り直して合わせるときは、ここの数を直す
 */

/**
 * 凡例の箱の端と行のあいだ（上下に置けば上と下、左右に置けば左と右）。左右に置いた凡例は、箱の外（描く範囲の側）にもこれだけ空ける。
 * 左右に置いた凡例は、上下にこの空きを取らずに並べる（LEGEND_SIDE_LIFT）が、入る行の数とページの分け方は箱の高さから上下にこれだけ除いて数える
 */
export const LEGEND_PAD = 4;
/** 凡例の行の高さのうち、文字の大きさに足す分 (px) */
export const LEGEND_ROW_EXTRA = 8;
/**
 * 上下に置いた凡例の行の高さ (px)。文字の大きさ ＋ 8px を四捨五入（8pt で 19px、10pt で 21px、20pt で 35px、40pt で 61px）。
 * 標準の一番上のグリッド線は、文字を 8pt から 40pt にすると 41px 下がり、行の高さの差（42px）と合う
 */
export function legendRowHeight(fontPx: number): number {
    return Math.round(fontPx + LEGEND_ROW_EXTRA);
}
/** 左右に置いた凡例の行の間隔が、上下に置いたときの行の高さより広い分 (px) */
export const LEGEND_SIDE_ROW_EXTRA = 4;
/**
 * 左右に置いた凡例の行の間隔 (px)。上下に置いたときの行の高さ ＋ 4px（10pt で 25px、20pt で 39px）。
 * 標準は 10pt で約 25px、20pt で約 39px だった（測った画面の縮みを、同じ画面で測った自作の行の間隔で補った値）
 */
export function legendSideRowHeight(fontPx: number): number {
    return legendRowHeight(fontPx) + LEGEND_SIDE_ROW_EXTRA;
}
/** 左右に置いた凡例のタイトルの行の高さ（文字の大きさに対する比。切り上げて px にする） */
export const LEGEND_SIDE_TITLE_RATIO = 1.5;
/**
 * 左右に置いた凡例のタイトルの行の高さ (px)。10pt で 20px、20pt で 40px。標準は、タイトルから 1 つ目の項目までが
 * 項目どうしの間隔より短く（10pt）、文字を大きくするとそれより速く広がる（20pt でほぼ同じ）
 */
export function legendSideTitleHeight(fontPx: number): number {
    return Math.ceil(fontPx * LEGEND_SIDE_TITLE_RATIO);
}
/**
 * 凡例の字のベースラインを、行の縦の真ん中から下げる高さ（文字の大きさに対する比）。印は行の縦の真ん中にそろえる。
 * 標準は、文字を大きくしても字の上の端がほとんど下がらず、下の端が文字の大きさの分だけ下がる（行の真ん中より字が下に寄る）
 */
export const LEGEND_TEXT_DROP = 0.45;
/** 凡例の字（項目・タイトル・まとまりの見出し）のベースライン。centerY は行の縦の真ん中 */
export function legendBaselineOf(centerY: number, fontPx: number): number {
    return centerY + fontPx * LEGEND_TEXT_DROP;
}
/** 凡例の字の下の端がベースラインから下りる深さ（文字の大きさに対する比。ほかの字の見積もりと同じ 0.22） */
export const LEGEND_TEXT_DESCENT = 0.22;
/**
 * 上下に置いた凡例の箱に、行の高さと上下の LEGEND_PAD のほかに足す高さ (px)。字の下の端（ベースライン ＋ 0.22 文字）が
 * 箱の下の端より下に出るときだけ、出る分を切り上げて足す（行・字・印の位置は変えず、箱の下を延ばす）。
 * 字の下の端は行の真ん中から 0.67 文字下で、行の高さの半分と LEGEND_PAD は文字 ＋ 8px の半分 ＋ 4px なので、文字がおよそ 35pt
 * （約 46.7px）を超えると足す（36pt で 1px、40pt で 2px、60pt で 6px）。行の高さを四捨五入するので、足す高さは文字の大きさにつれて
 * 行き来しながら増える（0.5pt 刻みでは 35pt まで 0 だが、34.7〜34.9pt でも 1px 足し、35.6〜35.8pt では 0 に戻る）
 */
export function legendDescentRoom(fontPx: number): number {
    const over = legendTextOverhang(fontPx, legendRowHeight(fontPx)) - LEGEND_PAD;
    // 計算の端数（1e-6 未満）で 1px 足さない
    return Math.max(0, Math.ceil(over - 1e-6));
}
/**
 * 行の字の下の端（ベースライン ＋ 0.22 文字）が、行の下の端より下に出る高さ (px。負なら行の中に収まる)。
 * 字の下の端は行の縦の真ん中から 0.67 文字下なので、行の高さの半分より 0.67 文字が大きいと出る
 */
export function legendTextOverhang(fontPx: number, rowHeight: number): number {
    return fontPx * (LEGEND_TEXT_DROP + LEGEND_TEXT_DESCENT) - rowHeight / 2;
}
/**
 * 左右に置いた凡例は、箱の上と下の LEGEND_PAD を取らずに寄せ、中身（タイトル・項目・矢印）をさらにこれだけ上げる (px)。
 * 標準の集合縦棒の凡例「右」「左」（上に寄せる）は、タイトルの字の上の端が、箱の上の 4px を取ったときより 4.6px 上（10pt・20pt）、
 * タイトルなしの 1 つ目の項目の字の上の端が 4.2px 上で、「右中央」は、タイトルと 1〜3 つ目の項目の 8 つの差（10pt・20pt）が
 * 空きを数えない真ん中より 0〜1.6px、平均 0.76px（項目だけの 6 つでは 0.6px）上だった
 * （2026-10-20 の Desktop）。上の空きを 0 にして 4px、残りの 0.5px は寄せ方によらず上げる。下に寄せるときも下の空きを取らない
 */
export const LEGEND_SIDE_LIFT = 0.5;
/*
 * 項目のあいだと、印と名前のあいだ。どちらも字の送り（名前の文字の四角）の端から取る。標準の集合縦棒の凡例「上」（タイトル「区分」、
 * 項目は 特別損益・売上・販管費）を 10・20・32pt で測った、字と印の画素の端どうしの値（2026-10-21 の Desktop）に合わせてある：
 * 項目のあいだ（前の名前の字の右の端 → 次の印の左の端）は 10pt 6.7・6.7、20pt 12.3・13.5、32pt 19.0・21.0、
 * 印 → 名前の字は 10pt 2.8〜3.6、20pt 6.0〜6.7、32pt 9.5〜10.7。
 * 字の送りと画素の端の差は、和文の字の左右の空きの見積もりで見た：IPAGothic の字の外形（字の送り 1 に対する空き）で、
 * 名前の頭の字の左の空きが 特 0.043・売 0.083・販 0.039、名前の末の字の右の空きが 益 0.0425・上 0.064。
 * この見積もりで、項目のあいだは 3 つの大きさの 6 つとも標準と 0.55px 以内、印 → 名前の 3 つの平均は標準の幅の中
 * （見積もりの字ごとの空きの開きが標準より大きいので、字ごとには幅の外へ 0.4px まで出る）。
 * 左右に置いた凡例の印 → 名前も同じ決まりにする（標準の左右の凡例のこのあいだは測っていない）
 */
/** 項目のあいだのうち、文字の大きさに掛ける分 */
export const LEGEND_ITEM_GAP_RATIO = 0.4;
/** 項目のあいだのうち、文字の大きさによらない分 (px) */
export const LEGEND_ITEM_GAP_BASE = 0.65;
/**
 * 項目のあいだ (px)。前の名前の字の送りの右の端から、次の印の左の端まで。文字の 0.4 倍 ＋ 0.65px
 * （8pt 4.9px、10pt 6.0px、20pt 11.3px、32pt 17.7px、40pt 22.0px）
 */
export function legendItemGap(fontPx: number): number {
    return fontPx * LEGEND_ITEM_GAP_RATIO + LEGEND_ITEM_GAP_BASE;
}
/** 印と名前のあいだ（文字の大きさに対する比） */
export const LEGEND_MARKER_GAP_RATIO = 0.175;
/**
 * 印と名前のあいだ (px)。印の見える右の端（組み方に渡す印の幅の右の端。線の印は lineInkSpan）から、名前の字の送りの左の端まで。文字の 0.175 倍
 * （8pt 1.9px、10pt 2.3px、20pt 4.7px、32pt 7.5px、40pt 9.3px）
 */
export function legendMarkerGap(fontPx: number): number {
    return fontPx * LEGEND_MARKER_GAP_RATIO;
}
/** タイトル（まとまりの見出し）と項目のあいだ */
export const LEGEND_TITLE_GAP = 8;
/** ページを送る矢印と項目のあいだ */
export const LEGEND_ARROW_GAP = 6;
/** 左右に置いた凡例のページを送る矢印の一辺（行の間隔に対する比）。上下に置いた凡例の矢印の一辺は行の高さ */
export const LEGEND_SIDE_ARROW_RATIO = 0.8;
/** 左右に置くとき、凡例が取ってよい幅の上限（ビジュアルの幅に対する比）。上限が 40px に満たない狭いビジュアルでも 40px までは取る（legendLayout.ts） */
export const LEGEND_SIDE_MAX_RATIO = 0.3;
/** 線の印（折れ線・目標の線などの短い線）の長さ。文字の大きさに対する比 */
export const LEGEND_LINE_GLYPH_RATIO = 1.6;
/** 横に引く線の印の幅 (px)。四角の印の一辺より短くしない */
export function legendLineGlyphWidth(fontPx: number): number {
    return Math.max(legendMarkerRadius(fontPx) * 2, Math.round(fontPx * LEGEND_LINE_GLYPH_RATIO));
}
/**
 * 線の端の形（stroke-linecap）が、線（と破線の 1 つ 1 つのダッシュ）の端から外へ出す長さ。丸・四角は線の幅の半分、平らは 0
 */
export function lineCapReach(cap: string | undefined, width: number): number {
    return cap === "round" || cap === "square" ? width / 2 : 0;
}

/** 端の数の誤差（模様の長さの足し算の端数）を、ちょうどの位置とみなす幅 (px) */
const DASH_EPS = 1e-6;

/**
 * stroke-dasharray の文字を、ダッシュと隙間の長さの並びにする（SVG の決まり）。空白・カンマで区切った数で、数の数が奇数なら 2 回並べる。
 * 指定が無い・数でないものや負の数が混ざる・合計が 0 なら null（実線で描かれる）
 */
export function dashPatternOf(dash: string | null | undefined): number[] | null {
    if (dash === null || dash === undefined) return null;
    const parts = String(dash).trim().split(/[\s,]+/).filter((p) => p !== "");
    if (!parts.length || parts[0] === "none") return null;
    const values = parts.map(Number);
    if (values.some((v) => !Number.isFinite(v) || v < 0)) return null;
    const total = values.reduce((a, b) => a + b, 0);
    if (!(total > 0)) return null;
    return values.length % 2 ? [...values, ...values] : values;
}

/**
 * 線の印の、線の向きに沿って見える範囲（線の始めの端 x1 を 0、終わりの端 x2 を length とした長さ）。端の形の出る分を含む。
 * 見える所が無ければ null（平らな端で長さ 0 のダッシュだけの模様など）。
 *
 * 組み方に渡す印の幅は、描いた印の見える左右の端までにする（印と名前のあいだ・項目のあいだを見える端から取る）。破線・点線は、
 * 模様が線の始めの端から始まり（offset は stroke-dashoffset）、最後のダッシュが線の終わりの端より手前で終わることがあるので、
 * その模様で見える端を数える：左の端は最初のダッシュの始め、右の端は線の中で最後に描かれるダッシュの終わり（どちらも線の端で
 * 切られる）に、端の形の出る分を足す。長さ 0 のダッシュは、端の形が丸・四角なら点・四角を描き、平らなら描かない。
 * 線の終わりちょうどで始まる（始めちょうどで終わる）ダッシュは、描かれるとみなす（丸・四角の端なら、その点の分だけ見える端を
 * 外へ取る。描かれなくても名前との間が広くなるだけで、名前には掛からない側）。平らな端では長さが 0 なので数えない
 */
export function lineInkSpan(length: number, width: number, cap: string | undefined, dash?: string | null, offset = 0): { start: number; end: number } | null {
    const ext = lineCapReach(cap, width);
    const pattern = dashPatternOf(dash);
    if (!pattern) return length > 0 || ext > 0 ? { start: -ext, end: length + ext } : null;
    const period = pattern.reduce((a, b) => a + b, 0);
    // 端の形があれば、線の端に触れるだけのダッシュ（長さ 0 の点）も数える。平らなら、線の中に長さのある所だけ
    const touch = ext > 0;
    const tol = touch ? DASH_EPS : -DASH_EPS;
    let start = Infinity;
    let end = -Infinity;
    let at = 0;
    for (let i = 0; i < pattern.length; i += 2) {
        const o = at;
        const d = pattern[i];
        at += pattern[i] + pattern[i + 1];
        if (d === 0 && !touch) continue;
        // 線の位置 s のダッシュは、模様の位置 (s + offset) mod period。このダッシュの k 回目は [kP + o − offset, kP + o + d − offset]
        const first = Math.ceil((offset - o - d - tol) / period);
        const last = Math.floor((length + offset - o + tol) / period);
        if (first > last) continue;
        start = Math.min(start, Math.max(0, first * period + o - offset));
        end = Math.max(end, Math.min(length, last * period + o + d - offset));
    }
    return start <= end ? { start: start - ext, end: end + ext } : null;
}
