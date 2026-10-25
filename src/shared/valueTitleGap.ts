/**
 * 縦向きのグラフの値の軸のタイトル（-90° に回した字）の置き方。標準の集合縦棒を Desktop で測った数に合わせる
 * （損益のデータ・616×336・タイトル 12pt・表示単位なし。字の画素の端どうし。一番広い目盛りの数字は「-1,000」）。
 * - 目盛りの列（目盛りの数字と上の単位の広いほう）の外の端の字から、タイトルの字の内側の端までを、標準のすき間にする
 *   （valueTitleInkGapOf。単位があるときも、単位を目盛りの列の一部とみて同じにする）
 * - タイトルの字の外側の端を、ビューの端（その側に凡例があれば、凡例の字か印の内側の端）から標準と同じだけ離す（valueTitleBaselineFromEdge）
 * - タイトルの文字の大きさが 12pt から変わると、標準は外側の空きとすき間を変える（VALUE_TITLE_OUTER_SLOPE_*・VALUE_TITLE_GAP_SLOPE_*）
 */
import { LEGEND_PAD } from "./legend";
import { TITLE_INK_ASCENT_OVERRIDES, TITLE_INK_DESCENT_OVERRIDES } from "./titleInkTable";

/**
 * -90° に回したタイトルの字の画素が、ベースラインから文字の下側（回すと右）へ出る分と、上側（左）へ出る分（文字の大きさに対する割合）。
 * 自前のビジュアルの Desktop の測り（タイトル 12pt の全角の字）から求めた。下の titleInkOf の既定の組の値
 */
export const TITLE_INK_DESCENT = 0.07;
export const TITLE_INK_ASCENT = 0.83;

/** タイトルの字の画素の、ベースラインから上側・下側へ出る分（文字の大きさに対する割合） */
export interface TitleInk {
    ascent: number;
    descent: number;
}

/**
 * 字の組ごとの見積もり。値は、いくつかの書体（IPAGothic・WenQuanYi・Carlito・Caladea・Liberation・DejaVu・Inter・FreeSans）を
 * ブラウザーの canvas で測った字の画素の上下の端（actualBoundingBox）を丸めたもの（和文の記号は IPAGothic を基準にする）。
 * 既定の組は Desktop の全角の字の測り
 * - 既定（上 0.83・下 0.07）：ASCII の字・空白、Latin-1 の記号（U+00A0〜U+00BF と × ÷）、一般の句読点（U+2000〜U+206F。… ‥ — ※ など）、
 *   文字のような記号（U+2100〜U+214F。℃ № など）、矢印（U+2190〜U+21FF）、図形（U+25A0〜U+25FF）、その他の記号（U+2600〜U+26FF）、
 *   漢字・かな、和文の記号（U+3000〜U+303F）、囲みの CJK（U+3200〜U+32FF。㈱ など）、全角の形（U+FF01〜U+FF9F）
 * - やや高い字（丸数字 U+2460〜U+24FF、CJK の単位 U+3300〜U+33FF、↑ ‡ “ ” ‘ ’、ASCII の ' " $、全角の ｜）：上 0.88・下 0.12
 * - 高い字（ASCII の |）：上 1.0
 * - 下に伸びる字（英字の g・j・p・q・y、J・Q、括弧、縦の線、区切り、記号の一部と、その全角、„ § ¶ ¿ ¡ ¢）：下 0.25
 * - 付け記号のある英字（U+00C0〜U+024F・U+1E00〜U+1EFF。× ÷ を除く）：上 1.0・下 0.25
 * - 上の範囲の外の字（ギリシャ文字・キリル文字・絵文字など）：上 1.05・下 0.32
 * - 既定の組の範囲の中でも、1 字ずつ測って既定より出る字（µ ¦ ¸ ² ¼ ‖ † ■ ○ ㈱ など）は、その字の組（shared/titleInkTable）
 * 太字なら上下とも太字の出（boldInkGrowthOf）を足す。下線があれば、下は下線の画素の下の端（underlineInkBelowOf）以上
 */
export const TITLE_LATIN_DESCENT = 0.25;
/**
 * 下線の下の端の、置き方の値（ベースラインから下へ。文字の大きさに対する割合）。標準の X 軸のタイトル（36pt）の下線の下の端は、下へ下りる字と
 * 同じ 0.25 文字とみて置いた位置より 3.1px 下で、この値で置くと 0.2px の差。下線の画素がこれより下へ出る大きさ・画面の倍率では、
 * 画素の下の端で置く（underlineInkBelowOf）
 */
export const TITLE_UNDERLINE_DESCENT = 0.19;

/** 画面の倍率（window.devicePixelRatio）。無い・読めないときは 1 */
export function screenScale(): number {
    try {
        const ratio = typeof window === "undefined" ? undefined : window.devicePixelRatio;
        return typeof ratio === "number" && Number.isFinite(ratio) && ratio > 0 ? ratio : 1;
    } catch {
        return 1;
    }
}

/**
 * 見積もりがみる、ページの拡大縮小の一番小さい倍率。Desktop はページの大きさに合わせてビジュアルを CSS の変形で拡大縮小することがあり、
 * その倍率はビジュアルに渡らない（update の viewport は幅と高さだけ。CSS の変形では window.devicePixelRatio も変わらない）。Chromium は変形のあとの
 * 画面の画素の上で下線と字を塗るので、画素 1 つの大きさ（CSS の px）は 1 ÷（devicePixelRatio × 拡大）。拡大を 0.5〜2 倍とみて、
 * 画素がいちばん大きい 0.5 倍で見積もる。測った拡大では外へ出ない（下線は 0.5・0.7・1・1.26・1.68・2、太字は 0.5〜2 の 19 通り）が、下線は
 * そのあいだの 0.52〜0.6 倍で、26〜30pt の下の端が最大 0.64px 外へ出る（画素への丸め方が拡大で変わるためとみられる。underlineInkBelowOf）。
 * 拡大が 0.5 倍より小さいと、もっと外へ出ることがある
 */
export const TITLE_MIN_PAGE_ZOOM = 0.5;

/** 見積もりがみる、画面の画素 1 つの大きさ（CSS の px）。scale は画面の倍率（devicePixelRatio）。拡大が TITLE_MIN_PAGE_ZOOM のときの大きさ */
export function worstPixelOf(scale = screenScale()): number {
    return 1 / (scale * TITLE_MIN_PAGE_ZOOM);
}

/**
 * 下線の画素の下の端の見積もり（ベースラインから下へ。px）。fontPx は文字の大きさ、scale は画面の倍率（devicePixelRatio）。
 * Chromium は下線（text-decoration の太さ・位置が auto）を、太さ t = max(1, 文字の px ÷ 10)、ベースラインの g = max(1, ceil(t ÷ 2)) 下から引く
 * （どちらも CSS の px。Linux の Chromium で、IPAGothic・Inter・Liberation Sans・DejaVu Sans の 9〜48pt × 倍率 1・2.52 を測った）。
 * 拡大が 1 なら、画素の下の端は g ＋ t より最大 1 画素下（帯に掛かる画素まで塗る）。拡大が整数でないと、帯の位置と太さが画面の画素にそろって
 * さらに最大 1 CSS px ほど下に出る（倍率 1 × 拡大 0.7 で g ＋ t ＋ 1 画素より 1.2px 下など。仕組みは確かめていない）。
 * 見積もりは g ＋ t ＋ 1 ＋ 1 画素（worstPixelOf。拡大 0.5 倍の画素）と、置き方の値（TITLE_UNDERLINE_DESCENT × 文字の px）の大きいほう。
 * IPAGothic の「金額」を、倍率 1・1.25・1.5・2・2.52 × 拡大 0.5・0.7・1・1.26・1.68・2（transform: scale）× 9〜48pt × 画面の画素の中の
 * ベースラインの位置 20 通りで描くと、見積もりはどれも画素の下の端より下だった（一番近い所で 0.017px。倍率 2.52 × 拡大 0.5 の 14pt）。
 * そのあいだの拡大では外へ出る形がある：26〜30pt で、倍率 1 × 拡大 0.52 は最大 0.64px（30pt。下の端 9.64px、見積もり 9.00px）、
 * 拡大 0.55 は 0.32〜0.54px、拡大 0.58 は 0.13px、拡大 0.6 は 0.18px まで（倍率 1〜2.52）。画面の画素の 3 分の 1 ほどなので
 * 見積もりは変えず、この拡大では下線の下の端がその分切れることがあるとした
 */
export function underlineInkBelowOf(fontPx: number, scale = screenScale()): number {
    const thickness = Math.max(1, fontPx / 10);
    const gap = Math.max(1, Math.ceil(thickness / 2));
    return Math.max(TITLE_UNDERLINE_DESCENT * fontPx, gap + thickness + 1 + worstPixelOf(scale));
}

/**
 * 太字の字の画素が上下に出る分のうち、文字の大きさに比べた分（文字の大きさに対する割合）。太字の書体の無い字は、Chromium（Skia）が字の輪郭を
 * 太らせて描くので、上下に 1/48〜1/64 文字ずつ出て、そのうえ画面の画素にそろえる分（最大 1 画素）が出る。見積もりは boldInkGrowthOf
 */
export const TITLE_BOLD_GROWTH = 1 / 48;

/**
 * 太字の字の画素が上下に出る分の見積もり（px）。文字の大きさの TITLE_BOLD_GROWTH と、画面の画素 1 つ（worstPixelOf。拡大 0.5 倍の画素）の和。
 * IPAGothic の合成の太字の「金額」を、下線の見積もりと同じ倍率 × 拡大 0.5〜2 の 19 通り（0.5〜0.95 の 0.05 刻みと 1・1.1・1.2・1.26・1.4・1.5・
 * 1.68・1.8・2。組の一部は測りが欠けている）× 9〜48pt × 画素の中の位置 20 通りで描き、太字と普通の字の画素の外の端（それぞれ 20 通りの一番外）の
 * 差を上下で見ると、どれもこの見積もりの内だった（一番近い所で 0.04px。下側の倍率 2.52 × 拡大 0.5 の 30pt）。
 * 斜体（字を横に傾ける）は上下の出を変えない（倍率 1・1.5・2.52 の測りで 0）
 */
export function boldInkGrowthOf(fontPx: number, scale = screenScale()): number {
    return TITLE_BOLD_GROWTH * fontPx + worstPixelOf(scale);
}
export const TITLE_TALL_ASCENT = 1.0;
export const TITLE_INK_SLIGHTLY_TALL: TitleInk = { ascent: 0.88, descent: 0.12 };
export const TITLE_INK_ACCENTED: TitleInk = { ascent: 1.0, descent: 0.25 };
export const TITLE_INK_OTHER: TitleInk = { ascent: 1.05, descent: 0.32 };
const DESCENDING_CHARS = new Set([..."gjpqyJQ()[]{}|,;_@$/\\", ..."ｇｊｐｑｙＪＱ｜，；＿＠＄", ..."\u201e\u00a7\u00b6\u00bf\u00a1\u00a2"]);
const TALL_CHARS = new Set(["|"]);
const SLIGHTLY_TALL_CHARS = new Set([..."'\"$｜", ..."\u2191\u2021\u201c\u201d\u2018\u2019"]);
const inRanges = (c: number, ranges: ReadonlyArray<readonly [number, number]>) => ranges.some(([lo, hi]) => c >= lo && c <= hi);
const DEFAULT_RANGES = [
    [0x09, 0x09],
    [0x20, 0x7e],
    [0xa0, 0xbf],
    [0xd7, 0xd7],
    [0xf7, 0xf7],
    [0x2000, 0x206f],
    [0x2100, 0x214f],
    [0x2190, 0x21ff],
    [0x25a0, 0x25ff],
    [0x2600, 0x26ff],
    [0x3000, 0x30ff],
    [0x31f0, 0x31ff],
    [0x3200, 0x32ff],
    [0x3400, 0x4dbf],
    [0x4e00, 0x9fff],
    [0xf900, 0xfaff],
    [0xff01, 0xff9f],
] as const;
const SLIGHTLY_TALL_RANGES = [
    [0x2460, 0x24ff],
    [0x3300, 0x33ff],
] as const;
const ACCENTED_RANGES = [
    [0xc0, 0x24f],
    [0x1e00, 0x1eff],
] as const;

/** 表（shared/titleInkTable）で c の字の組の値。無ければ 0 */
const overrideOf = (c: number, table: ReadonlyArray<readonly [number, number, number]>) => {
    for (const [lo, hi, value] of table) if (c >= lo && c <= hi) return value;
    return 0;
};

/** 字（ch）の組の見積もり */
function charInkOf(ch: string): TitleInk {
    const c = ch.codePointAt(0)!;
    let ink: TitleInk =
        SLIGHTLY_TALL_CHARS.has(ch) || inRanges(c, SLIGHTLY_TALL_RANGES)
            ? TITLE_INK_SLIGHTLY_TALL
            : inRanges(c, DEFAULT_RANGES)
              ? TITLE_INK_DEFAULT
              : inRanges(c, ACCENTED_RANGES)
                ? TITLE_INK_ACCENTED
                : TITLE_INK_OTHER;
    if (DESCENDING_CHARS.has(ch)) ink = { ascent: ink.ascent, descent: Math.max(ink.descent, TITLE_LATIN_DESCENT) };
    if (TALL_CHARS.has(ch)) ink = { ascent: Math.max(ink.ascent, TITLE_TALL_ASCENT), descent: ink.descent };
    return { ascent: Math.max(ink.ascent, overrideOf(c, TITLE_INK_ASCENT_OVERRIDES)), descent: Math.max(ink.descent, overrideOf(c, TITLE_INK_DESCENT_OVERRIDES)) };
}

/** titleInkOf に渡す字の形。fontPx は文字の大きさ（下線と太字の見積もりに使う）、bold は太字か、scale は画面の倍率（省くと screenScale） */
export interface TitleInkStyle {
    fontPx?: number;
    bold?: boolean;
    scale?: number;
}

/**
 * タイトルの字（text）の画素の上下の出の見積もり。字ごとの見積もりの最大。underline は下線を引くか、style は字の形。
 * 太字なら上下に、文字の大きさがあれば太字の出（boldInkGrowthOf）を文字の大きさで割った値、無ければ TITLE_BOLD_GROWTH を足す。
 * 下線は、文字の大きさがあれば下線の画素の下の端（underlineInkBelowOf）を文字の大きさで割った値、無ければ TITLE_UNDERLINE_DESCENT を、下の出の下限にする
 */
export function titleInkOf(text: string, underline = false, style: TitleInkStyle = {}): TitleInk {
    let ascent = TITLE_INK_ASCENT;
    let descent = TITLE_INK_DESCENT;
    for (const ch of text) {
        const ink = charInkOf(ch);
        ascent = Math.max(ascent, ink.ascent);
        descent = Math.max(descent, ink.descent);
    }
    const { fontPx } = style;
    const sized = fontPx !== undefined && fontPx > 0;
    if (style.bold) {
        const growth = sized ? boldInkGrowthOf(fontPx, style.scale ?? screenScale()) / fontPx : TITLE_BOLD_GROWTH;
        ascent += growth;
        descent += growth;
    }
    if (underline) {
        descent = Math.max(descent, sized ? underlineInkBelowOf(fontPx, style.scale ?? screenScale()) / fontPx : TITLE_UNDERLINE_DESCENT);
    }
    return { ascent, descent };
}

/** 既定の組の字だけのタイトルの見積もり（標準の測りの字） */
export const TITLE_INK_DEFAULT: TitleInk = { ascent: TITLE_INK_ASCENT, descent: TITLE_INK_DESCENT };

/**
 * 目盛りの数字の字の画素が、文字の四角の外の端から内側に入る分（目盛りの文字の大きさに対する割合）。標準の測りの一番広い数字は
 * 「-1,000」で、左の端は「-」、右の端は「0」。どちらも自前のビジュアルの Desktop の測りに 0.3px 以内で合う。
 * 字の形によらずこの値で置く（左の端が「1」の数字（「1,000」）は、字の画素が 0.12 文字ほど内側に入るので、すき間が 0.05 文字ほど広く見える）
 */
export const TICK_INK_INSET_LEFT = 0.07;
export const TICK_INK_INSET_RIGHT = 0.035;

/**
 * 標準の、一番広い目盛りの数字の字とタイトルの字のすき間（字の画素の端どうし）。tickPx は目盛りの文字の大きさ、titlePx はタイトルの
 * 文字の大きさ（px。省略すると 12pt）。
 * タイトル 12pt で、左の軸は 12〜36pt の 11 点（9.5〜11.9）に当てた直線（測った点との差は 0.6px 以内。20pt のほかは 0.5px 以内）。
 * 右の軸は 7.2 と「4.7 ＋ 0.075 × px」の大きいほう（9・12・16・24・28・32・36pt の 7.1・7.1・7.5・7.1・7.9・7.5・8.7 との差は 0.4px 以内。
 * 28〜36pt は 3 点の差の一番大きいのを小さくする線）。タイトルが 12pt と違えば、1px あたり VALUE_TITLE_GAP_SLOPE_* だけ変える
 */
export function valueTitleInkGapOf(tickPx: number, right: boolean, titlePx = VALUE_TITLE_BASE_PX): number {
    const atBase = right ? Math.max(7.2, 4.7 + 0.075 * tickPx) : 8.2 + 0.07 * tickPx;
    return atBase + (right ? VALUE_TITLE_GAP_SLOPE_RIGHT : VALUE_TITLE_GAP_SLOPE_LEFT) * (titlePx - VALUE_TITLE_BASE_PX);
}

/** タイトルの文字の大きさの基準（12pt の px）。すき間・外側の空きの決まりの値は、この大きさのタイトルの値 */
export const VALUE_TITLE_BASE_PX = 16;

/**
 * タイトルの文字の大きさが VALUE_TITLE_BASE_PX から 1px 変わるときの、標準の動き（px）。R12 の形（損益・616×336・境界線あり・凡例なし）で、
 * タイトル 12・20・24・36・48pt を Desktop で測った。外側（字の外側の端 → 枠）とすき間（字 → 一番広い目盛りの数字の字）の 12pt との差：
 * - 左の軸：外側は 20pt +1.2・24pt +2.4・36pt +5.2・48pt +7.1 で、1px あたり +0.15（差は 0.4px 以内）。すき間は 20pt −1.2・36pt −2.8
 *   （目盛り 12pt）と 24pt −1.1（目盛り 9pt）で、1px あたり −0.0875（差は 0.3px 以内）。ベースラインから目盛りの字まではほぼ一定
 * - 右の軸：外側は 20pt −1.5・36pt −1.9。1px あたり −0.14 で寄せ、字の下の端がビューの端に来たら止める（36pt は止めた所。差は 0.1px 以内）。
 *   すき間は 20pt +1.2・36pt +4.4 で、1px あたり +0.135（差は 0.3px 以内）
 * 12pt より小さいタイトルは測っていない。同じ直線をのばす（左の外側はビューの端の 0.1px で止まる）
 */
export const VALUE_TITLE_OUTER_SLOPE_LEFT = 0.15;
export const VALUE_TITLE_OUTER_SLOPE_RIGHT = -0.14;
export const VALUE_TITLE_GAP_SLOPE_LEFT = -0.0875;
export const VALUE_TITLE_GAP_SLOPE_RIGHT = 0.135;

/**
 * タイトルの字の画素と目盛りの数字の字の画素のすき間の下限（px）。標準の測りには無い守りの値で、見積もりでこれより寄るときだけ外へ出す。
 * タイトルの側に凡例があるときの、凡例の字・印とのすき間の下限にも使う
 */
export const VALUE_TITLE_MIN_INK_GAP = 2;

/**
 * タイトルのベースラインから、目盛りの列の文字の四角の外の端まで（列の側へ測る。左の軸はベースラインの右、右の軸は左）。
 * titlePx はタイトルの文字の大きさ、tickPx は目盛りの文字の大きさ（px）、ink はタイトルの字の見積もり（titleInkOf）。
 * 標準は字の四角で置くとみて、列の側へ出る字の画素（左の軸は字の下側、右の軸は上側）を既定の組の見積もりで見て、
 * 字の画素のすき間を valueTitleInkGapOf（タイトルの大きさで変わる）にする（字の形によらず標準の位置）。ink で見た字の画素のすき間が
 * VALUE_TITLE_MIN_INK_GAP を切るとき（下に伸びる字・高い字の大きいタイトル）だけ、そのすき間まで列から離す
 */
export function valueTitleBaselineToColumn(titlePx: number, tickPx: number, right: boolean, ink: TitleInk = TITLE_INK_DEFAULT): number {
    const inkInset = tickPx * (right ? TICK_INK_INSET_RIGHT : TICK_INK_INSET_LEFT);
    const standard = valueTitleInkGapOf(tickPx, right, titlePx) - inkInset + titlePx * (right ? TITLE_INK_ASCENT : TITLE_INK_DESCENT);
    const apart = VALUE_TITLE_MIN_INK_GAP - inkInset + titlePx * (right ? ink.ascent : ink.descent);
    return Math.max(standard, apart);
}

/**
 * ビューの端から、タイトルの字の外側の端まで（px。タイトル 12pt）。標準は枠から左 5.5・右 8.3 で、自前のビジュアルの測りから出した枠とビューの
 * あいだ（左 5.2〜5.6、右 6.3〜6.5）を引いた。左は標準でもほぼビューの端
 */
export const VALUE_TITLE_EDGE_INK_LEFT = 0.1;
export const VALUE_TITLE_EDGE_INK_RIGHT = 1.9;

/**
 * タイトルの側に凡例を置いたときの、凡例の内側の端から、タイトルの字の外側の端まで（px。タイトル 12pt）。標準は R12 の形で、
 * 凡例「左」は凡例の字の画素の右の端 → タイトルの字 10.3、凡例「右」はタイトルの字 → 凡例の印の左の端 2.4。
 * 凡例「左」の内側の端は、凡例の字の画素ではなく字の送りの右の端（凡例の箱の幅は、標準も 2 本も字の送りの幅で決め、字の画素は測らない）。
 * 標準の 10.3 は、送りの端からの空きに、測った形の一番長い名前「特別損益」の最後の字「益」の右の空き（字の画素の右の端 → 送りの端。
 * IPAGothic で 0.04 文字、凡例の文字 10pt で 0.53px）を足した画素どうしの値とみて、送りの端から 10.3 − 0.5 ＝ 9.8 にする
 * （名前の最後の字の右の空きが違えば、画素どうしの空きもその分変わる。標準も同じ）
 */
export const VALUE_TITLE_LEGEND_INK_LEFT = 9.8;
export const VALUE_TITLE_LEGEND_INK_RIGHT = 2.4;
/**
 * 凡例の分を除いた所の端から、凡例の字の送りの端か印の内側の端まで（px）。左右に置いた凡例は、箱の外（描く範囲の側）と箱の中に LEGEND_PAD ずつ空け、
 * 字と印をその内側から並べる（shared/legendLayout）。一番長い名前（かタイトル）の送りの幅が箱の幅を決めるので、内側の端はいつもこの所
 * （名前が「1 文字 ＋ …」でも列に入らず箱で切れる、まれな形を除く）
 */
export const LEGEND_INK_INSET = 2 * LEGEND_PAD;

/**
 * 凡例の分を除いた所の端（凡例が無ければビューの端）から、タイトルのベースラインまで（px。内側へ測る。凡例の側へ出るときは負）。
 * titlePx はタイトルの文字の大きさ、ink はタイトルの字の見積もり（titleInkOf）、legend はタイトルの側に凡例があるか。
 * 既定の組の字の外側の端（左の軸は字の上側、右の軸は下側）を、12pt で、凡例が無ければビューの端から VALUE_TITLE_EDGE_INK_*、
 * 凡例があれば凡例の字・印の内側の端から VALUE_TITLE_LEGEND_INK_* に置き、タイトルが大きいほど 1px あたり VALUE_TITLE_OUTER_SLOPE_* だけ動かす。
 * 守り：左の軸は、字の上の端（高い字があればその分）を端から VALUE_TITLE_EDGE_INK_LEFT より外へ出さない（凡例があっても、描く範囲の側に置く）。
 * 右の軸は、字の下の端（下に伸びる字・下線があればその分）を、凡例が無ければビューの中に、凡例があれば凡例の字・印から
 * VALUE_TITLE_MIN_INK_GAP の外に置く。既定の組の字だけのタイトルは、守りに掛からなければ標準の位置
 */
export function valueTitleBaselineFromEdge(titlePx: number, right: boolean, ink: TitleInk = TITLE_INK_DEFAULT, legend = false): number {
    const grow = titlePx - VALUE_TITLE_BASE_PX;
    if (!right) {
        const outer = (legend ? VALUE_TITLE_LEGEND_INK_LEFT - LEGEND_INK_INSET : VALUE_TITLE_EDGE_INK_LEFT) + VALUE_TITLE_OUTER_SLOPE_LEFT * grow;
        return Math.max(outer + TITLE_INK_ASCENT * titlePx, VALUE_TITLE_EDGE_INK_LEFT + ink.ascent * titlePx);
    }
    const outer = (legend ? VALUE_TITLE_LEGEND_INK_RIGHT - LEGEND_INK_INSET : VALUE_TITLE_EDGE_INK_RIGHT) + VALUE_TITLE_OUTER_SLOPE_RIGHT * grow;
    const floor = legend ? VALUE_TITLE_MIN_INK_GAP - LEGEND_INK_INSET : 0;
    return Math.max(outer + TITLE_INK_DESCENT * titlePx, floor + ink.descent * titlePx);
}
