/**
 * 軸の帯の決まり。標準の集合縦棒・集合横棒を Desktop で測って決めた（2026-10-20。全般のレスポンシブはオフ、項目名・目盛り 9pt、
 * 軸のタイトル 12pt）。測った数は字の画素の端で、ここでは字の送り（文字の四角）の端か、ベースラインの位置に直してある。
 *
 * - 名前の側（横向きのグラフの左）：縦の軸のタイトル（-90° に回した字）、名前の欄、描く範囲の左の端
 * - 下の帯：描く範囲の下の端、軸の字（縦向きの項目名、横向きの目盛りの数字）、軸のタイトル
 * - 横向きの棒のデータ ラベル（横書き）の背景の左右の余白
 *
 * 基準の端（ビューの左・下の端か、凡例の箱の端）と文字の大きさは作品の側で渡す
 */
import { TITLE_INK_ASCENT, TITLE_INK_DEFAULT, TITLE_INK_DESCENT, TitleInk, TitleInkStyle, VALUE_TITLE_MIN_INK_GAP, titleInkOf } from "./valueTitleGap";

/**
 * 名前の側のタイトルのベースライン（-90° に回すと縦の線）の x。基準の左の端から、字の画素の左の端（文字の上側）が基準の左の端に
 * 来る所。標準は字の画素の左の端をビューの左の端の 0〜1.5px 外に置いた（ビューの外は描けないので、端に置く）。titlePx はタイトルの文字の大きさ、
 * ink はタイトルの字の見積もり（titleInkOf）。既定の組の字（上 0.83 文字）より高い字があれば、その分だけ内側へ置き、字の上の端をビューに入れる
 */
export function sideTitleBaselineOf(titlePx: number, ink: TitleInk = TITLE_INK_DEFAULT): number {
    return titlePx * Math.max(TITLE_INK_ASCENT, ink.ascent);
}

/**
 * 名前の側のタイトルの字の画素の右の端（文字の下側）。基準の左の端から。ここから SIDE_TITLE_NAME_GAP 右に、名前の欄（か描く範囲）を置く。
 * 字の下側は既定の組の見積もり（下 0.07 文字）で見る（下に伸びる字があっても、名前の欄の位置は標準のまま）。ただし、見積もりの字の下の端と
 * 名前の字のすき間が VALUE_TITLE_MIN_INK_GAP を切る大きいタイトルだけ、そのすき間まで右へ寄せる
 */
export function sideTitleInkRightOf(titlePx: number, ink: TitleInk = TITLE_INK_DEFAULT): number {
    const below = Math.max(TITLE_INK_DESCENT * titlePx, ink.descent * titlePx - (SIDE_TITLE_NAME_GAP - VALUE_TITLE_MIN_INK_GAP));
    return sideTitleBaselineOf(titlePx, ink) + below;
}

/** タイトルの字の画素の右の端から、一番長い名前の字の送りの左の端まで (px)。標準は画素どうしで 9.9 */
export const SIDE_TITLE_NAME_GAP = 9.7;
/** タイトルを出さないとき、基準の左の端から、一番長い名前の字の送りの左の端まで (px)。標準は画素で 2.3 */
export const SIDE_NAME_LEAD = 2;
/** 名前の字の送りの右の端から、描く範囲の左の端まで (px)。標準は画素で 9.1 */
export const SIDE_NAME_PLOT_GAP = 9;
/**
 * 名前の側にタイトルも名前の欄も無いとき、基準の左の端から描く範囲の左の端まで (px)。標準は枠の左の端から 0 の線まで 9.4 で、
 * この置き方の 2 つの作品（9.5）と 0.1px で合う
 */
export const SIDE_BARE_LEFT = 4;
/**
 * 名前の欄を出さずにタイトルだけ出すとき、タイトルの字の画素の右の端から描く範囲の左の端まで (px)。標準は 11.5（12pt）。
 * SIDE_TITLE_NAME_GAP（9.7）で置いた 2 つの作品が 9.5 に見えた分（0.2）を見込んで足した
 */
export const SIDE_TITLE_PLOT_GAP = 11.7;

/**
 * 名前の側の、基準の左の端から描く範囲の左の端まで (px)。titlePx はタイトルの文字の大きさ（出さなければ null）、columnWidth は
 * 名前の欄の幅（左の端の SIDE_NAME_LEAD から、描く範囲の手前の SIDE_NAME_PLOT_GAP まで。名前の欄を出さなければ 0）。
 * 名前の欄を出さないときは、欄とその両側のすき間をまとめて除き、描く範囲を名前の欄の代わりにタイトルの右に置く（タイトルの字の右の端から
 * SIDE_TITLE_PLOT_GAP）。タイトルも無ければ SIDE_BARE_LEFT
 */
export function sidePlotLeftOf(titlePx: number | null, columnWidth: number, ink: TitleInk = TITLE_INK_DEFAULT): number {
    if (columnWidth > 0) return (titlePx === null ? 0 : sideTitleInkRightOf(titlePx, ink) + SIDE_TITLE_NAME_GAP - SIDE_NAME_LEAD) + columnWidth;
    return titlePx === null ? SIDE_BARE_LEFT : sideTitleInkRightOf(titlePx, ink) + SIDE_TITLE_PLOT_GAP;
}

/** 下の軸の字（縦向きの項目名、横向きの目盛りの数字）の 1 行目のベースライン：描く範囲の下の端から、字の大きさ ＋ この値 (px) */
export const AXIS_TEXT_DROP = 3.3;
/**
 * 下の軸の字の最後の行のベースラインから、その下のタイトルの行の上の端まで (px)。軸の字（項目名・目盛りの数字）を出さずにタイトルだけ
 * 出すときは、描く範囲の下の端からタイトルの行の上の端まで（字の行を除き、タイトルの上のすき間は残す）。字を出さない形は、標準の集合縦棒で
 * 描く範囲の下の端からタイトルの字の上の端まで 5.7（12pt。この置き方の見積もりは 6.3）。集合横棒は目盛りを消すと線も消えるので測れない
 */
export const AXIS_TEXT_BELOW = 6;
/**
 * 下の軸のタイトルを出さないときの、下の軸の字の最後の行のベースラインから基準の下の端まで (px)。縦向きの項目名（水平に並べた行）と、
 * 連続の X 軸の目盛り（同じ作りの字の行とみる。標準は測っていない）。
 * 標準は、和文の項目名の字の画素の下の端から枠の下の端まで 12.7（9pt）。前の置き方（ベースラインから 12・4）で描いた 2 つの作品の測り
 * （15.8・8.0）との差から出した（和文の字の画素はベースラインの 0.85px 下に出るとみる）。2 つの作品の値は 8.9 と 8.7 で、その間をとる
 */
export const COLUMN_TEXT_BOTTOM = 8.8;
/**
 * 下の軸のタイトルを出さないときの、下の軸の字の最後の行のベースラインから基準の下の端まで (px)。横向きの目盛りの数字の行。
 * 標準は、数字の字の画素の下の端（ベースライン）から枠の下の端まで 14.2（9pt）。同じく前の置き方の測り（16.8・9.2）との差から出した
 * （2 つの作品の値は 9.4 と 9.0）。縦向きより 0.4 広いのは、標準の測りの差から：枠からビューの下の端までを作品ごとに縦横で同じとみると、
 * 標準の横向きは 14.2 −（12.7 ＋ 0.85）＝ 0.65 広く、2 つの作品の枠の作りの値で出しても 0.3〜0.5 広い（タイトルの行の BAR_TITLE_BOTTOM も
 * 縦向きの COLUMN_TITLE_BOTTOM より広い）
 */
export const BAR_TEXT_BOTTOM = 9.2;

/**
 * 軸の字の「,」「;」の字の画素が、ベースラインから下へ出る分（文字の大きさに対する割合）。タイトルの見積もり（titleInkOf）の下に伸びる字の
 * 0.25 は g・p・y・括弧に合わせた値で、コンマには大きすぎる（字の形を測った値：Segoe UI と字の幅の同じ Selawik の標準 0.129・太字 0.148、
 * Liberation Sans 0.128〜0.155、DejaVu Sans 0.116〜0.142、Carlito 0.128〜0.135、Caladea 0.135、FreeSans 0.147、Inter 0.175〜0.193）。
 * 測った中の一番大きい値の上をとる。目盛りの数字（「2,000」）で多く出るので、軸の字だけこの値で見積もる
 */
export const AXIS_COMMA_DESCENT = 0.2;
const AXIS_COMMAS = /[,;]/g;

/**
 * 軸の字（項目名・目盛りの数字）text の画素が、ベースラインから下へ出る分（文字の大きさに対する割合）。「,」「;」は AXIS_COMMA_DESCENT、
 * ほかの字と下線・太字はタイトルと同じ見積もり（titleInkOf。既定の組の字だけなら 0.07）。style は字の形（文字の大きさ・太字。下線の画素の見積もりに使う）
 */
export function axisTextDescentOf(text: string, underline = false, style: TitleInkStyle = {}): number {
    const plain = text.replace(AXIS_COMMAS, "");
    const descent = titleInkOf(plain, underline, style).descent;
    return plain.length < text.length ? Math.max(descent, AXIS_COMMA_DESCENT) : descent;
}

/**
 * 下の軸のタイトルを出さないときの、一番下の軸の字の行のベースラインから基準の下の端まで (px)。base は COLUMN_TEXT_BOTTOM か BAR_TEXT_BOTTOM、
 * inkBelow は出す字の画素がベースラインから下へ出る分の一番大きい値 (px。axisTextDescentOf × 文字の大きさ)。下へ下りる字・コンマ・下線の
 * 下の端が base より下へ出るときだけ、字の下の端が基準の下の端に来るまで上げる（タイトルの行の bottomTitleBelowOf と同じ考え。既定の組の字
 * （下 0.07 文字）は、とても大きくなければ base のまま＝標準の位置）
 */
export function axisTextBottomOf(base: number, inkBelow: number): number {
    return Math.max(base, inkBelow);
}
/** 下の軸のタイトルの行の上の端から、タイトルのベースラインまで（タイトルの文字の大きさに対する割合。和文の入る字の上の広がり） */
export const AXIS_TITLE_ASCENT = 0.85;
/** 下の軸のタイトルのベースラインから、基準の下の端まで (px)。縦向きの X 軸のタイトル */
export const COLUMN_TITLE_BOTTOM = 3;
/** 下の軸のタイトルのベースラインから、基準の下の端まで (px)。横向きの値の軸のタイトル */
export const BAR_TITLE_BOTTOM = 4;

/**
 * 下の軸のタイトルのベースラインから、基準の下の端まで (px)。base は COLUMN_TITLE_BOTTOM か BAR_TITLE_BOTTOM、titlePx はタイトルの文字の大きさ、
 * ink はタイトルの字の見積もり（titleInkOf）。下へ下りる字（g・p・y・j・括弧など）や下線の下の端が base より下へ出るときだけ、
 * 字の下の端が基準の下の端に来るまで上げる（既定の組の字（下 0.07 文字）は、とても大きくなければ base のまま＝標準の位置）
 */
export function bottomTitleBelowOf(base: number, titlePx: number, ink: TitleInk = TITLE_INK_DEFAULT): number {
    return Math.max(base, ink.descent * titlePx);
}

/**
 * 横向きの棒のデータ ラベル（横書き）の背景の、字の送りの左右の端から外の余白 (px)。標準は背景を描いても字の位置を変えず、
 * 背景の四角を字の外へ広げる（一番右のラベルの背景の右の端は、背景なしの字の送りの右の端から 9pt で 4.3、24pt で 4.1 外）
 */
export const H_LABEL_BACKGROUND_PAD = 4;
