/**
 * 縦向きのグラフ（値が下から上）の、描く範囲の上の端と、その上の行（単位のラベル・ドリルの位置）、棒の上の外に置くデータ ラベルに
 * 空ける高さの決まり。標準の集合縦棒を Desktop で測って決めた（2026-10-20。全般のレスポンシブはオフ）。
 * どこを基準にするか（ビューの上の端か、上に置いた凡例の下か）と、文字の大きさは作品の側で渡す
 */

/** 描く範囲の上の端と、基準の上の端（ビューの上の端か、上に置いた凡例の下）のあいだ (px)。標準の集合縦棒（ラベルなし・単位なし）と同じ */
export const PLOT_TOP = 10;
/**
 * 基準の上の端と、上の端に置く文字（一番上の目盛りの数字・単位のラベル・ドリルの位置）の字の四角の上とのすき間の最小 (px)。
 * 標準の集合縦棒の 9pt の目盛りの数字（ラベルなしで、線が上の端から PLOT_TOP）と同じ
 */
export const TOP_TEXT_GAP = 4.6;
/** 目盛りの数字の字の四角が、線から上に出る高さ（文字の大きさに対する割合。ベースラインを線の 0.35 文字下に置き、字の上は 0.8 文字） */
export const TICK_TEXT_RISE = 0.45;
/** 単位のラベル・ドリルの位置のベースラインと、その下の基準の高さ（描く範囲の上の端か、目盛りを引く範囲の上の端）とのあいだ (px) */
export const TOP_ROW_RISE = 12;
/** 単位のラベルの、ベースラインから字の四角の上の端までの高さ（文字の大きさに対する割合） */
export const UNIT_TEXT_ASCENT = 0.8;
/** 単位のラベルの、ベースラインから字の四角の下の端までの深さ（文字の大きさに対する割合） */
export const UNIT_TEXT_DESCENT = 0.22;
/** 目盛りの列の上に置く単位のラベルの字の四角と、その下の一番上の目盛りの数字の字の四角とのすき間の最小 (px) */
export const COLUMN_BADGE_GAP = 3;
/** ドリルの位置の、ベースラインから字の四角の上の端までの高さ（文字の大きさに対する割合。和文の入る字の上の広がり） */
export const DRILL_TEXT_ASCENT = 0.85;
/** ドリルの位置の行が描く範囲の上に足す高さの上限（図の高さに対する割合）。超えるなら出さない（低い図で描く範囲が潰れないように） */
export const DRILL_ROW_SHARE_MAX = 0.12;
/** 値の軸の端に取るラベルの余白の、取ったあとの軸の長さに対する割合の上限（上と下を合わせて。余白で棒が潰れないように） */
export const LABEL_ROOM_SHARE_MAX = 0.6;
/** 棒の外の端に置くデータ ラベルの、字の四角（背景を描くなら背景の四角）の外に取る余白 (px)。背景を描かなくても取る */
export const OUTSIDE_LABEL_PAD = 3;

/**
 * 描く範囲の上の端（基準の上の端から）。一番上の目盛りの数字（大きさ tickPx。目盛りの数字を出さなければ 0）の字の四角の上が、
 * 上の端から TOP_TEXT_GAP より近くなるなら、その分だけ下げる
 */
export function plainTopOf(tickPx: number): number {
    return Math.max(PLOT_TOP, TOP_TEXT_GAP + tickPx * TICK_TEXT_RISE);
}

/**
 * 上の行に置く文字（大きさ px、ベースラインから字の四角の上の端までが ascent 文字）に要る、その行の下の基準の高さ（基準の上の端から）。
 * ベースラインを基準の高さの rise 上（既定は TOP_ROW_RISE）に置き、字の四角の上を TOP_TEXT_GAP 空ける
 */
export function topRowNeedOf(px: number, ascent: number, rise = TOP_ROW_RISE): number {
    return TOP_TEXT_GAP + px * ascent + rise;
}

/**
 * 目盛りの列の上に置く単位のラベル（大きさ unitPx）の、ベースラインとその下の一番上の目盛りの線とのあいだ (px)。
 * 既定は TOP_ROW_RISE。目盛りの数字（大きさ tickPx。出さなければ 0）と単位が大きく、字の四角どうしが COLUMN_BADGE_GAP より
 * 近づくなら、その分だけ上げる（既定の 9pt どうしでは TOP_ROW_RISE のまま）。目盛りの列の上に掛かる「プロット エリアの右上」の単位も、
 * 描く範囲の上の端からこの高さに置く（一番上の目盛りの線は、描く範囲の上の端か、それより下）
 */
export function columnBadgeRiseOf(tickPx: number, unitPx: number): number {
    return Math.max(TOP_ROW_RISE, tickPx * TICK_TEXT_RISE + unitPx * UNIT_TEXT_DESCENT + COLUMN_BADGE_GAP);
}

/**
 * 縦向きの棒の上の端の外に置く横書きのデータ ラベルの、字の四角と OUTSIDE_LABEL_PAD の上にさらに空ける高さ (px。size は一番上の行の文字の大きさ px)。
 * 標準の集合縦棒（ラベル 9〜36pt、1 行・背景なし）の、いちばん長い棒の端から描く範囲の上の端までは 9.6 ＋ 1.28 文字だった。
 * 棒の端から字の四角の上の端までは 8.3 ＋ 0.88 文字（数字の四角の下を棒の端から 8.3px、字の四角は数字の四角の下 0.08 文字から上 0.8 文字）なので、
 * 余白の 3px を除いた 0.4 文字 − 1.7px（0 より小さければ 0）
 */
export function outsideTopHeadroomOf(size: number): number {
    return Math.max(0, 0.4 * size - 1.7);
}
