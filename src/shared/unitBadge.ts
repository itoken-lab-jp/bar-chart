/**
 * 値の軸の単位のラベル（「(千円)」など）の、目盛りの列の幅の上限と、入らない単位の切り方
 */

/**
 * 目盛りの列を単位の幅まで広げても、描く範囲に残す幅（ビジュアルの幅に対する割合）。
 * 項目名の欄の上限（描く範囲に 4 分の 1 を残す）と同じ考え
 */
export const PLOT_WIDTH_SHARE_MIN = 0.25;

/** 目盛りの列 1 つ。幅は px */
export interface AxisColumn {
    /** 目盛りの数字の幅。上限を付けない（目盛りを出さなければ 0） */
    tick: number;
    /** 列の上に置く単位の幅（単位が無ければ 0） */
    unit: number;
    /** 列を目盛りの数字の幅から広げても、描く範囲が狭くならない幅（余白に下限があり、列がその内に収まる分）。省くと 0 */
    free?: number;
}

/**
 * 目盛りの列の幅（目盛りの数字と上の単位の広いほう）。目盛りの数字より広げるのは、描く範囲の幅がビジュアルの幅 visualWidth の
 * PLOT_WIDTH_SHARE_MIN を下回らない所まで。plotWidth は、どの列も目盛りの数字の幅にしたときの描く範囲の幅（列の幅が、free を超えて
 * 増えた分だけ描く範囲が狭くなる形で渡す）。列が 2 つ（左右の値の軸）なら、広げてよい幅を等しく分け、片方が使い切らなければ残りをもう片方に回す
 */
export function axisColumnWidths(columns: readonly AxisColumn[], plotWidth: number, visualWidth: number): number[] {
    let room = Math.max(0, plotWidth - visualWidth * PLOT_WIDTH_SHARE_MIN);
    const frees = columns.map((c) => Math.max(0, c.free ?? 0));
    // 描く範囲を狭める分（free の内は狭めない）
    const wants = columns.map((c, i) => Math.max(0, c.unit - c.tick - frees[i]));
    const extras = wants.map(() => 0);
    // 欲しい幅の小さい列から、残りを等しく分ける
    const order = wants.map((_, i) => i).sort((a, b) => wants[a] - wants[b]);
    order.forEach((i, k) => {
        extras[i] = Math.min(wants[i], room / (order.length - k));
        room -= extras[i];
    });
    return columns.map((c, i) => c.tick + Math.min(Math.max(0, c.unit - c.tick), frees[i] + extras[i]));
}

const OPEN = "(（[［";
const CLOSE = ")）]］";
/** 「…」の直前に残さない字：開きの括弧と、区切り（「:」と空白）。残すと「(千円（…)」「(単位:…)」のように、中身の無い形になる */
const DANGLING = /[(（[［{｛「『【〔〈《:：\s]+$/u;
/** 単位の名前より前に付ける前置き（「(単位: 千円)」の「単位: 」）。前置きだけが残る切り方はしない */
const PREFIX = /^単位\s*[:：]\s*/u;

/**
 * 幅 maxWidth に入らない単位を、末尾を「…」で切った字（幅は widthOf で測る。字の間を足すなどは使う側で決める）。
 * 括弧で囲む書き方（「(千円)」「(単位: 千円)」）なら、閉じの括弧は残して中を切る（「(千円（税…)」）。
 * 「…」の直前に開きの括弧・区切り（「:」と空白）が残るなら落とす（「(千円（…)」ではなく「(千円…)」）。
 * 前置き（「単位: 」）のあとの単位の名前が 1 文字も残らない形は出さない（「(単位:…)」にはせず、空）。
 * 字は符号位置で数えて切る（サロゲート ペアの字を半分にしない）。
 * 単位の名前の 1 文字と「…」も入らなければ空（単位を出さない。場所の上限を守る）
 */
export function truncateUnitText(text: string, maxWidth: number, widthOf: (shown: string) => number): string {
    if (!text || widthOf(text) <= maxWidth) return text;
    const chars = Array.from(text);
    const last = chars.length - 1;
    const bracketed = chars.length > 2 && OPEN.includes(chars[0]) && CLOSE.indexOf(chars[last]) === OPEN.indexOf(chars[0]);
    const [head, body, tail] = bracketed ? [chars[0], chars.slice(1, -1), chars[last]] : ["", chars, ""];
    const prefixLength = Array.from(PREFIX.exec(body.join(""))?.[0] ?? "").length;
    for (let n = body.length - 1; n > prefixLength; n--) {
        const kept = body.slice(0, n).join("").replace(DANGLING, "");
        // 前置きと、落とした括弧・区切りのほかに、単位の名前が 1 文字以上残る形だけ
        if (Array.from(kept).length <= prefixLength) continue;
        const shown = `${head}${kept}…${tail}`;
        if (widthOf(shown) <= maxWidth) return shown;
    }
    return "";
}
