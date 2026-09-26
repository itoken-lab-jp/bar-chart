import { UnitItem, formatValue } from "./units";

/**
 * マイナスと 0 の書き方。決算説明資料などの日本の帳票は、マイナスを ▲ で書き、変化なしを ±0 や -、
 * 丸めて 0 になるマイナスを ▲0 と書き分ける。モデルの書式文字列（#,0;▲#,0）はビジュアルが単位（万・億）を
 * 自前で割って整形すると効かないので、ここで付ける。
 *
 * 既定（SIGN_STYLE_DEFAULTS）は formatValue と同じ表示にする（マイナスは -、0 は 0、丸めて 0 のマイナスは -0）。
 * + を付けるときも、0 の書き方が「0」なら、プラスに + を足すだけ（丸めて 0 になるプラスは +0）。
 */

export const NEGATIVE_STYLES = {
    minus: "minus",
    triangle: "triangle",
    whiteTriangle: "whiteTriangle",
    parentheses: "parentheses",
} as const;

/** マイナスの書き方の選択肢。表示される形そのものを名前にする */
export const NEGATIVE_STYLE_ITEMS: UnitItem[] = [
    { value: NEGATIVE_STYLES.minus, displayName: "-1,234" },
    { value: NEGATIVE_STYLES.triangle, displayName: "▲1,234" },
    { value: NEGATIVE_STYLES.whiteTriangle, displayName: "△1,234" },
    { value: NEGATIVE_STYLES.parentheses, displayName: "(1,234)" },
];

export const ZERO_STYLES = {
    zero: "zero",
    plusMinus: "plusMinus",
    dash: "dash",
} as const;

/** 0（丸めて 0 になる値を含む）の書き方の選択肢 */
export const ZERO_STYLE_ITEMS: UnitItem[] = [
    { value: ZERO_STYLES.zero, displayName: "0" },
    { value: ZERO_STYLES.plusMinus, displayName: "±0" },
    { value: ZERO_STYLES.dash, displayName: "-" },
];

export interface SignStyle {
    /** マイナスの書き方（NEGATIVE_STYLES） */
    negative: string;
    /** 0 の書き方（ZERO_STYLES） */
    zero: string;
    /** 丸めて 0 になるマイナスに符号を残す（▲0）。切ると 0 の書き方にそろえる */
    negativeZero: boolean;
    /** プラスに + を付ける */
    plus: boolean;
}

export const SIGN_STYLE_DEFAULTS: SignStyle = {
    negative: NEGATIVE_STYLES.minus,
    zero: ZERO_STYLES.zero,
    negativeZero: true,
    plus: false,
};

const NEGATIVE_MARKS: Record<string, [string, string]> = {
    [NEGATIVE_STYLES.minus]: ["-", ""],
    [NEGATIVE_STYLES.triangle]: ["▲", ""],
    [NEGATIVE_STYLES.whiteTriangle]: ["△", ""],
    [NEGATIVE_STYLES.parentheses]: ["(", ")"],
};

/**
 * 数値を単位で割って書式化し、マイナスと 0 の書き方を当てる。suffix（単位の語・% など）は符号の内側に入る（▲1.2億、(1.2億)）。
 * 0 かどうかは丸めたあとの表示で決める（小数 0 桁の 0.4 は 0、-0.4 は丸めて 0 のマイナス）
 */
export function formatSigned(value: number, divisor: number, precision: string, style: Partial<SignStyle> = {}, suffix = ""): string {
    const s = { ...SIGN_STYLE_DEFAULTS, ...style };
    const abs = formatValue(Math.abs(value), divisor, precision);
    const negative = value < 0 || Object.is(value, -0);
    const [open, close] = NEGATIVE_MARKS[s.negative] ?? NEGATIVE_MARKS[NEGATIVE_STYLES.minus];
    // 数字の 1〜9 が 1 つも無ければ、表示は 0（NaN・∞ は数字が無くても 0 ではない）
    const shownZero = Number.isFinite(value) && !/[1-9]/.test(abs);
    if (shownZero && !(negative && s.negativeZero)) {
        if (s.zero === ZERO_STYLES.dash) return "-";
        if (s.zero === ZERO_STYLES.plusMinus) return `±${abs}${suffix}`;
        // 「0」：+ を付けるなら、丸めて 0 になるプラスは +0（前の版と同じ）。丸めて 0 のマイナスで符号を残さないときは 0
        return `${s.plus && value > 0 ? "+" : ""}${abs}${suffix}`;
    }
    if (negative) return `${open}${abs}${suffix}${close}`;
    return `${s.plus && value > 0 ? "+" : ""}${abs}${suffix}`;
}
