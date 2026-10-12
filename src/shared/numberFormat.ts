import { UnitItem, formatValue } from "./units";
import { readableKeepingHue } from "./color";

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

/** 単位で割って丸めた絶対値の文字と、マイナスか、表示が 0 か。formatSigned と shownSignOf が同じ丸めを使う */
function shownOf(value: number, divisor: number, precision: string): { abs: string; negative: boolean; shownZero: boolean } {
    const abs = formatValue(Math.abs(value), divisor, precision);
    // 数字の 1〜9 が 1 つも無ければ、表示は 0（NaN・∞ は数字が無くても 0 ではない）
    return { abs, negative: value < 0 || Object.is(value, -0), shownZero: Number.isFinite(value) && !/[1-9]/.test(abs) };
}

/**
 * 数値を単位で割って書式化し、マイナスと 0 の書き方を当てる。suffix（単位の語・% など）は符号の内側に入る（▲1.2億、(1.2億)）。
 * 0 かどうかは丸めたあとの表示で決める（小数 0 桁の 0.4 は 0、-0.4 は丸めて 0 のマイナス）
 */
export function formatSigned(value: number, divisor: number, precision: string, style: Partial<SignStyle> = {}, suffix = ""): string {
    const s = { ...SIGN_STYLE_DEFAULTS, ...style };
    const { abs, negative, shownZero } = shownOf(value, divisor, precision);
    const [open, close] = NEGATIVE_MARKS[s.negative] ?? NEGATIVE_MARKS[NEGATIVE_STYLES.minus];
    if (shownZero && !(negative && s.negativeZero)) {
        if (s.zero === ZERO_STYLES.dash) return "-";
        if (s.zero === ZERO_STYLES.plusMinus) return `±${abs}${suffix}`;
        // 「0」：+ を付けるなら、丸めて 0 になるプラスは +0（前の版と同じ）。丸めて 0 のマイナスで符号を残さないときは 0
        return `${s.plus && value > 0 ? "+" : ""}${abs}${suffix}`;
    }
    if (negative) return `${open}${abs}${suffix}${close}`;
    return `${s.plus && value > 0 ? "+" : ""}${abs}${suffix}`;
}

// --- 符号の色（良い・悪い） -------------------------------------------------------

/**
 * 表示の符号。formatSigned と同じ引数・同じ丸めで、見える符号を -1・0・1 で返す。丸めて 0 は 0（±0・- と同じく色を付けない）。
 * ただし丸めて 0 のマイナスで符号を残す（▲0）ときは -1（見た目の符号と色をそろえる）。NaN は 0
 */
export function shownSignOf(value: number, divisor: number, precision: string, style: Partial<SignStyle> = {}): -1 | 0 | 1 {
    if (Number.isNaN(value)) return 0;
    const s = { ...SIGN_STYLE_DEFAULTS, ...style };
    const { negative, shownZero } = shownOf(value, divisor, precision);
    if (shownZero) return negative && s.negativeZero ? -1 : 0;
    return negative ? -1 : 1;
}

/** 色の塗り方。none = 色を付けない、bad = 悪い（マイナス）だけ、both = 良いと悪い（プラスとマイナス） */
export const TONE_MODES = { none: "none", bad: "bad", both: "both" } as const;

/** 値の符号で塗る（プラスとマイナス）ときの選択肢の名前 */
export const SIGN_TONE_MODE_ITEMS: UnitItem[] = [
    { value: TONE_MODES.none, displayName: "色なし" },
    { value: TONE_MODES.bad, displayName: "マイナスだけ" },
    { value: TONE_MODES.both, displayName: "プラスとマイナス" },
];

/** 差の良し悪しで塗る（増えて良い・増えて悪い）ときの選択肢の名前 */
export const DIFF_TONE_MODE_ITEMS: UnitItem[] = [
    { value: TONE_MODES.none, displayName: "色なし" },
    { value: TONE_MODES.bad, displayName: "悪い差だけ" },
    { value: TONE_MODES.both, displayName: "良い差と悪い差" },
];

/** 良い・悪いの既定の色（Power BI の既定のテーマの良い・悪いの色） */
export const DEFAULT_GOOD_COLOR = "#1AAB40";
export const DEFAULT_BAD_COLOR = "#D64554";

/**
 * 良い・悪いの字が地に対して読める濃さ（4.5:1。小さい字の読みやすさの目安、WCAG AA）。テーマの良い・悪いの色が淡ければ
 * （Fluent 2 の #9FD89F・#EEACB2）、色相と彩度を保ったまま暗くする
 */
export const SENTIMENT_CONTRAST = 4.5;

/**
 * 符号で塗る文字の良い・悪いの色の既定。書式ペインで決めていなければ、テーマのセンチメント カラー（palette.positive・negative）を
 * 地に対して読める濃さにして使う。テーマの色が無い・読めなければ DEFAULT_GOOD_COLOR・DEFAULT_BAD_COLOR から作る
 */
export function sentimentTextColors(
    palette: { positive?: { value?: string }; negative?: { value?: string }; background?: { value?: string }; foreground?: { value?: string } } | undefined
): { good: string; bad: string } {
    const valid = (info: { value?: string } | undefined, fallback: string) => (info?.value && /^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(info.value) ? info.value : fallback);
    const background = valid(palette?.background, "#FFFFFF");
    const text = valid(palette?.foreground, "#252423");
    return {
        good: readableKeepingHue(valid(palette?.positive, DEFAULT_GOOD_COLOR), background, SENTIMENT_CONTRAST, text),
        bad: readableKeepingHue(valid(palette?.negative, DEFAULT_BAD_COLOR), background, SENTIMENT_CONTRAST, text),
    };
}

export type Tone = "good" | "bad";

/**
 * 良い・悪い。shownSign（shownSignOf の結果）に良い向き good（1 = 増えて良い、-1 = 増えて悪い、0 = 色を付けない）を掛けて決める。
 * 符号で塗るとき（プラスが良い）は good = 1。mode が bad なら良い側は null、none なら常に null
 */
export function toneOf(shownSign: number, good = 1, mode: string = TONE_MODES.both): Tone | null {
    if (mode === TONE_MODES.none || !shownSign || !good) return null;
    const tone: Tone = shownSign * good > 0 ? "good" : "bad";
    return tone === "good" && mode === TONE_MODES.bad ? null : tone;
}

/** 値の書式が %（構成比・率）か。% の書式のメジャーは、表示単位で割らず、単位ラベルも付けずに % で出す */
export function isPercentFormat(format: string | undefined): boolean {
    return /%/.test(format ?? "");
}

/** % の書式の小数の桁（0.0% なら 1、0% なら 0） */
export function percentDigitsOf(format: string | undefined): number {
    const match = /\.(0+)\s*%/.exec(format ?? "");
    return match ? match[1].length : 0;
}

/**
 * 割合（0.352）を % の文字（35.2%）にする。precision は小数点以下の桁数の設定（"auto" なら書式の桁 formatDigits）。
 * style を渡すと、マイナスと 0 の書き方（▲・±0 など）と + を当てる
 */
export function formatPercent(value: number, precision: string, formatDigits: number, style?: Partial<SignStyle>): string {
    const digits = precision === "auto" ? String(formatDigits) : precision;
    return `${style ? formatSigned(value * 100, 1, digits, style) : formatValue(value * 100, 1, digits)}%`;
}

/**
 * 書式ペインの良い・悪いの色を使うか。保存してある（saved）か、既定の色から変えてあれば使い、そうでなければ空（テーマの色を使う）。
 * 既定の色と同じ色を選んで保存したときも、保存があれば使う
 */
export function ownToneColor(value: string | undefined, saved: unknown, fallback: string): string {
    return value && (saved != null || value.toUpperCase() !== fallback.toUpperCase()) ? value : "";
}

/**
 * 書式ペインの良い・悪いの色に、実際に塗る色（テーマから作った色）を出す（保存はしない）。書式で決めていなければ置き換える
 */
export function showThemeToneColors(
    slices: { positiveColor: { value: { value: string } }; negativeColor: { value: { value: string } } },
    saved: Record<string, unknown> | undefined,
    tones: { good: string; bad: string }
): void {
    if (!ownToneColor(slices.positiveColor.value?.value, saved?.positiveColor, DEFAULT_GOOD_COLOR)) slices.positiveColor.value = { value: tones.good };
    if (!ownToneColor(slices.negativeColor.value?.value, saved?.negativeColor, DEFAULT_BAD_COLOR)) slices.negativeColor.value = { value: tones.bad };
}
