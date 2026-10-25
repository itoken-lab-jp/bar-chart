import { UnitItem, formatValue, type Grouping, groupingAt } from "./units";
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
    /**
     * メジャーの書式の負の節（sectionSignOf）。あれば negative の代わりに使う。丸めて 0 になる負の値にも当てる
     * （負の値の書き方。negativeZero を切れば、丸めて 0 の負の値は数だけ）
     */
    negativeSection?: SectionText;
    /**
     * メジャーの書式の 0 の節（sectionSignOf）。あれば zero の代わりに、ちょうど 0 の値だけに当てる。
     * 丸めて 0 になる値は、丸める前の符号で正の書き方か負の書き方（標準の書式と同じ）
     */
    zeroSection?: SectionText;
}

/** 書式の節の書き方。text は数を書かずにその字だけを出す（数の記号の無い節）。marks は数を前後の字で包む */
export type SectionText = { text: string } | { marks: [string, string] };

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
function shownOf(value: number, divisor: number, precision: string, grouping?: Grouping): { abs: string; negative: boolean; shownZero: boolean } {
    const abs = formatValue(Math.abs(value), divisor, precision, grouping);
    // 数字の 1〜9 が 1 つも無ければ、表示は 0（NaN・∞ は数字が無くても 0 ではない）
    return { abs, negative: value < 0 || Object.is(value, -0), shownZero: Number.isFinite(value) && !/[1-9]/.test(abs) };
}

/**
 * 数値を単位で割って書式化し、マイナスと 0 の書き方を当てる。suffix（単位の語・% など）は符号の内側に入る（▲1.2億、(1.2億)）。
 * 0 かどうかは丸めたあとの表示で決める（小数 0 桁の 0.4 は 0、-0.4 は丸めて 0 のマイナス）。
 * 書式の 0 の節（style.zeroSection）を当てるときだけは、ちょうど 0 の値が 0 の節で、丸めて 0 の値は丸める前の符号で書く。
 * grouping は桁区切りを付けるか（formatValue と同じ。省くと付ける）
 */
export function formatSigned(value: number, divisor: number, precision: string, style: Partial<SignStyle> = {}, suffix = "", grouping?: Grouping): string {
    const { abs, shownZero } = shownOf(value, divisor, precision, grouping);
    return signedText(value, abs, shownZero, { ...SIGN_STYLE_DEFAULTS, ...style }, suffix);
}

/** 絶対値の文字 abs（単位で割って丸めたもの）に、値の符号の書き方を当てる。formatSigned と目盛り（formatSectionTick）が使う */
function signedText(value: number, abs: string, shownZero: boolean, s: SignStyle, suffix: string): string {
    const negative = value < 0 || Object.is(value, -0);
    const plus = s.plus && value > 0 ? "+" : "";
    if (s.zeroSection) {
        // 書式の 0 の節は、ちょうど 0 の値だけ（-0 も 0）。丸めて 0 の負の値で符号を残さないときは数だけ
        if (value === 0) return sectionText(s.zeroSection, abs, suffix);
        if (shownZero && negative && !s.negativeZero) return `${abs}${suffix}`;
    } else if (shownZero && !(negative && s.negativeZero)) {
        if (s.zero === ZERO_STYLES.dash) return "-";
        if (s.zero === ZERO_STYLES.plusMinus) return `±${abs}${suffix}`;
        // 「0」：+ を付けるなら、丸めて 0 になるプラスは +0（前の版と同じ）。丸めて 0 のマイナスで符号を残さないときは 0
        return `${plus}${abs}${suffix}`;
    }
    if (negative) {
        if (s.negativeSection) return sectionText(s.negativeSection, abs, suffix);
        const [open, close] = NEGATIVE_MARKS[s.negative] ?? NEGATIVE_MARKS[NEGATIVE_STYLES.minus];
        return `${open}${abs}${suffix}${close}`;
    }
    return `${plus}${abs}${suffix}`;
}

function sectionText(section: SectionText, abs: string, suffix: string): string {
    return "text" in section ? section.text : `${section.marks[0]}${abs}${suffix}${section.marks[1]}`;
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

// --- 率の書式（% と ‰） ----------------------------------------------------------
// 判定は Power BI の書式（powerbi-visuals-utils-formattingutils）と同じにする。引用（"…"・'…'）と \ のあとの 1 字は
// 書式の記号ではなく字なので、0.0"%" や 0.0\% は 100 倍しない（35.2 を「35.2%」と出す。値がすでに % の数の率）。
// 引用の外の % は 100 倍、‰ は 1000 倍。; で分けた書式（正;負;0）は、桁と倍率を先頭（正の値）の節で決める

/** 書式の中の字：引用（"…"・'…'。閉じていなければ最後まで）と、\ で始まる 1 字 */
const LITERAL_PATTERN = /'[^']*'|"[^"]*"|\\.|'[^']*$|"[^"]*$|\\$/;

interface FormatToken {
    /** 書式に書いたままの文字（引用符・\ を含む） */
    raw: string;
    /** 表示される文字（字なら引用符と \ を外したもの） */
    text: string;
    /** 字（引用・エスケープの中）か */
    literal: boolean;
}

function tokensOf(format: string): FormatToken[] {
    const tokens: FormatToken[] = [];
    const pattern = new RegExp(LITERAL_PATTERN.source, "g");
    let last = 0;
    for (let match = pattern.exec(format); match; match = pattern.exec(format)) {
        if (match.index > last) tokens.push({ raw: format.slice(last, match.index), text: format.slice(last, match.index), literal: false });
        const raw = match[0];
        const closed = raw[0] !== "\\" && raw.length > 1 && raw.endsWith(raw[0]);
        tokens.push({ raw, text: raw.slice(1, closed ? -1 : undefined), literal: true });
        last = match.index + raw.length;
    }
    if (last < format.length) tokens.push({ raw: format.slice(last), text: format.slice(last), literal: false });
    return tokens;
}

/** 書式の字を、引用・エスケープの中の字（literal）と、書式の記号として読む字に分ける。字は引用符と \ を外して返す */
export function formatTokensOf(format: string): Array<{ text: string; literal: boolean }> {
    return tokensOf(format).map(({ text, literal }) => ({ text, literal }));
}

/** ; で分けた書式の節（正;負;0）。引用・エスケープの中の ; では分けない。節は書いたままの文字 */
export function formatSectionsOf(format: string): string[] {
    const sections = [""];
    for (const token of tokensOf(format)) {
        if (token.literal) {
            sections[sections.length - 1] += token.raw;
            continue;
        }
        const parts = token.raw.split(";");
        sections[sections.length - 1] += parts[0];
        sections.push(...parts.slice(1));
    }
    return sections;
}

// --- 書式の節（正;負;0）の符号の書き方 -----------------------------------------------
// 標準のビジュアルは、メジャーの書式の負の節・0 の節でデータ ラベルと目盛りを書く（#,0;(#,0);- なら負は「(320)」、0 は「-」）。
// このビジュアルは数を自前で書く（表示単位で割る・桁・単位の語）ので、節の数の記号の前後の字だけを読み、符号の書き方として当てる

/** 数の記号（桁・桁区切り・小数点・率の記号）。これより前の字と後ろの字が節の飾り */
const NUMBER_SYMBOLS = /[0#?,.%\u2030]/;

interface SectionParts {
    /** 最初の数の記号より前の字 */
    prefix: string;
    /** 最後の数の記号より後ろの字 */
    suffix: string;
    hasNumber: boolean;
    /** 節の字すべて（数の記号の無い節で出す） */
    text: string;
}

/** 節を、数の記号の前の字・後ろの字に分ける。字（引用・エスケープ）は引用符と \ を外す。数の記号のあいだの字は使わない */
function sectionPartsOf(section: string): SectionParts {
    const parts: SectionParts = { prefix: "", suffix: "", hasNumber: false, text: "" };
    for (const token of tokensOf(section)) {
        for (const ch of token.literal ? [token.text] : [...token.raw]) {
            parts.text += ch;
            if (!token.literal && NUMBER_SYMBOLS.test(ch)) {
                parts.hasNumber = true;
                parts.suffix = "";
            } else if (parts.hasNumber) parts.suffix += ch;
            else parts.prefix += ch;
        }
    }
    return parts;
}

/** text の先頭か末尾にある part を外す（正の節にもある字。"¥"#,0;"¥"-#,0 の ¥） */
function withoutPart(text: string, part: string): string {
    if (!part) return text;
    if (text.startsWith(part)) return text.slice(part.length);
    return text.endsWith(part) ? text.slice(0, -part.length) : text;
}

/** 負か 0 の節の書き方。数の記号が無ければ字だけ、あれば前後の字（正の節と同じ字は外す）で数を包む */
function sectionTextOf(parts: SectionParts, positive: SectionParts): SectionText {
    if (!parts.hasNumber) return { text: parts.text };
    return { marks: [withoutPart(parts.prefix, positive.prefix.replace(/\+/g, "")), withoutPart(parts.suffix, positive.suffix)] };
}

/**
 * メジャーの書式の節（正;負;0）から読んだ符号の書き方。節が 1 つ（; の無い書式）か、書式が無ければ null。
 * - 負の節：数の記号より前の字と後ろの字で負の値を包む（(#,0) は「(」と「)」、▲#,0 は「▲」、#,0- は後ろに「-」）。
 *   数の記号が無い節（空の節を含む）は、その字だけを出す（Power BI の書式ライブラリと同じ）
 * - 0 の節：数の記号が無ければその字（-・"ゼロ"）、あれば 0 を前後の字で包む。節が 2 つなら、0 は正の節で書く
 * - 正の節：前に付いた + だけを効かせる。ほかの字（"円" など）は付けないので、負と 0 の節からも、正の節の前後にある字は外す
 * 数そのものはこのビジュアルの書き方（表示単位・桁・単位の語・率の記号）で書き、印の内側に入れる。桁と倍率は先頭の節で決める（rateFormatOf）
 */
export function sectionSignOf(format: string | undefined): Partial<SignStyle> | null {
    if (!format) return null;
    const sections = formatSectionsOf(format);
    if (sections.length < 2) return null;
    const positive = sectionPartsOf(sections[0]);
    const plus = positive.prefix.includes("+");
    return {
        negativeSection: sectionTextOf(sectionPartsOf(sections[1]), positive),
        zeroSection: sections.length > 2 ? sectionTextOf(sectionPartsOf(sections[2]), positive) : { marks: [plus ? "+" : "", ""] },
        ...(plus ? { plus: true } : {}),
    };
}

/** 同じ軸に載る書式の節の書き方がすべて同じなら、その書き方。違えば（節の無い書式が混じっても）null。sameRateOf と同じ考え方 */
export function sameSectionSignOf(formats: Array<string | undefined>): Partial<SignStyle> | null {
    const styles = formats.map(sectionSignOf);
    const first = styles[0];
    if (!first) return null;
    const key = JSON.stringify(first);
    return styles.every((style) => style !== null && JSON.stringify(style) === key) ? first : null;
}

/**
 * データ ラベルの符号の書き方。書式ペインの「負の値の書き方」「0 の書き方」「+ を付ける」（plusSign）は、レポートに保存してあれば
 * （saved。dataView の objects の dataLabels）ペインを、保存が無ければメジャーの書式の節を使う（節も無ければペインの既定）。
 * + は、保存が無く書式の正の節に + があれば付け、無ければペインの値（既定を含む）のまま
 */
export function withSectionSign(
    pane: Partial<SignStyle>,
    sections: Partial<SignStyle> | null,
    saved: Record<string, unknown> | undefined
): Partial<SignStyle> {
    if (!sections) return pane;
    return {
        ...pane,
        ...(saved?.negativeStyle == null && sections.negativeSection ? { negativeSection: sections.negativeSection } : {}),
        ...(saved?.zeroStyle == null && sections.zeroSection ? { zeroSection: sections.zeroSection } : {}),
        ...(saved?.plusSign == null && sections.plus ? { plus: true } : {}),
    };
}

/**
 * 値の軸の目盛りに書式の節を当てる（標準と同じく、書式ペインの符号の書き方は使わない）。abs は目盛りの絶対値を書いた数、
 * suffix は単位の語・率の記号。節はラベルと同じく丸める前の値で選ぶ（ちょうど 0 の目盛りだけ 0 の節。目盛りは d3 の ticks で、
 * 0 の目盛りはちょうど 0。桁を絞って丸めて 0 になる目盛りは「0K」「(0K)」で、0 の節にしない）
 */
export function formatSectionTick(value: number, abs: string, sections: Partial<SignStyle>, suffix = ""): string {
    return signedText(value, abs, !/[1-9]/.test(abs), { ...SIGN_STYLE_DEFAULTS, ...sections }, suffix);
}

// --- 書式の桁区切り ---------------------------------------------------------------
// 判定は Power BI の書式（powerbi-visuals-utils-formattingutils の valueFormatter）と同じにする。
// ただし g・G（桁数なし）だけは標準のビジュアルに合わせる（standardGroupingOf）。
// 数を自前で書く（表示単位で割る・桁・単位の語）ときも、桁区切りだけは書式に従う（標準のビジュアルと同じ）

/** 1 文字と桁数の標準の書式（N0・C2・P1 など）。valueFormatter と同じ形 */
const STANDARD_FORMAT = /^[a-z]\d{0,2}$/i;

/**
 * 標準の書式（1 文字と桁数）の桁区切り。valueFormatter の字で確かめた：
 * D・E・F・R・X はなし。G0・G00 はあり、ほかの桁数（G2 など。有効桁）はなし。ほかの文字（N・C・P と、既定の書き方に落ちる
 * A・Q・Z など）はあり。
 * G の桁数なし（g・G。整数の列の既定で、書式ペインの「全般」）は、表示単位で割っても小数の桁を決めても桁区切りなし。
 * valueFormatter は小数の桁を渡すとき（割る・桁数を決める）に書式を #,0.00 に替えて桁区切りを付けるが、Desktop の標準の集合縦棒は
 * 「全般」の整数の列を表示単位「千」で「1250K」と書き、桁区切りを付けなかった（表示単位「なし」も「1250000」）。標準に合わせる。
 * 小数の桁を決めた形の標準は測っていない
 */
function standardGroupingOf(format: string): Grouping {
    if (/^g0+$/i.test(format)) return true;
    return !/^[defgrx]/i.test(format);
}

/**
 * 書式に桁区切りがあるか。先頭の節（正の値。; で分けた最初の節）の整数の部分（引用・エスケープの外の最初の小数点より前）で、
 * 最初の , が最後の桁の記号（0・#）より前にあれば桁区切りあり（#,0・#,##0・#,0.00・"¥"#,0）。
 * 桁の記号より後ろにしか無い ,（0, の末尾。千で割る印）は桁区切りではない（#,0, は前の , で桁区切りあり）。
 * 指数の書式は , があってもなし（valueFormatter は 1E+8・1.23e8 と書き、仮数は 1 桁）。指数かは valueFormatter と同じく、
 * 引用・エスケープの外の字に /e[+-]*[0#]/i が当たるか（E+0・E-00・E0・e00。符号は無くてもよく、e の後ろに 0 か # が要る）。
 * 当たらない e（"#,0 e-"・"#,0 Ex"・"#,0 one-off"）は字で、, の決め方のまま。valueFormatter は書式に e・E があると（字の中でも）
 * 表示単位で割らず、上書きの書式も渡さない（displayUnitSystem.js の hasScientitifcFormat）ので、表示単位・桁数によらず同じ決め方になる。
 * 標準の書式（1 文字と桁数）は standardGroupingOf。書式が無ければなし。
 * 型の既定の書式（valueFormatter.getFormatStringByColumn。小数の列は #,0.00、整数の列は g）を渡せば、書式の無い列もその書式で決まる。
 * 名前の付いた書式（Standard・Percent など）は numberFormatOf で直してから渡す
 */
export function groupingOf(format: string | undefined): Grouping {
    if (!format) return false;
    if (STANDARD_FORMAT.test(format)) return standardGroupingOf(format);
    // 字（引用・エスケープの中）は 1 字の目印に置き換え、記号として読む字だけを見る（rateFormatOf と同じ）
    const bare = tokensOf(formatSectionsOf(format)[0])
        .map((token) => (token.literal ? "\uE000" : token.raw))
        .join("");
    if (/e[+-]*[0#]/i.test(bare)) return false;
    const whole = bare.split(".")[0];
    const comma = whole.indexOf(",");
    return comma >= 0 && comma < Math.max(whole.lastIndexOf("0"), whole.lastIndexOf("#"));
}

/**
 * 1 つの数がいくつかの列から来るとき（同じ軸に載るメジャー・積み上げの合計など）の桁区切り。どれか 1 つの書式に桁区切りがあれば付ける
 * （書式が混ざれば付ける。前の版はいつも付けていて、付けて読み違えることは無い）。列が無ければ付けない。
 * 割る・桁数を決めたときだけ付く形（{ plain, decimals }）が混ざれば、plain と decimals を別々に「どれか」で決める
 */
export function someGroupingOf(formats: Array<string | undefined>): Grouping {
    const each = formats.map(groupingOf);
    const plain = each.some((g) => (typeof g === "boolean" ? g : g.plain));
    const decimals = each.some((g) => (typeof g === "boolean" ? g : g.decimals));
    return plain === decimals ? plain : { plain, decimals };
}

/**
 * value.toFixed(digits) の字に、grouping なら整数の部分に 3 桁ごとの , を入れる。丸めは toFixed のまま
 * （率の目盛りなど、toFixed で書いてきた数の桁区切りを書式に合わせる）。桁数を決めて書くので、{ plain, decimals } は decimals
 */
export function fixedNumber(value: number, digits: number, grouping: Grouping): string {
    const text = value.toFixed(digits);
    if (!groupingAt(grouping, 1, String(digits))) return text;
    const match = /^(-?)(\d+)(.*)$/.exec(text);
    if (!match) return text;
    return `${match[1]}${match[2].replace(/\B(?=(\d{3})+(?!\d))/g, ",")}${match[3]}`;
}

/** 率の書式：値に掛ける倍率（% は 100、‰ は 1000、字で書いた % は 1）、数のあとに付ける記号、小数の桁 */
export interface RateFormat {
    scale: number;
    mark: string;
    /** 先頭の節の小数点のあとの 0 の数（0.0% なら 1）。小数点以下の桁数が「自動」のときに使う */
    digits: number;
}

/** % の書式（100 倍して % を付ける） */
export const PERCENT_RATE: Pick<RateFormat, "scale" | "mark"> = { scale: 100, mark: "%" };

/**
 * 値の書式が率（% か ‰）なら、その倍率・記号・桁。率でなければ null。
 * 標準の書式 P（P1 なら小数 1 桁、P だけなら 2 桁）も %。
 * 字で書いた % か ‰ だけを付けた書式（0.0"%"・0.0\%）も率で、倍率は 1（値がすでに % の数。Power BI も 100 倍せずに % を付ける）
 */
export function rateFormatOf(format: string | undefined): RateFormat | null {
    if (!format) return null;
    const standard = /^[pP](\d{0,2})$/.exec(format);
    if (standard) return { scale: 100, mark: "%", digits: standard[1] ? Number(standard[1]) : 2 };
    // 字（引用・エスケープの中）は 1 字の目印に置き換え、記号として読む字だけを見る
    const tokens = tokensOf(formatSectionsOf(format)[0]);
    const bare = tokens.map((token) => (token.literal ? "\uE000" : token.raw)).join("");
    const dot = bare.indexOf(".");
    const digits = dot < 0 ? 0 : (/^[^eE]*/.exec(bare.slice(dot + 1))![0].match(/0/g) ?? []).length;
    const marks = bare.replace(/[^%\u2030]/g, "");
    if (marks) return { scale: (marks.includes("%") ? 100 : 1) * (marks.includes("\u2030") ? 1000 : 1), mark: marks, digits };
    // 字が % か ‰ だけなら倍率 1 の率（前の空白は残す：0.0" %" は「35.2 %」）。ほかの字も付いていれば（"%増"）数
    const text = tokens
        .filter((token) => token.literal)
        .map((token) => token.text)
        .join("")
        .trimEnd();
    return /^\s*[%\u2030]$/.test(text) ? { scale: 1, mark: text, digits } : null;
}

/** 値の書式が率（% か ‰。構成比・率）か。率の書式のメジャーは、表示単位で割らず、単位ラベルも付けずに率で出す */
export function isPercentFormat(format: string | undefined): boolean {
    return rateFormatOf(format) !== null;
}

/** 率の書式の小数の桁（0.0% なら 1、0% なら 0）。率でなければ 0 */
export function percentDigitsOf(format: string | undefined): number {
    return rateFormatOf(format)?.digits ?? 0;
}

/** いくつかの書式（同じ軸に載るメジャー）がすべて同じ率（倍率と記号）なら、その率（桁は先頭のもの）。そうでなければ null */
export function sameRateOf(formats: Array<string | undefined>): RateFormat | null {
    const rates = formats.map(rateFormatOf);
    const first = rates[0];
    if (!first) return null;
    return rates.every((rate) => rate !== null && rate.scale === first.scale && rate.mark === first.mark) ? first : null;
}

/**
 * 割合（0.352）を率の文字（35.2%）にする。precision は小数点以下の桁数の設定（"auto" なら書式の桁 formatDigits）。
 * style を渡すと、マイナスと 0 の書き方（▲・±0 など）と + を当てる。記号は符号の内側に入る（(2.0%)・▲2.0%。0 を - と書くなら -）。
 * rate は倍率と記号（既定は %。‰ の書式なら 1000 倍して ‰）。grouping は桁区切りを付けるか（省くと付ける）
 */
export function formatPercent(
    value: number,
    precision: string,
    formatDigits: number,
    style?: Partial<SignStyle>,
    rate: Pick<RateFormat, "scale" | "mark"> = PERCENT_RATE,
    grouping?: Grouping
): string {
    const digits = precision === "auto" ? String(formatDigits) : precision;
    return style
        ? formatSigned(value * rate.scale, 1, digits, style, rate.mark, grouping)
        : `${formatValue(value * rate.scale, 1, digits, grouping)}${rate.mark}`;
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
