import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import { BLANK_TEXT } from "./tooltip";

import PrimitiveValue = powerbi.PrimitiveValue;

/**
 * 全角で入れた数字・符号・桁区切り・小数点を半角にする（IME がオンのまま入れた「１，０００」「－５０」「１．５」など）。
 * マイナスは全角のハイフンマイナス・数学のマイナス・長音（日本語入力で - のキーを押すと「ー」になる）・ハイフンやダッシュの類
 * （‐ ‑ ‒ – ﹣）・半角の長音（ｰ）も読む。読点「、」は桁区切りにしない（「1、2」を 12 にしない）
 */
function toHalfWidth(text: string): string {
    return text
        .replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0))
        .replace(/[－−ー‐‑‒–﹣ｰ]/g, "-")
        .replace(/＋/g, "+")
        .replace(/，/g, ",")
        .replace(/．/g, ".");
}

/**
 * 桁区切りのカンマが正しい位置にある数の字（小数点より右には置かない）。
 * 3 桁ずつ（1,000）、インド式（最後が 3 桁、その前が 2 桁ずつ。10,00,000）、4 桁ずつ（1,0000）
 */
const GROUPED_NUMBERS = [
    /^[+-]?\d{1,3}(,\d{3})+(\.\d*)?$/,
    /^[+-]?\d{1,2}(,\d{2})*,\d{3}(\.\d*)?$/,
    /^[+-]?\d{1,4}(,\d{4})+(\.\d*)?$/,
];

/**
 * 書式ペインに入れた数の文字を数にする。全角も、桁区切りのカンマ（1,000）も読む。
 * カンマは正しい位置のときだけ桁区切りとして外す（「0,5」「12,34」「1,2,3」「1,5」は小数点のカンマと区別できないので数にしない）。
 * 空・数でない・有限でなければ null（自動や未設定の扱いは呼ぶ側で決める）
 */
export function parseNumberText(text: string | number | undefined | null): number | null {
    if (typeof text === "number") return Number.isFinite(text) ? text : null;
    let trimmed = toHalfWidth(String(text ?? "")).trim();
    if (!trimmed) return null;
    if (trimmed.includes(",")) {
        if (!GROUPED_NUMBERS.some((pattern) => pattern.test(trimmed))) return null;
        trimmed = trimmed.replace(/,/g, "");
    }
    const value = Number(trimmed);
    return Number.isFinite(value) ? value : null;
}

/**
 * DataView の値を数にする。空白（null・undefined・空の文字）と、数にならないもの・NaN・±Infinity は null。
 * 0 は 0（空白と区別する）。文字は書式ペインの入力と同じく全角と桁区切りも読む。日付は時刻の数、真偽は 1・0
 */
export function numberOf(raw: PrimitiveValue | undefined | unknown): number | null {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
    if (typeof raw === "string") return parseNumberText(raw);
    const value = Number(raw);
    return Number.isFinite(value) ? value : null;
}

/** 表示では空白にする値か（null・undefined・空の文字・NaN・±Infinity）。数として出す値に使う */
export function isBlankValue(raw: PrimitiveValue | undefined | unknown): boolean {
    return raw === null || raw === undefined || raw === "" || (typeof raw === "number" && !Number.isFinite(raw));
}

/** 名前として出すもの（カテゴリ・凡例など）で空白にする値か（null・undefined・空の文字）。NaN・±Infinity は空白にしない */
export function isBlankName(raw: PrimitiveValue | undefined | unknown): boolean {
    return raw === null || raw === undefined || raw === "";
}

/** 日付の書式が無い（書式が無い・数の書式だけ）ときに日付を出す書式 */
export const DEFAULT_DATE_FORMAT = "yyyy/MM/dd";

/** 1 文字の標準の日付の書式（.NET。Power BI は日時の列の「General Date」などをこの形で届ける） */
const STANDARD_DATE_FORMATS = "dDfFgGmMoOrRstTuUyY";

/** 名前の付いた日付の書式と、同じ意味の 1 文字（書式のライブラリは名前のままでは書けず、崩れた字になる） */
const NAMED_DATE_FORMATS: Record<string, string> = {
    "general date": "G",
    "long date": "D",
    "short date": "d",
    "long time": "T",
    "short time": "t",
};

/**
 * 名前の付いた数の書式と、数に当てる書式（日付の書式にしない）。書式のライブラリは名前のままでは書けず、書式の名前の字そのものを返す。
 * 意味の決まるものは同じ意味の書式に、決まらないもの（通貨の記号が地域で変わる Currency、Yes/No・True/False・On/Off）は書式なし（undefined）にする
 */
const NAMED_NUMBER_FORMATS: Record<string, string | undefined> = {
    "general number": undefined,
    fixed: "0.00",
    standard: "#,0.00",
    percent: "0.00%",
    scientific: "0.00E+00",
    currency: undefined,
    "yes/no": undefined,
    "true/false": undefined,
    "on/off": undefined,
};

const isNamedNumberFormat = (format: string): boolean => Object.prototype.hasOwnProperty.call(NAMED_NUMBER_FORMATS, format.trim().toLowerCase());

/** 数に当てる書式。名前の付いた数の書式は同じ意味の書式か書式なし（undefined）に直し、ほかはそのまま */
export function numberFormatOf(format: string | undefined): string | undefined {
    if (!format || !isNamedNumberFormat(format)) return format;
    return NAMED_NUMBER_FORMATS[format.trim().toLowerCase()];
}

/** 1 文字の d D f F g G r R は、日付の標準の書式でもあり数の標準の書式でもある（数にはそのまま当てる） */
const isNumberStandardToo = (format: string): boolean => format.length === 1 && "dDfFgGrR".includes(format);

/**
 * 数の値だけに当てる書式（目盛り・ラベルの欄の数など）。日付の書式は書式なし（undefined。数を書式の字そのものにしない）、
 * 名前の付いた数の書式は numberFormatOf で直す。valueTextOf・nameTextOf が数に当てる書式と同じ
 */
export function numberValueFormatOf(format: string | undefined): string | undefined {
    return dateFormatOf(format) && !isNumberStandardToo(format!) ? undefined : numberFormatOf(format);
}

/**
 * 日付・時刻に当てる書式。日付の書式でなければ null。
 * 1 文字は標準の日付の書式のときだけ、名前の付いた書式は同じ意味の 1 文字に直す。名前の付いた数の書式（Currency など）は日付にしない。
 * ほかは引用・エスケープ・角かっこの中（字として出す部分）を除いて、数の場所取り（0・#）があれば数の書式
 * （「#,0 kWh」「FY0」の単位の字を日付の字と読まない。秒の小数の ss.000 は除いて見る）、無ければ日付・時刻の字（y M d H h s）があるか。
 * 「D2」「N0」のような 1 文字と桁数の数の書式は日付にしない
 */
export function dateFormatOf(format: string | undefined): string | null {
    if (!format) return null;
    const key = format.trim().toLowerCase();
    const named = NAMED_DATE_FORMATS[key];
    if (named) return named;
    if (isNamedNumberFormat(key)) return null;
    if (format.length === 1) return STANDARD_DATE_FORMATS.includes(format) ? format : null;
    if (/^[A-Za-z]\d+$/.test(format)) return null;
    const bare = format.replace(/"[^"]*"|'[^']*'|\\.|\[[^\]]*\]/g, "");
    if (/[0#]/.test(bare.replace(/s\.0+/gi, "s"))) return null;
    return /[yMdHhs]|date|time/i.test(bare) ? format : null;
}

/** 書式なしで出す（数に日付の書式が来たとき、書式の字そのものを出さないため） */
const plainFormatter = valueFormatter.create({});

function textOf(
    format: string | undefined,
    blank: string,
    empty: string,
    isBlank: (raw: unknown) => boolean
): (raw: PrimitiveValue | undefined | unknown) => string {
    // 名前の付いた数の書式（Currency など）は、当てる前に同じ意味の書式か書式なしに直す
    const formatter = valueFormatter.create({ format: numberFormatOf(format) });
    const dateFormat = dateFormatOf(format);
    const dateFormatter = valueFormatter.create({ format: dateFormat ?? DEFAULT_DATE_FORMAT });
    // 数に日付の書式が来たら書式なし（1 文字の d D f F g G r R は数の標準の書式でもあるので、数にはそのまま当てる）
    const numberFormatter = dateFormat && !isNumberStandardToo(format!) ? plainFormatter : formatter;
    return (raw) => {
        if (raw === "") return empty;
        if (isBlank(raw)) return blank;
        if (raw instanceof Date) return Number.isFinite(raw.getTime()) ? dateFormatter.format(raw) : blank;
        if (typeof raw === "number") return numberFormatter.format(raw);
        return formatter.format(raw as PrimitiveValue);
    };
}

/**
 * 値を表示の文字にする関数を作る（行ごとに書式を作り直さないよう、書式 1 つにつき 1 回作る）。数として出す値に使う。
 * 空白（null・undefined・NaN・±Infinity）は blank、空の文字は empty（既定は blank と同じ）。数・文字・真偽は書式（valueFormatter）を通す。
 * 日付は書式に日付が無ければ yyyy/MM/dd（String() の「Wed Apr 01 2026 … GMT…」や、数の書式の文字そのものを出さない）。
 * 数に日付の書式が来たら、書式を当てずに数として出す
 */
export function valueTextOf(
    format: string | undefined,
    blank: string = BLANK_TEXT,
    empty: string = blank
): (raw: PrimitiveValue | undefined | unknown) => string {
    return textOf(format, blank, empty, isBlankValue);
}

/**
 * 名前として出すもの（カテゴリ・凡例・階層・イベントの名前など）を表示の文字にする関数を作る。
 * valueTextOf と同じだが、NaN・Infinity・-Infinity は空白にせず、互いに区別できる字（NaN・+Infinity・-Infinity）で出す
 */
export function nameTextOf(
    format: string | undefined,
    blank: string = BLANK_TEXT,
    empty: string = blank
): (raw: PrimitiveValue | undefined | unknown) => string {
    return textOf(format, blank, empty, isBlankName);
}

/** 列の書式の文字（モデルの書式。無ければ列の型から決まる既定） */
export function formatOfColumn(column: powerbi.DataViewMetadataColumn | undefined): string | undefined {
    return column ? valueFormatter.getFormatStringByColumn(column) : undefined;
}

/**
 * 保存された選択肢の値を、選択肢の値の並びから探す。Desktop は数字の形の値（"2024"・"0410"）を数で保存して数で届けるので、
 * まず文字の完全一致、無ければ数として同じもの（"0410" と 410）を探す。見つからなければ -1
 */
export function savedChoiceIndex(raw: unknown, values: readonly string[]): number {
    // 数で届いたときも、まず字の完全一致（410 なら "410"）。並び順で "0410" を先に取らない
    if (typeof raw === "string" || (typeof raw === "number" && Number.isFinite(raw))) {
        const exact = values.indexOf(String(raw));
        if (exact >= 0) return exact;
    }
    const target = typeof raw === "number" ? raw : typeof raw === "string" ? parseNumberText(raw) : null;
    if (target === null || !Number.isFinite(target)) return -1;
    return values.findIndex((value) => parseNumberText(value) === target);
}

/** 保存された選択肢の値を文字にする（数で届いた値も文字に戻す）。文字でも数でもなければ null */
export function savedChoiceText(raw: unknown): string | null {
    if (typeof raw === "string") return raw;
    if (typeof raw === "number" && Number.isFinite(raw)) return String(raw);
    return null;
}
