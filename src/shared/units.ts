/**
 * 表示単位（万・億など）。0〜12 の桁数（指数）で割り、語を付ける。
 * 日本語の表記は 十・百・千・万…兆、英語の表記は K・M・bn・T だけ（1000 の冪以外は語が無い）
 */

export interface UnitDefinition {
    exponent: number;
    divisor: number;
    /** 解決された語（日本語なら 万・億、英語なら K・M・bn・T） */
    unitWord: string;
    unitWordJa: string;
    /** 標準（英語）の語。1000 の冪以外は語が無いので空 */
    unitWordStd: string;
    displayName: string;
}

// 標準表記の語は K/M/bn/T だけ。「10K」「100M」のような語を数字の後ろに付けると
// 1.23 + 10K = 「1.2310K」と読み違えるため、1000 の冪以外は空にして toStandardUnitKey で読み替える
export const UNIT_DEFINITIONS: Record<string, UnitDefinition> = {
    "0": { exponent: 0, divisor: 1, unitWord: "", unitWordJa: "", unitWordStd: "", displayName: "なし" },
    "1": { exponent: 1, divisor: 10, unitWord: "十", unitWordJa: "十", unitWordStd: "", displayName: "十" },
    "2": { exponent: 2, divisor: 100, unitWord: "百", unitWordJa: "百", unitWordStd: "", displayName: "百" },
    "3": { exponent: 3, divisor: 1_000, unitWord: "千", unitWordJa: "千", unitWordStd: "K", displayName: "千" },
    "4": { exponent: 4, divisor: 10_000, unitWord: "万", unitWordJa: "万", unitWordStd: "", displayName: "万" },
    "5": { exponent: 5, divisor: 100_000, unitWord: "十万", unitWordJa: "十万", unitWordStd: "", displayName: "十万" },
    "6": { exponent: 6, divisor: 1_000_000, unitWord: "百万", unitWordJa: "百万", unitWordStd: "M", displayName: "百万" },
    "7": { exponent: 7, divisor: 10_000_000, unitWord: "千万", unitWordJa: "千万", unitWordStd: "", displayName: "千万" },
    "8": { exponent: 8, divisor: 100_000_000, unitWord: "億", unitWordJa: "億", unitWordStd: "", displayName: "億" },
    "9": { exponent: 9, divisor: 1_000_000_000, unitWord: "十億", unitWordJa: "十億", unitWordStd: "bn", displayName: "十億" },
    "10": { exponent: 10, divisor: 10_000_000_000, unitWord: "百億", unitWordJa: "百億", unitWordStd: "", displayName: "百億" },
    "11": { exponent: 11, divisor: 100_000_000_000, unitWord: "千億", unitWordJa: "千億", unitWordStd: "", displayName: "千億" },
    "12": { exponent: 12, divisor: 1_000_000_000_000, unitWord: "兆", unitWordJa: "兆", unitWordStd: "T", displayName: "兆" },
};

/** 自動単位で使う段（小さい順）。日本語は 千・万・百万・億・兆、標準は K・M・bn・T */
const AUTO_KEYS_JA = ["0", "3", "4", "6", "8", "12"];
const AUTO_KEYS_STD = ["0", "3", "6", "9", "12"];

/** precision が "auto" のときの小数桁の上限（formatValue と同じ） */
const AUTO_FRACTION_DIGITS = 2;

/** 表示に使う小数桁で丸めた後の |value / divisor| */
function roundedScaledAbs(abs: number, divisor: number, precision: string): number {
    const digits = precision === "auto" ? AUTO_FRACTION_DIGITS : Math.max(0, parseInt(precision, 10) || 0);
    const factor = Math.pow(10, digits);
    return Math.round((abs / divisor) * factor) / factor;
}

/**
 * 小さい順の divisor 列から abs に合う段の添字を返す。
 * 丸めた結果が次の段に届く（9,999.9 → 「10千」、99,999,999 → 「10,000万」）ときは次の段へ繰り上げる。
 */
export function pickUnitIndex(abs: number, divisors: number[], precision: string): number {
    let index = 0;
    for (let i = divisors.length - 1; i >= 0; i--) {
        if (abs >= divisors[i]) {
            index = i;
            break;
        }
    }
    while (
        index + 1 < divisors.length &&
        roundedScaledAbs(abs, divisors[index], precision) >= divisors[index + 1] / divisors[index]
    ) {
        index++;
    }
    return index;
}

/** 最大絶対値から自動単位（指数）を判定。丸め後の桁上がりも考慮する */
export function resolveAutoUnitKey(maxAbsValue: number, notation = "japanese", precision = "auto"): string {
    const keys = notation === "standard" ? AUTO_KEYS_STD : AUTO_KEYS_JA;
    const divisors = keys.map((k) => UNIT_DEFINITIONS[k].divisor);
    return keys[pickUnitIndex(Math.abs(maxAbsValue), divisors, precision)];
}

/**
 * 標準表記で使えるキーに読み替える。K/M/bn/T の語が無い桁（1,2,4,5,7,8,10,11）は
 * 1 つ下の 1000 の冪に落とす（例: 4 = 1万 → 3 = K で「12.31K」）。
 * 数値を大きく割ったのに語が付かず 10 倍・100 倍に読まれる事故を防ぐ。
 */
export function toStandardUnitKey(key: string): string {
    const def = UNIT_DEFINITIONS[key];
    if (!def) return "0";
    return String(Math.floor(def.exponent / 3) * 3);
}

/** 選ばれたキー（auto を含む）から単位を決める */
export function resolveUnit(unitTypeKey: string, maxAbsValue: number, notation = "japanese", precision = "auto"): UnitDefinition {
    const standard = notation === "standard";
    let key = unitTypeKey === "auto" ? resolveAutoUnitKey(maxAbsValue, notation, precision) : unitTypeKey;
    if (!UNIT_DEFINITIONS[key]) key = "0";
    if (standard) key = toStandardUnitKey(key);
    const base = UNIT_DEFINITIONS[key];
    return { ...base, unitWord: standard ? base.unitWordStd : base.unitWordJa };
}

/** 数値を単位で割って書式化する。precision が auto なら小数 2 桁まで */
export function formatValue(value: number, divisor: number, precision: string): string {
    const scaled = value / divisor;
    if (precision === "auto") {
        return scaled.toLocaleString("ja-JP", { maximumFractionDigits: AUTO_FRACTION_DIGITS, minimumFractionDigits: 0 });
    }
    const digits = Math.max(0, parseInt(precision, 10) || 0);
    return scaled.toLocaleString("ja-JP", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

/** 書式ペインの選択肢。表示名は文字列（powerbi.IEnumMember としてそのまま渡せる） */
export interface UnitItem {
    value: string;
    displayName: string;
}

/**
 * 書式ペインの表示単位の選択肢。value (0〜12 の桁数) はレポートに保存されるので変えない。表示名だけを単位の語にする。
 * 英語表記に K/M/bn/T の語が無い桁（万・億など）は、1 つ下の K/M/bn/T（十・百は単位なし）に読み替える（toStandardUnitKey）
 */
export const UNIT_TYPES: UnitItem[] = [
    { value: "auto", displayName: "自動" },
    { value: "0", displayName: "なし" },
    { value: "1", displayName: "十" },
    { value: "2", displayName: "百" },
    { value: "3", displayName: "千" },
    { value: "4", displayName: "万" },
    { value: "5", displayName: "十万" },
    { value: "6", displayName: "百万" },
    { value: "7", displayName: "千万" },
    { value: "8", displayName: "億" },
    { value: "9", displayName: "十億" },
    { value: "10", displayName: "百億" },
    { value: "11", displayName: "千億" },
    { value: "12", displayName: "兆" },
];

/** 単位の表記。value は保存済みレポートとの互換のため変えない（"standard" = K・M・bn・T の表記） */
export const UNIT_NOTATIONS: UnitItem[] = [
    { value: "japanese", displayName: "日本語（万・億）" },
    { value: "standard", displayName: "英語（K・M・bn）" },
];

/** 小数点以下の桁数 */
export const PRECISIONS: UnitItem[] = [
    { value: "auto", displayName: "自動" },
    { value: "0", displayName: "0" },
    { value: "1", displayName: "1" },
    { value: "2", displayName: "2" },
    { value: "3", displayName: "3" },
];
