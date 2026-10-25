/**
 * 数の字（ja-JP の桁区切りと小数の桁）。toLocaleString("ja-JP", …) と同じ字を返す。
 * toLocaleString は呼ぶたびに書式（Intl.NumberFormat）を作り直して遅いので、作った書式を桁の組ごとに覚えて使い回す
 */

const formats = new Map<string, Intl.NumberFormat>();

/**
 * value.toLocaleString("ja-JP", { minimumFractionDigits, maximumFractionDigits, useGrouping }) と同じ字。
 * useGrouping を省くと、省いた toLocaleString と同じ（ブラウザーの既定の桁区切り）
 */
export function localeNumber(value: number, minimumFractionDigits: number, maximumFractionDigits: number, useGrouping?: boolean): string {
    const key = `${minimumFractionDigits}|${maximumFractionDigits}|${useGrouping}`;
    let format = formats.get(key);
    if (!format) {
        format = new Intl.NumberFormat("ja-JP", { minimumFractionDigits, maximumFractionDigits, ...(useGrouping === undefined ? {} : { useGrouping }) });
        formats.set(key, format);
    }
    return format.format(value);
}
