import powerbi from "powerbi-visuals-api";

/** 読み込むカテゴリ（行）の上限。capabilities の dataReductionAlgorithm の top と同じ数にする */
export const ROW_LIMIT = 30000;
/** 読み込む凡例の値の上限。capabilities の dataReductionAlgorithm の top と同じ数にする */
export const SERIES_LIMIT = 2000;

/** 凡例の値が上限に達したときの警告 */
export const SERIES_TRUNCATED_TITLE = "すべての凡例の値を読み込めていません";
export const SERIES_TRUNCATED_NOTICE = "凡例の値が 2,000 に達したので、先頭の 2,000 で描いています。フィルターで絞ってください。";

/** ドリルダウンした位置をつなぐ区切り（「事業A ＞ 製品A1」） */
export const DRILL_PATH_SEPARATOR = " ＞ ";

export interface VisualWarning {
    title: string;
    detail: string;
}

/**
 * 警告のアイコンを出す。描き直すたびに消えるので、update のたびに呼ぶ。
 * アイコンは 1 つしか出せないので、題は先頭の警告、本文はすべての警告をつなぐ
 */
export function displayWarnings(host: powerbi.extensibility.visual.IVisualHost, warnings: Array<VisualWarning | null | undefined | false | "">): void {
    const shown = warnings.filter((w): w is VisualWarning => !!w);
    if (shown.length) host.displayWarningIcon(shown[0].title, shown.map((w) => w.detail).join(" "));
}
