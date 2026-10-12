import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

/** 値が空のとき入力欄に出す文字。標準の「外側のパディング」と同じ */
export const AUTO_PLACEHOLDER = "自動";

/**
 * 値を空（undefined）にでき、空のとき入力欄に「自動」と出す NumUpDown。
 * API の visuals.NumUpDown は placeholderText を持つが、formattingmodel 6.0.4 の
 * NumUpDown は options しか渡さないので、ここで足す。
 * 空のまま = 保存値なし = 自動。書式ペインの「既定値に戻す」でも空に戻る。
 */
export class AutoNumUpDown extends formattingSettings.NumUpDown {
    getFormattingComponent(objectName: string): powerbi.visuals.NumUpDown {
        return {
            ...super.getFormattingComponent(objectName),
            placeholderText: AUTO_PLACEHOLDER,
        };
    }
}

/** 選択肢から値で選ぶ。無ければ先頭 */
export const itemOf = (items: powerbi.IEnumMember[], value: string): powerbi.IEnumMember =>
    items.find((i) => i.value === value) ?? items[0];

/**
 * ドロップダウンの値（IEnumMember か文字）を文字にする。無いとき・保存値が選択肢に無くて value が空のときは fallback
 */
export function enumValue(value: unknown, fallback: string): string {
    if (value === undefined || value === null) return fallback;
    if (typeof value === "object") {
        const inner = (value as { value?: unknown }).value;
        return inner === undefined || inner === null ? fallback : String(inner);
    }
    return String(value);
}

/** ItemDropdown の値の文字列。保存値が items に無いと value は undefined のままになるので、そのときは fallback */
export function dropdownValue(slice: formattingSettings.ItemDropdown, fallback: string): string {
    return enumValue(slice.value, fallback);
}

/** 軸のタイトルのスタイル（標準と同じ 3 つ） */
export const TITLE_STYLES = {
    showTitleOnly: "showTitleOnly",
    showUnitOnly: "showUnitOnly",
    showBoth: "showBoth",
} as const;

export const TITLE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: TITLE_STYLES.showTitleOnly, displayName: "タイトルのみを表示" },
    { value: TITLE_STYLES.showUnitOnly, displayName: "単位のみを表示" },
    { value: TITLE_STYLES.showBoth, displayName: "両方を表示" },
];

/** データ ラベルの文字の向き */
export const LABEL_ORIENTATIONS = {
    horizontal: "horizontal",
    vertical: "vertical",
} as const;

export const LABEL_ORIENTATION_ITEMS: powerbi.IEnumMember[] = [
    { value: LABEL_ORIENTATIONS.horizontal, displayName: "横" },
    { value: LABEL_ORIENTATIONS.vertical, displayName: "縦" },
];
