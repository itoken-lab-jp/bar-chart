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
