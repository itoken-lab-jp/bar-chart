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
            // 型定義では省略できるメソッドだが、NumUpDown は実装を持つ
            ...super.getFormattingComponent!(objectName),
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

/**
 * 「…の色を一致させる」がオンなら、色の欄を隠す（一致させた色で塗るので、選んだ色は効かない）。オフなら出す。
 * 隠すのは書式ペインの表示だけで、保存値はそのまま残す
 */
export function hideColorWhenMatched(match: { value?: boolean }, ...colors: Array<{ visible?: boolean }>): void {
    const matched = match.value ?? false;
    for (const color of colors) color.visible = !matched;
}

/**
 * 組み立てた書式モデルの「設定の適用先」の項目と、その中の欄・まとまりの uid を番号で振り直す。
 * powerbi-visuals-utils-formattingmodel は、項目の uid を表示名から、欄の uid をカードと保存名から作るので、
 * 同じ名前の項目（同じ名前の系列が 2 つあるなど）や、カードの見出しのトグルと同じ保存名の欄が同じ uid になる
 * （Power BI はどちらの欄の値を出すかが決まらない）。uid は保存されないので、振り直しても保存値は変わらない
 */
export function numberContainerItems(model: powerbi.visuals.FormattingModel): powerbi.visuals.FormattingModel {
    for (const card of model.cards ?? []) {
        for (const group of (card as powerbi.visuals.FormattingCard).groups ?? []) {
            const container = (group as powerbi.visuals.FormattingGroup).container;
            container?.containerItems.forEach((item, i) => {
                item.uid = `${container.uid}-${i}`;
                const prefix = (target: { uid: string } | undefined) => {
                    if (target) target.uid = `${item.uid}-${target.uid}`;
                };
                item.slices?.forEach((slice) => prefix(slice));
                for (const inner of item.groups ?? []) {
                    prefix(inner);
                    prefix(inner.topLevelToggle);
                    inner.slices?.forEach((slice) => prefix(slice as { uid: string }));
                }
            });
        }
    }
    return model;
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
