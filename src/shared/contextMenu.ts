import powerbi from "powerbi-visuals-api";

type ISelectionId = powerbi.visuals.ISelectionId;

/**
 * 右クリックのメニュー（selectionManager.showContextMenu）に渡す選択 ID。ids は押した所が選ぶ行（選ぶときと同じもの）。
 * 1 行だけならその行（含める・除外・ドリルスルーがその行に効く）。複数の行をまとめた所（その他・小計・比較の増減など）と、
 * 行の無い所（背景・選べない合計）は {}（ビジュアルのメニュー。Microsoft の例と同じ）。先頭の 1 行を渡すとその行だけのメニューになり、
 * null はメニューが出ないホストがあるかもしれない。推測で、実機では確かめていない。どちらも渡さない
 */
export function contextMenuIdOf(ids: readonly ISelectionId[] | null | undefined): ISelectionId {
    return ids?.length === 1 ? ids[0] : ({} as ISelectionId);
}
