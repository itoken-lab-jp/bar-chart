import type * as React from "react";

/**
 * キーボードで選べる部品（棒・凡例の項目・軸のラベル）に付ける属性。
 * Tab で止まり、Enter・Space で選ぶ（Ctrl・⌘ を押しながらなら足す）、Esc で選択を外す。
 * 押しっぱなしの繰り返しでは選び直さない（選んだり外したりを繰り返さないように）
 */
export interface SelectKeyOptions {
    /** 読み上げの名前 */
    label: string;
    /** 選ぶ。multi は Ctrl・⌘ を押しながら */
    onSelect: (multi: boolean) => void;
    /** Esc で選択を外す。無ければ何もしない */
    onClear?: () => void;
    /** 選ばれているか（読み上げの「押されている」） */
    pressed?: boolean;
}

export interface SelectKeyProps {
    role: string;
    tabIndex: number;
    "aria-label": string;
    "aria-pressed"?: boolean;
    onKeyDown: (e: React.KeyboardEvent) => void;
}

export function selectKeyProps({ label, onSelect, onClear, pressed }: SelectKeyOptions): SelectKeyProps {
    return {
        role: "button",
        tabIndex: 0,
        "aria-label": label,
        ...(pressed === undefined ? {} : { "aria-pressed": pressed }),
        onKeyDown: selectKeyDown(onSelect, onClear),
    };
}

/** キーの扱いだけ（役割と名前は呼ぶ側が付けるとき）。Enter・Space で選ぶ（Ctrl・⌘ で足す、押し続けでは選び直さない）、Esc で外す */
export function selectKeyDown(onSelect: (multi: boolean) => void, onClear?: () => void): (e: React.KeyboardEvent) => void {
    return (e) => {
        if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            e.stopPropagation();
            if (!e.repeat) onSelect(e.ctrlKey || e.metaKey);
        } else if (e.key === "Escape" && onClear) {
            e.stopPropagation();
            onClear();
        }
    };
}
