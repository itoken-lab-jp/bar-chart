import type * as React from "react";

/**
 * キーボードで選べる部品（棒・凡例の項目・軸のラベル）に付ける属性。
 * Tab で止まり、Enter・Space で選ぶ（Ctrl・⌘ を押しながらなら足す）、Ctrl+Shift+C で選択を外す（Power BI の標準と同じキー）。
 * 押しっぱなしの繰り返しでは選び直さない（選んだり外したりを繰り返さないように）。
 * Esc は扱わない（止めもしない）。Power BI ではビジュアルから出るキー（Desktop では届かないことがある。前の Desktop の確認での観察で、手元のキーボードでは確かめていない）
 */
export interface SelectKeyOptions {
    /** 読み上げの名前 */
    label: string;
    /** 選ぶ。multi は Ctrl・⌘ を押しながら */
    onSelect: (multi: boolean) => void;
    /** Ctrl+Shift+C で選択を外す。無ければ何もしない */
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

/** キーの扱いだけ（役割と名前は呼ぶ側が付けるとき）。Enter・Space で選ぶ（Ctrl・⌘ で足す、押し続けでは選び直さない）、Ctrl+Shift+C で外す */
export function selectKeyDown(onSelect: (multi: boolean) => void, onClear?: () => void): (e: React.KeyboardEvent) => void {
    return (e) => {
        if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            e.stopPropagation();
            if (!e.repeat) onSelect(e.ctrlKey || e.metaKey);
        } else if (onClear && isClearKey(e.key, e.ctrlKey || e.metaKey, e.shiftKey, e.altKey, e.code)) {
            e.preventDefault();
            if (!e.repeat) onClear();
        }
    };
}

/**
 * 選択を外すキー（Ctrl+Shift+C。Mac は ⌘+Shift+C）。Power BI の標準のビジュアル・テーブルと同じ。
 * 字の配列によらず C の位置のキー（code が KeyC。キリル文字の С など）でも外す。Alt を押していれば外さない（Ctrl+Alt は AltGr にもなる）。
 * alt・code を渡さなければ、字（key）と Ctrl・Shift だけで見る。キーの処理から呼ぶときは e.altKey・e.code も渡す
 */
export function isClearKey(key: string, ctrl: boolean, shift: boolean, alt = false, code = ""): boolean {
    return ctrl && shift && !alt && (key.toLowerCase() === "c" || code === "KeyC");
}
