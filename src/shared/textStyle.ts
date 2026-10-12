import type * as React from "react";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

/** 文字の書式（書式ペインのフォントの部品の値）。size は pt */
export interface FontValues {
    family: string;
    size: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
}

/** 文字の大きさを書式ペインの範囲（8〜32pt）に丸める。値が無ければ fallback */
export function clampFontSize(size: number | null | undefined, fallback: number): number {
    return Math.max(8, Math.min(32, size ?? fallback));
}

/** フォントの部品（FontControl）の値。保存が無ければ family・size。大きさは 8〜32pt に丸める */
export function fontValuesOf(control: formattingSettings.FontControl, family: string, size: number): FontValues {
    return {
        family: control.fontFamily.value ?? family,
        size: clampFontSize(control.fontSize.value, size),
        bold: control.bold?.value ?? false,
        italic: control.italic?.value ?? false,
        underline: control.underline?.value ?? false,
    };
}

/**
 * 設定の中の文字の書式を、名前の前に付く語（prefix）で取り出す。prefix が空なら fontFamily・fontSize・bold・italic・underline、
 * "title" なら titleFontFamily・titleFontSize・titleBold・titleItalic・titleUnderline
 */
export function fontOf(settings: object, prefix = ""): FontValues {
    const s = settings as Record<string, unknown>;
    const key = (name: string) => (prefix ? `${prefix}${name[0].toUpperCase()}${name.slice(1)}` : name);
    return {
        family: String(s[key("fontFamily")] ?? ""),
        size: Number(s[key("fontSize")]) || 0,
        bold: !!s[key("bold")],
        italic: !!s[key("italic")],
        underline: !!s[key("underline")],
    };
}

/** SVG の文字のスタイル（大きさ・フォント・太字・斜体・下線）。色は呼ぶ側が足す */
export function fontStyleOf(font: Pick<FontValues, "family" | "size"> & Partial<FontValues>): React.CSSProperties {
    return {
        fontSize: `${font.size}pt`,
        fontFamily: font.family,
        fontWeight: font.bold ? "bold" : "normal",
        fontStyle: font.italic ? "italic" : "normal",
        textDecoration: font.underline ? "underline" : undefined,
    };
}

/** フォントの部品の値を、設定の項目の名前（fontFamily・fontSize・bold・italic・underline）で返す */
export function fontFields(control: formattingSettings.FontControl, family: string, size: number) {
    const f = fontValuesOf(control, family, size);
    return { fontFamily: f.family, fontSize: f.size, bold: f.bold, italic: f.italic, underline: f.underline };
}

/** タイトルのフォントの部品の値を、titleFontFamily・titleFontSize・titleBold・titleItalic・titleUnderline の名前で返す */
export function titleFontFields(control: formattingSettings.FontControl, family: string, size: number) {
    const f = fontValuesOf(control, family, size);
    return { titleFontFamily: f.family, titleFontSize: f.size, titleBold: f.bold, titleItalic: f.italic, titleUnderline: f.underline };
}
