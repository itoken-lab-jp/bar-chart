/**
 * ビジュアルの中の操作の部品（ボタン・選択・ダイアログ・メニューの行）の見た目。形と動きは Power BI の画面と同じ Fluent 2 に合わせ、
 * 色はレポートのテーマの文字と地の色から作る。Fluent 2 の灰（白の地と #242424 の文字のあいだの段）を、地から文字へ寄せる量で表すので、
 * 暗い地のテーマでも同じ濃さの段になる
 */
import { blend, luminanceOf } from "./color";

/** 角丸（Fluent 2 の borderRadiusMedium） */
export const CONTROL_RADIUS = "4px";

/** 色が変わる速さ（Fluent 2 の durationFaster と curveEasyEase） */
export const CONTROL_TRANSITION = "100ms cubic-bezier(0.33, 0, 0.67, 1)";

/** ダイアログ・メニューの影（Fluent 2 の shadow16） */
export const POPOVER_SHADOW = "0 0 2px rgba(0, 0, 0, 0.12), 0 8px 16px rgba(0, 0, 0, 0.14)";

/** 暗い地のダイアログ・メニューの影（Fluent 2 の暗いテーマの shadow16。明るい地より濃い） */
export const POPOVER_SHADOW_DARK = "0 0 2px rgba(0, 0, 0, 0.24), 0 8px 16px rgba(0, 0, 0, 0.28)";

/** グラフの上に浮くボタンの影（Fluent 2 の shadow4） */
export const FLOATING_SHADOW = "0 0 2px rgba(0, 0, 0, 0.12), 0 2px 4px rgba(0, 0, 0, 0.14)";

/**
 * 地から文字へ寄せる量。Fluent 2 の明るいテーマの値（地 #FFFFFF、文字 #242424）から割り出した
 * （例：枠 #D1D1D1 は (255 - 209) / (255 - 36) = 0.21）
 */
const STEPS = {
    hover: 0.045, // #F5F5F5 colorNeutralBackground1Hover
    selected: 0.091, // #EBEBEB colorNeutralBackground1Selected
    pressed: 0.142, // #E0E0E0 colorNeutralBackground1Pressed
    divider: 0.142, // #E0E0E0 colorNeutralStroke2
    stroke: 0.21, // #D1D1D1 colorNeutralStroke1
    strokeHover: 0.255, // #C7C7C7 colorNeutralStroke1Hover
    disabled: 0.3, // #BDBDBD colorNeutralForegroundDisabled
    strokeAccessible: 0.72, // #616161 colorNeutralStrokeAccessible
    subtle: 0.863, // #424242 colorNeutralForeground2
};

/**
 * 暗い地で、ダイアログの地をページの地から明るい側へ寄せる量。Fluent 2 の暗いテーマはページ（#1F1F1F）より
 * ダイアログ（#292929）を一段明るくして、影の見えない暗い地でも境目を見せる（(41 - 31) / (255 - 31) = 0.045）
 */
const DARK_SURFACE = 0.045;

export interface ControlColors {
    /** 文字（ボタン・選択・ダイアログの見出し） */
    text: string;
    /** 控えめな文字（枠なしのボタン、メニューの行、件数） */
    subtle: string;
    /** 地 */
    background: string;
    /** マウスを乗せた地・押した地・押してある（オン）の地 */
    hover: string;
    pressed: string;
    selected: string;
    /** 枠・マウスを乗せた枠・選択の下の線 */
    stroke: string;
    strokeHover: string;
    strokeAccessible: string;
    /** ダイアログの中の区切りの線 */
    divider: string;
    /** ダイアログ・メニューの地と、その中でマウスを乗せた地。暗い地ではページより一段明るい */
    surface: string;
    surfaceHover: string;
    /** ダイアログ・メニューの影 */
    popoverShadow: string;
    /** 押せない文字 */
    disabled: string;
    /** キーボードで移ったときの枠 */
    focus: string;
}

/** テーマの文字と地の色から、操作の部品の色を作る。ハイコントラストでは前景色と背景色の 2 色だけにする */
export function controlColorsOf(text: string, background: string, highContrast = false): ControlColors {
    if (highContrast) {
        return {
            text,
            subtle: text,
            background,
            hover: background,
            pressed: background,
            selected: background,
            stroke: text,
            strokeHover: text,
            strokeAccessible: text,
            divider: text,
            surface: background,
            surfaceHover: background,
            popoverShadow: "none",
            disabled: text,
            focus: text,
        };
    }
    const step = (amount: number) => blend(background, text, amount);
    // 地が字より暗ければ暗い地（読めない色は明るい地とみなす）
    const dark = (luminanceOf(background) ?? 1) < (luminanceOf(text) ?? 0);
    const surface = dark ? step(DARK_SURFACE) : background;
    return {
        text,
        subtle: step(STEPS.subtle),
        background,
        hover: step(STEPS.hover),
        pressed: step(STEPS.pressed),
        selected: step(STEPS.selected),
        stroke: step(STEPS.stroke),
        strokeHover: step(STEPS.strokeHover),
        strokeAccessible: step(STEPS.strokeAccessible),
        divider: step(STEPS.divider),
        surface,
        surfaceHover: blend(surface, text, STEPS.hover),
        popoverShadow: dark ? POPOVER_SHADOW_DARK : POPOVER_SHADOW,
        disabled: step(STEPS.disabled),
        focus: text,
    };
}

/** スタイルシートで受ける CSS 変数（--ctl-text など）。部品の入れ物の style に置く */
export function controlVars(colors: ControlColors): Record<string, string> {
    return {
        "--ctl-text": colors.text,
        "--ctl-subtle": colors.subtle,
        "--ctl-bg": colors.background,
        "--ctl-hover": colors.hover,
        "--ctl-pressed": colors.pressed,
        "--ctl-selected": colors.selected,
        "--ctl-stroke": colors.stroke,
        "--ctl-stroke-hover": colors.strokeHover,
        "--ctl-stroke-accessible": colors.strokeAccessible,
        "--ctl-divider": colors.divider,
        "--ctl-surface": colors.surface,
        "--ctl-surface-hover": colors.surfaceHover,
        "--ctl-disabled": colors.disabled,
        "--ctl-focus": colors.focus,
        "--ctl-radius": CONTROL_RADIUS,
        "--ctl-transition": CONTROL_TRANSITION,
        "--ctl-popover-shadow": colors.popoverShadow,
    };
}
