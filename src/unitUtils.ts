import { textMeasurementService } from "powerbi-visuals-utils-formattingutils";

/**
 * 日本語ビジネスレポート向け単位ユーティリティ
 * 0〜12の桁数（指数）ベースでスケーリングと表記を解決する
 */

export interface UnitDefinition {
    exponent: number;
    divisor: number;
    unitWord: string;   // 解決されたスケーリング語
    unitWordJa: string; // 日本語 (千, 万, 億, 兆)
    unitWordStd: string;// 標準 (K, M, bn, T)。1000 の冪以外は語が無いので空
    displayName: string;
}

// 標準表記の語は K/M/bn/T だけ。「10K」「100M」のような語を数字の後ろに付けると
// 1.23 + 10K = 「1.2310K」と読み違えるため、1000 の冪以外は空にして toStandardUnitKey で読み替える
export const UNIT_DEFINITIONS: Record<string, UnitDefinition> = {
    "0": { exponent: 0, divisor: 1, unitWord: "", unitWordJa: "", unitWordStd: "", displayName: "なし" },
    "1": { exponent: 1, divisor: 10, unitWord: "十", unitWordJa: "十", unitWordStd: "", displayName: "十" },
    "2": { exponent: 2, divisor: 100, unitWord: "百", unitWordJa: "百", unitWordStd: "", displayName: "百" },
    "3": { exponent: 3, divisor: 1_000, unitWord: "千", unitWordJa: "千", unitWordStd: "K", displayName: "千" },
    "4": { exponent: 4, divisor: 10_000, unitWord: "万", unitWordJa: "万", unitWordStd: "", displayName: "万" },
    "5": { exponent: 5, divisor: 100_000, unitWord: "十万", unitWordJa: "十万", unitWordStd: "", displayName: "十万" },
    "6": { exponent: 6, divisor: 1_000_000, unitWord: "百万", unitWordJa: "百万", unitWordStd: "M", displayName: "百万" },
    "7": { exponent: 7, divisor: 10_000_000, unitWord: "千万", unitWordJa: "千万", unitWordStd: "", displayName: "千万" },
    "8": { exponent: 8, divisor: 100_000_000, unitWord: "億", unitWordJa: "億", unitWordStd: "", displayName: "億" },
    "9": { exponent: 9, divisor: 1_000_000_000, unitWord: "十億", unitWordJa: "十億", unitWordStd: "bn", displayName: "十億" },
    "10": { exponent: 10, divisor: 10_000_000_000, unitWord: "百億", unitWordJa: "百億", unitWordStd: "", displayName: "百億" },
    "11": { exponent: 11, divisor: 100_000_000_000, unitWord: "千億", unitWordJa: "千億", unitWordStd: "", displayName: "千億" },
    "12": { exponent: 12, divisor: 1_000_000_000_000, unitWord: "兆", unitWordJa: "兆", unitWordStd: "T", displayName: "兆" },
};

/** 自動単位で使う段（小さい順）。日本語は 千・万・百万・億・兆、標準は K・M・bn・T */
const AUTO_KEYS_JA = ["0", "3", "4", "6", "8", "12"];
const AUTO_KEYS_STD = ["0", "3", "6", "9", "12"];

/** precision が "auto" のときの小数桁の上限（formatValue と同じ） */
const AUTO_FRACTION_DIGITS = 2;

/** 表示に使う小数桁で丸めた後の |value / divisor| */
function roundedScaledAbs(abs: number, divisor: number, precision: string): number {
    const digits = precision === "auto" ? AUTO_FRACTION_DIGITS : Math.max(0, parseInt(precision, 10) || 0);
    const factor = Math.pow(10, digits);
    return Math.round((abs / divisor) * factor) / factor;
}

/**
 * 小さい順の divisor 列から abs に合う段の添字を返す。
 * 丸めた結果が次の段に届く（9,999.9 → 「10千」、99,999,999 → 「10,000万」）ときは次の段へ繰り上げる。
 */
function pickUnitIndex(abs: number, divisors: number[], precision: string): number {
    let index = 0;
    for (let i = divisors.length - 1; i >= 0; i--) {
        if (abs >= divisors[i]) {
            index = i;
            break;
        }
    }
    while (
        index + 1 < divisors.length &&
        roundedScaledAbs(abs, divisors[index], precision) >= divisors[index + 1] / divisors[index]
    ) {
        index++;
    }
    return index;
}

/**
 * 最大絶対値から自動単位（指数）を判定。丸め後の桁上がりも考慮する
 */
export function resolveAutoUnitKey(
    maxAbsValue: number,
    notation: string = "japanese",
    precision: string = "auto"
): string {
    const keys = notation === "standard" ? AUTO_KEYS_STD : AUTO_KEYS_JA;
    const divisors = keys.map((k) => UNIT_DEFINITIONS[k].divisor);
    return keys[pickUnitIndex(Math.abs(maxAbsValue), divisors, precision)];
}

/**
 * 標準表記で使えるキーに読み替える。K/M/bn/T の語が無い桁（1,2,4,5,7,8,10,11）は
 * 1 つ下の 1000 の冪に落とす（例: 4 = 1万 → 3 = K で「12.31K」）。
 * 数値を大きく割ったのに語が付かず 10 倍・100 倍に読まれる事故を防ぐ。
 */
export function toStandardUnitKey(key: string): string {
    const def = UNIT_DEFINITIONS[key];
    if (!def) return "0";
    return String(Math.floor(def.exponent / 3) * 3);
}

/**
 * 選択されたキー（auto含む）から実際の単位定義を解決
 */
export function resolveUnit(
    unitTypeKey: string,
    maxAbsValue: number,
    notation: string = "japanese",
    precision: string = "auto"
): UnitDefinition {
    const isStd = notation === "standard";
    let key = unitTypeKey === "auto" ? resolveAutoUnitKey(maxAbsValue, notation, precision) : unitTypeKey;
    if (!UNIT_DEFINITIONS[key]) key = "0";
    if (isStd) key = toStandardUnitKey(key);
    const base = UNIT_DEFINITIONS[key];
    return {
        ...base,
        unitWord: isStd ? base.unitWordStd : base.unitWordJa,
    };
}

/**
 * 数値をスケールしフォーマットする
 */
export function formatValue(value: number, divisor: number, precision: string): string {
    const scaled = value / divisor;
    if (precision === "auto") {
        return scaled.toLocaleString("ja-JP", {
            maximumFractionDigits: 2,
            minimumFractionDigits: 0,
        });
    }
    const digits = Math.max(0, parseInt(precision, 10) || 0);
    return scaled.toLocaleString("ja-JP", {
        minimumFractionDigits: digits,
        maximumFractionDigits: digits,
    });
}

/** 値ごとに単位を選ぶときの段（小さい順） */
const DYNAMIC_LADDER_JA = [
    { divisor: 1, word: "" },
    { divisor: 1e3, word: "千" },
    { divisor: 1e4, word: "万" },
    { divisor: 1e8, word: "億" },
    { divisor: 1e12, word: "兆" },
];
const DYNAMIC_LADDER_STD = [
    { divisor: 1, word: "" },
    { divisor: 1e3, word: "K" },
    { divisor: 1e6, word: "M" },
    { divisor: 1e9, word: "bn" },
    { divisor: 1e12, word: "T" },
];

/**
 * 対数スケールなど桁が大きく跨がる場合に、値ごとの桁に応じた動的スケーリング・単位付与を行う。
 * 単位は丸めた後の値で決め直す（99,999,999 → 「1億」、9,999.9 → 「1万」、標準 999,999 → 「1M」）
 */
export function formatDynamicValue(
    value: number,
    notation: string = "japanese",
    precision: string = "auto",
    withUnit: boolean = true
): string {
    const abs = Math.abs(value);
    if (abs === 0) return "0";

    if (!withUnit) {
        return formatValue(value, 1, precision);
    }

    const ladder = notation === "standard" ? DYNAMIC_LADDER_STD : DYNAMIC_LADDER_JA;
    const rung = ladder[pickUnitIndex(abs, ladder.map((r) => r.divisor), precision)];
    const num = formatValue(value, rung.divisor, precision);
    return rung.word ? `${num}${rung.word}` : num;
}

/**
 * スケーリング語とユーザー入力単位を合成する
 * @param unitWord 指数倍率に対応する語（例: "百万", "億"）
 * @param unitText ユーザー入力の単位（例: "円", "人", "件"）
 * @param includeDisplayUnit スケーリング語を含めるかどうか（既定: true）
 */
export function composeUnitText(
    unitWord: string,
    unitText: string,
    includeDisplayUnit: boolean = true
): string {
    const word = includeDisplayUnit ? (unitWord || "") : "";
    const custom = (unitText || "").trim();
    return `${word}${custom}`;
}

/**
 * 合成された単位文字列からバッジ用表記を組み立てる
 * @param composed 合成単位（例: "百万円", "人"）
 * @param style 表記スタイル ("parentheses" | "withPrefix")
 */
export function formatUnitBadge(
    composed: string,
    style: string = "parentheses"
): string {
    const trimmed = (composed || "").trim();
    if (!trimmed) return "";
    if (style === "withPrefix") {
        return `(単位: ${trimmed})`;
    }
    return `(${trimmed})`;
}

/**
 * 軸上バッジのテキストを解決する
 */
export function resolveBadgeText(spec: {
    unitShow: boolean;
    unitPosition: string; // "valueAxisTop" | "plotTopRight" | "none"
    unitIncludeDisplayUnit: boolean;
    unitStyle: string;
    unitWord: string;
    unitText: string;
}): string {
    if (!spec.unitShow || spec.unitPosition === "none") {
        return "";
    }
    const composed = composeUnitText(spec.unitWord, spec.unitText, spec.unitIncludeDisplayUnit);
    return formatUnitBadge(composed, spec.unitStyle);
}

/**
 * 背景色（16進カラーコード）に対する最適な文字色（白または黒）を返す（WCAG 2.1 相対輝度基準）。
 */
export function contrastingText(backgroundColor: string): string {
    const hex = backgroundColor.replace("#", "");
    const full = hex.length === 3 ? hex.split("").map((c) => c + c).join("") : hex;
    const value = Number.parseInt(full, 16);
    if (!Number.isFinite(value)) return "#252423";

    const channel = (shift: number) => {
        const c = ((value >> shift) & 0xff) / 255;
        return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    };
    // WCAG の相対輝度 (sRGB)
    const luminance = 0.2126 * channel(16) + 0.7152 * channel(8) + 0.0722 * channel(0);
    return luminance > 0.4 ? "#252423" : "#FFFFFF";
}

// --- データラベルの配置 --------------------------------------------------------

export interface FontSpec {
    family: string;
    size: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
}

export interface PlacedLabel {
    fits: boolean;
    fitsVertically: boolean;
    fitsAlongBar: boolean;
    outside: boolean;
    y: number; // 基準点 anchorY
    width: number;
    height: number;
    box: { x: number; y: number; width: number; height: number };
}

/** フォントサイズに対するベースラインより上/下の割合。背景の高さを出すのに使う */
const ASCENT_RATIO = 0.8;
const DESCENT_RATIO = 0.22;
/** 背景と文字のあいだの余白 */
export const LABEL_PADDING = 3;

/**
 * テキスト幅の計測
 */
export function measureTextWidth(text: string, font: FontSpec): number {
    try {
        if (typeof document !== "undefined") {
            const width = textMeasurementService.measureSvgTextWidth({
                text,
                fontFamily: font.family,
                fontSize: `${font.size}px`,
                fontWeight: font.bold ? "bold" : "normal",
                fontStyle: font.italic ? "italic" : "normal",
            });
            if (width > 0) return width;
        }
    } catch {
        // フォールバック
    }
    // 文字種別（全角/半角）を考慮した簡易メトリクスフォールバック
    let w = 0;
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        w += (code >= 0x20 && code <= 0x7e) ? font.size * 0.6 : font.size * 1.0;
    }
    return w;
}

/**
 * データラベルの配置（位置・向き・入りきらないときの外側への逃がし）を決める
 */
export function placeLabel(spec: {
    position: string; // "auto" | "insideTop" | "insideCenter" | "insideBottom" | "outsideEnd"
    top: number;
    bottom: number;
    barWidth: number;
    value: string;
    font: FontSpec;
    vertical: boolean;
    overflow: boolean;
}): PlacedLabel {
    const {
        position,
        top,
        bottom,
        barWidth,
        value,
        font,
        vertical,
        overflow,
    } = spec;

    const valueWidth = measureTextWidth(value, font);
    const lineY = font.size * 0.35;

    const left = -valueWidth / 2;
    const right = valueWidth / 2;
    const topEdge = lineY - font.size * ASCENT_RATIO;
    const bottomEdge = lineY + font.size * DESCENT_RATIO;

    const box = {
        x: left - LABEL_PADDING,
        y: topEdge - LABEL_PADDING,
        width: right - left + LABEL_PADDING * 2,
        height: bottomEdge - topEdge + LABEL_PADDING * 2,
    };

    const neededY = (vertical ? box.width : box.height) + 2;
    const neededX = (vertical ? box.height : box.width) + 2;
    const fitsAlongBar = barWidth >= neededX;
    const barHeight = Math.max(0, bottom - top);
    const fitsVertically = barHeight >= neededY;
    const fits = fitsVertically && fitsAlongBar;

    const requestedOutside = position === "outsideEnd";
    const outside = requestedOutside || (!fits && overflow && fitsAlongBar);

    let anchorY: number;
    if (outside) {
        anchorY = top - LABEL_PADDING - bottomEdge;
    } else if (position === "insideTop" || position === "auto") {
        anchorY = top + neededY / 2;
    } else if (position === "insideBottom") {
        anchorY = bottom - neededY / 2;
    } else {
        anchorY = (top + bottom) / 2;
    }

    return {
        fits,
        fitsVertically,
        fitsAlongBar,
        outside,
        y: anchorY,
        width: box.width,
        height: box.height,
        box,
    };
}

