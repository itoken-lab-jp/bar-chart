import powerbi from "powerbi-visuals-api";

/** 凡例を描く位置（辺と寄せ）。保存値は下の LEGEND_POSITION_ITEMS（標準と同じ値）で、ここへ読み替える */
export const LEGEND_POSITIONS = {
    topLeft: "topLeft",
    topCenter: "topCenter",
    topRight: "topRight",
    bottomLeft: "bottomLeft",
    bottomCenter: "bottomCenter",
    bottomRight: "bottomRight",
    leftTop: "leftTop",
    leftCenter: "leftCenter",
    leftBottom: "leftBottom",
    rightTop: "rightTop",
    rightCenter: "rightCenter",
    rightBottom: "rightBottom",
} as const;
export type LegendPosition = (typeof LEGEND_POSITIONS)[keyof typeof LEGEND_POSITIONS];

/**
 * 凡例の位置の保存値は、標準のビジュアルとレポートテーマと同じ値（Top・Bottom など）にする。
 * 基本テーマ（Fluent 2 は "Bottom"）やカスタムテーマの値がそのまま届く。左下・右下は標準に無いので独自の値
 */
const LEGEND_POSITION_PLACEMENTS: Record<string, LegendPosition> = {
    Top: LEGEND_POSITIONS.topLeft,
    TopCenter: LEGEND_POSITIONS.topCenter,
    TopRight: LEGEND_POSITIONS.topRight,
    Bottom: LEGEND_POSITIONS.bottomLeft,
    BottomCenter: LEGEND_POSITIONS.bottomCenter,
    BottomRight: LEGEND_POSITIONS.bottomRight,
    Left: LEGEND_POSITIONS.leftTop,
    LeftCenter: LEGEND_POSITIONS.leftCenter,
    LeftBottom: LEGEND_POSITIONS.leftBottom,
    Right: LEGEND_POSITIONS.rightTop,
    RightCenter: LEGEND_POSITIONS.rightCenter,
    RightBottom: LEGEND_POSITIONS.rightBottom,
};

/** 12 通り。既定は上詰め (左) */
export const LEGEND_POSITION_ITEMS: powerbi.IEnumMember[] = [
    { value: "Top", displayName: "上詰め (左)" },
    { value: "TopCenter", displayName: "上詰め (中央)" },
    { value: "TopRight", displayName: "上詰め (右)" },
    { value: "Bottom", displayName: "下詰め (左)" },
    { value: "BottomCenter", displayName: "下詰め (中央)" },
    { value: "BottomRight", displayName: "下詰め (右)" },
    { value: "Left", displayName: "左上" },
    { value: "LeftCenter", displayName: "左中央" },
    { value: "LeftBottom", displayName: "左下" },
    { value: "Right", displayName: "右上" },
    { value: "RightCenter", displayName: "右中央" },
    { value: "RightBottom", displayName: "右下" },
];

/** 保存値を標準の値にする。前の版の保存値（topLeft など）も読み替える。知らない値は undefined */
export function standardLegendPosition(value: unknown): string | undefined {
    if (typeof value !== "string") return undefined;
    if (value in LEGEND_POSITION_PLACEMENTS) return value;
    return Object.keys(LEGEND_POSITION_PLACEMENTS).find((k) => LEGEND_POSITION_PLACEMENTS[k] === value);
}

/** 保存値を、描画で使う位置（topLeft など）にする。知らない値は上詰め (左) */
export function legendPlacementValue(value: unknown): LegendPosition {
    return LEGEND_POSITION_PLACEMENTS[standardLegendPosition(value) ?? "Top"];
}

export type LegendSide = "top" | "bottom" | "left" | "right";
export type LegendAlign = "start" | "center" | "end";

/** "topLeft" のような位置の値を、置く辺と寄せ方に分ける。知らない値は上詰め (左) */
export function legendPlacement(position: string): { side: LegendSide; align: LegendAlign } {
    const m = /^(top|bottom|left|right)(Left|Center|Right|Top|Bottom)$/.exec(position);
    if (!m) return { side: "top", align: "start" };
    const align: LegendAlign = m[2] === "Center" ? "center" : m[2] === "Left" || m[2] === "Top" ? "start" : "end";
    return { side: m[1] as LegendSide, align };
}
