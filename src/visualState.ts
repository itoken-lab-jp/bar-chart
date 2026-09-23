/**
 * レポート閲覧者が触った UI 状態を Power BI 側に保存する。
 *
 * ページを移動するとビジュアルのインスタンスは破棄されるので、React の state に
 * 持っているだけでは戻ってきたときに初期化される。`host.persistProperties` で
 * レポートに書き戻し、`dataView.metadata.objects` から読み直す。
 *
 * 保存するのは累計の ON/OFF・押したときの書式の値・区切りの 3 つで、JSON を使わず text プロパティを素で持つ。
 */
import powerbi from "powerbi-visuals-api";

import DataView = powerbi.DataView;

/** metadata.objects に保存する内部状態オブジェクト名。書式ペインには出さない */
export const VISUAL_STATE_OBJECT = "visualState";

export const VISUAL_STATE_PROPERTIES = {
    cumulative: "cumulative",
    cumulativeBase: "cumulativeBase",
    cumulativeReset: "cumulativeReset",
    cumulativeResetBase: "cumulativeResetBase",
} as const;

export interface VisualState {
    /** 閲覧者が選んだ累計の ON/OFF。null = まだ触っていない */
    cumulative: boolean | null;
    /**
     * 閲覧者が押した時点の書式設定の値。
     * 作成者が書式を変えたらこれとずれるので、そのとき上書きを捨てる。
     * これが無いと、一度でも押されたレポートは書式ペインから累計を変えられなくなる。
     */
    cumulativeBase: boolean | null;
    /** グラフ上で選んだ累計の区切り。null = 未選択 */
    cumulativeReset: string | null;
    /**
     * 閲覧者が区切りを選んだ時点の、書式の区切り。作成者が書式の区切りを変えたらこれとずれるので、
     * そのとき閲覧者の区切りを捨てる（cumulativeBase と同じ考え方）
     */
    cumulativeResetBase: string | null;
}

export interface PersistedVisualState {
    /** 保存直後の古い update を見分けるためのキー */
    raw: string;
    state: VisualState;
}

export const EMPTY_VISUAL_STATE: VisualState = {
    cumulative: null,
    cumulativeBase: null,
    cumulativeReset: null,
    cumulativeResetBase: null,
};

export function serializeVisualState(state: VisualState): string {
    return JSON.stringify({
        cumulative: state.cumulative,
        cumulativeBase: state.cumulativeBase,
        cumulativeReset: state.cumulativeReset,
        cumulativeResetBase: state.cumulativeResetBase,
    });
}

/** persistProperties に渡す text プロパティ群。null は空文字で表す */
export function toPersistedProperties(state: VisualState): Record<string, string> {
    return {
        [VISUAL_STATE_PROPERTIES.cumulative]: boolToText(state.cumulative),
        [VISUAL_STATE_PROPERTIES.cumulativeBase]: boolToText(state.cumulativeBase),
        [VISUAL_STATE_PROPERTIES.cumulativeReset]: state.cumulativeReset ?? "",
        [VISUAL_STATE_PROPERTIES.cumulativeResetBase]: state.cumulativeResetBase ?? "",
    };
}

/** metadata.objects から保存済みの状態を読む。壊れていれば null を返して既定に戻す */
export function readVisualState(dataView: DataView | undefined): PersistedVisualState | null {
    const raw = dataView?.metadata?.objects?.[VISUAL_STATE_OBJECT];
    if (!raw || typeof raw !== "object") return null;

    const object = raw as unknown as Record<string, unknown>;
    const cumulative = textToBool(object[VISUAL_STATE_PROPERTIES.cumulative]);
    const cumulativeBase = textToBool(object[VISUAL_STATE_PROPERTIES.cumulativeBase]);
    const cumulativeReset = textToText(object[VISUAL_STATE_PROPERTIES.cumulativeReset]);
    const cumulativeResetBase = textToText(object[VISUAL_STATE_PROPERTIES.cumulativeResetBase]);
    if (cumulative === undefined || cumulativeBase === undefined || cumulativeReset === undefined || cumulativeResetBase === undefined) return null;

    const state: VisualState = { cumulative, cumulativeBase, cumulativeReset, cumulativeResetBase };
    return { raw: serializeVisualState(state), state };
}

/**
 * 実際に使う累計の ON/OFF を決める。
 *
 * 閲覧者の操作を優先するが、作成者が書式ペインで値を変えたら (= base とずれたら)
 * そちらを新しい既定として扱い、閲覧者の上書きは捨てる。
 */
export function effectiveCumulative(state: VisualState | null, setting: boolean): boolean {
    if (!state || state.cumulative === null) return setting;
    if (state.cumulativeBase !== setting) return setting;
    return state.cumulative;
}

/**
 * 作成者が書式を変えて古くなった閲覧者の操作を捨てた状態を返す。変わらなければ同じオブジェクトを返す。
 * 捨てた状態は保存し直す（残すと、作成者が書式を元に戻したときに閲覧者の操作が生き返る）
 */
export function withoutStale(state: VisualState, cumulativeSetting: boolean, resetSetting: string): VisualState {
    const staleCumulative = state.cumulative !== null && state.cumulativeBase !== cumulativeSetting;
    // 選んだときの書式の区切りが無ければ（区切りの基準を持たない形で保存された状態）比べようが無いので捨てない
    const staleReset = state.cumulativeReset !== null && state.cumulativeResetBase !== null && state.cumulativeResetBase !== resetSetting;
    if (!staleCumulative && !staleReset) return state;
    return {
        cumulative: staleCumulative ? null : state.cumulative,
        cumulativeBase: staleCumulative ? null : state.cumulativeBase,
        cumulativeReset: staleReset ? null : state.cumulativeReset,
        cumulativeResetBase: staleReset ? null : state.cumulativeResetBase,
    };
}

function boolToText(value: boolean | null): string {
    return value === null ? "" : value ? "true" : "false";
}

/** 空文字 = 未設定 (null)。想定外の値は undefined にして呼び出し側で捨てさせる */
function textToBool(value: unknown): boolean | null | undefined {
    if (value === undefined || value === "") return null;
    if (typeof value === "boolean") return value;
    if (value === "true") return true;
    if (value === "false") return false;
    return undefined;
}


function textToText(value: unknown): string | null | undefined {
    if (value === undefined || value === "") return null;
    return typeof value === "string" ? value : undefined;
}
