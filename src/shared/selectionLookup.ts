/**
 * 選んだ ID（ISelectionId）の中にあるかを引く関数を作る。選んだ ID を 1 つずつ equals で比べると、描くたびに
 * 「点の数 × 選んだ数」になる（「その他」や区切りを押すと、選んだ ID は行の数だけある）。getKey() で同じキーの ID を引き、
 * その中だけを equals で比べる。答えは selected.some((s) => s.equals(id)) と同じ（equals が true なら getKey() も同じ、とみなす）
 */

/** getKey と equals を持つ ID（ISelectionId） */
export interface KeyedId<T> {
    getKey(): string;
    equals(other: T): boolean;
}

export function selectionLookup<T extends KeyedId<T>>(selected: readonly T[]): (id: T | null | undefined) => boolean {
    const byKey = new Map<string, T[]>();
    for (const s of selected) {
        const key = s.getKey();
        const bucket = byKey.get(key);
        if (bucket) bucket.push(s);
        else byKey.set(key, [s]);
    }
    return (id) => {
        if (!id) return false;
        const bucket = byKey.get(id.getKey());
        return !!bucket && bucket.some((s) => s.equals(id));
    };
}
