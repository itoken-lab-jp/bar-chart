/**
 * 「画像としてコピー」のボタン。グラフの入れ物（position を持つ要素）の直下に置く。
 *
 * - マウスを乗せたとき・キーボードで来たときだけ、右下に小さく出す。入れ物の中のスクロールバーが右下にあれば、その内側に避ける
 * - 押すと、入れ物の中の見えているグラフを画像にしてクリップボードに入れ、「コピーしました」を短く出す
 * - クリップボードに書けるのはクリックの操作の中だけなので、画像は入れ物にマウスが入ったとき（とボタンに来たとき）に作っておき、
 *   押したらその場で書く。中身が変わったとき（stamp が変わったとき）とスクロールしたときに捨てる。作っておけなかったときは
 *   押してから作って書き、ブラウザーが止めたら作った画像を覚えて「もう一度押してください」と出す（次に押せばその場で書く）
 * - 押して入るのは画像入りの HTML（Office は図として貼る）。画像しか受け取らない所（チャットなど）へは、ボタンをドラッグすると
 *   画像のファイルとして持ち出せる。ブラウザーのメニューが出る所（Power BI サービス）では、ボタンの右クリックで「画像をコピー」も使える。
 *   どちらも、作っておいた画像の <img> をボタンに見えないように重ねて的にする
 */
import * as React from "react";

import { ChartImage, IMAGE_SCALE, PreparedImage, SKIP_ATTRIBUTE, chartImage, writeImage } from "./copyImage";

export interface CopyImageButtonProps {
    /** 貼った先の代わりの文字（画像の alt と、文字しか受けない所へ貼ったときの文字） */
    alt: string;
    /**
     * 中身の目印。どれかが変わったら、作っておいた画像を捨てる（データ・大きさ・選択など、描いたものが変わる値を並べる）。
     * 並べる数はいつも同じにする
     */
    stamp: readonly unknown[];
    /** 画像の地の色（グラフの背景） */
    background?: string;
    /** ボタンの色（ハイコントラストのときに渡す） */
    colors?: { foreground: string; background: string };
    /** ボタンの右クリックでブラウザーのメニュー（「画像をコピー」）を出すか。false なら入れ物に任せる（Power BI のメニュー） */
    browserMenu?: boolean;
}

type Status = "idle" | "done" | "again" | "failed";

const MESSAGES: Record<Status, string> = {
    idle: "",
    done: "コピーしました",
    again: "もう一度押してください",
    failed: "コピーできませんでした",
};

const HOST_CLASS = "copy-image-host";
const BUTTON_CLASS = "copy-image-button";
const LIVE_CLASS = "copy-image-live";
/** 入れ物の hover で出す。スタイルはこの部品の中で閉じる（ビジュアルのスタイルシートに頼らない） */
const CSS = `
.${BUTTON_CLASS}{position:absolute;z-index:10;display:flex;align-items:center;gap:4px;height:24px;padding:0 6px;
border:1px solid var(--copy-image-border);border-radius:2px;background:var(--copy-image-bg);color:var(--copy-image-fg);
font:12px "Segoe UI",sans-serif;cursor:pointer;opacity:0;pointer-events:none;transition:opacity .15s}
.${HOST_CLASS}:hover>.${BUTTON_CLASS},.${BUTTON_CLASS}:focus-visible,.${BUTTON_CLASS}[data-status]{opacity:1;pointer-events:auto}
.${BUTTON_CLASS}:hover{filter:brightness(.96)}
.${BUTTON_CLASS}:focus-visible{outline:2px solid var(--copy-image-fg);outline-offset:1px}
.${BUTTON_CLASS} svg{display:block}
.${BUTTON_CLASS} img{position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:inherit}
.${LIVE_CLASS}{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
`;
const DONE_MS = 1500;
/** ドラッグの影の幅（px） */
const DRAG_PREVIEW_WIDTH = 240;
/** DownloadURL に入れる data URL の上限（ブラウザーが URL として扱う長さ。超えたら画像のドラッグだけに任せる） */
const MAX_DOWNLOAD_URL = 2_000_000;

/** ファイル名に使えない記号を除く（DownloadURL は「種類:名前:URL」の形なので : も除く） */
const fileNameOf = (alt: string) => (alt.replace(/[\\/:*?"<>|]/g, "_").trim() || "chart") + ".png";
const EDGE = 4;

/** 入れ物の右下にかかるスクロールバーの幅（右下に置くものを、その内側へずらす量） */
function scrollbarInsets(root: HTMLElement): { right: number; bottom: number } {
    const outer = root.getBoundingClientRect();
    let right = 0;
    let bottom = 0;
    for (const el of Array.from(root.querySelectorAll<HTMLElement>("div"))) {
        const horizontal = el.offsetHeight - el.clientHeight;
        const vertical = el.offsetWidth - el.clientWidth;
        if (horizontal <= 0 && vertical <= 0) continue;
        const style = getComputedStyle(el);
        const box = el.getBoundingClientRect();
        if (horizontal > 0 && /auto|scroll/.test(style.overflowX) && Math.abs(box.bottom - outer.bottom) < 2) bottom = Math.max(bottom, horizontal);
        if (vertical > 0 && /auto|scroll/.test(style.overflowY) && Math.abs(box.right - outer.right) < 2) right = Math.max(right, vertical);
    }
    return { right, bottom };
}

export function CopyImageButton({ alt, stamp, background = "#ffffff", colors, browserMenu = false }: CopyImageButtonProps): React.JSX.Element {
    const ref = React.useRef<HTMLButtonElement>(null);
    const cache = React.useRef(new PreparedImage());
    const pending = React.useRef<{ started: number; promise: Promise<ChartImage | null> } | null>(null);
    const [status, setStatus] = React.useState<Status>("idle");
    const [insets, setInsets] = React.useState({ right: 0, bottom: 0 });
    /** ドラッグ・右クリックの的に重ねる画像（作っておいた画像） */
    const [imageUrl, setImageUrl] = React.useState<string | null>(null);

    const invalidate = React.useCallback(() => {
        cache.current.invalidate();
        setImageUrl(null);
    }, []);

    /** 作っておいた画像を返す。無ければ作る（同じ中身を作っている途中なら、それを待つ） */
    const prepare = React.useCallback((): Promise<ChartImage | null> => {
        const root = ref.current?.parentElement;
        if (!root) return Promise.resolve(null);
        const ready = cache.current.current;
        if (ready) return Promise.resolve(ready);
        const started = cache.current.begin();
        if (pending.current && pending.current.started === started) return pending.current.promise;
        const promise = chartImage(root, background).then(
            (image) => {
                if (cache.current.store(started, image)) setImageUrl(image?.url ?? null);
                // 作れなかったら、次に頼まれたときに作り直す
                if (!image && pending.current?.promise === promise) pending.current = null;
                return image;
            },
            (err: unknown) => {
                if (pending.current?.promise === promise) pending.current = null;
                throw err;
            }
        );
        pending.current = { started, promise };
        return promise;
    }, [background]);

    /** 失敗しても何もしない（押したときに作り直す） */
    const prepareQuietly = React.useCallback((): void => {
        prepare().catch((): null => null);
    }, [prepare]);

    /** 捨てたあと、マウスが乗ったままなら作り直しておく（ドラッグ・右クリックの的が消えたままにならないように） */
    const refresh = React.useCallback((): void => {
        invalidate();
        const root = ref.current?.parentElement;
        if (root?.matches(":hover")) prepareQuietly();
    }, [invalidate, prepareQuietly]);

    // 中身が変わったら、作っておいた画像は古い（ボタン自身の表示の描き直しでは捨てない）。描いた直後に捨てる
    React.useLayoutEffect(() => refresh(), [...stamp, background]);

    React.useEffect(() => {
        const root = ref.current?.parentElement;
        if (!root) return undefined;
        root.classList.add(HOST_CLASS);
        // 中のスクロール（捕まえて拾う）で見えている範囲が変わる
        const onScroll = () => refresh();
        root.addEventListener("pointerenter", prepareQuietly);
        root.addEventListener("scroll", onScroll, true);
        return () => {
            root.removeEventListener("pointerenter", prepareQuietly);
            root.removeEventListener("scroll", onScroll, true);
            root.classList.remove(HOST_CLASS);
        };
    }, [prepareQuietly, refresh]);

    // スクロールバーは描き直しで出たり消えたりするので、描き直すたびに測る（同じなら描き直さない）
    React.useLayoutEffect(() => {
        const root = ref.current?.parentElement;
        if (!root) return;
        const next = scrollbarInsets(root);
        setInsets((prev) => (prev.right === next.right && prev.bottom === next.bottom ? prev : next));
    });

    React.useEffect(() => {
        if (status === "idle") return undefined;
        const timer = window.setTimeout(() => setStatus("idle"), DONE_MS);
        return () => window.clearTimeout(timer);
    }, [status]);

    const onClick = (e: React.MouseEvent) => {
        // 背景のクリック（選択の解除）にしない
        e.stopPropagation();
        const ready = cache.current.current;
        if (ready) {
            setStatus(writeImage(ready, alt) ? "done" : "failed");
            return;
        }
        prepare()
            .then((image) => {
                if (!image) return setStatus("failed");
                // ブラウザーが止めたら、覚えた画像で次に押したときにその場で書く
                setStatus(writeImage(image, alt) ? "done" : "again");
            })
            .catch(() => setStatus("failed"));
    };

    /** ドラッグ：画像のファイルとして渡す（DownloadURL も付ける）。影はグラフを縮めた画像 */
    const onDragStart = (e: React.DragEvent<HTMLImageElement>) => {
        e.stopPropagation();
        const url = imageUrl;
        if (!url || !e.dataTransfer) return;
        e.dataTransfer.effectAllowed = "copy";
        if (url.length <= MAX_DOWNLOAD_URL) e.dataTransfer.setData("DownloadURL", `image/png:${fileNameOf(alt)}:${url}`);
        const preview = document.createElement("img");
        preview.src = url;
        const natural = e.currentTarget.naturalWidth / IMAGE_SCALE || DRAG_PREVIEW_WIDTH;
        const width = Math.min(DRAG_PREVIEW_WIDTH, natural);
        preview.width = width;
        preview.style.cssText = "position:fixed;left:-10000px;top:0";
        document.body.appendChild(preview);
        // カーソルを影の真ん中に（左上だと、落とす先を影で隠しやすい）
        e.dataTransfer.setDragImage(preview, width / 2, (width * e.currentTarget.naturalHeight) / (e.currentTarget.naturalWidth || 1) / 2);
        window.setTimeout(() => preview.remove(), 0);
    };

    const message = MESSAGES[status];
    const fg = colors?.foreground ?? "#252423";
    const style = {
        right: EDGE + insets.right,
        bottom: EDGE + insets.bottom,
        "--copy-image-fg": fg,
        "--copy-image-bg": colors?.background ?? "rgba(255,255,255,.92)",
        "--copy-image-border": colors ? fg : "#c8c6c4",
    } as React.CSSProperties;
    return (
        <>
            <style>{CSS}</style>
            <button
                ref={ref}
                type="button"
                className={BUTTON_CLASS}
                style={style}
                data-status={status === "idle" ? undefined : status}
                {...{ [SKIP_ATTRIBUTE]: "" }}
                aria-label="画像としてコピー"
                title={
                    "画像としてコピー（PowerPoint などに貼れます）。ドラッグすると画像のファイルとして持ち出せます" +
                    (browserMenu ? "。右クリックの「画像をコピー」も使えます" : "")
                }
                onClick={onClick}
                onFocus={prepareQuietly}
                onPointerEnter={prepareQuietly}
                onPointerDown={(e) => e.stopPropagation()}
            >
                <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
                    <rect x="5" y="1.5" width="9" height="10" rx="1" fill="none" stroke="currentColor" strokeWidth="1.3" />
                    <path d="M3.5 4.5H3a1 1 0 0 0-1 1V14a1 1 0 0 0 1 1h7a1 1 0 0 0 1-1v-.5" fill="none" stroke="currentColor" strokeWidth="1.3" />
                </svg>
                {message && <span aria-hidden="true">{message}</span>}
                {imageUrl && (
                    <img
                        src={imageUrl}
                        alt=""
                        aria-hidden="true"
                        draggable
                        onDragStart={onDragStart}
                        // ブラウザーのメニューを出す所では、入れ物（Power BI のメニューを出して既定を止める）に渡さない
                        onContextMenu={browserMenu ? (e) => e.stopPropagation() : undefined}
                    />
                )}
            </button>
            {/* 読み上げ：いつも置いておき、中の文字を変える（押したあとに足した要素は読まれないことが多い） */}
            <span className={LIVE_CLASS} role="status" aria-live="polite" {...{ [SKIP_ATTRIBUTE]: "" }}>
                {message}
            </span>
        </>
    );
}
