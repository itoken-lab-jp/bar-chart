/**
 * グラフを画像としてクリップボードに入れる。
 *
 * - 画像：要素の中の見えている SVG を、画面の位置のまま canvas に重ねて描き、PNG にする。スクロールする箱の中は、
 *   見えている範囲で切る。外を切らない（overflow: visible の）SVG は、箱の外へ出た中身も描く。スタイルシートで付けた見た目は、
 *   SVG を単独で画像にすると効かないので、計算済みのスタイルを写す。
 *   HTML で描いた部分は入らない（ボタン・お知らせなど）。文字だけの HTML は TEXT_ATTRIBUTE を付けると 1 行の文字として描く。
 *   SVG ごとに別の画像にするので、ほかの SVG に書いた定義（<defs> の模様など）は url(#id) で参照できない。フォントは
 *   端末に入っているものだけ（文書で読み込んだフォントは画像の SVG では使えない）。選択・ハイライト・フォーカスは見えているまま写る
 * - クリップボード：ビジュアルは sandbox の iframe の中で動き、Async Clipboard API（navigator.clipboard）は
 *   Permissions Policy で止まる。execCommand("copy") で起こした copy イベントに、画像を入れた HTML を書く。
 *   PowerPoint・Word などは HTML の画像を図として貼る。execCommand はクリックの操作の中でしか通らない。
 *   クリップボードに画像の形式（PNG・ビットマップ）で入れる道はビジュアルのコードからは無いので、画像しか受け取らない所へは、
 *   画像の要素（<img>）のドラッグ（画像のファイルとして渡る）と、ブラウザーの右クリックの「画像をコピー」で渡す
 */

export interface ChartImage {
    /** PNG の data URL */
    url: string;
    /** 画面での大きさ（CSS px）。画像の画素はこの scale 倍 */
    width: number;
    height: number;
}

export interface Rect {
    left: number;
    top: number;
    right: number;
    bottom: number;
}

/** この属性を持つ要素の中は画像に入れない（コピーのボタンなど） */
export const SKIP_ATTRIBUTE = "data-copy-image-skip";
/** この属性を持つ HTML の要素は、中の文字を 1 行の文字として描く（SVG の外に置いた軸のタイトルなど） */
export const TEXT_ATTRIBUTE = "data-copy-image-text";

/** 画素の倍率。貼った先で拡大してもにじまないように */
export const IMAGE_SCALE = 2;
const SVG_NS = "http://www.w3.org/2000/svg";

/** 画像に写すスタイル。SVG の見た目に効くものだけ */
const STYLE_PROPERTIES = [
    "fill",
    "fill-opacity",
    "fill-rule",
    "stroke",
    "stroke-width",
    "stroke-opacity",
    "stroke-dasharray",
    "stroke-dashoffset",
    "stroke-linecap",
    "stroke-linejoin",
    "opacity",
    "visibility",
    "display",
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "text-decoration",
    "text-anchor",
    "dominant-baseline",
    "alignment-baseline",
    "letter-spacing",
    "shape-rendering",
    "paint-order",
    "font-variant-numeric",
    "stroke-miterlimit",
    "stop-color",
    "stop-opacity",
    "marker-start",
    "marker-mid",
    "marker-end",
    "writing-mode",
    "vector-effect",
];

/** 2 つの矩形の重なり。重ならなければ null */
export function intersect(a: Rect, b: Rect): Rect | null {
    const r = { left: Math.max(a.left, b.left), top: Math.max(a.top, b.top), right: Math.min(a.right, b.right), bottom: Math.min(a.bottom, b.bottom) };
    return r.right > r.left && r.bottom > r.top ? r : null;
}

/** 貼る HTML。画素は scale 倍なので、幅・高さに画面での大きさを書く */
export function imageHtml(image: ChartImage, alt: string): string {
    const escaped = alt.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    return `<img src="${image.url}" width="${Math.round(image.width)}" height="${Math.round(image.height)}" alt="${escaped}">`;
}

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
});

function crc32(bytes: Uint8Array): number {
    let c = 0xffffffff;
    for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
}

/**
 * PNG に解像度（pHYs）を書く。画素は画面の IMAGE_SCALE 倍なので、96 dpi × 倍率にすると、貼った先が画面と同じ大きさで置く
 * （書かないと、画素の数のまま大きく貼る所がある）。IHDR の直後に入れる。PNG でなければそのまま返す
 */
export function withResolution(png: Uint8Array, dpi: number): Uint8Array {
    const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    const typeAt = (at: number) => String.fromCharCode(png[at], png[at + 1], png[at + 2], png[at + 3]);
    if (png.length < 33 || signature.some((b, i) => png[i] !== b) || typeAt(12) !== "IHDR") return png;
    // すでに解像度があれば足さない（PNG の決まりで 1 つまで）
    for (let at = 8; at + 8 <= png.length; ) {
        const length = new DataView(png.buffer, png.byteOffset + at, 4).getUint32(0);
        const type = typeAt(at + 4);
        if (type === "pHYs") return png;
        if (type === "IDAT" || type === "IEND") break;
        at += 12 + length;
    }
    const ihdrEnd = 8 + 8 + 13 + 4;
    const perMeter = Math.round(dpi / 0.0254);
    const chunk = new Uint8Array(4 + 4 + 9 + 4);
    const view = new DataView(chunk.buffer);
    view.setUint32(0, 9);
    chunk.set([0x70, 0x48, 0x59, 0x73], 4);
    view.setUint32(8, perMeter);
    view.setUint32(12, perMeter);
    chunk[16] = 1;
    view.setUint32(17, crc32(chunk.subarray(4, 17)));
    const out = new Uint8Array(png.length + chunk.length);
    out.set(png.subarray(0, ihdrEnd), 0);
    out.set(chunk, ihdrEnd);
    out.set(png.subarray(ihdrEnd), ihdrEnd + chunk.length);
    return out;
}

export function dataUrlWithResolution(url: string, dpi: number): string {
    const prefix = "data:image/png;base64,";
    if (!url.startsWith(prefix)) return url;
    const bin = atob(url.slice(prefix.length));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const out = withResolution(bytes, dpi);
    let text = "";
    for (let i = 0; i < out.length; i += 0x8000) text += String.fromCharCode(...out.subarray(i, i + 0x8000));
    return prefix + btoa(text);
}

/**
 * canvas を解像度つきの PNG の data URL にする。PNG にするのは toBlob（ブラウザーが描画の外で縮める）。toDataURL はその場で縮め、
 * 大きな表（7,000 × 6,500 画素ほど）では 1 秒以上ほかの操作が止まった。toBlob が無い・失敗したときは toDataURL
 */
export async function canvasPng(canvas: HTMLCanvasElement): Promise<string> {
    const blob = await new Promise<Blob | null>((resolve) => {
        try {
            canvas.toBlob(resolve, "image/png");
        } catch {
            resolve(null);
        }
    });
    if (!blob) return dataUrlWithResolution(canvas.toDataURL("image/png"), 96 * IMAGE_SCALE);
    // 解像度はバイトのまま足してから data URL にする（data URL を開いて足し直すと、大きな画像で文字の作り直しが重い）
    const png = withResolution(new Uint8Array(await blob.arrayBuffer()), 96 * IMAGE_SCALE);
    return new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = () => reject(reader.error);
        reader.readAsDataURL(new Blob([png.slice()], { type: "image/png" }));
    });
}

export const rectOf = (el: Element): Rect => {
    const r = el.getBoundingClientRect();
    return { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
};

/**
 * 箱の、スクロールバーを除いた四角（枠線の外の端まで）。右と下のスクロールバーの幅だけ縮める。はみ出しを切らない要素と、幅の分からない要素（SVG の要素など）は箱のまま。
 * 位置は拡大したあと、clientWidth などは拡大する前の値で来るので、箱の倍率を掛ける
 */
export function withoutScrollbars(el: Element): Rect {
    const box = rectOf(el);
    const html = el as Partial<HTMLElement>;
    if (!html.offsetWidth || !html.offsetHeight || html.clientWidth === undefined || html.clientHeight === undefined) return box;
    const scaleX = (box.right - box.left) / html.offsetWidth;
    const scaleY = (box.bottom - box.top) / html.offsetHeight;
    const style = getComputedStyle(el);
    // スクロールバーを持てるのは、はみ出しを切る箱だけ（インラインの要素は clientWidth が 0 なので、ほかの要素では幅を引かない）
    if (style.overflowX === "visible" && style.overflowY === "visible") return box;
    const width = (side: string) => Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0;
    const vertical = Math.max(0, html.offsetWidth - html.clientWidth - width("left") - width("right"));
    const horizontal = Math.max(0, html.offsetHeight - html.clientHeight - width("top") - width("bottom"));
    return { ...box, right: box.right - Math.round(vertical) * scaleX, bottom: box.bottom - Math.round(horizontal) * scaleY };
}

/**
 * はみ出しを切る箱が中身を切る四角：枠線の内側（padding box）から、スクロールバーを除いたもの。スクロールバーの下に隠れた中身
 * （バーの帯にかかる行・線、箱に切られる前提で外へ伸ばした地）を写さない。幅の分からない要素は箱のまま
 */
export function clipRectOf(el: Element, style?: Pick<CSSStyleDeclaration, "overflowX" | "overflowY"> & { overflowClipMargin?: string }): Rect {
    const box = rectOf(el);
    // overflow: clip に overflow-clip-margin: border-box なら、枠線の外の端で切る（枠線に重ねて描いた中身も見える。枠線の上に地の色の覆いを
    // 重ねて線の一部を消す表で、枠線の内側で切ると覆いが切られて線が残った）。余白の長さは使っていない
    if (style?.overflowX === "clip" && style.overflowY === "clip" && style.overflowClipMargin?.includes("border-box")) return box;
    const html = el as Partial<HTMLElement>;
    if (!html.offsetWidth || !html.offsetHeight || html.clientWidth === undefined || html.clientHeight === undefined) return box;
    const scaleX = (box.right - box.left) / html.offsetWidth;
    const scaleY = (box.bottom - box.top) / html.offsetHeight;
    const left = box.left + (html.clientLeft ?? 0) * scaleX;
    const top = box.top + (html.clientTop ?? 0) * scaleY;
    return { left, top, right: left + html.clientWidth * scaleX, bottom: top + html.clientHeight * scaleY };
}

/** 要素から root までの、はみ出しを切る箱（の中身を見せる所）と root で切った見えている範囲 */
export function visibleRect(el: Element, root: HTMLElement): Rect | null {
    let rect: Rect | null = rectOf(el);
    for (let node = el.parentElement; node && rect; node = node.parentElement) {
        if (node === root) {
            rect = intersect(rect, rectOf(node));
            break;
        }
        const style = getComputedStyle(node);
        if (style.overflowX !== "visible" || style.overflowY !== "visible") rect = intersect(rect, clipRectOf(node, style));
    }
    return rect;
}

/**
 * root の中の要素の見えている範囲（visibleRect）と、中身がはみ出して見える範囲（overflowRect）を、祖先ごとに覚えて求める。答えは 2 つの関数と同じ。
 * 要素ごとに root までたどると、要素の数 × 深さだけ計算スタイルと四角を読み、大きな表では画像を作るのに秒かかった
 */
export function clipTracker(root: HTMLElement): { visible: (el: Element) => Rect | null; overflow: (el: Element) => Rect | null } {
    const inner = new Map<Element, Rect | null>();
    // node の子孫が見える範囲：node から root までの、はみ出しを切る箱（root は箱そのもの）の重なり
    const within = (node: Element): Rect | null => {
        const cached = inner.get(node);
        if (cached !== undefined) return cached;
        let rect: Rect | null;
        if (node === root || !node.parentElement) rect = rectOf(node);
        else {
            rect = within(node.parentElement);
            if (rect) {
                const style = getComputedStyle(node);
                if (style.overflowX !== "visible" || style.overflowY !== "visible") rect = intersect(rect, clipRectOf(node, style));
            }
        }
        inner.set(node, rect);
        return rect;
    };
    return {
        visible: (el) => {
            if (el === root || !el.parentElement) return visibleRect(el, root);
            const clip = within(el.parentElement);
            return clip && intersect(rectOf(el), clip);
        },
        overflow: (el) => (el === root || !el.parentElement ? overflowRect(el, root) : within(el.parentElement)),
    };
}

/** 計算済みのスタイルを、写しの同じ位置の要素へ書く */
function inlineStyles(source: SVGSVGElement, copy: SVGSVGElement): void {
    const from = [source, ...Array.from(source.querySelectorAll("*"))];
    const to = [copy, ...Array.from(copy.querySelectorAll("*"))];
    from.forEach((el, i) => {
        const target = to[i] as SVGElement | undefined;
        if (!target || !(el instanceof SVGElement)) return;
        const computed = getComputedStyle(el);
        const text = STYLE_PROPERTIES.map((p) => `${p}:${computed.getPropertyValue(p)}`).join(";");
        target.setAttribute("style", text);
    });
}

/** 要素の中身がはみ出して見える範囲。要素から root までの、はみ出しを切る箱（と root）の重なり（要素の箱では切らない） */
export function overflowRect(el: Element, root: HTMLElement): Rect | null {
    let rect: Rect | null = rectOf(root);
    for (let node = el.parentElement; node && rect && node !== root; node = node.parentElement) {
        const style = getComputedStyle(node);
        if (style.overflowX !== "visible" || style.overflowY !== "visible") rect = intersect(rect, clipRectOf(node, style));
    }
    return rect;
}

/**
 * SVG を画面の位置のまま canvas に描く。ふつうは SVG の箱の大きさの画像にして、見えている範囲（visibleRect）で切る。
 * SVG の外を切らない（overflow: visible の）SVG は、箱の外へ出た中身も画面に見えるので、外の SVG（はみ出して見える範囲 overflowRect の大きさ）の
 * 中に写しを同じ位置に入れて画像にし、はみ出して見える範囲で切る
 */
export async function drawSvg(ctx: CanvasRenderingContext2D, svg: SVGSVGElement, root: HTMLElement, origin: Rect): Promise<void> {
    const box = rectOf(svg);
    const svgStyle = getComputedStyle(svg);
    const overflowing = svgStyle.overflowX === "visible" && svgStyle.overflowY === "visible";
    const visible = overflowing ? overflowRect(svg, root) : visibleRect(svg, root);
    const width = box.right - box.left;
    const height = box.bottom - box.top;
    if (!visible || width <= 0 || height <= 0) return;
    const copy = svg.cloneNode(true) as SVGSVGElement;
    inlineStyles(svg, copy);
    copy.setAttribute("xmlns", SVG_NS);
    copy.setAttribute("width", String(width));
    copy.setAttribute("height", String(height));
    copy.style.removeProperty("width");
    copy.style.removeProperty("height");
    copy.style.removeProperty("min-width");
    copy.style.removeProperty("position");
    // 画像の SVG と、それを置く canvas の上の四角
    let picture: SVGSVGElement = copy;
    let at: Rect = box;
    if (overflowing) {
        picture = document.createElementNS(SVG_NS, "svg");
        picture.setAttribute("xmlns", SVG_NS);
        picture.setAttribute("width", String(visible.right - visible.left));
        picture.setAttribute("height", String(visible.bottom - visible.top));
        copy.setAttribute("x", String(box.left - visible.left));
        copy.setAttribute("y", String(box.top - visible.top));
        copy.style.setProperty("overflow", "visible");
        picture.appendChild(copy);
        at = visible;
    }
    const image = new Image();
    image.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(picture));
    await image.decode();
    ctx.save();
    ctx.beginPath();
    ctx.rect(visible.left - origin.left, visible.top - origin.top, visible.right - visible.left, visible.bottom - visible.top);
    ctx.clip();
    ctx.drawImage(image, at.left - origin.left, at.top - origin.top, at.right - at.left, at.bottom - at.top);
    ctx.restore();
}

/** 印を付けた HTML の文字を、画面の位置・フォント・色で 1 行に描く */
function drawText(ctx: CanvasRenderingContext2D, el: HTMLElement, root: HTMLElement, origin: Rect): void {
    const text = el.textContent?.trim();
    const visible = visibleRect(el, root);
    if (!text || !visible) return;
    const box = rectOf(el);
    const style = getComputedStyle(el);
    if (style.visibility === "hidden" || style.display === "none") return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(visible.left - origin.left, visible.top - origin.top, visible.right - visible.left, visible.bottom - visible.top);
    ctx.clip();
    ctx.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    ctx.fillStyle = style.color;
    ctx.textBaseline = "middle";
    ctx.fillText(text, box.left - origin.left, (box.top + box.bottom) / 2 - origin.top);
    ctx.restore();
}

/** root の中の見えているグラフを画像にする。描くものが無ければ null */
export async function chartImage(root: HTMLElement, background = "#ffffff"): Promise<ChartImage | null> {
    const origin = rectOf(root);
    const width = origin.right - origin.left;
    const height = origin.bottom - origin.top;
    const svgs = Array.from(root.querySelectorAll("svg")).filter((svg) => !svg.parentElement?.closest("svg") && !svg.closest(`[${SKIP_ATTRIBUTE}]`));
    if (width <= 0 || height <= 0 || svgs.length === 0) return null;
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(width * IMAGE_SCALE);
    canvas.height = Math.round(height * IMAGE_SCALE);
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.scale(IMAGE_SCALE, IMAGE_SCALE);
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, width, height);
    // 後ろにある SVG から描く（文書の順が重なりの順）
    for (const svg of svgs) await drawSvg(ctx, svg, root, origin);
    // 文字だけの HTML は SVG より前に出ている（前面に重ねた軸のタイトルなど）
    for (const el of Array.from(root.querySelectorAll<HTMLElement>(`[${TEXT_ATTRIBUTE}]`))) {
        if (!el.closest(`[${SKIP_ATTRIBUTE}]`)) drawText(ctx, el, root, origin);
    }
    return { url: await canvasPng(canvas), width, height };
}

/**
 * 作っておいた画像。中身が変わったら invalidate で捨てる。作り始めたあとに捨てられたら、できた画像は覚えない
 * （作っている途中にスクロール・描き直しがあったとき、古い画像を書かない）
 */
export class PreparedImage {
    private generation = 0;
    private image: ChartImage | null = null;

    invalidate(): void {
        this.generation++;
        this.image = null;
    }

    /** 作り始める。返した番号を store に渡す */
    begin(): number {
        return this.generation;
    }

    /** 作り始めてから捨てられていなければ覚える。覚えたら true */
    store(started: number, image: ChartImage | null): boolean {
        if (started !== this.generation) return false;
        this.image = image;
        return true;
    }

    get current(): ChartImage | null {
        return this.image;
    }

    /** 作り始めた番号が今も新しいか */
    isCurrent(started: number): boolean {
        return started === this.generation;
    }
}

/** スクロールが止まったとみなすまでの時間 (ms)。止まってから画像を作り直す */
export const SCROLL_SETTLE_MS = 250;

/**
 * スクロールしたときの作り置きの扱い。見えている範囲が変わるので、作っておいた画像はスクロールのたびにその場で捨てる（軽い）。
 * 作り直し（見えている要素のスタイルを全部写す。重い）は、スクロールが止まってから 1 回だけにする
 */
export function scrollRefresher(invalidate: () => void, rebuild: () => void, delay = SCROLL_SETTLE_MS): { onScroll: () => void; dispose: () => void } {
    let timer: ReturnType<typeof setTimeout> | null = null;
    const cancel = () => {
        if (timer !== null) clearTimeout(timer);
        timer = null;
    };
    return {
        onScroll: () => {
            invalidate();
            cancel();
            timer = setTimeout(() => {
                timer = null;
                rebuild();
            }, delay);
        },
        dispose: cancel,
    };
}

/** powerbi.common.CustomVisualHostEnv のビット（const enum なので数で持つ） */
const HOST_ENV_WEB = 1 << 0;
const HOST_ENV_PUBLISH_TO_WEB = 1 << 1;
const HOST_ENV_DESKTOP = 1 << 2;
const HOST_ENV_MOBILE = 1 << 6;

/**
 * ブラウザーの右クリックのメニュー（「画像をコピー」）を出せる所か。ブラウザーで開くサービス（と Web に公開）だけ。
 * Desktop・モバイル・ほかのアプリへの埋め込み・分からないときは、ブラウザーのメニューが出ないことがあるので、Power BI のメニューに任せる
 * （外すと右クリックが何も出さなくなる）
 */
export function hasBrowserMenu(hostEnv: number | undefined): boolean {
    const env = hostEnv ?? 0;
    return (env & (HOST_ENV_WEB | HOST_ENV_PUBLISH_TO_WEB)) !== 0 && (env & (HOST_ENV_DESKTOP | HOST_ENV_MOBILE)) === 0;
}

/** 画像をクリップボードに書く。クリックの操作の中で呼ぶ。書けなければ false（操作の外、ブラウザーが止めた） */
export function writeImage(image: ChartImage, alt: string): boolean {
    const html = imageHtml(image, alt);
    let wrote = false;
    const onCopy = (e: ClipboardEvent) => {
        if (!e.clipboardData) return;
        e.clipboardData.setData("text/html", html);
        e.clipboardData.setData("text/plain", alt);
        e.preventDefault();
        wrote = true;
    };
    document.addEventListener("copy", onCopy, true);
    try {
        return document.execCommand("copy") && wrote;
    } catch {
        return false;
    } finally {
        document.removeEventListener("copy", onCopy, true);
    }
}
