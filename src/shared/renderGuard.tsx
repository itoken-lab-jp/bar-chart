import * as React from "react";
import { flushSync } from "react-dom";
import powerbi from "powerbi-visuals-api";

/** 描けなかったときの知らせ */
export const RENDER_FAILED_MESSAGE = "このビジュアルを描けませんでした";

/** 知らせの見た目。Power BI の案内に近い灰色の字（#605e5c・12px）で、真ん中に出す */
const NOTICE_STYLE: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    height: "100%",
    padding: 12,
    boxSizing: "border-box",
    color: "#605e5c",
    fontFamily: '"Segoe UI", wf_segoe-ui_normal, helvetica, arial, sans-serif',
    fontSize: 12,
    lineHeight: 1.6,
    textAlign: "center",
};

/** 描けなかったときに、前の絵の代わりに出す知らせ */
export function RenderFailedNotice(): React.ReactElement {
    return (
        <div className="render-failed" role="alert" style={NOTICE_STYLE}>
            <p style={{ margin: 0 }}>{RENDER_FAILED_MESSAGE}</p>
        </div>
    );
}

interface BoundaryProps {
    children?: React.ReactNode;
    /** 子の描画（描いたあとの効果を含む）で出た例外 */
    onError?: (error: unknown) => void;
}

interface BoundaryState {
    failed: boolean;
    /** 失敗したときの子。子が変わったら（次に描くとき）もう一度描く */
    children: React.ReactNode;
}

/**
 * 子の描画の例外を受ける境界。受けたら子を消して知らせを出し、onError に渡す。
 * React の描画の中の例外は、描かせた側の try に届かない（境界が無いと、何も描かれないまま残る）
 */
export class RenderBoundary extends React.Component<BoundaryProps, BoundaryState> {
    public state: BoundaryState = { failed: false, children: this.props.children };

    public static getDerivedStateFromProps(props: BoundaryProps, state: BoundaryState): Partial<BoundaryState> | null {
        return props.children === state.children ? null : { failed: false, children: props.children };
    }

    public static getDerivedStateFromError(): Partial<BoundaryState> {
        return { failed: true };
    }

    public componentDidCatch(error: unknown): void {
        this.props.onError?.(error);
    }

    public render(): React.ReactNode {
        return this.state.failed ? <RenderFailedNotice /> : this.props.children;
    }
}

/** 描く先。React の root の render だけを使う（試験では偽物を渡す） */
export interface RenderTarget {
    render(children: React.ReactNode): void;
}

/** update の成否を知らせる先（レンダリング イベント） */
export type RenderingEvents = Pick<powerbi.extensibility.IVisualEventService, "renderingStarted" | "renderingFinished" | "renderingFailed">;

/** renderingFailed に渡す理由。文字にできない値（プロトタイプの無いもの、toString が投げるもの）でも投げない */
function reasonOf(error: unknown): string {
    try {
        return String(error);
    } catch {
        return "update failed";
    }
}

/** update の中で、描いたあとの効果から描き直しを頼まれて続けて描く回数の上限。頼みが止まらなくても画面を固めない */
const MAX_SYNC_PASSES = 5;

/**
 * 描画を守る。
 * - update：renderingStarted を送り、update の中身を走らせ、描き終えてから renderingFinished を 1 回送る。
 *   中身の例外と、その間に境界が受けた描画の例外は失敗にして、前の絵を消して知らせを出し、renderingFailed を送る
 * - paint：境界で包んで描く。update の中では描き終えるまで待つ（flushSync）。描いたあとの効果から描き直しを頼まれたら
 *   （値の軸の高さが決まったなど）、今の描画を終えてから続けて描く。update の外（選択・見る人の操作）は今までどおり React に任せる
 */
export class GuardedRenderer {
    private failure: { error: unknown } | null = null;
    private updating = false;
    private painting = false;
    private next: React.ReactNode | null = null;

    constructor(
        private readonly target: RenderTarget,
        private readonly sync: (fn: () => void) => void = flushSync
    ) {}

    private readonly catchError = (error: unknown): void => {
        this.failure ??= { error };
    };

    public paint(element: React.ReactNode): void {
        const node = <RenderBoundary onError={this.catchError}>{element}</RenderBoundary>;
        if (!this.updating) {
            this.target.render(node);
            return;
        }
        this.next = node;
        if (this.painting) return;
        this.painting = true;
        try {
            for (let pass = 0; this.next !== null; pass++) {
                const current = this.next;
                this.next = null;
                if (pass >= MAX_SYNC_PASSES) {
                    this.target.render(current);
                    break;
                }
                this.sync(() => this.target.render(current));
            }
        } finally {
            this.painting = false;
            this.next = null;
        }
    }

    /** update の 1 回分。描けたら true */
    public update(events: RenderingEvents, options: powerbi.extensibility.visual.VisualUpdateOptions, work: () => void): boolean {
        events.renderingStarted(options);
        this.failure = null;
        this.updating = true;
        try {
            work();
        } catch (error) {
            this.failure ??= { error };
        } finally {
            this.updating = false;
        }
        const failure = this.failure;
        if (failure === null) {
            events.renderingFinished(options);
            return true;
        }
        console.error("update failed", failure.error);
        // 知らせを出すときに投げても（前の絵を外したときの後片付けが投げるなど）、renderingFailed は必ず送る。
        // 後片付けが投げると React は root の中を空にするので、もう一度だけ出し直す（空になった root には出せる）
        for (let attempt = 0; attempt < 2; attempt++) {
            try {
                this.showFailure();
                break;
            } catch (error) {
                console.error("could not show the failure notice", error);
            }
        }
        events.renderingFailed(options, reasonOf(failure.error));
        return false;
    }

    /** 前の絵を消して知らせを出す */
    public showFailure(): void {
        this.sync(() => this.target.render(<RenderFailedNotice />));
    }
}
