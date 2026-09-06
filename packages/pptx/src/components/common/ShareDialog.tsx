// 分享弹窗：调用宿主注入的分享实现生成链接，展示并支持复制
// 分享内容以点击"分享"时刻的文档快照为准
import { useEffect, useRef, useState, type JSX } from 'react';
import { Check, Copy, Loader2, RefreshCw, Share2, TriangleAlert } from 'lucide-react';
import { getPptxShareHandler } from '../../core/share/shareBridge';
import type { PptxShareResult } from '../../core/share/shareBridge';
import type { LoadedDoc } from '../../core/editor/persistence';

type ShareState =
  | { phase: 'loading' }
  | { phase: 'done' }
  | { phase: 'error'; message: string };

interface ShareDialogProps {
  open: boolean;
  /** 点击"分享"时刻的文档快照（由调用方捕获，弹窗打开期间不随编辑变化） */
  doc: LoadedDoc | null;
  onClose: () => void;
}

/** 复制到剪贴板：优先 Clipboard API，失败降级 execCommand */
async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    /* 继续尝试降级方案 */
  }
  const input = document.createElement('textarea');
  input.value = text;
  input.style.position = 'fixed';
  input.style.opacity = '0';
  document.body.appendChild(input);
  input.select();
  const ok = document.execCommand('copy');
  document.body.removeChild(input);
  return ok;
}

export function ShareDialog({ open, doc, onClose }: ShareDialogProps): JSX.Element | null {
  const [state, setState] = useState<ShareState>({ phase: 'loading' });
  const [copied, setCopied] = useState(false);
  // 分享结果（链接 + 宿主业务文案）与复制态
  const [result, setResult] = useState<PptxShareResult | null>(null);
  // 仅在 open/attempt 变化时重新发起分享（onClose 每次渲染都是新引用，经 ref 读取避免重跑）
  const [attempt, setAttempt] = useState(0);
  const docRef = useRef(doc);
  docRef.current = doc;
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;
  const copiedTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setState({ phase: 'loading' });
    setResult(null);
    setCopied(false);

    const handler = getPptxShareHandler();
    const target = docRef.current;
    if (!handler || !target) {
      setState({ phase: 'error', message: '当前环境不支持分享' });
      return;
    }
    void handler(target)
      .then((res) => {
        if (!cancelled) {
          setResult(res);
          setState({ phase: 'done' });
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) {
          setState({ phase: 'error', message: err instanceof Error ? err.message : '分享失败，请稍后重试' });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [open, attempt]);

  // Esc 关闭（捕获阶段拦截，避免触发编辑器快捷键）
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onCloseRef.current();
      }
    };
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [open]);

  useEffect(() => () => clearTimeout(copiedTimerRef.current), []);

  if (!open) return null;

  const handleCopy = async () => {
    if (state.phase !== 'done' || !result) return;
    if (await copyText(result.url)) {
      setCopied(true);
      clearTimeout(copiedTimerRef.current);
      copiedTimerRef.current = setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div
      className="fixed inset-0 z-[200] flex items-center justify-center bg-black/40 p-4"
      role="presentation"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onCloseRef.current();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="pptx-share-dialog-title"
        className="w-full max-w-md overflow-hidden rounded-xl border border-gray-200 bg-white shadow-xl"
      >
        <div className="flex items-center justify-between border-b border-gray-100 px-5 py-3">
          <div className="flex items-center gap-2">
            <Share2 size={16} className="text-[#d14424]" />
            <h2 id="pptx-share-dialog-title" className="text-base font-semibold text-gray-800">
              分享演示文稿
            </h2>
          </div>
        </div>

        <div className="px-5 py-4">
          {state.phase === 'loading' && (
            <div className="flex items-center justify-center gap-2 py-8 text-sm text-gray-500">
              <Loader2 size={16} className="animate-spin" />
              正在生成分享链接...
            </div>
          )}

          {state.phase === 'error' && (
            <div className="py-4 text-center">
              <TriangleAlert size={24} className="mx-auto mb-2 text-[#ff8800]" />
              <p className="mb-4 text-sm text-gray-700">{state.message}</p>
              <button
                type="button"
                onClick={() => setAttempt((n) => n + 1)}
                className="inline-flex items-center gap-1.5 rounded-md border border-gray-200 px-3 py-1.5 text-sm text-gray-700 hover:bg-gray-50"
              >
                <RefreshCw size={14} />
                重试
              </button>
            </div>
          )}

          {state.phase === 'done' && result && (
            <>
              <div className="flex gap-2">
                <input
                  type="text"
                  readOnly
                  value={result.url}
                  onFocus={(e) => e.currentTarget.select()}
                  aria-label="分享链接"
                  className="min-w-0 flex-1 rounded-md border border-gray-200 bg-gray-50 px-3 py-2 text-sm text-gray-700 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => void handleCopy()}
                  className={`inline-flex shrink-0 items-center gap-1.5 rounded-md px-3 py-2 text-sm font-medium transition-colors ${
                    copied ? 'bg-[#fbeae5] text-[#d14424]' : 'bg-[#d14424] text-white hover:bg-[#b93a1d]'
                  }`}
                >
                  {copied ? <Check size={14} /> : <Copy size={14} />}
                  {copied ? '已复制' : '复制'}
                </button>
              </div>
              {result.tips && result.tips.length > 0 && (
                <ul className="mt-4 space-y-1.5 rounded-lg bg-gray-50 px-3.5 py-3 text-xs leading-relaxed text-gray-500">
                  {result.tips.map((tip) => (
                    <li key={tip}>· {tip}</li>
                  ))}
                </ul>
              )}
            </>
          )}
        </div>

        <div className="flex justify-end border-t border-gray-100 px-5 py-3">
          <button
            type="button"
            onClick={onCloseRef.current}
            className="rounded-md bg-[#d14424] px-4 py-1.5 text-sm font-medium text-white hover:bg-[#b93a1d]"
          >
            完成
          </button>
        </div>
      </div>
    </div>
  );
}
