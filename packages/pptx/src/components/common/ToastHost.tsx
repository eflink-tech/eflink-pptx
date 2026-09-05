// 轻量消息提示
import { useToastStore } from '../../store/uiStore'

export function ToastHost() {
  const toasts = useToastStore((s) => s.toasts)
  if (!toasts.length) return null
  return (
    <div className="pointer-events-none fixed left-1/2 top-14 z-[300] flex -translate-x-1/2 flex-col items-center gap-2">
      {toasts.map((t) => (
        <div
          key={t.id}
          className={`pointer-events-auto rounded-md px-4 py-2 text-sm text-white shadow-lg ${
            t.kind === 'success' ? 'bg-green-600' : t.kind === 'error' ? 'bg-red-500' : 'bg-gray-800'
          }`}
        >
          {t.text}
        </div>
      ))}
    </div>
  )
}
