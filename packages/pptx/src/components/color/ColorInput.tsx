// 统一颜色输入：色块按钮 + 弹出式调色板（主题色 / 标准色 / 最近 / HSV 高级）
// 使用 React Portal 将弹出层渲染到 body，避免被 overflow 父容器裁剪
// 自动检测视口边界，防止被裁剪
import { createPortal } from 'react-dom'
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { ColorPicker } from './ColorPicker'

export interface ColorInputProps {
  value: string | undefined
  onChange: (v: string) => void
  allowEmpty?: boolean
  small?: boolean
  /** localStorage key，按使用场景区分最近颜色；默认 'pptx-color-recent' */
  recentStorageKey?: string
  /** 传入后调色板显示"默认"按钮，点击恢复到该颜色 */
  defaultColor?: string
}

const DEFAULT_RECENT_KEY = 'pptx-color-recent'

interface PopoverRect {
  top: number
  left: number
}

export function ColorInput({
  value,
  onChange,
  allowEmpty,
  small,
  recentStorageKey = DEFAULT_RECENT_KEY,
  defaultColor,
}: ColorInputProps) {
  const [open, setOpen] = useState(false)
  const btnRef = useRef<HTMLButtonElement>(null)
  const popoverRef = useRef<HTMLDivElement>(null)
  const [rect, setRect] = useState<PopoverRect | null>(null)

  // 计算按钮位置，并根据弹出层实际尺寸调整防止溢出视口
  const updateRect = useCallback(() => {
    const btn = btnRef.current
    const popover = popoverRef.current
    if (!btn) return

    const btnR = btn.getBoundingClientRect()
    const vw = window.innerWidth
    const vh = window.innerHeight

    // 弹出层宽度（已渲染时取实际值，否则估算 260px）
    const pw = popover?.offsetWidth ?? 260
    const ph = popover?.offsetHeight ?? 300

    // 水平：优先左对齐按钮；溢出右边界时向左偏移
    let left = btnR.left
    if (left + pw > vw - 8) left = vw - pw - 8
    if (left < 8) left = 8

    // 垂直：优先按钮下方；溢出下边界时放到按钮上方
    let top = btnR.bottom + 4
    if (top + ph > vh - 8) top = btnR.top - ph - 4
    if (top < 8) top = 8

    setRect({ top, left })
  }, [])

  useLayoutEffect(() => {
    if (!open) return
    // 第一帧用按钮位置，下一帧测量弹出层实际尺寸再微调
    const btn = btnRef.current
    if (btn) {
      const r = btn.getBoundingClientRect()
      setRect({ top: r.bottom + 4, left: r.left })
    }
    // 等 DOM 渲染后测量实际尺寸并调整
    const raf = requestAnimationFrame(() => updateRect())
    window.addEventListener('scroll', updateRect, true)
    window.addEventListener('resize', updateRect)
    return () => {
      cancelAnimationFrame(raf)
      window.removeEventListener('scroll', updateRect, true)
      window.removeEventListener('resize', updateRect)
    }
  }, [open, updateRect])

  // 点击外部关闭
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      const popover = document.getElementById('color-input-popover')
      if (popover?.contains(e.target as Node)) return
      if (btnRef.current?.contains(e.target as Node)) return
      setOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [open])

  // 当前显示颜色：有效值取 value，否则白色
  const display = value && value !== '#00000000' ? value : '#ffffff'

  return (
    <span className="inline-flex items-center gap-1">
      <button
        ref={btnRef}
        type="button"
        title={value || '选择颜色'}
        onClick={() => setOpen((v) => !v)}
        className={`${small ? 'size-5' : 'size-6'} shrink-0 cursor-pointer rounded border border-gray-200 p-0.5`}
        style={{ backgroundColor: display.slice(0, 7) }}
      />
      {allowEmpty && (
        <button
          type="button"
          className="rounded px-1 text-xs text-gray-400 hover:bg-gray-100"
          onClick={() => onChange('#00000000')}
          title="清除"
        >
          ✕
        </button>
      )}

      {open && rect && createPortal(
        <div
          ref={popoverRef}
          id="color-input-popover"
          className="fixed z-[9999]"
          style={{ top: rect.top, left: rect.left }}
        >
          <ColorPicker
            open
            onOpenChange={setOpen}
            value={value}
            recentStorageKey={recentStorageKey}
            onSelect={(color, options) => {
              onChange(color)
              if (options?.close !== false) setOpen(false)
            }}
            onDefault={() => {
              if (defaultColor) onChange(defaultColor)
              setOpen(false)
            }}
          />
        </div>,
        document.body,
      )}
    </span>
  )
}
