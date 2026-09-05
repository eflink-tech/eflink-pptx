// MathLive 公式编辑器：将 <math-field> web component 封装为 React 受控组件
import { useEffect, useRef } from 'react'
import type { MathfieldElement } from 'mathlive'
import 'mathlive'
import 'mathlive/static.css'

// React 19 的 JSX 命名空间不含自定义元素，需补充 <math-field> 声明。
// 此处只能引用全局命名空间（React.*）与内联结构类型：命名导入的类型在
// 模块增强的 d.ts 产出中会被视作私有名（TS4033）
declare module 'react' {
  namespace JSX {
    interface IntrinsicElements {
      'math-field': React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement> & {
        ref?: { current: HTMLElement | null }
      }
    }
  }
}

interface MathFieldEditorProps {
  /** LaTeX 源码（受控） */
  value: string
  /** 值变化回调 */
  onChange: (latex: string) => void
  className?: string
}

export function MathFieldEditor({ value, onChange, className }: MathFieldEditorProps) {
  const ref = useRef<MathfieldElement>(null)

  // 初始化：禁用虚拟键盘 + 设置 smartFence
  useEffect(() => {
    const el = ref.current
    if (!el) return
    // 新版 mathlive 类型中已无 virtualKeyboardMode 属性，走 attribute 设置
    el.setAttribute('virtual-keyboard-mode', 'off')
    el.smartFence = true
  }, [])

  // 监听 input 事件 → 实时同步 LaTeX 值
  useEffect(() => {
    const el = ref.current
    if (!el) return
    const handler = () => onChange(el.value)
    el.addEventListener('input', handler)
    return () => el.removeEventListener('input', handler)
  }, [onChange])

  return <math-field ref={ref} className={className}>{value}</math-field>
}
