// 富文本浮动格式工具条：编辑元素时显示，通过 richTextController 操作当前编辑器
import { useEffect, useState } from 'react'
import {
  Bold, Italic, Underline, Strikethrough, Superscript, Subscript,
  AlignLeft, AlignCenter, AlignRight, AlignJustify, List, ListOrdered,
  RemoveFormatting, Link2, Code, Eraser,
} from 'lucide-react'
import { richTextController } from './RichTextEditor'
import { ColorInput } from '../color/ColorInput'

const FONT_FAMILIES = [
  { value: '', label: '字体' },
  { value: 'Microsoft YaHei', label: '微软雅黑' },
  { value: 'SimSun', label: '宋体' },
  { value: 'SimHei', label: '黑体' },
  { value: 'KaiTi', label: '楷体' },
  { value: 'Arial', label: 'Arial' },
  { value: 'Times New Roman', label: 'Times New Roman' },
  { value: 'Georgia', label: 'Georgia' },
  { value: 'Courier New', label: 'Courier New' },
]

const FONT_SIZES = [12, 14, 16, 18, 20, 24, 28, 32, 40, 48, 56, 72, 96]

function TB({ icon, label, onClick, active }: {
  icon: React.ReactNode
  label: string
  onClick: () => void
  active?: boolean
}) {
  return (
    <button
      title={label}
      data-active={active ? '1' : undefined}
      className={`flex h-7 w-7 items-center justify-center rounded ${active ? 'bg-[#fbeae5] text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon}
    </button>
  )
}

export function TextFormatToolbar() {
  const [state, setState] = useState({
    bold: false, italic: false, underline: false, strike: false,
    superscript: false, subscript: false, code: false,
    align: 'left' as string, fontFamily: '', fontSize: 0,
    canLink: false,
  })

  // 订阅选区变化刷新按钮态
  useEffect(() => {
    const refresh = () => {
      richTextController.exec((editor) => {
        setState({
          bold: editor.isActive('bold'),
          italic: editor.isActive('italic'),
          underline: editor.isActive('underline'),
          strike: editor.isActive('strike'),
          superscript: editor.isActive('superscript'),
          subscript: editor.isActive('subscript'),
          code: editor.isActive('code'),
          align: editor.getAttributes('paragraph').textAlign ?? 'left',
          fontFamily: editor.getAttributes('textStyle').fontFamily ?? '',
          fontSize: parseInt(editor.getAttributes('textStyle').fontSize ?? '0', 10) || 0,
          canLink: editor.isActive('link'),
        })
      })
    }
    document.addEventListener('selectionchange', refresh)
    const timer = window.setInterval(refresh, 500)
    return () => {
      document.removeEventListener('selectionchange', refresh)
      window.clearInterval(timer)
    }
  }, [])

  const exec = (fn: Parameters<typeof richTextController.exec>[0]) => richTextController.exec(fn)

  const setLink = () => {
    const url = window.prompt('链接地址（留空取消）')
    exec((editor) => {
      if (url) editor.chain().focus().setLink({ href: url }).run()
      else editor.chain().focus().unsetLink().run()
    })
  }

  return (
    <div
      className="pptx-text-toolbar absolute left-1/2 top-2 z-20 flex -translate-x-1/2 flex-wrap items-center gap-0.5 rounded-lg border border-gray-200 bg-white px-2 py-1 shadow-lg"
      onMouseDown={(e) => e.preventDefault()}
      data-testid="text-toolbar"
    >
      <select
        value={state.fontFamily}
        onChange={(e) => exec((ed) => (e.target.value ? ed.chain().focus().setFontFamily(e.target.value).run() : ed.chain().focus().unsetFontFamily().run()))}
        className="h-7 rounded border border-gray-200 px-1 text-xs"
      >
        {FONT_FAMILIES.map((f) => <option key={f.value} value={f.value}>{f.label}</option>)}
      </select>

      <select
        value={state.fontSize}
        onChange={(e) => exec((ed) => ed.chain().focus().setFontSize(`${e.target.value}px`).run())}
        className="h-7 w-16 rounded border border-gray-200 px-1 text-xs"
        title="字号（px）"
      >
        {!state.fontSize && <option value={0}>字号</option>}
        {FONT_SIZES.map((s) => <option key={s} value={s}>{s}</option>)}
      </select>

      <div className="mx-0.5 h-5 w-px bg-gray-200" />

      <TB icon={<Bold size={14} />} label="加粗" active={state.bold} onClick={() => exec((ed) => ed.chain().focus().toggleBold().run())} />
      <TB icon={<Italic size={14} />} label="斜体" active={state.italic} onClick={() => exec((ed) => ed.chain().focus().toggleItalic().run())} />
      <TB icon={<Underline size={14} />} label="下划线" active={state.underline} onClick={() => exec((ed) => ed.chain().focus().toggleUnderline().run())} />
      <TB icon={<Strikethrough size={14} />} label="删除线" active={state.strike} onClick={() => exec((ed) => ed.chain().focus().toggleStrike().run())} />
      <TB icon={<Code size={14} />} label="代码" active={state.code} onClick={() => exec((ed) => ed.chain().focus().toggleCode().run())} />
      <TB icon={<Superscript size={14} />} label="上标" active={state.superscript} onClick={() => exec((ed) => ed.chain().focus().toggleSuperscript().run())} />
      <TB icon={<Subscript size={14} />} label="下标" active={state.subscript} onClick={() => exec((ed) => ed.chain().focus().toggleSubscript().run())} />

      <span className="mx-0.5 flex items-center gap-1">
        <ColorInput
          value={undefined}
          onChange={(v) => exec((ed) => ed.chain().focus().setColor(v).run())}
          small
          recentStorageKey="pptx-text-recent"
        />
        <ColorInput
          value={undefined}
          onChange={(v) => exec((ed) => ed.chain().focus().toggleHighlight({ color: v }).run())}
          small
          recentStorageKey="pptx-highlight-recent"
        />
      </span>

      <div className="mx-0.5 h-5 w-px bg-gray-200" />

      <TB icon={<AlignLeft size={14} />} label="左对齐" active={state.align === 'left'} onClick={() => exec((ed) => ed.chain().focus().setTextAlign('left').run())} />
      <TB icon={<AlignCenter size={14} />} label="居中" active={state.align === 'center'} onClick={() => exec((ed) => ed.chain().focus().setTextAlign('center').run())} />
      <TB icon={<AlignRight size={14} />} label="右对齐" active={state.align === 'right'} onClick={() => exec((ed) => ed.chain().focus().setTextAlign('right').run())} />
      <TB icon={<AlignJustify size={14} />} label="两端对齐" active={state.align === 'justify'} onClick={() => exec((ed) => ed.chain().focus().setTextAlign('justify').run())} />

      <div className="mx-0.5 h-5 w-px bg-gray-200" />

      <TB icon={<List size={14} />} label="无序列表" onClick={() => exec((ed) => ed.chain().focus().toggleBulletList().run())} />
      <TB icon={<ListOrdered size={14} />} label="有序列表" onClick={() => exec((ed) => ed.chain().focus().toggleOrderedList().run())} />
      <TB icon={<Link2 size={14} />} label="链接" active={state.canLink} onClick={setLink} />
      <TB icon={<Eraser size={14} />} label="清除颜色" onClick={() => exec((ed) => ed.chain().focus().unsetColor().run())} />
      <TB icon={<RemoveFormatting size={14} />} label="清除格式" onClick={() => exec((ed) => ed.chain().focus().unsetAllMarks().run())} />
    </div>
  )
}
