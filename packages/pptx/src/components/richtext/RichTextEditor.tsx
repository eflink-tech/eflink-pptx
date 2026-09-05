// TipTap 富文本编辑器（画布内嵌编辑）
// 通过 richTextController 让格式工具条操作当前激活的编辑器实例
import { useEffect, useRef } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import Document from '@tiptap/extension-document'
import Paragraph from '@tiptap/extension-paragraph'
import Text from '@tiptap/extension-text'
import Bold from '@tiptap/extension-bold'
import Italic from '@tiptap/extension-italic'
import Underline from '@tiptap/extension-underline'
import Strike from '@tiptap/extension-strike'
import Superscript from '@tiptap/extension-superscript'
import Subscript from '@tiptap/extension-subscript'
import { TextStyle, FontFamily, FontSize } from '@tiptap/extension-text-style'
import Color from '@tiptap/extension-color'
import Highlight from '@tiptap/extension-highlight'
import TextAlign from '@tiptap/extension-text-align'
import BulletList from '@tiptap/extension-bullet-list'
import OrderedList from '@tiptap/extension-ordered-list'
import ListItem from '@tiptap/extension-list-item'
import Blockquote from '@tiptap/extension-blockquote'
import Code from '@tiptap/extension-code'
import Link from '@tiptap/extension-link'
import HardBreak from '@tiptap/extension-hard-break'
import Dropcursor from '@tiptap/extension-dropcursor'
import Gapcursor from '@tiptap/extension-gapcursor'

/** 当前激活编辑器注册器（供外部工具条调用命令） */
class RichTextController {
  private editor: Editor | null = null

  register(editor: Editor | null): void {
    this.editor = editor
  }

  get active(): Editor | null {
    return this.editor
  }

  /** 有编辑器激活时执行命令；返回是否已消费 */
  exec(fn: (editor: Editor) => void): boolean {
    if (!this.editor) return false
    fn(this.editor)
    return true
  }
}

export const richTextController = new RichTextController()

interface Props {
  elementId: string
  initialHTML: string
  onChange: (html: string) => void
  placeholder?: string
}

export function RichTextEditor({ elementId, initialHTML, onChange }: Props) {
  const onChangeRef = useRef(onChange)
  onChangeRef.current = onChange

  const editor = useEditor({
    extensions: [
      Document, Paragraph, Text, Bold, Italic, Underline, Strike,
      Superscript, Subscript, Code, HardBreak, Dropcursor, Gapcursor,
      TextStyle, FontFamily, FontSize, Color, Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ['paragraph'] }),
      BulletList, OrderedList, ListItem, Blockquote,
      Link.configure({ openOnClick: false, autolink: true }),
    ],
    content: initialHTML,
    autofocus: 'end',
    editorProps: {
      attributes: {
        class: 'pptx-prosemirror',
      },
      // 中文等 IME 组字确认时 Enter 不应拆出新段落，否则文字会“往下跳”
      handleKeyDown: (view, event) => {
        if (event.key === 'Enter' && view.composing) return true
        return false
      },
    },
    onUpdate: ({ editor: ed }) => onChangeRef.current(ed.getHTML()),
  }, [elementId])

  useEffect(() => {
    richTextController.register(editor)
    return () => richTextController.register(null)
  }, [editor, elementId])

  return (
    <EditorContent
      editor={editor}
      className="pptx-tiptap size-full outline-none"
    />
  )
}
