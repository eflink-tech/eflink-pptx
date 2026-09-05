// 快捷键说明
import { Modal } from './ModalHost'

const GROUPS: Array<{ title: string; items: Array<[string, string]> }> = [
  {
    title: '通用',
    items: [
      ['Ctrl+S', '保存文档'],
      ['Ctrl+Z / Ctrl+Shift+Z', '撤销 / 重做'],
      ['Ctrl+C / Ctrl+X / Ctrl+V', '复制 / 剪切 / 粘贴'],
      ['Ctrl+D', '原地复制'],
      ['Ctrl+A', '全选元素'],
      ['Delete / Backspace', '删除选中元素'],
      ['Esc', '取消选择 / 关闭弹窗'],
    ],
  },
  {
    title: '编辑',
    items: [
      ['方向键', '微移 1px'],
      ['Shift + 方向键', '微移 10px'],
      ['双击文本', '进入富文本编辑'],
      ['Ctrl+滚轮', '缩放画布'],
      ['Alt + 拖拽', '临时关闭吸附'],
      ['Shift + 缩放柄', '等比缩放'],
      ['Shift + 点击', '追加选中 / 取消'],
    ],
  },
  {
    title: '对齐 / 组合',
    items: [
      ['Ctrl+G', '组合 / 取消组合'],
      ['Ctrl+L / Ctrl+R', '左对齐 / 右对齐'],
      ['Ctrl+T / Ctrl+B', '顶对齐 / 底对齐'],
    ],
  },
  {
    title: '页面与放映',
    items: [
      ['PageUp / PageDown', '上一页 / 下一页'],
      ['放映中：← → / 空格', '翻页'],
      ['放映中：B / W', '黑屏 / 白屏'],
      ['放映中：Esc', '退出放映'],
      ['放映中：右键', '上一页'],
    ],
  },
]

export function HotkeyDialog() {
  return (
    <Modal title="快捷键" width={600}>
      <div className="grid grid-cols-2 gap-x-6 gap-y-4">
        {GROUPS.map((g) => (
          <div key={g.title}>
            <div className="mb-1.5 text-xs font-semibold text-gray-700">{g.title}</div>
            {g.items.map(([key, desc]) => (
              <div key={key} className="flex items-center justify-between border-b border-gray-50 py-1 text-xs">
                <span className="rounded bg-gray-100 px-1.5 py-0.5 font-mono text-[11px] text-gray-600">{key}</span>
                <span className="text-gray-500">{desc}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    </Modal>
  )
}
