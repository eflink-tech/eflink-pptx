// 通用模态框壳 + 模态分发
import { useUIStore } from '../../store/uiStore'
import { TemplateDialog } from './TemplateDialog'
import { ThemeDialog } from './ThemeDialog'
import { ExportDialog } from './ExportDialog'
import { ImportDialog } from './ImportDialog'
import { FindReplaceDialog } from './FindReplaceDialog'
import { HotkeyDialog } from './HotkeyDialog'
import { FileManagerDialog } from './FileManagerDialog'

export function Modal({ title, children, width = 520 }: { title: string; children: React.ReactNode; width?: number }) {
  const closeModal = useUIStore((s) => s.closeModal)
  return (
    <div className="fixed inset-0 z-[150] flex items-center justify-center bg-black/40" onMouseDown={(e) => { if (e.target === e.currentTarget) closeModal() }}>
      <div className="modal-pop max-h-[85vh] overflow-hidden rounded-xl bg-white shadow-2xl" style={{ width }}>
        <div className="flex items-center justify-between border-b border-gray-100 px-5 py-3">
          <span className="text-sm font-semibold text-gray-700">{title}</span>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100" onClick={closeModal}>✕</button>
        </div>
        <div className="max-h-[75vh] overflow-y-auto p-5">{children}</div>
      </div>
    </div>
  )
}

export function ModalHost() {
  const modal = useUIStore((s) => s.modal)
  if (!modal) return null
  switch (modal) {
    case 'template': return <TemplateDialog />
    case 'theme': return <ThemeDialog />
    case 'export': return <ExportDialog />
    case 'import': return <ImportDialog />
    case 'findReplace': return <FindReplaceDialog />
    case 'hotkey': return <HotkeyDialog />
    case 'fileManager': return <FileManagerDialog />
    default: return null
  }
}
