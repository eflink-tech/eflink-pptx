// 应用布局：顶栏 + 左缩略图 + 画布 + 右面板 + 底栏 + 模态/放映/Toast
import { useUIStore } from '../../store/uiStore'
import { useGlobalHotkeys } from '../../core/editor/hotkeys'
import { TopBar } from './TopBar'
import { BottomBar } from './BottomBar'
import { ThumbnailPanel } from '../thumbnails/ThumbnailPanel'
import { EditableCanvas } from '../canvas/EditableCanvas'
import { StylePanel } from '../panels/StylePanel'
import { AIChatPanel } from '../ai/AIChatPanel'
import { ModalHost } from '../dialogs/ModalHost'
import { Player } from '../player/Player'
import { ToastHost } from '../common/ToastHost'

export function AppLayout() {
  useGlobalHotkeys()
  const thumbnailsVisible = useUIStore((s) => s.thumbnailsVisible)
  const aiPanelVisible = useUIStore((s) => s.aiPanelVisible)
  const playerMode = useUIStore((s) => s.playerMode)

  return (
    <div className="flex h-full flex-col bg-white">
      <TopBar />
      <div className="flex min-h-0 flex-1">
        {thumbnailsVisible && <ThumbnailPanel />}
        <div className="min-w-0 flex-1">
          <EditableCanvas />
        </div>
        {aiPanelVisible && <AIChatPanel />}
        <StylePanel />
      </div>
      <BottomBar />
      <ModalHost />
      {playerMode !== 'off' && <Player />}
      <ToastHost />
    </div>
  )
}
