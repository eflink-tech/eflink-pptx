// 快捷键说明
import { useI18n } from '@eflink-tech/office-kit'
import { Modal } from './ModalHost'

interface HotkeyGroup {
  titleKey: string
  items: Array<[key: string, descKey: string]>
}

const GROUPS: HotkeyGroup[] = [
  {
    titleKey: 'pptx.hotkey.group.general',
    items: [
      ['Ctrl+S', 'pptx.hotkey.save'],
      ['Ctrl+Z / Ctrl+Shift+Z', 'pptx.hotkey.undoRedo'],
      ['Ctrl+C / Ctrl+X / Ctrl+V', 'pptx.hotkey.clipboard'],
      ['Ctrl+D', 'pptx.hotkey.duplicate'],
      ['Ctrl+A', 'pptx.hotkey.selectAll'],
      ['Delete / Backspace', 'pptx.hotkey.delete'],
      ['Esc', 'pptx.hotkey.esc'],
    ],
  },
  {
    titleKey: 'pptx.hotkey.group.edit',
    items: [
      ['pptx.hotkey.key.arrows', 'pptx.hotkey.nudge1'],
      ['Shift + pptx.hotkey.key.arrows', 'pptx.hotkey.nudge10'],
      ['pptx.hotkey.key.doubleClickText', 'pptx.hotkey.editText'],
      ['pptx.hotkey.key.ctrlWheel', 'pptx.hotkey.zoom'],
      ['pptx.hotkey.key.altDrag', 'pptx.hotkey.snap'],
      ['pptx.hotkey.key.shiftGrab', 'pptx.hotkey.ratio'],
      ['pptx.hotkey.key.shiftClick', 'pptx.hotkey.addSelect'],
    ],
  },
  {
    titleKey: 'pptx.hotkey.group.align',
    items: [
      ['Ctrl+G', 'pptx.hotkey.groupAction'],
      ['Ctrl+L / Ctrl+R', 'pptx.hotkey.alignLR'],
      ['Ctrl+T / Ctrl+B', 'pptx.hotkey.alignTB'],
    ],
  },
  {
    titleKey: 'pptx.hotkey.group.pages',
    items: [
      ['PageUp / PageDown', 'pptx.hotkey.prevNext'],
      ['pptx.hotkey.key.playArrows', 'pptx.hotkey.playNav'],
      ['pptx.hotkey.key.playBW', 'pptx.hotkey.playBW'],
      ['pptx.hotkey.key.playEsc', 'pptx.hotkey.playExit'],
      ['pptx.hotkey.key.playRightClick', 'pptx.hotkey.playPrev'],
    ],
  },
]

/**
 * 快捷键列可能整串是文案 key（pptx.hotkey.key.arrows）、纯组合键（Ctrl+Z）
 * 或两者混合（Shift + pptx.hotkey.key.arrows）。按 ' + ' / ' / ' 分段保留分隔符，
 * 仅翻译 pptx.* 段再原样拼回，'Ctrl+Z' 这类无空格写法不受影响。
 */
function renderKey(key: string, t: (key: string) => string): string {
  return key
    .split(/( \+ | \/ )/)
    .map((part) => (part.startsWith('pptx.') ? t(part) : part))
    .join('')
}

export function HotkeyDialog() {
  const { t } = useI18n()
  return (
    <Modal title={t('pptx.hotkey.title')} width={600}>
      <div className="grid grid-cols-2 gap-x-6 gap-y-4">
        {GROUPS.map((g) => (
          <div key={g.titleKey}>
            <div className="mb-1.5 text-xs font-semibold text-gray-700">{t(g.titleKey)}</div>
            {g.items.map(([key, descKey]) => (
              <div key={key} className="flex items-center justify-between border-b border-gray-50 py-1 text-xs">
                <span className="rounded bg-gray-100 px-1.5 py-0.5 font-mono text-[11px] text-gray-600">{renderKey(key, t)}</span>
                <span className="text-gray-500">{t(descKey)}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    </Modal>
  )
}
