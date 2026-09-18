# PPTX 顶栏主菜单整理实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 引入 ☰ 主菜单（级联二级，对齐 Word/Excel 交互），收拢低频入口，并彻底删除「文档管理」功能。

**Architecture:** 新增 `MainMenu.tsx` 组件（参照 eflink-word 的 `ToolbarMenu.tsx` 裁剪为二级级联），`TopBar.tsx` 精简为品牌区 + 主菜单 + 高频图标（撤销/重做/插入/AI/放映）。分享逻辑随「分享」菜单项从 TopBar 迁入 MainMenu。

**Tech Stack:** React 19 + Zustand + lucide-react + Tailwind；vitest + @testing-library/react（jsdom）。

**Spec:** `docs/superpowers/specs/2026-09-19-pptx-topbar-mainmenu-design.md`

**工作目录:** 所有命令在 `eflink-pptx/packages/pptx/` 下执行（除非注明仓库根）。

---

### Task 1: 删除「文档管理」功能（4 处）

**Files:**
- Modify: `src/store/uiStore.ts:6`（ModalName 类型）
- Modify: `src/components/dialogs/ModalHost.tsx`（import + case）
- Modify: `src/components/layout/TopBar.tsx:82`（入口按钮）
- Delete: `src/components/dialogs/FileManagerDialog.tsx`

- [ ] **Step 1: 确认 FileManagerDialog 无其他引用**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx/packages/pptx
grep -rn "FileManagerDialog\|fileManager" src --include="*.ts*"
```

Expected: 仅 4 处命中（uiStore.ts 类型、ModalHost.tsx、TopBar.tsx、FileManagerDialog.tsx 本体）。若出现其他命中，停止并上报。

- [ ] **Step 2: 从 ModalName 类型中移除 `'fileManager'`**

`src/store/uiStore.ts` 第 4-6 行改为：

```ts
export type ModalName =
  | 'template' | 'theme' | 'export' | 'import' | 'findReplace'
  | 'hotkey' | 'aiSettings' | 'aiHistory' | 'about' | null
```

- [ ] **Step 3: 移除 ModalHost 的 import 与 case**

`src/components/dialogs/ModalHost.tsx`：删除第 9 行 `import { FileManagerDialog } from './FileManagerDialog'` 和第 36 行 `case 'fileManager': return <FileManagerDialog />`。

- [ ] **Step 4: 移除 TopBar 入口按钮**

`src/components/layout/TopBar.tsx`：删除第 82 行

```tsx
<ToolButton icon={<FolderOpen size={17} />} label="文档管理" onClick={() => ui.openModal('fileManager')} />
```

- [ ] **Step 5: 删除组件文件**

```bash
rm src/components/dialogs/FileManagerDialog.tsx
```

- [ ] **Step 6: 验证无残留 + 类型检查通过**

```bash
grep -rn "FileManagerDialog\|fileManager" src --include="*.ts*"
# Expected: 无输出
npx tsc --noEmit -p tsconfig.json
# Expected: 无错误
```

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "refactor: 移除 pptx 文档管理功能（入口/弹窗/类型）"
```

---

### Task 2: 新建 MainMenu 主菜单组件（TDD）

**Files:**
- Test: `src/components/menus/MainMenu.test.tsx`
- Create: `src/components/menus/MainMenu.tsx`

- [ ] **Step 1: 写失败测试**

创建 `src/components/menus/MainMenu.test.tsx`：

```tsx
// 主菜单组件测试：分类展开 / 二级菜单 / 叶子分类触发
import { describe, expect, it } from 'vitest'
import { fireEvent, render, screen } from '@testing-library/react'
import { MainMenu } from './MainMenu'
import { useUIStore } from '../../store/uiStore'

describe('MainMenu', () => {
  it('点击后展开五个一级分类', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    expect(screen.getByText('文件')).toBeTruthy()
    expect(screen.getByText('设计')).toBeTruthy()
    expect(screen.getByText('视图')).toBeTruthy()
    expect(screen.getByText('放映')).toBeTruthy()
    expect(screen.getByText('快捷键')).toBeTruthy()
  })

  it('悬停文件分类展开二级菜单（新建文档/保存）', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.mouseEnter(screen.getByText('文件'))
    expect(screen.getByText('新建文档')).toBeTruthy()
    expect(screen.getByText('保存')).toBeTruthy()
  })

  it('点击快捷键叶子分类打开快捷键弹窗', () => {
    render(<MainMenu />)
    fireEvent.click(screen.getByTestId('main-menu'))
    fireEvent.click(screen.getByText('快捷键'))
    expect(useUIStore.getState().modal).toBe('hotkey')
  })
})
```

- [ ] **Step 2: 运行测试确认失败**

```bash
npx vitest run src/components/menus/MainMenu.test.tsx
```

Expected: FAIL（`Cannot find module './MainMenu'`）。

- [ ] **Step 3: 实现 MainMenu 组件**

创建 `src/components/menus/MainMenu.tsx`：

```tsx
// 主菜单：级联二级菜单（对齐 word/excel 交互——悬停一级分类展开二级，带 ✓ 开关态与快捷键提示）
// 分类按 eflink 实际功能定制：文件 / 设计 / 视图 / 放映 / 快捷键（叶子）；仅列真实可用功能
import { Fragment, useEffect, useRef, useState } from 'react'
import {
  Check, ChevronDown, ChevronRight, Download, Eye, FilePlus2, FileText,
  Grid3x3, Keyboard, LayoutTemplate, Menu, MessageCircle, MonitorPlay,
  MonitorSpeaker, Palette, PanelLeft, Save, Search, Share2, Upload,
  type LucideIcon,
} from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { createDoc, saveDoc, type LoadedDoc } from '../../core/editor/persistence'
import { getPptxShareHandler } from '../../core/share/shareBridge'
import { ShareDialog } from '../common/ShareDialog'

interface MenuLeaf {
  key: string
  label: string
  icon?: LucideIcon
  shortcut?: string
  /** 开关型菜单项：选中时右侧显示 ✓ */
  checked?: boolean
  action?: () => void
}

interface MenuCategory {
  key: string
  label: string
  icon: LucideIcon
  /** 分组之间渲染分隔线 */
  groups: MenuLeaf[][]
  /** 叶子分类（无子菜单时直接点击触发，如快捷键） */
  leaf?: MenuLeaf
}

export function MainMenu() {
  const [open, setOpen] = useState(false)
  const [activeCat, setActiveCat] = useState<string | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const ui = useUIStore()

  // 分享弹窗（doc 为点击"分享"时刻的文档快照，弹窗期间编辑不影响本次分享内容）
  const [shareOpen, setShareOpen] = useState(false)
  const [shareDoc, setShareDoc] = useState<LoadedDoc | null>(null)
  // 分享前强制保存：先落库最新内容，再捕获当前文档；保存失败则中止分享（避免分享远端旧数据）
  const openShare = async () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast('文档未初始化，无法分享', 'error'); return }
    try {
      await saveDoc(s.docId, s.docName, s.presentation)
      s.markSaved()
    } catch {
      useToastStore.getState().toast('保存失败，无法分享', 'error')
      return
    }
    setShareDoc({ id: s.docId, name: s.docName, presentation: s.presentation })
    setShareOpen(true)
  }

  function closeAll() {
    setOpen(false)
    setActiveCat(null)
  }

  // Escape 关闭整组菜单；点击组件外部关闭
  useEffect(() => {
    if (!open) return
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) closeAll()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeAll()
    }
    window.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])

  const saveNow = () => {
    const s = useEditorStore.getState()
    if (!s.docId) { useToastStore.getState().toast('文档未初始化，无法保存', 'error'); return }
    saveDoc(s.docId, s.docName, s.presentation)
      .then(() => { s.markSaved(); useToastStore.getState().toast('已保存', 'success') })
      .catch(() => useToastStore.getState().toast('保存失败，请重试', 'error'))
  }

  const createNew = async () => {
    const doc = await createDoc('未命名演示文稿')
    useEditorStore.getState().loadDocument(doc)
  }

  const play = (mode: 'playing' | 'presenter') => {
    useUIStore.getState().setPlayerMode(mode, useEditorStore.getState().slideIndex)
  }

  // 分享入口仅在宿主注入分享实现后出现（纯组件独立运行时不显示）
  const canShare = getPptxShareHandler() !== null

  const categories: MenuCategory[] = [
    {
      key: 'file',
      label: '文件',
      icon: FileText,
      groups: [
        [
          { key: 'new', label: '新建文档', icon: FilePlus2, action: () => void createNew() },
          { key: 'save', label: '保存', icon: Save, shortcut: 'Ctrl S', action: saveNow },
        ],
        [
          { key: 'import', label: '导入（PPTX/JSON）…', icon: Upload, action: () => ui.openModal('import') },
          { key: 'export', label: '导出…', icon: Download, action: () => ui.openModal('export') },
        ],
        ...(canShare ? [[
          { key: 'share', label: '分享', icon: Share2, action: () => void openShare() },
          { key: 'feedback', label: '反馈', icon: MessageCircle, action: () => window.open('/contact', '_blank') },
        ] as MenuLeaf[]] : []),
      ],
    },
    {
      key: 'design',
      label: '设计',
      icon: Palette,
      groups: [
        [
          { key: 'template', label: '模板库…', icon: LayoutTemplate, action: () => ui.openModal('template') },
          { key: 'theme', label: '主题配色…', icon: Palette, action: () => ui.openModal('theme') },
        ],
      ],
    },
    {
      key: 'view',
      label: '视图',
      icon: Eye,
      groups: [
        [
          { key: 'grid', label: '网格', icon: Grid3x3, checked: ui.gridVisible, action: () => ui.toggleGrid() },
          { key: 'thumbnails', label: '缩略图面板', icon: PanelLeft, checked: ui.thumbnailsVisible, action: () => ui.toggleThumbnails() },
        ],
        [
          { key: 'findReplace', label: '查找替换…', icon: Search, shortcut: 'Ctrl F', action: () => ui.openModal('findReplace') },
        ],
      ],
    },
    {
      key: 'play',
      label: '放映',
      icon: MonitorPlay,
      groups: [
        [
          { key: 'play', label: '放映', icon: MonitorPlay, action: () => play('playing') },
          { key: 'presenter', label: '演讲者视图', icon: MonitorSpeaker, action: () => play('presenter') },
        ],
      ],
    },
    {
      key: 'hotkey',
      label: '快捷键',
      icon: Keyboard,
      /** 快捷键是叶子分类：点击直接打开对话框，不需要子菜单 */
      groups: [],
      leaf: {
        key: 'hotkey-list',
        label: '查看全部快捷键',
        icon: Keyboard,
        action: () => ui.openModal('hotkey'),
      },
    },
  ]

  const renderLeaf = (leaf: MenuLeaf) => (
    <button
      key={leaf.key}
      type="button"
      onClick={() => {
        leaf.action?.()
        closeAll()
      }}
      className="flex w-full items-center gap-2.5 px-3 py-[7px] text-left text-[13px] text-[#1f2329] transition-colors hover:bg-[#f2f3f4]"
    >
      {leaf.icon ? <leaf.icon size={16} className="shrink-0 text-[#51565f]" /> : <span className="w-4 shrink-0" />}
      <span className="min-w-0 flex-1 truncate">{leaf.label}</span>
      {leaf.checked && <Check size={13} className="shrink-0 text-[#1f2329]" />}
      {leaf.shortcut && <span className="shrink-0 text-xs text-[#8f959e]">{leaf.shortcut}</span>}
    </button>
  )

  return (
    <div ref={rootRef} className="relative shrink-0">
      <button
        type="button"
        data-testid="main-menu"
        title="主菜单"
        onClick={() => {
          setOpen((v) => !v)
          setActiveCat(null)
        }}
        className={`flex items-center gap-0.5 rounded-md px-2 py-1.5 text-xs transition-colors ${
          open ? 'bg-gray-100 text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'
        }`}
      >
        <Menu size={17} />
        <ChevronDown size={12} />
      </button>

      {open && (
        <div className="absolute left-0 top-full z-[9999] mt-1 w-[132px] rounded-lg border border-black/[0.08] bg-white py-1 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
          {categories.map((cat) => (
            <div
              key={cat.key}
              className="relative"
              onMouseEnter={() => {
                // 叶子分类（快捷键）不展开子菜单
                if (!cat.leaf) setActiveCat(cat.key)
              }}
            >
              <button
                type="button"
                onClick={() => {
                  // 叶子分类直接触发
                  if (cat.leaf) {
                    cat.leaf.action?.()
                    closeAll()
                    return
                  }
                  setActiveCat(cat.key)
                }}
                className={`flex w-full items-center justify-between px-3 py-2 text-left text-[13px] transition-colors ${
                  activeCat === cat.key
                    ? 'bg-[#f2f3f4] text-[#1f2329]'
                    : 'text-[#1f2329] hover:bg-[#f2f3f4]'
                } ${cat.leaf ? 'cursor-pointer' : ''}`}
              >
                <span className="flex items-center gap-2">
                  <cat.icon size={16} className="shrink-0 text-[#51565f]" />
                  {cat.label}
                </span>
                {!cat.leaf && <ChevronRight size={12} className="shrink-0 text-[#8f959e]" />}
              </button>

              {activeCat === cat.key && (
                <div className="absolute left-full top-0 z-[9999] ml-0.5 w-[224px] rounded-lg border border-black/[0.08] bg-white py-1 shadow-[0_4px_16px_rgba(0,0,0,0.12)]">
                  {cat.groups.map((group, gi) => (
                    <Fragment key={gi}>
                      {gi > 0 && <div className="my-1 h-px bg-black/[0.05]" />}
                      {group.map((leaf) => renderLeaf(leaf))}
                    </Fragment>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      <ShareDialog open={shareOpen} doc={shareDoc} onClose={() => setShareOpen(false)} />
    </div>
  )
}
```

- [ ] **Step 4: 运行测试确认通过**

```bash
npx vitest run src/components/menus/MainMenu.test.tsx
```

Expected: PASS（3 个用例全绿）。

- [ ] **Step 5: Commit**

```bash
git add src/components/menus/MainMenu.tsx src/components/menus/MainMenu.test.tsx
git commit -m "feat: 新增 pptx 主菜单组件（文件/设计/视图/放映/快捷键）"
```

---

### Task 3: TopBar 精简并接入 MainMenu

**Files:**
- Modify: `src/components/layout/TopBar.tsx`（整体重写，167 行 → 约 90 行）

- [ ] **Step 1: 重写 TopBar.tsx**

用以下内容**整体替换** `src/components/layout/TopBar.tsx`：

```tsx
// 顶部工具栏：品牌区 + 主菜单 + 高频图标（撤销/重做/插入/AI/放映）
// 文件类与低频功能入口已收进 MainMenu（文件/设计/视图/放映/快捷键）
import { ArrowLeft, Undo2, Redo2, MonitorPlay, Sparkles } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import { getEditorBackHref } from '../../core/editor/chrome'
import { InsertMenu } from '../menus/InsertMenu'
import { MainMenu } from '../menus/MainMenu'
import logoUrl from '../../assets/pptx-eflink-logo.png'

function ToolButton({ icon, label, onClick, disabled, active }: {
  icon: React.ReactNode
  label: string
  onClick: () => void
  disabled?: boolean
  active?: boolean
}) {
  return (
    <button
      className={`flex items-center gap-1 rounded-md px-2 py-1.5 text-xs transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${
        active ? 'bg-[#fbeae5] text-[#d14424]' : 'text-gray-600 hover:bg-gray-100'
      }`}
      title={label}
      onClick={onClick}
      disabled={disabled}
    >
      {icon}
    </button>
  )
}

export function TopBar() {
  const undoDepth = useEditorStore((s) => s.history.length)
  const redoDepth = useEditorStore((s) => s.future.length)
  const ui = useUIStore()
  const backHref = getEditorBackHref()

  return (
    <div className="flex h-12 shrink-0 items-center gap-1 border-b border-gray-200 bg-white px-3" data-testid="topbar">
      {backHref && (
        <a
          href={backHref}
          title="返回"
          className="mr-1 flex size-8 shrink-0 items-center justify-center rounded-md text-gray-500 transition-colors hover:bg-gray-100 hover:text-gray-700"
        >
          <ArrowLeft size={18} />
        </a>
      )}
      <img src={logoUrl} alt="易飞演示文稿" className="mr-1.5 h-8 w-8 rounded-full" />
      <span className="mr-2 text-base font-bold text-[#d14424]">易飞演示文稿</span>

      <MainMenu />

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <ToolButton icon={<Undo2 size={17} />} label="撤销" disabled={undoDepth === 0} onClick={() => useEditorStore.getState().undo()} />
      <ToolButton icon={<Redo2 size={17} />} label="重做" disabled={redoDepth === 0} onClick={() => useEditorStore.getState().redo()} />

      <div className="mx-1 h-6 w-px bg-gray-200" />
      <div className="flex-1" />

      <InsertMenu />

      <div className="flex-1" />

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <button
        className={`flex items-center gap-1 rounded-md px-2.5 py-1.5 text-xs ${ui.aiPanelVisible ? 'bg-[#d14424] text-white' : 'bg-[#fbeae5] text-[#d14424] hover:bg-[#f6d9d0]'}`}
        title="AI 助手"
        onClick={() => ui.toggleAIPanel()}
        data-testid="ai-toggle"
      >
        <Sparkles size={15} />
      </button>

      <div className="mx-1 h-6 w-px bg-gray-200" />

      <button
        className="ml-1 flex items-center gap-1 rounded-md bg-[#d14424] px-3 py-1.5 text-xs text-white hover:bg-[#b93a1d]"
        title="放映"
        onClick={() => ui.setPlayerMode('playing', useEditorStore.getState().slideIndex)}
        data-testid="play"
      >
        <MonitorPlay size={15} />
      </button>

      {/* 文档名与保存状态指示在底部状态栏（BottomBar） */}
    </div>
  )
}
```

说明（重写时同步移除的死代码/迁出逻辑，执行者无需再处理）：
- 原第 120-130 行的隐藏 `<input type="file">`（`fileInputRef`）为死代码——无任何地方调用 `fileInputRef.current.click()`，直接不保留；图片上传由 `InsertMenu.insertImage` 自建 input 完成
- 分享逻辑（`openShare`/`ShareDialog`/`shareBridge`）与保存/新建逻辑已迁入 `MainMenu`
- 原顶部「新建/文档管理/保存/模板库/主题配色/导入/导出/查找替换/快捷键/网格/缩略图/演讲者视图」图标全部移除

- [ ] **Step 2: 类型检查 + 全量测试**

```bash
npx tsc --noEmit -p tsconfig.json
# Expected: 无错误
npx vitest run
# Expected: 全量 PASS（含 MainMenu 3 个新用例）
```

- [ ] **Step 3: lint**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx
npx oxlint packages/pptx/src/components/layout/TopBar.tsx packages/pptx/src/components/menus/MainMenu.tsx
# Expected: 无 error
```

- [ ] **Step 4: Commit**

```bash
git add packages/pptx/src/components/layout/TopBar.tsx
git commit -m "refactor: 顶栏精简为品牌区+主菜单+高频图标，低频入口收进主菜单"
```

---

### Task 4: 全量验证 + 手动验收

**Files:** 无代码改动（验证任务）

- [ ] **Step 1: 残留检查**

```bash
cd /Users/apple/Documents/myf-project/eflink.tech/eflink-pptx
grep -rn "FileManagerDialog\|fileManager\|文档管理" packages/pptx/src --include="*.ts*"
# Expected: 无输出
```

- [ ] **Step 2: 全量类型检查 + 测试（仓库根）**

```bash
pnpm typecheck && pnpm test
# Expected: 均通过
```

- [ ] **Step 3: Playwright 手动验收**

启动 dev：`cd eflink-pptx && pnpm dev:demo`（端口见 demo 配置），逐项验证：

1. ☰ 主菜单展开，五个分类齐全；悬停展开二级、Escape/点击外部可关闭
2. 文件：新建文档生效；保存出现「已保存」toast；导入/导出弹窗能打开
3. 设计：模板库/主题配色弹窗能打开
4. 视图：网格/缩略图开关有 ✓ 态且生效；查找替换弹窗能打开
5. 放映：放映/演讲者视图能进入
6. 快捷键：点击打开快捷键弹窗
7. Ctrl+S / Ctrl+F 快捷键仍生效
8. 顶栏保留图标：撤销/重做（禁用态正确）、插入 9 件套、AI、放映主按钮
9. 顶栏与弹窗无「文档管理」残留

- [ ] **Step 4: 最终提交（如有手动验收微调）**

```bash
git add -A
git commit -m "chore: 主菜单整理收尾"
```

---

## 自查记录

- **Spec 覆盖:** 顶栏布局（Task 3）、五分类菜单（Task 2）、删除文档管理 4 处（Task 1）、验证方式（Task 4）——全部有对应任务
- **类型一致性:** `MenuLeaf`/`MenuCategory` 仅在 MainMenu.tsx 内部使用；`LoadedDoc`、`getPptxShareHandler`、`createDoc`/`saveDoc`/`markSaved` 均为既有导出；modal 名称均为 uiStore 既有值（移除 fileManager 后）
- **占位符:** 无 TBD/TODO；所有代码步骤含完整代码
