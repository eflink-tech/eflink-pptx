// AI 助手面板：设置 / 对话 / 生成大纲与整套页面 / 选中元素润色 / 会话历史
import { useEffect, useRef, useState } from 'react'
import { Settings, Trash2, X, Wand2, FileDown, Sparkles, History, ChevronLeft } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { aiSettingsStore, aiDb, BUILTIN_TEMPLATES } from '../../ai/db'
import { errorMessage, sendMessage, type ChatMessage } from '../../ai/aiService'
import {
  buildDeckPrompt, parseDeck, parseOutline,
} from '../../ai/systemPrompt'
import { applyGeneratedDeck, applyOutline } from '../../ai/actionExecutor'
import type { AISettings, Conversation, OutlineResult } from '../../ai/types'
import { genId } from '../../core/utils/id'

type PanelView = 'chat' | 'settings' | 'history'

export function AIChatPanel() {
  const ui = useUIStore()
  const [view, setView] = useState<PanelView>('chat')
  const [settings, setSettings] = useState<AISettings>({ baseUrl: '', apiKey: '', model: '' })
  const [configured, setConfigured] = useState(false)
  const [messages, setMessages] = useState<Array<{ role: 'user' | 'assistant'; content: string }>>([])
  const [input, setInput] = useState('')
  const [busy, setBusy] = useState(false)
  const [outline, setOutline] = useState<OutlineResult | null>(null)
  const [topic, setTopic] = useState('')
  const [pageCount, setPageCount] = useState(8)
  const [genView, setGenView] = useState(false)
  const [convId, setConvId] = useState<string | null>(null)
  const [historyList, setHistoryList] = useState<Conversation[]>([])
  const abortRef = useRef<AbortController | null>(null)
  const listRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    aiSettingsStore.getSettings().then((s) => {
      setSettings(s)
      setConfigured(Boolean(s.baseUrl && s.apiKey && s.model))
    })
  }, [])

  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight })
  }, [messages, busy])

  // 会话持久化：消息变化时写入 Dexie
  useEffect(() => {
    if (!messages.length) return
    const id = convId ?? (() => { const nid = genId('conv-'); setConvId(nid); return nid })()
    const now = Date.now()
    void aiDb.conversations.put({
      id,
      title: messages[0]?.content.replace(/\n/g, ' ').slice(0, 24) || '新对话',
      messages: messages.map((m) => ({ ...m, time: now })),
      createdAt: now,
      updatedAt: now,
    })
  }, [messages, convId])

  const loadHistory = () => {
    void aiDb.conversations.orderBy('updatedAt').reverse().toArray().then(setHistoryList)
  }

  const openConversation = (conv: Conversation) => {
    setMessages(conv.messages.map((m) => ({ role: m.role, content: m.content })))
    setConvId(conv.id)
    setView('chat')
  }

  const deleteConversation = (id: string) => {
    void aiDb.conversations.delete(id).then(loadHistory)
  }

  const saveSettings = async () => {
    await aiSettingsStore.saveSettings(settings)
    setConfigured(Boolean(settings.baseUrl && settings.apiKey && settings.model))
    setView('chat')
    useToastStore.getState().toast('AI 设置已保存', 'success')
  }

  const doSend = async (userText: string, onText?: (text: string) => string) => {
    if (busy) return
    const s = await aiSettingsStore.getSettings()
    if (!s.baseUrl || !s.apiKey || !s.model) {
      setView('settings')
      return
    }
    const text = onText ? onText(userText) : userText
    const next: Array<{ role: 'user' | 'assistant'; content: string }> = [...messages, { role: 'user', content: text }]
    setMessages(next)
    setInput('')
    setBusy(true)
    const controller = new AbortController()
    abortRef.current = controller

    const chat: ChatMessage[] = [
      { role: 'system', content: '你是易飞演示文稿的AI助手，回答简洁专业，用中文。' },
      ...next.map((m) => ({ role: m.role, content: m.content }) as ChatMessage),
    ]

    try {
      await sendMessage({
        settings: s,
        messages: chat,
        signal: controller.signal,
        onContentChunk: (_chunk, full) => {
          setMessages([...next, { role: 'assistant', content: full }])
        },
      })
    } catch (error) {
      setMessages([...next, { role: 'assistant', content: `⚠️ ${errorMessage(error)}` }])
    } finally {
      setBusy(false)
      abortRef.current = null
    }
  }

  const generateOutline = async () => {
    if (!topic.trim()) {
      useToastStore.getState().toast('请输入主题', 'error')
      return
    }
    const tpl = BUILTIN_TEMPLATES[0]
    const prompt = tpl.content
      .replace('{topic}', topic.trim())
      .replace('{count}', String(pageCount))
    const s = await aiSettingsStore.getSettings()
    setBusy(true)
    setMessages((prev) => [...prev, { role: 'user', content: `生成大纲：${topic.trim()}（${pageCount} 页）` }])
    try {
      const { content } = await sendMessage({
        settings: s,
        messages: [
          { role: 'system', content: '你是专业的演示文稿策划师，只输出 JSON。' },
          { role: 'user', content: prompt },
        ],
      })
      const result = parseOutline(content)
      setOutline(result)
      setMessages((prev) => [...prev, { role: 'assistant', content: `已生成「${result.topic}」大纲，共 ${result.slides.length} 页。\n${result.slides.map((sl, i) => `${i + 1}. ${sl.title}`).join('\n')}` }])
    } catch (error) {
      useToastStore.getState().toast(errorMessage(error), 'error')
      setMessages((prev) => [...prev, { role: 'assistant', content: `⚠️ ${errorMessage(error)}` }])
    } finally {
      setBusy(false)
    }
  }

  const generateDeck = async () => {
    if (!outline) return
    const store = useEditorStore.getState()
    const prompt = buildDeckPrompt(outline, store.presentation.theme.colors)
    const s = await aiSettingsStore.getSettings()
    setBusy(true)
    setMessages((prev) => [...prev, { role: 'user', content: '根据大纲生成整套页面版式…' }])
    try {
      const { content } = await sendMessage({
        settings: s,
        messages: [
          { role: 'system', content: '你是专业的 PPT 版式设计师，只输出 JSON。' },
          { role: 'user', content: prompt },
        ],
        onContentChunk: (_c, full) => {
          setMessages((prev) => {
            const base = [...prev]
            if (base[base.length - 1]?.role === 'assistant') base[base.length - 1] = { role: 'assistant', content: `生成中…（${full.length} 字）` }
            return base
          })
        },
      })
      const slides = parseDeck(content)
      applyGeneratedDeck(slides)
      setGenView(false)
      useToastStore.getState().toast(`已应用 ${slides.length} 页 AI 生成结果`, 'success')
      setMessages((prev) => [...prev, { role: 'assistant', content: `已应用 ${slides.length} 页到当前文档。` }])
    } catch (error) {
      useToastStore.getState().toast(errorMessage(error), 'error')
      setMessages((prev) => [...prev, { role: 'assistant', content: `⚠️ ${errorMessage(error)}` }])
    } finally {
      setBusy(false)
    }
  }

  /** 对选中文本元素执行模板操作（润色/扩写/翻译），发送后可用「应用」按钮回写 */
  const runTemplateOnSelection = async (templateId: string) => {
    const store = useEditorStore.getState()
    const slide = store.presentation.slides[store.slideIndex]
    const el = slide?.elements.find((e) => e.id === store.selectedIds[0])
    if (!el || el.type !== 'text') {
      useToastStore.getState().toast('请先选中文本元素', 'error')
      return
    }
    const rawText = el.content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
    const tpl = BUILTIN_TEMPLATES.find((t) => t.id === templateId)
    if (!tpl) return
    await doSend(rawText, (text) => tpl.content.replace('{text}', text))
  }

  // 对话结束后若为模板操作，提供「应用到选中元素」按钮
  const lastAssistant = [...messages].reverse().find((m) => m.role === 'assistant')
  const canApplyToSelection = Boolean(
    lastAssistant && !busy && selectedIsText(),
  )

  function selectedIsText(): boolean {
    const store = useEditorStore.getState()
    const slide = store.presentation.slides[store.slideIndex]
    const el = slide?.elements.find((e) => e.id === store.selectedIds[0])
    return el?.type === 'text'
  }

  const applyLastToSelection = () => {
    if (!lastAssistant) return
    const store = useEditorStore.getState()
    const text = lastAssistant.content
    store.pushHistory()
    store.updateElements([store.selectedIds[0]], (d) => {
      if (d.type !== 'text') return
      d.content = text.split('\n').filter(Boolean).map((line) => `<p>${line}</p>`).join('') || `<p>${text}</p>`
    })
    useToastStore.getState().toast('已应用到选中元素', 'success')
  }

  return (
    <div className="flex h-full w-[320px] shrink-0 flex-col border-l border-gray-200 bg-white" data-testid="ai-panel">
      {/* 头部 */}
      <div className="flex h-10 shrink-0 items-center justify-between border-b border-gray-100 px-3">
        <span className="flex items-center gap-1 text-sm font-medium text-[#d14424]">
          <Sparkles size={15} />
          AI 助手
        </span>
        <div className="flex items-center gap-1">
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title="会话历史" onClick={() => { loadHistory(); setView(view === 'history' ? 'chat' : 'history') }}>
            <History size={15} />
          </button>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title="AI 设置" onClick={() => setView(view === 'settings' ? 'chat' : 'settings')}>
            <Settings size={15} />
          </button>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title="新对话" onClick={() => { setMessages([]); setOutline(null); setConvId(null); setView('chat') }}>
            <Trash2 size={15} />
          </button>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" onClick={() => ui.toggleAIPanel()}>
            <X size={15} />
          </button>
        </div>
      </div>

      {view === 'history'
        ? (
            <div className="flex-1 overflow-y-auto p-3" data-testid="ai-history">
              {view === 'history' && (
                <button className="mb-2 flex items-center gap-1 text-xs text-gray-500 hover:text-gray-700" onClick={() => setView('chat')}>
                  <ChevronLeft size={14} />
                  返回对话
                </button>
              )}
              {historyList.length === 0 && <div className="py-6 text-center text-xs text-gray-400">暂无历史会话</div>}
              <div className="space-y-1.5">
                {historyList.map((conv) => (
                  <div key={conv.id} className="flex items-center gap-2 rounded-lg border border-gray-200 px-3 py-2">
                    <button className="flex-1 truncate text-left text-xs text-gray-700" onClick={() => openConversation(conv)}>
                      {conv.title}
                    </button>
                    <span className="shrink-0 text-[10px] text-gray-400">{new Date(conv.updatedAt).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })}</span>
                    <button className="shrink-0 rounded p-0.5 text-gray-400 hover:bg-red-50 hover:text-red-500" title="删除" onClick={() => deleteConversation(conv.id)}>🗑</button>
                  </div>
                ))}
              </div>
            </div>
          )
        : view === 'settings'
        ? (
            <div className="flex-1 overflow-y-auto p-3">
              <div className="mb-3 text-xs text-gray-500">
                配置任意 OpenAI 兼容接口（如通义千问、DeepSeek、Kimi）。密钥仅保存在本地浏览器（IndexedDB），不会上传。
              </div>
              {([['baseUrl', '接口地址', 'https://dashscope.aliyuncs.com/compatible-mode/v1'], ['apiKey', 'API Key', 'sk-…'], ['model', '模型', 'qwen-plus']] as const).map(([key, label, placeholder]) => (
                <div key={key} className="mb-2">
                  <div className="mb-1 text-xs text-gray-500">{label}</div>
                  <input
                    type={key === 'apiKey' ? 'password' : 'text'}
                    value={settings[key]}
                    placeholder={placeholder}
                    onChange={(e) => setSettings({ ...settings, [key]: e.target.value })}
                    className="w-full rounded border border-gray-200 px-2 py-1.5 text-xs focus:border-[#d14424] focus:outline-none"
                  />
                </div>
              ))}
              <button className="mt-2 w-full rounded-md bg-[#d14424] py-1.5 text-xs text-white hover:bg-[#b93a1d]" onClick={() => void saveSettings()}>
                保存设置
              </button>
            </div>
          )
        : (
            <>
              {/* 生成 PPT 流程 */}
              <div className="shrink-0 border-b border-gray-100 p-2">
                {!genView
                  ? (
                      <button
                        className="flex w-full items-center justify-center gap-1 rounded-md bg-[#fbeae5] py-1.5 text-xs text-[#d14424] hover:bg-[#f6d9d0]"
                        onClick={() => setGenView(true)}
                        data-testid="ai-gen-toggle"
                      >
                        <Wand2 size={14} />
                        AI 生成 PPT
                      </button>
                    )
                  : (
                      <div className="rounded-md bg-[#faf7f5] p-2">
                        <div className="mb-1.5 flex items-center justify-between text-xs text-gray-500">
                          <span>生成流程</span>
                          <button className="text-gray-400 hover:text-gray-600" onClick={() => setGenView(false)}>收起</button>
                        </div>
                        <input
                          value={topic}
                          onChange={(e) => setTopic(e.target.value)}
                          placeholder="输入主题，如：2026 产品规划"
                          className="mb-1.5 w-full rounded border border-gray-200 px-2 py-1.5 text-xs focus:border-[#d14424] focus:outline-none"
                        />
                        <div className="mb-1.5 flex items-center gap-2 text-xs text-gray-500">
                          页数
                          <input
                            type="number"
                            min={3}
                            max={20}
                            value={pageCount}
                            onChange={(e) => setPageCount(Math.min(20, Math.max(3, Number(e.target.value) || 8)))}
                            className="w-16 rounded border border-gray-200 px-1.5 py-1 text-xs"
                          />
                          <button
                            className="ml-auto rounded bg-[#d14424] px-3 py-1 text-white hover:bg-[#b93a1d] disabled:opacity-50"
                            disabled={busy}
                            onClick={() => void generateOutline()}
                            data-testid="ai-gen-outline"
                          >
                            生成大纲
                          </button>
                        </div>
                        {outline && (
                          <div className="rounded border border-gray-200 bg-white p-2">
                            <div className="mb-1 text-xs font-medium text-gray-700">大纲 · {outline.topic}（{outline.slides.length} 页）</div>
                            <ol className="mb-2 list-decimal space-y-0.5 pl-4 text-xs text-gray-500">
                              {outline.slides.map((s, i) => <li key={i}>{s.title}</li>)}
                            </ol>
                            <div className="flex gap-1.5">
                              <button
                                className="flex-1 rounded border border-gray-200 py-1 text-xs text-gray-600 hover:bg-gray-50"
                                onClick={() => { applyOutline(outline); useToastStore.getState().toast('已应用大纲到文档', 'success') }}
                              >
                                仅应用大纲
                              </button>
                              <button
                                className="flex-1 rounded bg-[#d14424] py-1 text-xs text-white hover:bg-[#b93a1d] disabled:opacity-50"
                                disabled={busy}
                                onClick={() => void generateDeck()}
                                data-testid="ai-gen-deck"
                              >
                                生成整套页面
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
              </div>

              {/* 消息列表 */}
              <div ref={listRef} className="flex-1 space-y-2 overflow-y-auto p-3" data-testid="ai-messages">
                {configured
                  ? null
                  : (
                      <div className="rounded-md bg-[#fbeae5] p-2 text-xs text-[#b93a1d]">
                        尚未配置 AI 接口，点击右上角 ⚙️ 进行设置。
                      </div>
                    )}
                {messages.map((m, i) => (
                  <div key={i} className={`rounded-md px-2.5 py-2 text-xs leading-relaxed whitespace-pre-wrap ${m.role === 'user' ? 'ml-6 bg-[#d14424] text-white' : 'mr-2 bg-gray-100 text-gray-700'}`}>
                    {m.content}
                  </div>
                ))}
                {canApplyToSelection && (
                  <button
                    className="flex w-full items-center justify-center gap-1 rounded-md border border-[#d14424] py-1.5 text-xs text-[#d14424] hover:bg-[#fbeae5]"
                    onClick={applyLastToSelection}
                  >
                    <FileDown size={13} />
                    应用最新结果到选中元素
                  </button>
                )}
              </div>

              {/* 选中元素快捷操作 */}
              {selectedIsText() && (
                <div className="flex shrink-0 flex-wrap gap-1 border-t border-gray-100 px-2 pt-2">
                  {BUILTIN_TEMPLATES.filter((t) => t.id !== 'tpl-outline').map((t) => (
                    <button
                      key={t.id}
                      className="rounded-full border border-gray-200 px-2 py-0.5 text-[11px] text-gray-500 hover:border-[#d14424] hover:text-[#d14424] disabled:opacity-50"
                      disabled={busy}
                      onClick={() => void runTemplateOnSelection(t.id)}
                    >
                      {t.name}
                    </button>
                  ))}
                </div>
              )}

              {/* 输入区 */}
              <div className="flex shrink-0 items-end gap-1.5 p-2">
                <textarea
                  value={input}
                  onChange={(e) => setInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault()
                      if (input.trim()) void doSend(input.trim())
                    }
                  }}
                  rows={2}
                  placeholder="输入消息，Enter 发送"
                  className="flex-1 resize-none rounded-md border border-gray-200 px-2 py-1.5 text-xs focus:border-[#d14424] focus:outline-none"
                  data-testid="ai-input"
                />
                {busy
                  ? (
                      <button className="rounded-md bg-gray-200 px-3 py-2 text-xs text-gray-600" onClick={() => abortRef.current?.abort()}>
                        停止
                      </button>
                    )
                  : (
                      <button
                        className="rounded-md bg-[#d14424] px-3 py-2 text-xs text-white hover:bg-[#b93a1d] disabled:opacity-50"
                        disabled={!input.trim()}
                        onClick={() => input.trim() && void doSend(input.trim())}
                      >
                        发送
                      </button>
                    )}
              </div>
            </>
          )}
    </div>
  )
}
