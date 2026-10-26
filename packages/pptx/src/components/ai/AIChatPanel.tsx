// AI 助手面板：设置 / 对话 / 生成大纲与整套页面 / 选中元素润色 / 会话历史
import { useEffect, useRef, useState } from 'react'
import { useI18n } from '@eflink-tech/office-kit'
import { Settings, Trash2, X, Wand2, FileDown, Sparkles, History, ChevronLeft } from 'lucide-react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore, useToastStore } from '../../store/uiStore'
import { aiSettingsStore, aiDb, BUILTIN_TEMPLATES } from '../../ai/db'
import { errorMessage, sendMessage, type ChatMessage } from '../../ai/aiService'
import {
  buildDeckPrompt, parseDeck, parseOutline,
} from '../../ai/systemPrompt'
import { applyGeneratedDeck, applyOutline } from '../../ai/actionExecutor'
import { ensureAiLegacyImport } from '../../storage/legacy'
import type { AISettings, Conversation, OutlineResult } from '../../ai/types'
import { genId } from '../../core/utils/id'

type PanelView = 'chat' | 'settings' | 'history'

export function AIChatPanel() {
  const { t } = useI18n()
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
    // 首次读取 AI 数据前先做旧库一次性导入（ensureAiLegacyImport 内部吞掉失败，不阻塞编辑器）
    void (async () => {
      await ensureAiLegacyImport()
      const s = await aiSettingsStore.getSettings()
      setSettings(s)
      setConfigured(Boolean(s.baseUrl && s.apiKey && s.model))
    })().catch(() => { /* 读取失败则保持默认设置 */ })
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
      title: messages[0]?.content.replace(/\n/g, ' ').slice(0, 24) || t('pptx.ai.newChat'),
      messages: messages.map((m) => ({ ...m, time: now })),
      createdAt: now,
      updatedAt: now,
    })
  }, [messages, convId, t])

  const loadHistory = () => {
    void (async () => {
      await ensureAiLegacyImport()
      setHistoryList(await aiDb.conversations.orderBy('updatedAt').reverse().toArray())
    })().catch(() => { /* 读取失败则保持当前列表 */ })
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
    useToastStore.getState().toast(t('pptx.ai.settingsSaved'), 'success')
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
      { role: 'system', content: t('pptx.ai.systemPrompt') },
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
      useToastStore.getState().toast(t('pptx.ai.enterTopic'), 'error')
      return
    }
    const tpl = BUILTIN_TEMPLATES[0]
    const prompt = tpl.content
      .replace('{topic}', topic.trim())
      .replace('{count}', String(pageCount))
    const s = await aiSettingsStore.getSettings()
    setBusy(true)
    setMessages((prev) => [...prev, { role: 'user', content: t('pptx.ai.userGenOutlineMsg', { topic: topic.trim(), count: pageCount }) }])
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
      setMessages((prev) => [...prev, { role: 'assistant', content: `${t('pptx.ai.genOutlineDone', { topic: result.topic, count: result.slides.length })}\n${result.slides.map((sl, i) => `${i + 1}. ${sl.title}`).join('\n')}` }])
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
    setMessages((prev) => [...prev, { role: 'user', content: t('pptx.ai.userGenSlidesMsg') }])
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
            if (base[base.length - 1]?.role === 'assistant') base[base.length - 1] = { role: 'assistant', content: t('pptx.ai.generating', { count: full.length }) }
            return base
          })
        },
      })
      const slides = parseDeck(content)
      applyGeneratedDeck(slides)
      setGenView(false)
      useToastStore.getState().toast(t('pptx.ai.appliedSlides', { count: slides.length }), 'success')
      setMessages((prev) => [...prev, { role: 'assistant', content: t('pptx.ai.appliedSlidesMsg', { count: slides.length }) }])
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
      useToastStore.getState().toast(t('pptx.ai.needTextSelection'), 'error')
      return
    }
    const rawText = el.content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
    const tpl = BUILTIN_TEMPLATES.find((item) => item.id === templateId)
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
    useToastStore.getState().toast(t('pptx.ai.appliedSelection'), 'success')
  }

  return (
    <div className="flex h-full w-[320px] shrink-0 flex-col border-l border-gray-200 bg-white" data-testid="ai-panel">
      {/* 头部 */}
      <div className="flex h-10 shrink-0 items-center justify-between border-b border-gray-100 px-3">
        <span className="flex items-center gap-1 text-sm font-medium text-[#d14424]">
          <Sparkles size={15} />
          {t('pptx.ai.title')}
        </span>
        <div className="flex items-center gap-1">
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title={t('pptx.ai.history')} onClick={() => { loadHistory(); setView(view === 'history' ? 'chat' : 'history') }}>
            <History size={15} />
          </button>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title={t('pptx.ai.settings')} onClick={() => setView(view === 'settings' ? 'chat' : 'settings')}>
            <Settings size={15} />
          </button>
          <button className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600" title={t('pptx.ai.newChat')} onClick={() => { setMessages([]); setOutline(null); setConvId(null); setView('chat') }}>
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
                  {t('pptx.ai.back')}
                </button>
              )}
              {historyList.length === 0 && <div className="py-6 text-center text-xs text-gray-400">{t('pptx.ai.noHistory')}</div>}
              <div className="space-y-1.5">
                {historyList.map((conv) => (
                  <div key={conv.id} className="flex items-center gap-2 rounded-lg border border-gray-200 px-3 py-2">
                    <button className="flex-1 truncate text-left text-xs text-gray-700" onClick={() => openConversation(conv)}>
                      {conv.title}
                    </button>
                    <span className="shrink-0 text-[10px] text-gray-400">{new Date(conv.updatedAt).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })}</span>
                    <button className="shrink-0 rounded p-0.5 text-gray-400 hover:bg-red-50 hover:text-red-500" title={t('pptx.ai.delete')} onClick={() => deleteConversation(conv.id)}>🗑</button>
                  </div>
                ))}
              </div>
            </div>
          )
        : view === 'settings'
        ? (
            <div className="flex-1 overflow-y-auto p-3">
              <div className="mb-3 text-xs text-gray-500">
                {t('pptx.ai.settingsHint')}
              </div>
              {([['baseUrl', t('pptx.ai.baseUrl'), 'https://dashscope.aliyuncs.com/compatible-mode/v1'], ['apiKey', 'API Key', 'sk-…'], ['model', t('pptx.ai.model'), 'qwen-plus']] as const).map(([key, label, placeholder]) => (
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
                {t('pptx.ai.saveSettings')}
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
                        {t('pptx.ai.genTitle')}
                      </button>
                    )
                  : (
                      <div className="rounded-md bg-[#faf7f5] p-2">
                        <div className="mb-1.5 flex items-center justify-between text-xs text-gray-500">
                          <span>{t('pptx.ai.genFlow')}</span>
                          <button className="text-gray-400 hover:text-gray-600" onClick={() => setGenView(false)}>{t('pptx.ai.collapse')}</button>
                        </div>
                        <input
                          value={topic}
                          onChange={(e) => setTopic(e.target.value)}
                          placeholder={t('pptx.ai.topicPlaceholder')}
                          className="mb-1.5 w-full rounded border border-gray-200 px-2 py-1.5 text-xs focus:border-[#d14424] focus:outline-none"
                        />
                        <div className="mb-1.5 flex items-center gap-2 text-xs text-gray-500">
                          {t('pptx.ai.pageCount')}
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
                            {t('pptx.ai.genOutline')}
                          </button>
                        </div>
                        {outline && (
                          <div className="rounded border border-gray-200 bg-white p-2">
                            <div className="mb-1 text-xs font-medium text-gray-700">{t('pptx.ai.outline', { topic: outline.topic, count: outline.slides.length })}</div>
                            <ol className="mb-2 list-decimal space-y-0.5 pl-4 text-xs text-gray-500">
                              {outline.slides.map((s, i) => <li key={i}>{s.title}</li>)}
                            </ol>
                            <div className="flex gap-1.5">
                              <button
                                className="flex-1 rounded border border-gray-200 py-1 text-xs text-gray-600 hover:bg-gray-50"
                                onClick={() => { applyOutline(outline); useToastStore.getState().toast(t('pptx.ai.applyOutlineDone'), 'success') }}
                              >
                                {t('pptx.ai.applyOutline')}
                              </button>
                              <button
                                className="flex-1 rounded bg-[#d14424] py-1 text-xs text-white hover:bg-[#b93a1d] disabled:opacity-50"
                                disabled={busy}
                                onClick={() => void generateDeck()}
                                data-testid="ai-gen-deck"
                              >
                                {t('pptx.ai.genSlides')}
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
                        {t('pptx.ai.noConfig')}
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
                    {t('pptx.ai.applyToSelection')}
                  </button>
                )}
              </div>

              {/* 选中元素快捷操作 */}
              {selectedIsText() && (
                <div className="flex shrink-0 flex-wrap gap-1 border-t border-gray-100 px-2 pt-2">
                  {BUILTIN_TEMPLATES.filter((tpl) => tpl.id !== 'tpl-outline').map((tpl) => (
                    <button
                      key={tpl.id}
                      className="rounded-full border border-gray-200 px-2 py-0.5 text-[11px] text-gray-500 hover:border-[#d14424] hover:text-[#d14424] disabled:opacity-50"
                      disabled={busy}
                      onClick={() => void runTemplateOnSelection(tpl.id)}
                    >
                      {tpl.name}
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
                  placeholder={t('pptx.ai.inputPlaceholder')}
                  className="flex-1 resize-none rounded-md border border-gray-200 px-2 py-1.5 text-xs focus:border-[#d14424] focus:outline-none"
                  data-testid="ai-input"
                />
                {busy
                  ? (
                      <button className="rounded-md bg-gray-200 px-3 py-2 text-xs text-gray-600" onClick={() => abortRef.current?.abort()}>
                        {t('pptx.ai.stop')}
                      </button>
                    )
                  : (
                      <button
                        className="rounded-md bg-[#d14424] px-3 py-2 text-xs text-white hover:bg-[#b93a1d] disabled:opacity-50"
                        disabled={!input.trim()}
                        onClick={() => input.trim() && void doSend(input.trim())}
                      >
                        {t('pptx.ai.send')}
                      </button>
                    )}
              </div>
            </>
          )}
    </div>
  )
}
