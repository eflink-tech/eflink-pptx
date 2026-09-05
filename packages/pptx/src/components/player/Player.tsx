// 放映模式：全屏放映 + 演讲者视图
import { useCallback, useEffect, useState } from 'react'
import { useEditorStore } from '../../store/editorStore'
import { useUIStore } from '../../store/uiStore'
import { SlideRenderer } from '../canvas/SlideRenderer'
import type { Slide } from '../../types/slides'

/** 为元素动画计算样式 */
function animStyle(slide: Slide): Record<string, React.CSSProperties> {
  const map: Record<string, React.CSSProperties> = {}
  let autoIndex = 0
  for (const el of slide.elements) {
    if (!el.anim) continue
    if (el.anim.type === 'in') {
      let delay = el.anim.delay
      if (el.anim.trigger === 'afterPrevious') {
        delay += autoIndex * 600
        autoIndex++
      }
      map[el.id] = {
        animation: `anim-in-${el.anim.effect} ${el.anim.duration}ms ${delay}ms both`,
      }
    } else if (el.anim.type === 'out') {
      map[el.id] = { animation: `anim-out-${el.anim.effect} ${el.anim.duration}ms ${el.anim.delay}ms both` }
    } else {
      map[el.id] = { animation: `anim-em-${el.anim.effect} ${el.anim.duration}ms ${el.anim.delay}ms both` }
    }
  }
  return map
}

function FullscreenPlayer() {
  const presentation = useEditorStore((s) => s.presentation)
  const [index, setIndex] = useState(() => useUIStore.getState().playerStartIndex)
  const [scale, setScale] = useState(1)
  const [flash, setFlash] = useState<'black' | 'white' | null>(null)

  const slide = presentation.slides[index]
  const slideW = presentation.width
  const slideH = Math.round(presentation.width / presentation.viewportRatio)

  useEffect(() => {
    const onResize = () => setScale(Math.min(window.innerWidth / slideW, window.innerHeight / slideH))
    onResize()
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [slideW, slideH])

  // 进入全屏；全屏退出时（无论 ESC / F11 / 代码触发）自动关闭放映
  useEffect(() => {
    let exited = false
    void document.documentElement.requestFullscreen?.().catch(() => undefined)
    const onFsChange = () => {
      if (!document.fullscreenElement && !exited) {
        exited = true
        useUIStore.getState().setPlayerMode('off')
      }
    }
    document.addEventListener('fullscreenchange', onFsChange)
    return () => {
      document.removeEventListener('fullscreenchange', onFsChange)
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined)
    }
  }, [])

  const next = useCallback(() => {
    if (flash) { setFlash(null); return }
    setIndex((i) => Math.min(i + 1, presentation.slides.length - 1))
  }, [flash, presentation.slides.length])

  const prev = useCallback(() => setIndex((i) => Math.max(0, i - 1)), [])

  const exit = useCallback(() => {
    if (document.fullscreenElement) {
      void document.exitFullscreen().catch(() => undefined)
    }
    useUIStore.getState().setPlayerMode('off')
  }, [])

  /** 点击：命中带链接的元素则跳转；音视频控件点击不翻页；否则翻页 */
  const handleStageClick = useCallback((e: React.MouseEvent) => {
    if (flash) { setFlash(null); return }
    const target = e.target as HTMLElement
    // 音视频自带控件，交给媒体自身处理
    if (target.closest('video, audio')) return
    const node = target.closest('[data-el-id]')
    if (node) {
      const id = node.getAttribute('data-el-id')
      const el = presentation.slides[index]?.elements.find((x) => x.id === id)
      const link = el?.link
      if (link) {
        if (link.type === 'web') {
          window.open(link.target, '_blank', 'noopener')
          return
        }
        const targetIndex = presentation.slides.findIndex((s) => s.id === link.target)
        if (targetIndex >= 0) {
          setIndex(targetIndex)
          return
        }
      }
    }
    next()
  }, [flash, index, next, presentation.slides])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      switch (e.key) {
        case 'Escape': exit(); break
        case 'ArrowRight': case ' ': case 'PageDown': case 'Enter': e.preventDefault(); next(); break
        case 'ArrowLeft': case 'PageUp': e.preventDefault(); prev(); break
        case 'b': case 'B': setFlash((f) => (f ? null : 'black')); break
        case 'w': case 'W': setFlash((f) => (f ? null : 'white')); break
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [exit, next, prev])

  if (!slide) return null
  const anims = animStyle(slide)
  const transition = slide.transition
  const transClass = transition && transition.preset !== 'none' ? `trans-${transition.preset}` : ''

  return (
    <div
      className="fixed inset-0 z-[200] flex items-center justify-center bg-black"
      onClick={handleStageClick}
      onContextMenu={(e) => { e.preventDefault(); prev() }}
      data-testid="player"
    >
      {flash
        ? <div className={`absolute inset-0 ${flash === 'black' ? 'bg-black' : 'bg-white'}`} />
        : (
            <div
              key={slide.id}
              className={`relative ${transClass}`}
              style={{ width: slideW, height: slideH, transform: `scale(${scale})`, transformOrigin: 'center' }}
            >
              <SlideRenderer slide={slide} width={slideW} height={slideH} />
              <style>{`
                .pptx-element { animation-fill-mode: both; }
                ${Object.entries(anims).map(([id, style]) => `
                  .player-anim [data-el-id="${id}"] { ${cssText(style)} }
                `).join('\n')}
              `}</style>
            </div>
          )}
      <div className="absolute bottom-3 right-4 text-xs text-white/50">
        {index + 1} / {presentation.slides.length} · 点击/空格下一页 · 左键返回 · Esc 退出
      </div>
    </div>
  )
}

function cssText(style: React.CSSProperties): string {
  return Object.entries(style)
    .map(([k, v]) => `${k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`)}:${v}`)
    .join(';')
}

/** 演讲者视图 */
function PresenterView() {
  const presentation = useEditorStore((s) => s.presentation)
  const [index, setIndex] = useState(() => useUIStore.getState().playerStartIndex)
  const [elapsed, setElapsed] = useState(0)
  const [scale, setScale] = useState(1)

  const slide = presentation.slides[index]
  const nextSlide = presentation.slides[index + 1]
  const slideW = presentation.width
  const slideH = Math.round(presentation.width / presentation.viewportRatio)

  useEffect(() => {
    const onResize = () => {
      const availW = window.innerWidth - 320 - 48
      const availH = window.innerHeight - 90
      if (availW <= 0 || availH <= 0) return
      setScale(Math.min(availW / slideW, availH / slideH))
    }
    onResize()
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [slideW, slideH])

  useEffect(() => {
    const timer = window.setInterval(() => setElapsed((v) => v + 1), 1000)
    return () => window.clearInterval(timer)
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') useUIStore.getState().setPlayerMode('off')
      if (e.key === 'ArrowRight' || e.key === ' ') setIndex((i) => Math.min(i + 1, presentation.slides.length - 1))
      if (e.key === 'ArrowLeft') setIndex((i) => Math.max(0, i - 1))
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [presentation.slides.length])

  if (!slide) return null
  const anims = animStyle(slide)
  const mm = String(Math.floor(elapsed / 60)).padStart(2, '0')
  const ss = String(elapsed % 60).padStart(2, '0')

  return (
    <div className="fixed inset-0 z-[200] flex flex-col bg-[#1e1e22] p-4" data-testid="presenter">
      <div className="mb-2 flex items-center justify-between text-xs text-gray-400">
        <span>演讲者视图 · {index + 1} / {presentation.slides.length}</span>
        <span className="font-mono text-base text-white" data-testid="presenter-timer">{mm}:{ss}</span>
        <button className="rounded bg-white/10 px-3 py-1 hover:bg-white/20" onClick={() => useUIStore.getState().setPlayerMode('off')}>退出</button>
      </div>
      <div className="flex min-h-0 flex-1 gap-3">
        <div className="flex flex-1 items-center justify-center overflow-hidden rounded-lg bg-black">
          <div key={slide.id} className="player-anim" style={{ width: slideW, height: slideH, transform: `scale(${scale})`, transformOrigin: 'center' }}>
            <SlideRenderer slide={slide} width={slideW} height={slideH} />
            <style>{`
              ${Object.entries(anims).map(([id, style]) => `
                .player-anim [data-el-id="${id}"] { ${cssText(style)} }
              `).join('\n')}
            `}</style>
          </div>
        </div>
        <div className="flex w-[280px] shrink-0 flex-col gap-3">
          <div className="flex-1 overflow-y-auto rounded-lg bg-white/5 p-3 text-sm leading-relaxed text-gray-200" data-testid="presenter-notes">
            {slide.note || '（本页无备注）'}
          </div>
          <div className="h-[130px] overflow-hidden rounded-lg bg-white/10">
            {nextSlide
              ? (
                  <div style={{ width: 280, height: Math.round(280 / presentation.viewportRatio) }}>
                    <div style={{ transform: `scale(${280 / slideW})`, transformOrigin: 'top left' }}>
                      <SlideRenderer slide={nextSlide} width={slideW} height={slideH} staticMode />
                    </div>
                  </div>
                )
              : <div className="flex h-full items-center justify-center text-xs text-gray-500">已经是最后一页</div>}
          </div>
        </div>
      </div>
    </div>
  )
}

export function Player() {
  const mode = useUIStore((s) => s.playerMode)
  if (mode === 'playing') return <FullscreenPlayer />
  if (mode === 'presenter') return <PresenterView />
  return null
}
