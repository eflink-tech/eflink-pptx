import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createPresentation } from '../../types/slides'
import { loadStartupDoc, setPptxStorageBackend, writeMirror } from './persistence'

const LAST_DOC_KEY = 'eflink-pptx-last-doc'
const MIRROR_KEY = 'eflink-pptx-mirror'

/** 登录态（自定义后端）存储桩：get 桩可按用例改写 */
const backendGet = vi.fn()
const backend = {
  put: vi.fn(),
  get: backendGet,
  remove: vi.fn(),
  list: vi.fn(async () => []),
}

describe('loadStartupDoc（后端模式）镜像兜底', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    localStorage.clear()
    setPptxStorageBackend(backend)
  })
  afterEach(() => {
    setPptxStorageBackend(null)
  })

  it('指针文档加载不到（新建空文档）：不得把其他文档的镜像内容当作草稿恢复', async () => {
    localStorage.setItem(LAST_DOC_KEY, 'doc-B')
    // 镜像是「上次编辑的另一篇文档 A」的完整内容（历史上订阅器曾无条件写入）
    localStorage.setItem(
      MIRROR_KEY,
      JSON.stringify({ id: 'doc-A', name: '文档A', presentation: createPresentation('s-a') }),
    )
    backendGet.mockResolvedValue(undefined) // get('doc-B') → 新建文档无内容
    const doc = await loadStartupDoc()
    expect(doc.id).not.toBe('doc-A')
    expect(doc.name).toBe('未命名演示文稿')
  })

  it('镜像属于同一文档（id 与指针一致）：作为其未保存草稿恢复', async () => {
    localStorage.setItem(LAST_DOC_KEY, 'doc-B')
    localStorage.setItem(
      MIRROR_KEY,
      JSON.stringify({ id: 'doc-B', name: '文档B', presentation: createPresentation('s-b') }),
    )
    backendGet.mockResolvedValue(undefined)
    const doc = await loadStartupDoc()
    expect(doc.id).toBe('doc-B')
    expect(doc.presentation.slides[0].id).toBe('s-b')
  })

  it('writeMirror 在后端模式只更新指针，不写内容镜像', () => {
    writeMirror('doc-B', '文档B', createPresentation('s-b'))
    expect(localStorage.getItem(LAST_DOC_KEY)).toBe('doc-B')
    expect(localStorage.getItem(MIRROR_KEY)).toBeNull()
  })

  it('writeMirror 在本地模式写内容镜像（游客防崩溃丢稿语义不变）', () => {
    setPptxStorageBackend(null)
    writeMirror('doc-local', '本地文稿', createPresentation('s-l'))
    expect(localStorage.getItem(LAST_DOC_KEY)).toBe('doc-local')
    const mirror = JSON.parse(localStorage.getItem(MIRROR_KEY) ?? 'null')
    expect(mirror.id).toBe('doc-local')
    expect(mirror.presentation.slides[0].id).toBe('s-l')
  })
})
