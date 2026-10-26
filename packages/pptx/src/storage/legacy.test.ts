import { beforeEach, describe, expect, it } from 'vitest'
import { migrateLegacyLocalStorage } from './legacy'
import { LAST_DOC_KEY, LEGACY_LAST_DOC_KEY, LEGACY_MIRROR_KEY, MIRROR_KEY } from './keys'

describe('migrateLegacyLocalStorage（efLink → DocuLuma）', () => {
  beforeEach(() => localStorage.clear())

  it('旧键值搬到新键，旧键被删除（真改名，不留下重复副本）', () => {
    localStorage.setItem(LEGACY_LAST_DOC_KEY, 'doc-1')
    localStorage.setItem(LEGACY_MIRROR_KEY, JSON.stringify({ id: 'doc-1', name: '旧文档' }))

    migrateLegacyLocalStorage()

    expect(localStorage.getItem(LAST_DOC_KEY)).toBe('doc-1')
    expect(JSON.parse(localStorage.getItem(MIRROR_KEY) ?? 'null')).toEqual({ id: 'doc-1', name: '旧文档' })
    expect(localStorage.getItem(LEGACY_LAST_DOC_KEY)).toBeNull()
    expect(localStorage.getItem(LEGACY_MIRROR_KEY)).toBeNull()
  })

  it('新键已有值时不覆盖，仅清理旧键', () => {
    localStorage.setItem(LEGACY_LAST_DOC_KEY, 'doc-old')
    localStorage.setItem(LAST_DOC_KEY, 'doc-new')

    migrateLegacyLocalStorage()

    expect(localStorage.getItem(LAST_DOC_KEY)).toBe('doc-new')
    expect(localStorage.getItem(LEGACY_LAST_DOC_KEY)).toBeNull()
  })
})
