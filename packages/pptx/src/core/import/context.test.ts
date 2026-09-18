import { describe, expect, it } from 'vitest'
import { IDENTITY_XFORM, addSkipped, mapX, mapY } from './context'
import type { ImportReport } from '../../types/slides'

describe('context', () => {
  it('IDENTITY_XFORM 被冻结且为恒等变换', () => {
    expect(Object.isFrozen(IDENTITY_XFORM)).toBe(true)
    expect(mapX(IDENTITY_XFORM, 10)).toBe(10)
    expect(mapY(IDENTITY_XFORM, 10)).toBe(10)
  })

  it('addSkipped 按类别累计计数', () => {
    const report: ImportReport = { skipped: {} }
    addSkipped(report, 'chart')
    addSkipped(report, 'chart')
    addSkipped(report, 'oleObject')
    expect(report.skipped).toEqual({ chart: 2, oleObject: 1 })
  })
})
