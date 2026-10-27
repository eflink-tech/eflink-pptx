// 品牌可切换：宿主按部署站点选择 DocuLuma 或 Eflink 品牌（logo + 产品名）。
// 只影响界面观感；本地存储命名空间始终是 doculuma:*，与品牌切换无关。
import type { OfficeHeaderBrand } from '@eflink-tech/office-kit'
import doculumaLogo from './assets/pptx-doculuma-logo.png'
import eflinkLogo from './assets/pptx-eflink-logo.png'

export type EditorBrandId = 'doculuma' | 'eflink'

export const EDITOR_BRAND_IDS = ['doculuma', 'eflink'] as const

const BUILTIN_BRANDS: Record<EditorBrandId, Required<OfficeHeaderBrand>> = {
  doculuma: { name: 'DocuLuma Slides', logo: doculumaLogo },
  eflink: { name: '易飞演示文稿', logo: eflinkLogo },
}

/** 宽松解析品牌参数（'Eflink'/'EFLINK' 均可），缺省 doculuma */
export function resolveEditorBrandId(input?: string | null): EditorBrandId {
  return (input ?? '').trim().toLowerCase() === 'eflink' ? 'eflink' : 'doculuma'
}

export function builtinEditorBrand(id: EditorBrandId): Required<OfficeHeaderBrand> {
  return BUILTIN_BRANDS[id]
}
