import { useEffect, useState } from 'react'
import { PptxEditor, builtinEditorBrand, resolveEditorBrandId, type EditorBrandId, type PptxLocale } from '@eflink-tech/pptx'

const LOCALES: { value: PptxLocale; label: string }[] = [
  { value: 'zh-CN', label: '中文' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'pt-BR', label: 'Português' },
]

const BRANDS: { value: EditorBrandId; label: string }[] = [
  { value: 'doculuma', label: 'DocuLuma' },
  { value: 'eflink', label: 'Eflink' },
]

/** ?locale=en 直达指定语言，便于逐语言验收 */
function readLocale(): PptxLocale {
  const value = new URLSearchParams(window.location.search).get('locale')
  return LOCALES.find((item) => item.value === value)?.value ?? 'zh-CN'
}

/** ?brand=eflink 直达指定品牌，便于逐品牌验收 */
function readBrand(): EditorBrandId {
  return resolveEditorBrandId(new URLSearchParams(window.location.search).get('brand'))
}

/** 独立 demo：整幅编辑器，自动打开上次文档（组件内行为） */
export default function App() {
  const [locale, setLocale] = useState<PptxLocale>(readLocale)
  const [brand, setBrand] = useState<EditorBrandId>(readBrand)

  const handleLocaleChange = (next: PptxLocale) => {
    const url = new URL(window.location.href)
    url.searchParams.set('locale', next)
    window.history.replaceState(null, '', url)
    setLocale(next)
  }

  const handleBrandChange = (next: EditorBrandId) => {
    const url = new URL(window.location.href)
    url.searchParams.set('brand', next)
    window.history.replaceState(null, '', url)
    setBrand(next)
  }

  useEffect(() => {
    document.title = builtinEditorBrand(brand).name
  }, [brand])

  return (
    <>
      <PptxEditor key={`${locale}-${brand}`} locale={locale} branding={brand} />
      <select
        value={brand}
        aria-label="Brand"
        onChange={(e) => handleBrandChange(e.target.value as EditorBrandId)}
        className="fixed bottom-3 right-3 z-[100000] rounded-md border border-black/10 bg-white px-2 py-1 text-xs text-[#1f2329] shadow-md outline-none"
      >
        {BRANDS.map((item) => (
          <option key={item.value} value={item.value}>
            {item.label}
          </option>
        ))}
      </select>
      <select
        value={locale}
        aria-label="Language"
        onChange={(e) => handleLocaleChange(e.target.value as PptxLocale)}
        className="fixed bottom-9 right-3 z-[100000] rounded-md border border-black/10 bg-white px-2 py-1 text-xs text-[#1f2329] shadow-md outline-none"
      >
        {LOCALES.map((item) => (
          <option key={item.value} value={item.value}>
            {item.label}
          </option>
        ))}
      </select>
    </>
  )
}
