import { useState } from 'react'
import { PptxEditor, type PptxLocale } from '@eflink-tech/pptx'

const LOCALES: { value: PptxLocale; label: string }[] = [
  { value: 'zh-CN', label: '中文' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'pt-BR', label: 'Português' },
]

function readLocale(): PptxLocale {
  const value = new URLSearchParams(window.location.search).get('locale')
  return LOCALES.find((item) => item.value === value)?.value ?? 'zh-CN'
}

function LanguageSwitcher({
  locale,
  onChange,
}: {
  locale: PptxLocale
  onChange: (locale: PptxLocale) => void
}) {
  return (
    <select
      value={locale}
      onChange={(e) => onChange(e.target.value as PptxLocale)}
      aria-label="Language"
      className="fixed bottom-9 right-3 z-[100000] rounded-md border border-black/10 bg-white px-2 py-1 text-xs text-[#1f2329] shadow-md outline-none"
    >
      {LOCALES.map((item) => (
        <option key={item.value} value={item.value}>
          {item.label}
        </option>
      ))}
    </select>
  )
}

export default function App() {
  const [locale, setLocale] = useState<PptxLocale>(readLocale)

  const handleLocaleChange = (next: PptxLocale) => {
    const url = new URL(window.location.href)
    url.searchParams.set('locale', next)
    window.history.replaceState(null, '', url)
    setLocale(next)
  }

  return (
    <>
      <PptxEditor locale={locale} />
      <LanguageSwitcher locale={locale} onChange={handleLocaleChange} />
    </>
  )
}
