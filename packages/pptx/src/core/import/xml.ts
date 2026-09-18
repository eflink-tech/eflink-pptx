/** XML 安全访问助手：所有 OOXML 节点访问必须经由此模块。
 *
 * 约定：
 * - 节点访问助手（attr/directChild/...）不抛异常；parseXML 对解析失败显式抛错。
 * - 助手按字面限定名（含前缀，如 "a:ln"）匹配，调用方须传源文档完整限定名。
 * - `mc:AlternateContent` 的 Choice/Fallback 可能含同名重复元素，递归查找时调用方需注意去重/取舍。
 */

export function attr(el: Element | null, name: string): string | null {
  return el?.getAttribute(name) ?? null
}

export function directChild(parent: Element, name: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (child.nodeName === name) return child
  }
  return null
}

export function directChildren(parent: Element, name: string): Element[] {
  return Array.from(parent.children).filter((c) => c.nodeName === name)
}

export function descendants(parent: Element, name: string): Element[] {
  return Array.from(parent.getElementsByTagName(name))
}

export function firstDescendant(parent: Element, name: string): Element | null {
  return parent.getElementsByTagName(name)[0] ?? null
}

export function parseXML(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`XML 解析失败（前 80 字符）：${xml.slice(0, 80)}`)
  }
  return doc
}
