/** XML 安全访问助手：所有 OOXML 节点访问必须经由此模块。
 *
 * 约定：
 * - 节点访问助手（attr/directChild/...）不抛异常；parseXML 对解析失败显式抛错。
 * - 助手优先按字面限定名（含前缀，如 "a:ln"）匹配；未命中时退化为 localName 匹配，
 *   以兼容 strict OOXML 与非常规前缀（两者 localName 一致）。
 * - `mc:AlternateContent` 的 Choice/Fallback 可能含同名重复元素，递归查找时调用方需注意去重/取舍。
 */

export function attr(el: Element | null, name: string): string | null {
  return el?.getAttribute(name) ?? null
}

/** 限定名匹配：字面一致优先，其次按 localName（忽略前缀） */
function nameMatches(el: Element, name: string): boolean {
  if (el.nodeName === name) return true
  const local = name.slice(name.indexOf(':') + 1)
  return el.localName === local
}

export function directChild(parent: Element, name: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (nameMatches(child, name)) return child
  }
  return null
}

export function directChildren(parent: Element, name: string): Element[] {
  return Array.from(parent.children).filter((c) => nameMatches(c, name))
}

export function descendants(parent: Element, name: string): Element[] {
  const exact = Array.from(parent.getElementsByTagName(name))
  if (exact.length) return exact
  const local = name.slice(name.indexOf(':') + 1)
  const all = Array.from(parent.getElementsByTagName('*'))
  return all.filter((el) => el.localName === local)
}

export function firstDescendant(parent: Element, name: string): Element | null {
  return descendants(parent, name)[0] ?? null
}

export function parseXML(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`XML 解析失败（前 80 字符）：${xml.slice(0, 80)}`)
  }
  return doc
}

/** p:graphicFrame 的内容类型：a:graphicData@uri 末段（table/chart/diagram/ole…），缺失返回 '' */
export function graphicDataKind(frame: Element): string {
  const uri = attr(firstDescendant(frame, 'a:graphicData'), 'uri') ?? ''
  return uri.slice(uri.lastIndexOf('/') + 1)
}

/** 是否 PresentationML 命名空间（过渡与 strict 两套 URI） */
export function isPresentationml(el: Element): boolean {
  const ns = el.namespaceURI
  return ns === 'http://schemas.openxmlformats.org/presentationml/2006/main'
    || ns === 'http://purl.oclc.org/ooxml/presentationml/main'
}

/** mc:AlternateContent：优先 mc:Fallback（旧客户端可渲染内容，如 3D 模型的平面位图），否则 mc:Choice。
 *  返回实际参与渲染的子节点列表；Choice/Fallback 可嵌套 AlternateContent，由调用方递归处理。 */
export function alternateContentChildren(ac: Element): Element[] {
  const fallback = directChild(ac, 'mc:Fallback')
  const fallbackChildren = fallback ? Array.from(fallback.children) : []
  if (fallbackChildren.length) return fallbackChildren
  const choice = directChild(ac, 'mc:Choice')
  return choice ? Array.from(choice.children) : []
}

/** mc:Choice@Requires（如 "am3d"、"dgm"、"a14"）；用于报告降级类别 */
export function alternateContentRequires(ac: Element): string {
  return attr(directChild(ac, 'mc:Choice'), 'Requires') ?? ''
}
