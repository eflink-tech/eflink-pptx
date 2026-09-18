/** XML 安全访问助手：所有 OOXML 节点访问必须经由此模块，WPS 非标准结构不抛异常 */

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
  return descendants(parent, name)[0] ?? null
}

export function parseXML(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml')
}
