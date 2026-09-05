/** 生成短随机 id（页面/元素用） */
export function genId(prefix = ''): string {
  const rand = Math.random().toString(36).slice(2, 8)
  const time = Date.now().toString(36).slice(-4)
  return `${prefix}${rand}${time}`
}
