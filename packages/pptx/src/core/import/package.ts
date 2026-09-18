// src/core/import/package.ts
/** 部件模型：zip 解压 + OOXML 关系（rels）图 + 媒体读取 */
import JSZip from 'jszip'
import { parseXML } from './xml'

export interface PartRel {
  /** 包内绝对路径（External 关系为原始目标） */
  target: string
  /** 关系 Type URI（如 .../relationships/image） */
  type?: string
  mode?: string
}

/** 相对路径解析：'../media/a.png'（基于 'ppt/slides'）→ 'ppt/media/a.png' */
export function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = `${baseDir}/${target}`.split('/')
  const out: string[] = []
  for (const part of parts) {
    if (part === '..') out.pop()
    else if (part && part !== '.') out.push(part)
  }
  return out.join('/')
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp', tiff: 'image/tiff',
  mp4: 'video/mp4', webm: 'video/webm', mov: 'video/quicktime',
  mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', wma: 'audio/x-ms-wma',
}

export class PptxPackage {
  private zip: JSZip
  private relsCache = new Map<string, Map<string, PartRel>>()
  private mediaCache = new Map<string, string>()

  private constructor(zip: JSZip) {
    this.zip = zip
  }

  static async load(file: Blob): Promise<PptxPackage> {
    return new PptxPackage(await JSZip.loadAsync(file))
  }

  async text(path: string): Promise<string | null> {
    return (await this.zip.file(path)?.async('text')) ?? null
  }

  /** zip 内全部文件路径列表（如 sldIdLst 缺失时按 slideN.xml 枚举页面兜底） */
  paths(): string[] {
    return Object.keys(this.zip.files)
  }

  /** 部件的 rels：'ppt/slides/slide1.xml' → 解析 'ppt/slides/_rels/slide1.xml.rels'，target 已解析为包内绝对路径。
   * 根部件（无 '/'）按约定解析 '_rels/<name>.rels'。
   * 返回值不得修改（内部缓存引用）。 */
  async rels(partPath: string): Promise<Map<string, PartRel>> {
    const cached = this.relsCache.get(partPath)
    if (cached) return cached
    const map = new Map<string, PartRel>()
    const idx = partPath.lastIndexOf('/')
    const baseDir = idx >= 0 ? partPath.slice(0, idx) : ''
    const name = idx >= 0 ? partPath.slice(idx + 1) : partPath
    const relPath = idx >= 0 ? `${baseDir}/_rels/${name}.rels` : `_rels/${name}.rels`
    const xml = await this.text(relPath)
    if (xml) {
      const doc = parseXML(xml)
      for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
        const id = rel.getAttribute('Id')
        const target = rel.getAttribute('Target')
        if (!id || !target) continue
        const mode = rel.getAttribute('TargetMode') ?? undefined
        map.set(id, {
          target: mode === 'External' ? target : resolveTarget(baseDir, target),
          type: rel.getAttribute('Type') ?? undefined,
          mode,
        })
      }
    }
    this.relsCache.set(partPath, map)
    return map
  }

  /** rId → 包内绝对路径（External 或缺失返回 null） */
  async relTarget(partPath: string, rId: string): Promise<string | null> {
    const rel = (await this.rels(partPath)).get(rId)
    if (!rel || rel.mode === 'External') return null
    return rel.target
  }

  /** 包内媒体文件 → data URL（带缓存） */
  async mediaDataUrl(packagePath: string): Promise<string | undefined> {
    const cached = this.mediaCache.get(packagePath)
    if (cached) return cached
    const file = this.zip.file(packagePath)
    if (!file) return undefined
    try {
      const base64 = await file.async('base64')
      const ext = packagePath.split('.').pop()?.toLowerCase() ?? ''
      const mime = MIME_BY_EXT[ext] ?? 'application/octet-stream'
      const dataUrl = `data:${mime};base64,${base64}`
      this.mediaCache.set(packagePath, dataUrl)
      return dataUrl
    } catch {
      return undefined
    }
  }
}
