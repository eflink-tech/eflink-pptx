// 分享桥接：宿主注入分享实现，包内只负责分享入口 UI
// 与 setPptxStorageBackend 同一注入模式：登录态、接口调用、链接拼装都由宿主决定，
// 组件库保持可独立运行：未注入 handler 时分享入口不渲染，包内零业务策略、零后端依赖
import type { LoadedDoc } from '../editor/persistence';

/** 分享结果：url 为完整分享链接；tips 为宿主业务文案（如有效期说明），弹窗按序展示 */
export interface PptxShareResult {
  url: string;
  tips?: string[];
}

/**
 * 分享实现：接收当前文档（含最新演示文稿数据），返回分享结果。
 * 抛错时分享弹窗展示错误信息并提供重试。
 */
export type PptxShareHandler = (doc: LoadedDoc) => Promise<PptxShareResult>;

let shareHandler: PptxShareHandler | null = null;

/** 注册分享实现（宿主须在编辑器挂载前调用；传 null 撤销并隐藏分享入口） */
export function setPptxShareHandler(handler: PptxShareHandler | null): void {
  shareHandler = handler;
}

export function getPptxShareHandler(): PptxShareHandler | null {
  return shareHandler;
}
