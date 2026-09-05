/** 图层排列图标（叠放方块样式，对齐 Office/WPS 视觉） */

function LayerIcon({ size, children }: { size: number; children: React.ReactNode }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" aria-hidden>
      {children}
    </svg>
  )
}

/** 上移一层：后层实心 + 前层描边 + 弧形箭头 */
export function IconLayerForward({ size = 14 }: { size?: number }) {
  return (
    <LayerIcon size={size}>
      <rect x="1" y="1" width="7.5" height="7.5" rx="0.75" fill="currentColor" fillOpacity="0.4" />
      <rect x="6.5" y="6.5" width="7.5" height="7.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <path d="M11.5 12.2c1.2-.8 2-2 2.3-3.5" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
      <path d="M13.8 8.7V11M13.8 8.7H11.3" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
    </LayerIcon>
  )
}

/** 下移一层：前层实心 + 后层描边 + 弧形箭头 */
export function IconLayerBackward({ size = 14 }: { size?: number }) {
  return (
    <LayerIcon size={size}>
      <rect x="1" y="1" width="7.5" height="7.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <rect x="6.5" y="6.5" width="7.5" height="7.5" rx="0.75" fill="currentColor" fillOpacity="0.4" />
      <path d="M11.5 6.8c1.2.8 2 2 2.3 3.5" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
      <path d="M13.8 10.3V7.8M13.8 10.3H11.3" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" />
    </LayerIcon>
  )
}

/** 置顶：三层叠放，最前层实心 */
export function IconLayerToFront({ size = 14 }: { size?: number }) {
  return (
    <LayerIcon size={size}>
      <rect x="7" y="7" width="6.5" height="6.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <rect x="4" y="4" width="6.5" height="6.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <rect x="1" y="1" width="6.5" height="6.5" rx="0.75" fill="currentColor" fillOpacity="0.4" stroke="currentColor" strokeWidth="1.2" />
    </LayerIcon>
  )
}

/** 置底：三层叠放，最后层实心 */
export function IconLayerToBack({ size = 14 }: { size?: number }) {
  return (
    <LayerIcon size={size}>
      <rect x="1" y="1" width="6.5" height="6.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <rect x="4" y="4" width="6.5" height="6.5" rx="0.75" stroke="currentColor" strokeWidth="1.2" fill="white" />
      <rect x="7" y="7" width="6.5" height="6.5" rx="0.75" fill="currentColor" fillOpacity="0.4" stroke="currentColor" strokeWidth="1.2" />
    </LayerIcon>
  )
}
