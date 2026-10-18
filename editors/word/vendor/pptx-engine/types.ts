// 摘自 genoffice packages/pptx-engine/src/types.ts（Apache-2.0）：custgeom.ts 只用到这个类型
export interface CustomGeometry {
  /** Main fill+stroke path */
  path?: string
  /** stroke="0" subpath: fill only (e.g. the pie sector of arc shapes) */
  fillPath?: string
  /** fill="none" subpath: stroke only (e.g. guide arcs) */
  strokePath?: string
}
