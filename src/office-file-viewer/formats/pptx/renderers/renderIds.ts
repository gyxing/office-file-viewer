// renderIds 生成稳定且安全的 SVG id，避免主视口和缩略图同时渲染时发生引用冲突。
/** 根据输入构建 `buildRendererId` 返回的标准化结果。 */
export function buildRendererId(
  renderKey: string,
  elementId: string,
  suffix: string,
) {
  return `${renderKey}-${elementId}-${suffix}`.replace(/[^a-zA-Z0-9_-]/g, '-');
}
