/** 图片预览相对适应窗口状态的最小缩放倍数。 */
export const OFFICE_IMAGE_PREVIEW_MIN_ZOOM = 0.1;
/** 图片预览相对适应窗口状态的最大缩放倍数。 */
export const OFFICE_IMAGE_PREVIEW_MAX_ZOOM = 5;
/** 图片预览按钮和滚轮每次调整的缩放步长。 */
export const OFFICE_IMAGE_PREVIEW_ZOOM_STEP = 0.1;

/** 将缩放值限制到图片预览支持的范围并消除浮点误差。 */
export function normalizeOfficeImagePreviewZoom(zoom: number) {
  const clamped = Math.min(
    OFFICE_IMAGE_PREVIEW_MAX_ZOOM,
    Math.max(OFFICE_IMAGE_PREVIEW_MIN_ZOOM, zoom),
  );
  return Math.round(clamped * 10) / 10;
}

/** 根据图片旋转后的包围尺寸计算适应预览区域的基础缩放。 */
export function calculateOfficeImageFitScale(
  imageWidth: number,
  imageHeight: number,
  viewportWidth: number,
  viewportHeight: number,
  rotation: number,
) {
  if (
    imageWidth <= 0 ||
    imageHeight <= 0 ||
    viewportWidth <= 0 ||
    viewportHeight <= 0
  ) {
    return 1;
  }
  const swapsAxes = Math.abs(rotation % 180) === 90;
  const rotatedWidth = swapsAxes ? imageHeight : imageWidth;
  const rotatedHeight = swapsAxes ? imageWidth : imageHeight;
  return Math.min(
    1,
    Math.max(
      0.01,
      Math.min(viewportWidth / rotatedWidth, viewportHeight / rotatedHeight),
    ),
  );
}
