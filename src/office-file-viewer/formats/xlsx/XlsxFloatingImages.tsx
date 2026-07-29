// XlsxFloatingImages 渲染锚定在工作表画布上的浮动图片。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../services/resource-store';
import type { XlsxImage, XlsxSheet } from '../../services/xlsx/types';
import {
  getXlsxMeasuredAnchorRect,
  XLSX_ROW_HEADER_WIDTH,
  type XlsxMeasuredAnchorRect,
  type XlsxSheetMetrics,
} from './sheetRenderUtils';

/** 定义 XlsxFloatingImages 组件可接收的属性。 */
type XlsxFloatingImagesProps = {
  /** 当前关联的工作表模型。 */
  sheet: XlsxSheet;
  /** 浏览器最终表格布局对应的工作表指标。 */
  metrics: XlsxSheetMetrics;
};

/** 渲染 XlsxFloatingImage 组件。 */
function XlsxFloatingImage({
  image,
  rect,
  columnHeaderHeight,
}: {
  /** 当前关联的图片资源或图片模型。 */
  image: XlsxImage;
  /** 按浏览器最终表格布局重算的锚点矩形。 */
  rect: XlsxMeasuredAnchorRect;
  /** 浏览器最终计算出的列标题行高度。 */
  columnHeaderHeight: number;
}) {
  const messages = useOfficeFileViewerMessages();
  const source = useMemo<OfficeResourceSource>(
    () =>
      typeof image.src === 'string'
        ? { kind: 'url', url: image.src }
        : image.src,
    [image.src],
  );
  const resource = useOfficeResourceUrl(source);
  const imageStyle = useMemo<CSSProperties>(
    () => ({
      left: XLSX_ROW_HEADER_WIDTH + rect.x,
      top: columnHeaderHeight + rect.y,
      width: rect.width,
      height: rect.height,
    }),
    [columnHeaderHeight, rect.height, rect.width, rect.x, rect.y],
  );

  return (
    <img
      className="office-file-xlsx-sheet-grid__floating-image"
      src={resource.url}
      alt={image.alt ?? ''}
      title={image.name}
      style={imageStyle}
      onError={(event) => {
        event.currentTarget.setAttribute(
          'aria-label',
          messages.spreadsheet.imageLoadFailed(image.alt),
        );
        event.currentTarget.setAttribute('data-load-error', 'true');
      }}
      aria-busy={resource.loading || undefined}
      data-load-error={resource.error ? 'true' : undefined}
    />
  );
}

const MemoXlsxFloatingImage = memo(XlsxFloatingImage);

/** 渲染 XlsxFloatingImagesComponent 组件。 */
function XlsxFloatingImagesComponent({
  sheet,
  metrics,
}: XlsxFloatingImagesProps) {
  const positionedImages = useMemo(
    () =>
      sheet.images.map((image) => ({
        image,
        rect: getXlsxMeasuredAnchorRect(sheet, metrics, image),
      })),
    [metrics, sheet],
  );

  return (
    <>
      {positionedImages.map(({ image, rect }) => (
        <MemoXlsxFloatingImage
          key={image.id}
          image={image}
          rect={rect}
          columnHeaderHeight={metrics.columnHeaderHeight}
        />
      ))}
    </>
  );
}

export const XlsxFloatingImages = memo(XlsxFloatingImagesComponent);
