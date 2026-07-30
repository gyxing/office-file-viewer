import React, { memo, useMemo } from 'react';
import type { DocxShapeItem } from '../../services/docx/types';
import { useOfficeResourceUrl } from '../../services/resource-store/useOfficeResourceUrl';
import { DocxChartView } from './DocxChartView';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';

type DocxShapeItemRendererProps = {
  item: DocxShapeItem;
};

function resolveShapePath(item: DocxShapeItem) {
  if (item.path) return item.path;
  if (item.kind !== 'ellipse') return undefined;
  return `M ${item.width / 2} 0 A ${item.width / 2} ${item.height / 2} 0 1 0 ${
    item.width / 2
  } ${item.height} A ${item.width / 2} ${item.height / 2} 0 1 0 ${
    item.width / 2
  } 0`;
}

/** 渲染单个形状子项，使每个背景资源拥有独立且合法的 Hook 生命周期。 */
function DocxShapeItemRendererComponent({ item }: DocxShapeItemRendererProps) {
  const image = useOfficeResourceUrl(item.imageSrc);
  const path = useMemo(() => resolveShapePath(item), [item]);
  const drawAsSvg = Boolean(path) || item.kind === 'line';
  const justifyContent =
    item.textVerticalAlign === 'middle'
      ? 'center'
      : item.textVerticalAlign === 'bottom'
      ? 'flex-end'
      : 'flex-start';

  return (
    <div
      className="office-file-docx-shape__item"
      data-resource-error={image.error ? 'true' : undefined}
      style={{
        left: item.left,
        top: item.top,
        ...(item.fitShapeToText
          ? { minWidth: item.width, minHeight: item.height }
          : { width: item.width, height: item.height }),
        justifyContent,
        background: drawAsSvg ? undefined : item.fillColor,
        backgroundImage: image.url ? `url(${image.url})` : undefined,
        backgroundSize: item.imageSrc ? '100% 100%' : undefined,
        backgroundRepeat: item.imageSrc ? 'no-repeat' : undefined,
        border: drawAsSvg ? undefined : item.border,
        borderRadius: item.borderRadius,
        paddingTop: item.paddingTop,
        paddingRight: item.paddingRight,
        paddingBottom: item.paddingBottom,
        paddingLeft: item.paddingLeft,
        whiteSpace: item.noWrap ? 'nowrap' : undefined,
        overflowWrap: item.noWrap ? 'normal' : undefined,
        wordBreak: item.noWrap ? 'normal' : undefined,
      }}
    >
      {path ? (
        <svg
          className="office-file-docx-shape__svg"
          viewBox={
            item.viewBox ??
            `0 0 ${Math.max(1, item.width)} ${Math.max(1, item.height)}`
          }
          preserveAspectRatio="none"
        >
          <path
            d={path}
            fill={item.fillColor ?? 'none'}
            stroke={item.strokeColor ?? 'none'}
            strokeWidth={item.strokeWidth}
            strokeDasharray={item.strokeDasharray}
            vectorEffect="non-scaling-stroke"
          />
        </svg>
      ) : null}
      {(item.blocks ?? item.paragraphs)?.map((block) =>
        block.type === 'table' ? (
          <DocxTableBlock
            key={block.id}
            block={block}
            availableWidth={item.width}
            maximumWidth={item.width}
          />
        ) : block.type === 'chart' ? (
          <div key={block.id} className="office-file-docx-table-block__chart">
            <DocxChartView block={block} zoom={100} />
          </div>
        ) : (
          <DocxParagraph key={block.id} block={block} compact asDiv />
        ),
      )}
    </div>
  );
}

export const DocxShapeItemRenderer = memo(DocxShapeItemRendererComponent);
