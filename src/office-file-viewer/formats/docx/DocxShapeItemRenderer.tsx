import React, { memo, useMemo } from 'react';
import type { DocxShapeItem } from '../../services/docx/types';
import { useOfficeResourceUrl } from '../../services/resource-store/useOfficeResourceUrl';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { DocxChartView } from './DocxChartView';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';

/** DOCX 形状子项渲染器属性。 */
type DocxShapeItemRendererProps = {
  /** 当前处理的项目。 */
  item: DocxShapeItem;
  /** 查找结果对应的顶层正文块标识。 */
  searchBlockId: string;
};

/** 形状边框通过独立图层绘制，避免 CSS 边框挤占 Office 声明的文字区域。 */
type DocxShapeBorderStyle = React.CSSProperties & {
  '--office-file-docx-shape-border'?: string;
  '--office-file-docx-shape-border-offset'?: string;
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

/** 将 OOXML 子图形的旋转与翻转转换为 CSS 变换。 */
function resolveShapeTransform(item: DocxShapeItem) {
  const transforms: string[] = [];
  if (item.rotation && Number.isFinite(item.rotation)) {
    const rotation = ((item.rotation % 360) + 360) % 360;
    transforms.push(`rotate(${rotation}deg)`);
  }
  if (item.flipH) transforms.push('scaleX(-1)');
  if (item.flipV) transforms.push('scaleY(-1)');
  return transforms.length ? transforms.join(' ') : undefined;
}

/** 渲染单个形状子项，使每个背景资源拥有独立且合法的 Hook 生命周期。 */
function DocxShapeItemRendererComponent({
  item,
  searchBlockId,
}: DocxShapeItemRendererProps) {
  const image = useOfficeResourceUrl(item.imageSrc);
  const path = useMemo(() => resolveShapePath(item), [item]);
  const pathLayers = useMemo(
    () =>
      item.pathLayers?.length
        ? item.pathLayers
        : path
        ? [
            {
              path,
              fillColor: item.fillColor,
              strokeColor: item.strokeColor,
              strokeWidth: item.strokeWidth,
              strokeDasharray: item.strokeDasharray,
            },
          ]
        : [],
    [
      item.fillColor,
      item.pathLayers,
      item.strokeColor,
      item.strokeDasharray,
      item.strokeWidth,
      path,
    ],
  );
  const transform = useMemo(() => resolveShapeTransform(item), [item]);
  const drawAsSvg = pathLayers.length > 0 || item.kind === 'line';
  const hasCssBorder = !drawAsSvg && Boolean(item.border);
  const borderStyle: DocxShapeBorderStyle = hasCssBorder
    ? {
        '--office-file-docx-shape-border': item.border,
        '--office-file-docx-shape-border-offset':
          String(-((item.strokeWidth ?? 1) / 2)) + 'px',
      }
    : {};
  const hyperlinkProps = useOfficeHyperlink<HTMLDivElement>({
    hyperlink: item.hyperlink,
    source: { type: 'shape', id: item.id },
  });
  const justifyContent =
    item.textVerticalAlign === 'middle'
      ? 'center'
      : item.textVerticalAlign === 'bottom'
      ? 'flex-end'
      : 'flex-start';

  return (
    <div
      {...hyperlinkProps}
      className={[
        'office-file-docx-shape__item',
        item.noWrap && 'office-file-docx-shape__item--no-wrap',
        item.clipVerticalOverflow &&
          'office-file-docx-shape__item--clip-vertical',
        hasCssBorder && 'office-file-docx-shape__item--bordered',
      ]
        .filter(Boolean)
        .join(' ')}
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
        ...borderStyle,
        borderRadius: item.borderRadius,
        paddingTop: item.paddingTop,
        paddingRight: item.paddingRight,
        paddingBottom: item.paddingBottom,
        paddingLeft: item.paddingLeft,
        whiteSpace: item.noWrap ? 'nowrap' : undefined,
        overflowWrap: item.noWrap ? 'normal' : undefined,
        wordBreak: item.noWrap ? 'normal' : undefined,
        transform,
        transformOrigin: transform ? 'center center' : undefined,
      }}
    >
      {pathLayers.length ? (
        <svg
          className="office-file-docx-shape__svg"
          viewBox={
            item.viewBox ??
            `0 0 ${Math.max(1, item.width)} ${Math.max(1, item.height)}`
          }
          preserveAspectRatio="none"
        >
          {pathLayers.map((layer, index) => (
            <path
              key={`${item.id}-path-${index}`}
              d={layer.path}
              fill={layer.fillColor ?? 'none'}
              stroke={layer.strokeColor ?? 'none'}
              strokeWidth={layer.strokeWidth}
              strokeDasharray={layer.strokeDasharray}
              vectorEffect="non-scaling-stroke"
            />
          ))}
        </svg>
      ) : null}
      {(item.blocks ?? item.paragraphs)?.map((block) =>
        block.type === 'table' ? (
          <DocxTableBlock
            key={block.id}
            block={block}
            availableWidth={item.width}
            maximumWidth={item.width}
            searchBlockId={searchBlockId}
          />
        ) : block.type === 'chart' ? (
          <div key={block.id} className="office-file-docx-table-block__chart">
            <DocxChartView block={block} zoom={100} />
          </div>
        ) : (
          <DocxParagraph
            key={block.id}
            block={block}
            compact
            asDiv
            insideShape
            searchBlockId={searchBlockId}
          />
        ),
      )}
    </div>
  );
}

export const DocxShapeItemRenderer = memo(DocxShapeItemRendererComponent);
