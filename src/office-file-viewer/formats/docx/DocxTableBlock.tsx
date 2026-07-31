// DocxTableBlock 渲染 DOCX 表格块，包括列宽、单元格边框、内边距和嵌套段落/图表。
import React, { memo } from 'react';
import type {
  DocxParagraphBlock,
  DocxTableBlock as DocxTableBlockModel,
  DocxTableCell,
} from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { DocxParagraph } from './DocxParagraph';
import { calculatePositionStyle } from './positionUtils';

/** DOCX 表格边线与内容区域之间的视觉修正量，单位为标准化渲染像素。 */
const DOCX_TABLE_EDGE_OFFSET = 7;

/** DOCX表格内容块组件属性。 */
type DocxTableBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxTableBlockModel;
  /** 当前可用宽度，单位为标准化渲染像素。 */
  availableWidth?: number;
  /** 表格在当前页面或容器内允许占用的最大物理宽度。 */
  maximumWidth?: number;
};

/** 渲染DOCX表格内容块。 */
function DocxTableBlockComponent({
  block,
  availableWidth,
  maximumWidth,
}: DocxTableBlockProps) {
  const resolveVerticalPadding = (
    value: number | undefined,
    hasExplicitRowHeight: boolean,
    cell: DocxTableCell,
  ) => {
    // Word 的自动行高会包含字体度量留白；仅在源文件未固定行高和内边距时补偿浏览器差异。
    const paragraph = cell.blocks.find(
      (item): item is DocxParagraphBlock => item.type === 'paragraph',
    );
    const paragraphFontSize = paragraph?.style?.fontSize ?? 14;
    const paragraphLineHeight =
      paragraph?.type === 'paragraph' && paragraph.lineHeight !== undefined
        ? paragraph.lineHeight > 4
          ? paragraph.lineHeight
          : paragraphFontSize * paragraph.lineHeight
        : paragraphFontSize * 1.2;
    const defaultVerticalPadding = block.insideShape
      ? 2
      : hasExplicitRowHeight
      ? 0
      : paragraphLineHeight <= 22
      ? 4.5
      : 3.5;
    return block.insideShape
      ? Math.max(value ?? 0, defaultVerticalPadding)
      : value ?? defaultVerticalPadding;
  };
  const marginLeft =
    block.align === 'center'
      ? 'auto'
      : block.align === 'right'
      ? 'auto'
      : -DOCX_TABLE_EDGE_OFFSET;
  const marginRight =
    block.align === 'center' ? 'auto' : block.align === 'right' ? 0 : 'auto';
  const totalColumns =
    block.columns?.reduce((sum, width) => sum + width, 0) ?? block.width ?? 0;
  // OOXML 的自动表宽常写成 0，真实宽度应回退到 tblGrid 的列宽总和。
  const resolvedTableWidth =
    block.width && block.width > 0
      ? block.width
      : totalColumns > 0
      ? totalColumns
      : undefined;
  const shouldFit = Boolean(
    !block.position &&
      resolvedTableWidth &&
      maximumWidth &&
      resolvedTableWidth > maximumWidth + DOCX_TABLE_EDGE_OFFSET * 2,
  );
  const constrainedTableWidth =
    shouldFit && maximumWidth ? maximumWidth : resolvedTableWidth;
  const tableWidth = block.position
    ? resolvedTableWidth ?? availableWidth ?? '100%'
    : shouldFit
    ? constrainedTableWidth
    : constrainedTableWidth ?? availableWidth ?? '100%';
  const overflowWidth =
    !block.position &&
    availableWidth &&
    constrainedTableWidth &&
    constrainedTableWidth > availableWidth
      ? constrainedTableWidth - availableWidth
      : 0;
  const overflowMarginLeft =
    block.align === 'center'
      ? -overflowWidth / 2
      : block.align === 'right'
      ? -overflowWidth
      : -DOCX_TABLE_EDGE_OFFSET;
  const positionStyle = calculatePositionStyle(block.position);

  return (
    <div
      className="office-file-docx-table-block"
      style={{
        ...positionStyle,
        position: block.position ? positionStyle.position : 'relative',
        top: block.position ? positionStyle.top : block.visualOffsetTop,
        zIndex: block.position ? positionStyle.zIndex : 1,
        marginTop: block.position ? undefined : block.marginTop,
        maxWidth: block.position ? 'none' : undefined,
      }}
    >
      <table
        className="office-file-docx-table-block__table"
        style={{
          width: tableWidth,
          marginLeft: block.position
            ? 0
            : overflowWidth
            ? overflowMarginLeft
            : marginLeft,
          marginRight: block.position ? 0 : marginRight,
        }}
      >
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col
                key={`${block.id}-col-${index}`}
                style={{
                  width:
                    shouldFit && totalColumns > 0
                      ? `${(width / totalColumns) * 100}%`
                      : width,
                }}
              />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => {
            const hasExplicitRowHeight = row.height !== undefined;
            return (
              <tr
                key={row.id}
                style={{
                  height: row.height,
                }}
              >
                {row.cells.map((cell) => (
                  <td
                    key={cell.id}
                    className="office-file-docx-table-block__cell"
                    colSpan={
                      cell.colSpan && cell.colSpan > 1
                        ? cell.colSpan
                        : undefined
                    }
                    rowSpan={
                      cell.rowSpan && cell.rowSpan > 1
                        ? cell.rowSpan
                        : undefined
                    }
                    style={{
                      borderTop:
                        cell.borderTop ??
                        (cell.hasBorderTop ? 'none' : '1px solid #cfd7e3'),
                      borderRight:
                        cell.borderRight ??
                        (cell.hasBorderRight ? 'none' : '1px solid #cfd7e3'),
                      borderBottom:
                        cell.borderBottom ??
                        (cell.hasBorderBottom ? 'none' : '1px solid #cfd7e3'),
                      borderLeft:
                        cell.borderLeft ??
                        (cell.hasBorderLeft ? 'none' : '1px solid #cfd7e3'),
                      paddingTop: resolveVerticalPadding(
                        cell.paddingTop,
                        hasExplicitRowHeight,
                        cell,
                      ),
                      paddingRight: cell.paddingRight ?? 7,
                      paddingBottom: resolveVerticalPadding(
                        cell.paddingBottom,
                        hasExplicitRowHeight,
                        cell,
                      ),
                      paddingLeft: cell.paddingLeft ?? 7,
                      width: shouldFit ? undefined : cell.width,
                      verticalAlign: cell.verticalAlign,
                      background: cell.backgroundColor ?? '#fff',
                      wordBreak: cell.noWrap ? 'normal' : 'break-word',
                      overflowWrap: cell.noWrap ? 'normal' : 'anywhere',
                      whiteSpace: cell.noWrap ? 'nowrap' : undefined,
                    }}
                  >
                    {cell.blocks.map((item) =>
                      item.type === 'chart' ? (
                        <div
                          key={item.id}
                          className="office-file-docx-table-block__chart"
                        >
                          <OfficeChartView
                            chart={item.chart}
                            width={item.width}
                            height={item.height}
                            zoom={100}
                          />
                        </div>
                      ) : item.type === 'table' ? (
                        <DocxTableBlockComponent
                          key={item.id}
                          block={item}
                          availableWidth={cell.width ?? availableWidth}
                          maximumWidth={
                            cell.width ?? availableWidth ?? maximumWidth
                          }
                        />
                      ) : (
                        <DocxParagraph
                          key={item.id}
                          block={item}
                          compact
                          asDiv
                        />
                      ),
                    )}
                  </td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export const DocxTableBlock = memo(DocxTableBlockComponent);
