// DocxTableBlock 渲染 DOCX 表格块，包括列宽、单元格边框、内边距和嵌套段落/图表。
import React, { memo } from 'react';
import type {
  DocxParagraphBlock,
  DocxTableBlock as DocxTableBlockModel,
  DocxTableCell,
} from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { DocxParagraph } from './DocxParagraph';
import { shouldSuppressDocxContextualSpacing } from './docxParagraphSpacing';
import { calculatePositionStyle } from './positionUtils';

/** DOCX 表格边线与内容区域之间的视觉修正量，单位为标准化渲染像素。 */
const DOCX_TABLE_EDGE_OFFSET = 7;

/** 读取 CSS 边框宽度，供 OOXML 单元格边距转换为内容区内边距。 */
function readCssBorderWidth(border: string | undefined) {
  if (!border || border === 'none') return 0;
  const width = Number.parseFloat(border);
  return Number.isFinite(width) ? width : 0;
}

/** Word 从网格线中心计算单元格边距，CSS 需扣除落在内容区内侧的边框宽度。 */
function resolveHorizontalCellPadding(
  margin: number | undefined,
  border: string | undefined,
) {
  return Math.max(
    0,
    (margin ?? DOCX_TABLE_EDGE_OFFSET) - readCssBorderWidth(border),
  );
}

/** 抵消浏览器把亚像素边框提升为一个布局像素造成的逐行累计误差。 */
function resolveThinBorderLayoutCompensation(border: string | undefined) {
  const width = readCssBorderWidth(border);
  return width > 0 && width < 1 ? 1 - width : 0;
}

/** DOCX表格内容块组件属性。 */
type DocxTableBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocxTableBlockModel;
  /** 当前可用宽度，单位为标准化渲染像素。 */
  availableWidth?: number;
  /** 表格在当前页面或容器内允许占用的最大物理宽度。 */
  maximumWidth?: number;
  /** 查找结果对应的顶层正文块标识。 */
  searchBlockId?: string;
};

/** 渲染DOCX表格内容块。 */
function DocxTableBlockComponent({
  block,
  availableWidth,
  maximumWidth,
  searchBlockId = block.sourceBlockId ?? block.id,
}: DocxTableBlockProps) {
  const resolveVerticalPadding = (
    value: number | undefined,
    hasExplicitRowHeight: boolean,
    cell: DocxTableCell,
  ) => {
    // 源文件显式声明的 0 也具有语义，不能再被浏览器补偿值覆盖。
    if (value !== undefined) return value;
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
    return defaultVerticalPadding;
  };
  // 文本框内边距已经定义表格左边界，正文页的视觉补偿不能再次叠加。
  const tableEdgeOffset = block.insideShape ? 0 : DOCX_TABLE_EDGE_OFFSET;
  const marginLeft =
    block.align === 'center'
      ? 'auto'
      : block.align === 'right'
      ? 'auto'
      : -tableEdgeOffset;
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
      resolvedTableWidth > maximumWidth + tableEdgeOffset * 2,
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
      : -tableEdgeOffset;
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
      data-office-word-block-id={searchBlockId}
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
            const fragment = row.fragment;
            return (
              <tr
                key={row.id}
                style={{
                  height: fragment?.height ?? row.height,
                }}
              >
                {row.cells.map((cell) => {
                  const paddingTop = resolveVerticalPadding(
                    cell.paddingTop,
                    hasExplicitRowHeight,
                    cell,
                  );
                  const paddingBottom = resolveVerticalPadding(
                    cell.paddingBottom,
                    hasExplicitRowHeight,
                    cell,
                  );
                  const borderRight =
                    cell.borderRight ??
                    (cell.hasBorderRight ? 'none' : '1px solid #cfd7e3');
                  const borderLeft =
                    cell.borderLeft ??
                    (cell.hasBorderLeft ? 'none' : '1px solid #cfd7e3');
                  const paddingRight = resolveHorizontalCellPadding(
                    cell.paddingRight,
                    borderRight,
                  );
                  const paddingLeft = resolveHorizontalCellPadding(
                    cell.paddingLeft,
                    borderLeft,
                  );
                  const content = cell.blocks.map((item, itemIndex) =>
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
                        searchBlockId={searchBlockId}
                      />
                    ) : (
                      <DocxParagraph
                        key={item.id}
                        block={item}
                        compact
                        asDiv
                        suppressSpacingBefore={shouldSuppressDocxContextualSpacing(
                          item,
                          cell.blocks[itemIndex - 1],
                        )}
                        suppressSpacingAfter={
                          itemIndex === cell.blocks.length - 1 ||
                          shouldSuppressDocxContextualSpacing(
                            item,
                            cell.blocks[itemIndex + 1],
                          )
                        }
                        searchBlockId={searchBlockId}
                      />
                    ),
                  );
                  const borderTop =
                    fragment && fragment.offset > 0.5
                      ? 'none'
                      : cell.borderTop ??
                        (cell.hasBorderTop ? 'none' : '1px solid #cfd7e3');
                  const borderBottom =
                    fragment &&
                    fragment.offset + fragment.height <
                      fragment.sourceHeight - 0.5
                      ? 'none'
                      : cell.borderBottom ??
                        (cell.hasBorderBottom ? 'none' : '1px solid #cfd7e3');
                  const borderLayoutCompensation =
                    resolveThinBorderLayoutCompensation(borderBottom);
                  return (
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
                        borderTop,
                        borderRight,
                        borderBottom,
                        borderLeft,
                        paddingTop: fragment ? 0 : paddingTop,
                        paddingRight: fragment ? 0 : paddingRight,
                        paddingBottom: fragment ? 0 : paddingBottom,
                        paddingLeft: fragment ? 0 : paddingLeft,
                        width: shouldFit ? undefined : cell.width,
                        height: fragment?.height,
                        verticalAlign: fragment ? 'top' : cell.verticalAlign,
                        background: cell.backgroundColor ?? '#fff',
                        wordBreak: cell.noWrap ? 'normal' : 'break-word',
                        overflowWrap: cell.noWrap ? 'normal' : 'anywhere',
                        whiteSpace: cell.noWrap ? 'nowrap' : undefined,
                      }}
                    >
                      {fragment ? (
                        <div
                          className="office-file-docx-table-block__fragment-clip"
                          style={{ height: fragment.height }}
                        >
                          <div
                            className="office-file-docx-table-block__fragment-content"
                            style={{
                              top: -fragment.offset,
                              height: fragment.sourceHeight,
                              paddingTop,
                              paddingRight,
                              paddingBottom,
                              paddingLeft,
                              justifyContent:
                                cell.verticalAlign === 'middle'
                                  ? 'center'
                                  : cell.verticalAlign === 'bottom'
                                  ? 'flex-end'
                                  : undefined,
                            }}
                          >
                            {content}
                          </div>
                        </div>
                      ) : (
                        <div
                          className="office-file-docx-table-block__cell-content"
                          style={{
                            marginBottom: borderLayoutCompensation
                              ? -borderLayoutCompensation
                              : undefined,
                          }}
                        >
                          {content}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export const DocxTableBlock = memo(DocxTableBlockComponent);
