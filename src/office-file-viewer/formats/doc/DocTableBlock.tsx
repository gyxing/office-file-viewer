// DocTableBlock 渲染 DOC 表格块，包括单元格文字、边框和列宽。
import React, { memo } from 'react';
import type { DocTableBlock as DocTableBlockModel } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

/** DOC表格内容块组件属性。 */
type DocTableBlockProps = {
  /** 当前负责处理或渲染的内容块。 */
  block: DocTableBlockModel;
};

/** 渲染DOC表格内容块。 */
function DocTableBlockComponent({ block }: DocTableBlockProps) {
  const columnCount = Math.max(...block.rows.map((row) => row.cells.length), 1);
  const borderColor = block.style?.borderColor ?? '#cfd7e3';
  const totalColumns =
    block.columns?.reduce((sum, width) => sum + width, 0) ?? 0;
  const hasRowSpans = block.rows.some((row) =>
    row.cells.some((cell) => Boolean(cell.rowSpan && cell.rowSpan > 1)),
  );
  const tableTopOffset = block.width ? 8 : 0;
  const marginTop = tableTopOffset + (block.spacingBefore ?? 0);
  const marginLeft =
    block.width && block.align === 'center'
      ? `calc((100% - ${block.width}px) / 2)`
      : block.width && block.align === 'right'
      ? `calc(100% - ${block.width}px)`
      : block.offsetLeft;
  const wrapperStyle = {
    width: block.width,
    marginLeft,
    marginTop: marginTop || undefined,
    marginBottom: block.spacingAfter,
  };

  return (
    <div
      className={`office-file-doc-table${
        block.width ? ' office-file-doc-table--document-grid' : ''
      }`}
      style={wrapperStyle}
    >
      <table className="office-file-doc-table__table" style={{ width: '100%' }}>
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col
                key={`${block.id}-col-${index}`}
                style={{
                  width: totalColumns
                    ? `${(width / totalColumns) * 100}%`
                    : width,
                }}
              />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => {
            // 空白模板行没有文本可撑高，沿用 Word 表格的标准占位行高。
            const emptyRowHeight =
              block.width &&
              row.cells.every(
                (cell) => !cell.text.trim() && !cell.inlines?.length,
              )
                ? 32
                : undefined;
            return (
              <tr key={row.id} style={{ height: row.height ?? emptyRowHeight }}>
                {row.cells.map((cell) => (
                  <td
                    key={cell.id}
                    className="office-file-doc-table__cell"
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
                      borderTop: cell.borderTop ?? `1px solid ${borderColor}`,
                      borderRight:
                        cell.borderRight ?? `1px solid ${borderColor}`,
                      borderBottom:
                        cell.borderBottom ?? `1px solid ${borderColor}`,
                      borderLeft: cell.borderLeft ?? `1px solid ${borderColor}`,
                      width: cell.width,
                      verticalAlign: cell.verticalAlign ?? 'top',
                      ...docTextStyleToCss(cell.style),
                      // Chromium 会把宋体粗体字面横向栅格化到整数像素，微调字距以匹配 Word 的字体度量。
                      letterSpacing:
                        block.width &&
                        (cell.style?.fontWeight ?? 0) >= 700 &&
                        /(?:SimSun|宋体)/i.test(cell.style?.fontFamily ?? '')
                          ? -1.25
                          : undefined,
                      // 结构化 DOC 表格按 Word 自动行高补偿浏览器字体度量。
                      paddingTop:
                        cell.style?.paddingTop ??
                        (block.width ? 3.5 : undefined),
                      paddingBottom:
                        cell.style?.paddingBottom ??
                        (block.width ? 3.5 : undefined),
                    }}
                  >
                    <DocInlineContent
                      inlines={cell.inlines}
                      fallback={cell.text}
                      sourceId={cell.id}
                      wordTableLineBreaks={Boolean(block.width)}
                    />
                  </td>
                ))}
                {!hasRowSpans &&
                row.cells.reduce(
                  (count, cell) => count + Math.max(1, cell.colSpan ?? 1),
                  0,
                ) < columnCount
                  ? Array.from({
                      length:
                        columnCount -
                        row.cells.reduce(
                          (count, cell) =>
                            count + Math.max(1, cell.colSpan ?? 1),
                          0,
                        ),
                    }).map((_, index) => (
                      <td
                        key={`${row.id}-empty-${index}`}
                        className="office-file-doc-table__cell"
                        style={{ border: `1px solid ${borderColor}` }}
                      />
                    ))
                  : null}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export const DocTableBlock = memo(DocTableBlockComponent);
