// DocTableBlock 渲染 DOC 表格块，包括单元格文字、边框和列宽。
import React, { memo } from 'react';
import type { DocTableBlock as DocTableBlockModel } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

/** 定义 DocTableBlock 组件可接收的属性。 */
type DocTableBlockProps = {
  /** DocTableBlockProps 当前负责渲染的文档块模型。 */
  block: DocTableBlockModel;
};

/** 渲染 DocTableBlockComponent 组件。 */
function DocTableBlockComponent({ block }: DocTableBlockProps) {
  const columnCount = Math.max(...block.rows.map((row) => row.cells.length), 1);
  const borderColor = block.style?.borderColor ?? '#cfd7e3';
  const totalColumns =
    block.columns?.reduce((sum, width) => sum + width, 0) ?? 0;
  const hasRowSpans = block.rows.some((row) =>
    row.cells.some((cell) => Boolean(cell.rowSpan && cell.rowSpan > 1)),
  );

  const wrapperStyle =
    block.width && block.align === 'center'
      ? {
          width: block.width,
          marginLeft: `calc((100% - ${block.width}px) / 2)`,
        }
      : undefined;

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
          {block.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  className="office-file-doc-table__cell"
                  colSpan={
                    cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined
                  }
                  rowSpan={
                    cell.rowSpan && cell.rowSpan > 1 ? cell.rowSpan : undefined
                  }
                  style={{
                    borderTop: cell.borderTop ?? `1px solid ${borderColor}`,
                    borderRight: cell.borderRight ?? `1px solid ${borderColor}`,
                    borderBottom:
                      cell.borderBottom ?? `1px solid ${borderColor}`,
                    borderLeft: cell.borderLeft ?? `1px solid ${borderColor}`,
                    width: cell.width,
                    verticalAlign: cell.verticalAlign ?? 'top',
                    ...docTextStyleToCss(cell.style),
                  }}
                >
                  <DocInlineContent
                    inlines={cell.inlines}
                    fallback={cell.text}
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
                        (count, cell) => count + Math.max(1, cell.colSpan ?? 1),
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
          ))}
        </tbody>
      </table>
    </div>
  );
}

export const DocTableBlock = memo(DocTableBlockComponent);
