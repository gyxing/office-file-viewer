// TableRenderer 渲染 PPTX 表格元素，包括单元格填充、边框和文字样式。
import React, { memo } from 'react';
import type {
  TableElement,
  TextRun,
  TextStyle,
} from '../../../services/pptx/types';
import { useOfficeHyperlink } from '../../../shared/hyperlink';

/** 表格渲染器组件属性。 */
type TableRendererProps = {
  /** 当前处理或渲染的演示文稿元素。 */
  element: TableElement;
  /** 是否允许表格对象和内部文字响应链接交互。 */
  interactive: boolean;
};

function colorWithOpacity(color?: string, opacity?: number) {
  if (!color || opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

/** 渲染表格内单个文字片段并绑定可选链接。 */
function TableTextRun({
  run,
  cellStyle,
  sourceId,
  interactive,
}: {
  /** 当前文字片段。 */
  run: TextRun;
  /** 当前单元格继承的基础文字样式。 */
  cellStyle?: TextStyle;
  /** 当前片段在文稿中的稳定来源标识。 */
  sourceId: string;
  /** 是否允许当前文字片段响应链接交互。 */
  interactive: boolean;
}) {
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: run.hyperlink,
    source: { type: 'text', id: sourceId },
    interactive,
  });
  return (
    <span
      {...hyperlinkProps}
      style={{
        color: run.style?.color ?? cellStyle?.color,
        fontFamily: run.style?.fontFamily ?? cellStyle?.fontFamily,
        fontSize: run.style?.fontSize ?? cellStyle?.fontSize,
        fontWeight: run.style?.bold || cellStyle?.bold ? 600 : 400,
        fontStyle: run.style?.italic || cellStyle?.italic ? 'italic' : 'normal',
        textDecoration:
          [
            run.style?.underline || cellStyle?.underline ? 'underline' : '',
            run.style?.strike && run.style.strike !== 'none'
              ? 'line-through'
              : '',
          ]
            .filter(Boolean)
            .join(' ') || 'none',
        letterSpacing: run.style?.charSpace ?? cellStyle?.charSpace ?? 0,
      }}
    >
      {run.text}
    </span>
  );
}

/** 渲染表格渲染器。 */
function TableRendererComponent({ element, interactive }: TableRendererProps) {
  const columnWidths = element.columnWidths ?? [];
  const rowHeights = element.rowHeights ?? [];
  const hyperlinkProps = useOfficeHyperlink<HTMLTableElement>({
    hyperlink: element.hyperlink,
    source: { type: 'shape', id: element.id },
    interactive,
  });
  return (
    <table
      {...hyperlinkProps}
      style={{
        position: 'absolute',
        left: element.x,
        top: element.y,
        width: element.width,
        height: element.height,
        borderCollapse: 'collapse',
        tableLayout: 'fixed',
        color: '#172033',
        fontFamily: 'inherit',
        fontSize: 12,
        background: 'transparent',
      }}
    >
      <tbody>
        {element.rows.map((row, rowIndex) => (
          <tr
            key={rowIndex}
            style={{
              height: rowHeights[rowIndex],
            }}
          >
            {row.map((cell, cellIndex) => (
              <td
                key={cellIndex}
                style={{
                  width: columnWidths[cellIndex],
                  background: colorWithOpacity(
                    cell.backgroundColor ?? undefined,
                    cell.backgroundOpacity,
                  ),
                  borderStyle: cell.borderColor ? 'solid' : 'none',
                  borderColor:
                    colorWithOpacity(
                      cell.borderColor ?? undefined,
                      cell.borderOpacity,
                    ) ?? 'transparent',
                  borderWidth: cell.borderWidth ?? 1,
                  padding: `${cell.margins?.top ?? 0}px ${
                    cell.margins?.right ?? 0
                  }px ${cell.margins?.bottom ?? 0}px ${
                    cell.margins?.left ?? 0
                  }px`,
                  verticalAlign: cell.verticalAlign ?? 'middle',
                  overflow: 'hidden',
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-word',
                }}
              >
                {cell.paragraphs?.length
                  ? cell.paragraphs.map((paragraph, paragraphIndex) => (
                      <div
                        key={paragraphIndex}
                        style={{
                          textAlign:
                            paragraph.style?.align ??
                            cell.style?.align ??
                            'left',
                          lineHeight:
                            paragraph.style?.lineHeight ??
                            cell.style?.lineHeight ??
                            1.2,
                          whiteSpace: 'inherit',
                        }}
                      >
                        {paragraph.runs.map((run, runIndex) => (
                          <TableTextRun
                            key={runIndex}
                            run={run}
                            cellStyle={cell.style}
                            sourceId={`${element.id}:${rowIndex}:${cellIndex}:${paragraphIndex}:${runIndex}`}
                            interactive={interactive}
                          />
                        ))}
                      </div>
                    ))
                  : cell.text}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export const TableRenderer = memo(TableRendererComponent);
