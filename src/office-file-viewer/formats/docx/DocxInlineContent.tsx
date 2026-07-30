// DocxInlineContent 渲染 DOCX 段落内的文本、换行、图片、图表和形状。
import React, { memo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxImage } from './DocxImage';
import { DocxInlineChart } from './DocxInlineChart';
import { DocxShape } from './DocxShape';
import { buildDocxTextStyle } from './docxRenderUtils';

/** 定义 DocxInlineContent 组件可接收的属性。 */
type DocxInlineContentProps = {
  /** DocxInlineContentProps 当前负责渲染的行内内容模型。 */
  inline: DocxInline;
};

/** 渲染 DocxInlineContentComponent 组件。 */
function DocxInlineContentComponent({ inline }: DocxInlineContentProps) {
  if (inline.type === 'break') return <br />;
  if (inline.type === 'tab')
    return (
      <span
        className="office-file-docx-inline-tab"
        style={buildDocxTextStyle(inline.style)}
        aria-hidden="true"
      />
    );
  if (inline.type === 'image') return <DocxImage inline={inline} />;
  if (inline.type === 'chart') return <DocxInlineChart inline={inline} />;
  if (inline.type === 'shape') return <DocxShape inline={inline} />;
  return (
    <span style={buildDocxTextStyle(inline.style, { includeBackground: true })}>
      {inline.text}
    </span>
  );
}

export const DocxInlineContent = memo(DocxInlineContentComponent);
