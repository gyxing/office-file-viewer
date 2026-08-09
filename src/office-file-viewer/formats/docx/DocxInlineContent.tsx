// DocxInlineContent 渲染 DOCX 段落内的文本、换行、图片、图表和形状。
import React, { createContext, memo, useContext } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { DocxImage } from './DocxImage';
import { DocxInlineChart } from './DocxInlineChart';
import { DocxShape } from './DocxShape';
import { buildDocxTextStyle } from './docxRenderUtils';

/** 控制当前 DOCX 是否按源设置压缩东亚标点，仅在声明该规则的文档中启用。 */
export const DocxCharacterSpacingContext = createContext(false);

/** 需要按东亚排版规则压缩的标点匹配表达式。 */
const COMPRESSIBLE_PUNCTUATION_PATTERN = /([，。；：！？、（）【】《》“”‘’])/u;

/** 将需要压缩的东亚标点拆成独立行内节点，避免缩窄普通正文字符。 */
function renderCompressedPunctuation(text: string) {
  return text.split(COMPRESSIBLE_PUNCTUATION_PATTERN).map((part, index) =>
    COMPRESSIBLE_PUNCTUATION_PATTERN.test(part) ? (
      <span
        key={`punctuation-${index}`}
        className="office-file-docx-compressed-punctuation"
      >
        {part}
      </span>
    ) : (
      part
    ),
  );
}

/** DOCX 行内内容组件属性。 */
type DocxInlineContentProps = {
  /** 当前负责渲染的行内内容模型。 */
  inline: DocxInline;
  /** 当前行内内容在所属文档中的稳定标识。 */
  sourceId: string;
};

/** 渲染 DOCX 段落中的行内文字、图片和图表。 */
function DocxInlineContentComponent({
  inline,
  sourceId,
}: DocxInlineContentProps) {
  const compressPunctuation = useContext(DocxCharacterSpacingContext);
  const hyperlinkProps = useOfficeHyperlink<HTMLSpanElement>({
    hyperlink: inline.type === 'text' ? inline.hyperlink : undefined,
    source: { type: 'text', id: sourceId },
  });
  if (inline.type === 'break') return <br />;
  if (inline.type === 'bookmark') {
    return (
      <span
        className="office-file-word-bookmark"
        data-office-word-bookmark={inline.name}
        data-office-word-bookmark-id={inline.markerId}
        aria-hidden="true"
      />
    );
  }
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
    <span
      {...hyperlinkProps}
      style={buildDocxTextStyle(inline.style, { includeBackground: true })}
    >
      {compressPunctuation
        ? renderCompressedPunctuation(inline.text)
        : inline.text}
    </span>
  );
}

export const DocxInlineContent = memo(DocxInlineContentComponent);
