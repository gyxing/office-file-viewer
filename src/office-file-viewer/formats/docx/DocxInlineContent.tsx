// DocxInlineContent 渲染 DOCX 段落内的文本、换行、图片、图表和形状。
import React, { createContext, memo, useContext } from 'react';
import type { DocxInline } from '../../services/docx/types';
import type { OfficeFontFamilyResolver } from '../../services/fonts/types';
import { useOfficeHyperlink } from '../../shared/hyperlink';
import { OfficeSearchHighlightedText } from '../search/OfficeSearchContext';
import { DocxImage } from './DocxImage';
import { DocxInlineChart } from './DocxInlineChart';
import { DocxShape } from './DocxShape';
import {
  buildDocxTextStyle,
  resolveDocxLatinFontFamily,
} from './docxRenderUtils';

/** 控制当前 DOCX 是否按源设置压缩东亚标点。 */
export const DocxCharacterSpacingContext = createContext(false);

/** 需要按东亚排版规则压缩的全部标点。 */
const COMPRESSIBLE_PUNCTUATION = new Set(
  Array.from('，。；：！？、（）【】《》“”‘’'),
);

/** Word 禁止出现在行首的东亚标点。 */
const CLOSING_PUNCTUATION = new Set(Array.from('，。；：！？、）】》”’'));

/** Word 禁止出现在行尾的东亚标点。 */
const OPENING_PUNCTUATION = new Set(Array.from('（【《“‘'));

/** 渲染可按 Word 东亚规则调整占位的标点。 */
function renderCompressedPunctuationCharacter(
  character: string,
  key: string,
  isTextEnd = false,
): React.ReactElement {
  const className = [
    'office-file-docx-compressed-punctuation',
    character === '《'
      ? 'office-file-docx-compressed-punctuation--book-title-opening'
      : '',
    isTextEnd && CLOSING_PUNCTUATION.has(character)
      ? 'office-file-docx-compressed-punctuation--line-end'
      : '',
  ]
    .filter(Boolean)
    .join(' ');
  return (
    <span key={key} className={className}>
      {character}
    </span>
  );
}

/** Word 东亚字体提示下一个空格占用的目标字宽。 */
const DOCX_EAST_ASIA_SPACE_WIDTH_EM = 0.5;

/** DOCX 文本间距渲染选项。 */
type DocxTextSpacingOptions = {
  /** 是否压缩东亚标点。 */
  compressPunctuation: boolean;
  /** 是否按东亚字体度量保留空格。 */
  preserveEastAsiaSpaces: boolean;
  /** 西文字符使用的独立字体样式。 */
  latinTextStyle?: React.CSSProperties;
  /** 当前运行末尾闭标点是否可以作为段落行尾悬挂。 */
  allowLineEndHanging: boolean;
};

/** 可按 OOXML 西文字体独立渲染的基础拉丁字符。 */
const DOCX_LATIN_CHARACTER_PATTERN = /[\u0000-\u024f]/;

/** 按字符脚本拆分同一 OOXML 运行，避免中文字体覆盖西文字体。 */
function renderScriptAwareText(
  text: string,
  keyPrefix: string,
  latinTextStyle?: React.CSSProperties,
) {
  if (!latinTextStyle) return [text];
  const segments: Array<{ latin: boolean; text: string }> = [];
  Array.from(text).forEach((character) => {
    const latin = DOCX_LATIN_CHARACTER_PATTERN.test(character);
    const last = segments[segments.length - 1];
    if (last?.latin === latin) {
      last.text += character;
    } else {
      segments.push({ latin, text: character });
    }
  });
  return segments.map((segment, index) =>
    segment.latin ? (
      <span key={keyPrefix + '-latin-' + index} style={latinTextStyle}>
        {segment.text}
      </span>
    ) : (
      <React.Fragment key={keyPrefix + '-east-asia-' + index}>
        {segment.text}
      </React.Fragment>
    ),
  );
}

/** 还原东亚标点禁则和源文档保留空格。 */
function renderOfficeTextSpacing(
  text: string,
  options: DocxTextSpacingOptions,
): React.ReactNode[] {
  const characters = Array.from(text);
  const nodes: React.ReactNode[] = [];
  let plainText = '';

  const flushPlainText = () => {
    if (!plainText) return;
    nodes.push(
      ...renderScriptAwareText(
        plainText,
        'plain-' + nodes.length,
        options.latinTextStyle,
      ),
    );
    plainText = '';
  };

  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];

    if (options.preserveEastAsiaSpaces && character === ' ') {
      let spaceEnd = index + 1;
      while (spaceEnd < characters.length && characters[spaceEnd] === ' ') {
        spaceEnd += 1;
      }
      const spaceCount = spaceEnd - index;
      const wordSpacing =
        Math.max(0, spaceCount - 1) * DOCX_EAST_ASIA_SPACE_WIDTH_EM;

      flushPlainText();
      nodes.push(
        <span
          key={`east-asia-space-${index}`}
          className="office-file-docx-east-asia-space"
          data-office-docx-space-count={spaceCount}
          style={{ wordSpacing: wordSpacing + 'em' }}
        >
          {' '}
        </span>,
      );
      index = spaceEnd - 1;
      continue;
    }
    if (!options.compressPunctuation) {
      plainText += character;
      continue;
    }

    let cursor = index + 1;

    while (
      cursor < characters.length &&
      CLOSING_PUNCTUATION.has(characters[cursor])
    ) {
      cursor += 1;
    }
    if (!COMPRESSIBLE_PUNCTUATION.has(character) && cursor > index + 1) {
      flushPlainText();
      nodes.push(
        <span
          key={`closing-punctuation-${index}`}
          className="office-file-docx-punctuation-no-break"
        >
          {renderScriptAwareText(
            character,
            'closing-base-' + index,
            options.latinTextStyle,
          )}
          {characters
            .slice(index + 1, cursor)
            .map((punctuation, offset) =>
              renderCompressedPunctuationCharacter(
                punctuation,
                `closing-${index}-${offset}`,
                options.allowLineEndHanging &&
                  index + offset + 1 === characters.length - 1,
              ),
            )}
        </span>,
      );
      index = cursor - 1;
      continue;
    }

    if (OPENING_PUNCTUATION.has(character)) {
      while (
        cursor < characters.length &&
        OPENING_PUNCTUATION.has(characters[cursor])
      ) {
        cursor += 1;
      }
      if (cursor < characters.length) {
        flushPlainText();
        nodes.push(
          <span
            key={`opening-punctuation-${index}`}
            className="office-file-docx-punctuation-no-break"
          >
            {characters
              .slice(index, cursor)
              .map((punctuation, offset) =>
                renderCompressedPunctuationCharacter(
                  punctuation,
                  `opening-${index}-${offset}`,
                ),
              )}
            {renderScriptAwareText(
              characters[cursor],
              'opening-base-' + index,
              options.latinTextStyle,
            )}
          </span>,
        );
        index = cursor;
        continue;
      }
    }

    if (COMPRESSIBLE_PUNCTUATION.has(character)) {
      flushPlainText();
      nodes.push(
        renderCompressedPunctuationCharacter(
          character,
          `punctuation-${index}`,
          options.allowLineEndHanging && index === characters.length - 1,
        ),
      );
    } else {
      plainText += character;
    }
  }

  flushPlainText();
  return nodes;
}
/** DOCX 行内内容组件属性。 */
type DocxInlineContentProps = {
  /** 当前负责渲染的行内内容模型。 */
  inline: DocxInline;
  /** 当前行内内容在所属文档中的稳定标识。 */
  sourceId: string;
  /** 查找结果对应的顶层正文块标识。 */
  searchBlockId: string;
  /** 当前文档会话统一的字体链解析函数。 */
  resolveFontFamily: OfficeFontFamilyResolver;
  /** 当前行内内容是否为段落最后一个文字运行。 */
  isParagraphEnd: boolean;
};

/** 渲染 DOCX 段落中的行内文字、图片和图表。 */
function DocxInlineContentComponent({
  inline,
  sourceId,
  searchBlockId,
  resolveFontFamily,
  isParagraphEnd,
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
        style={buildDocxTextStyle(inline.style, undefined, resolveFontFamily)}
        aria-hidden="true"
      />
    );
  if (inline.type === 'image') return <DocxImage inline={inline} />;
  if (inline.type === 'chart') return <DocxInlineChart inline={inline} />;
  if (inline.type === 'shape') {
    return <DocxShape inline={inline} searchBlockId={searchBlockId} />;
  }
  const preserveEastAsiaSpaces = Boolean(
    inline.preserveSpace && inline.style?.fontHint === 'eastAsia',
  );
  const latinFontFamily = resolveDocxLatinFontFamily(
    inline.style,
    resolveFontFamily,
  );
  const latinTextStyle: React.CSSProperties | undefined = latinFontFamily
    ? {
        fontFamily: latinFontFamily,
        fontWeight:
          inline.style?.bold === true
            ? 700
            : inline.style?.bold === false
            ? 400
            : undefined,
        // 西文字体不继承宋体的描边粗体和小字号字宽补偿。
        WebkitTextStroke: 0,
        letterSpacing: 0,
      }
    : undefined;
  const renderText =
    compressPunctuation || preserveEastAsiaSpaces || latinTextStyle
      ? (text: string) =>
          renderOfficeTextSpacing(text, {
            compressPunctuation,
            preserveEastAsiaSpaces,
            latinTextStyle,
            allowLineEndHanging: isParagraphEnd,
          })
      : undefined;

  return (
    <span
      {...hyperlinkProps}
      style={{
        ...buildDocxTextStyle(
          inline.style,
          { includeBackground: true },
          resolveFontFamily,
        ),
        // 东亚连续空格已拆为可换行占位；其余 xml:space 文本才需要浏览器保留空白。
        whiteSpace:
          inline.advanceWidth !== undefined
            ? 'nowrap'
            : inline.preserveSpace && !preserveEastAsiaSpaces
            ? 'pre-wrap'
            : undefined,
        display: inline.advanceWidth !== undefined ? 'inline-block' : undefined,
        width: inline.advanceWidth,
      }}
    >
      <OfficeSearchHighlightedText
        text={inline.text}
        target={{ kind: 'word', blockId: searchBlockId }}
        renderText={renderText}
      />
    </span>
  );
}

export const DocxInlineContent = memo(DocxInlineContentComponent);
