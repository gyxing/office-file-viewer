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
  /** 是否启用东亚文字与西文之间的自动间距。 */
  autoSpaceLatin: boolean;
  /** 是否启用东亚文字与数字之间的自动间距。 */
  autoSpaceNumber: boolean;
  /** 当前文字运行之前紧邻的可见字符。 */
  previousTextCharacter?: string;
  /** 当前文字运行之后紧邻的可见字符。 */
  nextTextCharacter?: string;
};

/** 可按 OOXML 西文字体独立渲染的基础拉丁字符。 */
const DOCX_LATIN_CHARACTER_PATTERN = /[\u0000-\u024f]/;

/** Word 自动间距规则识别的东亚表意文字。 */
const DOCX_EAST_ASIA_CHARACTER_PATTERN =
  /[\u3040-\u30ff\u3100-\u312f\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff]/;

/** Word 自动间距规则识别的西文字母。 */
const DOCX_LATIN_LETTER_PATTERN = /[A-Za-z\u00c0-\u024f]/;

/** Word 自动间距规则识别的半角数字。 */
const DOCX_NUMBER_PATTERN = /[0-9]/;

/** 判断相邻字符之间是否需要补足 Word 的四分之一字宽。 */
function needsDocxAutoSpacing(
  leftCharacter: string | undefined,
  rightCharacter: string | undefined,
  options: Pick<DocxTextSpacingOptions, 'autoSpaceLatin' | 'autoSpaceNumber'>,
) {
  if (!leftCharacter || !rightCharacter) return false;
  const leftIsEastAsia = DOCX_EAST_ASIA_CHARACTER_PATTERN.test(leftCharacter);
  const rightIsEastAsia = DOCX_EAST_ASIA_CHARACTER_PATTERN.test(rightCharacter);
  if (leftIsEastAsia === rightIsEastAsia) return false;
  const westernCharacter = leftIsEastAsia ? rightCharacter : leftCharacter;
  return (
    (options.autoSpaceLatin &&
      DOCX_LATIN_LETTER_PATTERN.test(westernCharacter)) ||
    (options.autoSpaceNumber && DOCX_NUMBER_PATTERN.test(westernCharacter))
  );
}

/** 检查当前文字运行及其相邻运行是否存在需要自动留白的脚本边界。 */
function hasDocxAutoSpacingBoundary(
  text: string,
  options: Pick<
    DocxTextSpacingOptions,
    | 'autoSpaceLatin'
    | 'autoSpaceNumber'
    | 'previousTextCharacter'
    | 'nextTextCharacter'
  >,
) {
  const characters = Array.from(text);
  if (!characters.length) return false;
  if (
    needsDocxAutoSpacing(options.previousTextCharacter, characters[0], options)
  ) {
    return true;
  }
  for (let index = 1; index < characters.length; index += 1) {
    if (
      needsDocxAutoSpacing(characters[index - 1], characters[index], options)
    ) {
      return true;
    }
  }
  return needsDocxAutoSpacing(
    characters[characters.length - 1],
    options.nextTextCharacter,
    options,
  );
}

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
    const previousCharacter =
      index > 0 ? characters[index - 1] : options.previousTextCharacter;
    if (needsDocxAutoSpacing(previousCharacter, character, options)) {
      flushPlainText();
      nodes.push(
        <span
          key={`auto-spacing-${index}`}
          className="office-file-docx-auto-spacing"
          aria-hidden="true"
        >
          {' '}
        </span>,
      );
    }

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
    const nextCharacter =
      index + 1 < characters.length
        ? characters[index + 1]
        : options.nextTextCharacter;
    if (
      character === '+' &&
      DOCX_EAST_ASIA_CHARACTER_PATTERN.test(nextCharacter ?? '')
    ) {
      // Word 允许西文运算符后的东亚文字另起一行，浏览器默认禁则会额外撑高窄表格。
      plainText += character;
      flushPlainText();
      nodes.push(<wbr key={`word-break-${index}`} />);
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
  if (
    needsDocxAutoSpacing(
      characters[characters.length - 1],
      options.nextTextCharacter,
      options,
    )
  ) {
    nodes.push(
      <span
        key="auto-spacing-end"
        className="office-file-docx-auto-spacing"
        aria-hidden="true"
      >
        {' '}
      </span>,
    );
  }
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
  /** 是否启用东亚文字与西文之间的自动间距。 */
  autoSpaceLatin: boolean;
  /** 是否启用东亚文字与数字之间的自动间距。 */
  autoSpaceNumber: boolean;
  /** 当前文字运行之前紧邻的可见字符。 */
  previousTextCharacter?: string;
  /** 当前文字运行之后紧邻的可见字符。 */
  nextTextCharacter?: string;
};

/** 渲染 DOCX 段落中的行内文字、图片和图表。 */
function DocxInlineContentComponent({
  inline,
  sourceId,
  searchBlockId,
  resolveFontFamily,
  isParagraphEnd,
  autoSpaceLatin,
  autoSpaceNumber,
  previousTextCharacter,
  nextTextCharacter,
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
        // 西文字体不继承宋体的小字号字宽补偿。
        letterSpacing: 0,
      }
    : undefined;
  const hasAutoSpacing = hasDocxAutoSpacingBoundary(inline.text, {
    autoSpaceLatin,
    autoSpaceNumber,
    previousTextCharacter,
    nextTextCharacter,
  });
  const isNumberingSpacer =
    inline.advanceWidth !== undefined && inline.text.length === 0;
  const renderText =
    !isNumberingSpacer &&
    (compressPunctuation ||
      preserveEastAsiaSpaces ||
      latinTextStyle ||
      hasAutoSpacing)
      ? (text: string, startOffset: number, endOffset: number) => {
          const precedingCharacters =
            startOffset > 0
              ? Array.from(inline.text.slice(0, startOffset))
              : undefined;
          const followingCharacters =
            endOffset < inline.text.length
              ? Array.from(inline.text.slice(endOffset))
              : undefined;
          return renderOfficeTextSpacing(text, {
            compressPunctuation,
            preserveEastAsiaSpaces,
            latinTextStyle,
            allowLineEndHanging:
              isParagraphEnd && endOffset === inline.text.length,
            autoSpaceLatin,
            autoSpaceNumber,
            previousTextCharacter:
              precedingCharacters?.[precedingCharacters.length - 1] ??
              previousTextCharacter,
            nextTextCharacter: followingCharacters?.[0] ?? nextTextCharacter,
          });
        }
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
        // 编号容器会建立独立行盒，必须清除继承的负首行缩进，避免编号被二次左移。
        textIndent: inline.advanceWidth !== undefined ? 0 : undefined,
        // 悬挂编号只负责预留横向前进宽度；零高行盒避免编号标签把 24px 正文行距向上取整为 25px。
        height: inline.advanceWidth !== undefined ? 0 : undefined,
        lineHeight: inline.advanceWidth !== undefined ? 0 : undefined,
        position: inline.advanceWidth !== undefined ? 'relative' : undefined,
        // Word 的悬挂编号基线略低于同段正文，使用相对位移还原且不参与行盒计算。
        top: inline.advanceWidth !== undefined ? '0.12em' : undefined,
      }}
    >
      {isNumberingSpacer ? null : (
        <OfficeSearchHighlightedText
          text={inline.text}
          target={{ kind: 'word', blockId: searchBlockId }}
          renderText={renderText}
        />
      )}
    </span>
  );
}

export const DocxInlineContent = memo(DocxInlineContentComponent);
