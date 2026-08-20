import type { CSSProperties } from 'react';
import React from 'react';
import { UserIcon } from '../../shared/ui/OfficeIcons';
import type { WordMarkupCalloutLayout } from './wordMarkupCalloutLayout';

/** Word 页侧批注或修订标记属性。 */
type WordMarkupCalloutProps = {
  /** 当前标记的页面布局和源语义。 */
  callout: WordMarkupCalloutLayout;
  /** 修订类别的本地化文案。 */
  revisionKindLabel?: string;
  /** 源文件缺少作者时使用的文案。 */
  unknownAuthorLabel: string;
  /** 批注或修订缺少正文时使用的文案。 */
  emptyContentLabel: string;
  /** 当前标记是否由正文悬停或点击激活。 */
  emphasized: boolean;
  /** 用户选择页侧标记时触发。 */
  onSelect(): void;
};

/** 批注和修订在 Word 标记区使用的稳定语义色。 */
const WORD_MARKUP_COLORS = {
  comment: '#1677d2',
  revision: '#e5484d',
} as const;

/** 渲染 Word 原生风格的页侧标记及虚线连接。 */
export function WordMarkupCallout({
  callout,
  revisionKindLabel,
  unknownAuthorLabel,
  emptyContentLabel,
  emphasized,
  onSelect,
}: WordMarkupCalloutProps) {
  const color = WORD_MARKUP_COLORS[callout.type];
  const style = {
    top: callout.calloutTop,
    left: callout.calloutLeft,
    width: callout.width,
    height: callout.height,
    '--office-file-word-markup-color': color,
  } as CSSProperties;
  const connectorStyle = {
    '--office-file-word-markup-color': color,
  } as CSSProperties;
  const elbowX =
    callout.anchorX + (callout.calloutLeft - callout.anchorX) * 0.62;
  const body = callout.excerpt || emptyContentLabel;
  const author = callout.author || unknownAuthorLabel;
  const accessibleLabel =
    callout.type === 'comment'
      ? `${author} ${body}`
      : `${author} ${revisionKindLabel ?? ''} ${body}`.trim();

  return (
    <>
      <svg
        className={[
          'office-file-word-markup-connector',
          `office-file-word-markup-connector--${callout.type}`,
        ].join(' ')}
        data-office-word-markup-connector={callout.key}
        data-emphasized={emphasized ? 'true' : 'false'}
        style={connectorStyle}
        aria-hidden="true"
      >
        <path
          d={`M ${callout.anchorX} ${callout.anchorY} L ${elbowX} ${callout.anchorY} L ${callout.calloutLeft} ${callout.connectorY}`}
        />
      </svg>
      <button
        type="button"
        className={[
          'office-file-word-markup-callout',
          `office-file-word-markup-callout--${callout.type}`,
        ].join(' ')}
        aria-label={accessibleLabel}
        data-office-word-markup-callout={callout.key}
        data-office-word-comment-callout={
          callout.type === 'comment' ? callout.id : undefined
        }
        data-office-word-revision-callout={
          callout.type === 'revision' ? callout.id : undefined
        }
        data-emphasized={emphasized ? 'true' : 'false'}
        data-resolved={callout.resolved ? 'true' : 'false'}
        style={style}
        onClick={onSelect}
      >
        {callout.type === 'comment' ? (
          <span
            className="office-file-word-markup-callout__avatar"
            aria-hidden="true"
          >
            <UserIcon />
          </span>
        ) : null}
        <span className="office-file-word-markup-callout__content">
          <strong>{author}</strong>
          <span>
            {callout.type === 'revision' && revisionKindLabel
              ? `${revisionKindLabel}：${body}`
              : body}
          </span>
        </span>
      </button>
    </>
  );
}
