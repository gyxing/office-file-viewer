// DocxPageFrame 提供 DOCX 页面框架，负责页宽、页高、页边距、边框和缩放。
import type { CSSProperties, ReactNode } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxPage } from '../../services/docx/types';

/** DOCX页面框架组件属性。 */
type DocxPageFrameProps = {
  /** 当前关联的页面模型。 */
  page: DocxPage;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前组件包含的子节点。 */
  children: ReactNode;
  /** 当前页面的页眉内容。 */
  header?: ReactNode;
  /** 当前页面的页脚内容。 */
  footer?: ReactNode;
};

/** 渲染DOCX页面框架。 */
function DocxPageFrameComponent({
  page,
  zoom,
  children,
  header,
  footer,
}: DocxPageFrameProps) {
  const scale = zoom / 100;
  // DOCX 的边框和页边距来自文档本身，放在 article 上才能随页面坐标系一起缩放。
  const shellStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width * scale,
      height: page.minHeight * scale,
    }),
    [page.minHeight, page.width, scale],
  );
  const articleStyle = useMemo<CSSProperties>(
    () =>
      ({
        position: 'relative',
        '--office-file-docx-page-width': `${page.width}px`,
        '--office-file-docx-page-height': `${page.minHeight}px`,
        '--office-file-docx-page-margin-top': `${page.marginTop}px`,
        '--office-file-docx-page-margin-right': `${page.marginRight}px`,
        '--office-file-docx-page-margin-bottom': `${page.marginBottom}px`,
        '--office-file-docx-page-margin-left': `${page.marginLeft}px`,
        width: page.width,
        minHeight: page.minHeight,
        padding: `${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px`,
        borderTop: page.borderTop,
        borderRight: page.borderRight,
        borderBottom: page.borderBottom,
        borderLeft: page.borderLeft,
        transform: `scale(${scale})`,
      } as CSSProperties),
    [
      page.borderBottom,
      page.borderLeft,
      page.borderRight,
      page.borderTop,
      page.marginBottom,
      page.marginLeft,
      page.marginRight,
      page.marginTop,
      page.minHeight,
      page.width,
      scale,
    ],
  );
  const regionStyle = useMemo(
    () => ({
      left: page.marginLeft,
      width: page.width - page.marginLeft - page.marginRight,
    }),
    [page.marginLeft, page.marginRight, page.width],
  );

  return (
    <div className="office-file-docx-page-frame" style={shellStyle}>
      <article
        className="office-file-docx-page-frame__article"
        style={articleStyle}
      >
        {header ? (
          <div
            className="office-file-docx-page-frame__header"
            style={{
              ...regionStyle,
              top: page.headerDistance ?? page.marginTop / 2,
            }}
          >
            {header}
          </div>
        ) : null}
        {children}
        {footer ? (
          <div
            className="office-file-docx-page-frame__footer"
            style={{
              ...regionStyle,
              bottom: page.footerDistance ?? page.marginBottom / 2,
            }}
          >
            {footer}
          </div>
        ) : null}
      </article>
    </div>
  );
}

export const DocxPageFrame = memo(DocxPageFrameComponent);
