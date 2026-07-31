// DocPageFrame 提供 DOC 降级预览的页面框架，负责页面尺寸、页边距和缩放。
import type { CSSProperties, ReactNode } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocImage, DocPage } from '../../services/doc/types';

/** DOC页面框架组件属性。 */
type DocPageFrameProps = {
  /** 当前关联的页面模型。 */
  page: DocPage;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** 当前组件包含的子节点。 */
  children: ReactNode;
  /** 当前物理页需要显示的页眉徽标。 */
  headerImage?: DocImage;
  /** 当前物理页需要显示的页脚页码文本。 */
  footerText?: string;
};

/** 渲染DOC页面框架。 */
function DocPageFrameComponent({
  page,
  zoom,
  children,
  headerImage,
  footerText,
}: DocPageFrameProps) {
  const scale = zoom / 100;
  // 外层占位使用缩放后的尺寸，内层 article 保留原始 Word 坐标系并用 transform 缩放。
  const shellStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width * scale,
      height: page.minHeight * scale,
    }),
    [page.minHeight, page.width, scale],
  );
  const articleStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width,
      minHeight: page.minHeight,
      padding: `${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px`,
      transform: `scale(${scale})`,
    }),
    [
      page.marginBottom,
      page.marginLeft,
      page.marginRight,
      page.marginTop,
      page.minHeight,
      page.width,
      scale,
    ],
  );

  return (
    <div className="office-file-doc-page-frame" style={shellStyle}>
      <article
        className="office-file-doc-page-frame__article"
        style={articleStyle}
      >
        {headerImage ? (
          <img
            className="office-file-doc-page-frame__header-image"
            src={headerImage.src}
            alt=""
            loading="lazy"
            decoding="async"
          />
        ) : null}
        {children}
        {footerText ? (
          <div className="office-file-doc-page-frame__footer">{footerText}</div>
        ) : null}
      </article>
    </div>
  );
}

export const DocPageFrame = memo(DocPageFrameComponent);
