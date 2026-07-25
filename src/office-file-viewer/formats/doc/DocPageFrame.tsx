// DocPageFrame 提供 DOC 降级预览的页面框架，负责页面尺寸、页边距和缩放。
import type { CSSProperties, ReactNode } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocPage } from '../../services/doc/types';

/** 定义 DocPageFrame 组件可接收的属性。 */
type DocPageFrameProps = {
  /** DocPageFrameProps 当前关联的页面模型。 */
  page: DocPage;
  /** 当前预览缩放比例。 */
  zoom: number;
  /** DocPageFrameProps 包含并负责布局的 React 子节点。 */
  children: ReactNode;
};

/** 渲染 DocPageFrameComponent 组件。 */
function DocPageFrameComponent({ page, zoom, children }: DocPageFrameProps) {
  const scale = zoom / 100;
  // 外层占位使用缩放后的尺寸，内层 article 保留原始 Word 坐标系并用 transform 缩放。
  const shellStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width * scale,
      minHeight: page.minHeight * scale,
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
        {children}
      </article>
    </div>
  );
}

export const DocPageFrame = memo(DocPageFrameComponent);
