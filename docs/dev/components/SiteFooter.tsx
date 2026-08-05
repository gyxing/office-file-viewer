import React from 'react';
import { SITE_ROOT, type WebsiteContent } from './home-content';

type SiteFooterProps = {
  /** 当前语言对应的站点文案。 */
  content: WebsiteContent;
};

/** 渲染首页与文档页共用的站点品牌页脚。 */
export function SiteFooter({ content }: SiteFooterProps) {
  return (
    <footer className="office-viewer-site-footer">
      <a className="office-viewer-site-brand" href={SITE_ROOT}>
        <span className="office-viewer-site-brand-mark" aria-hidden="true">
          <i />
          <i />
          <i />
        </span>
        <span>Office File Viewer</span>
      </a>
      <p>{content.footer.description}</p>
      <div>
        <span>{content.footer.license}</span>
        <span>{content.footer.builtWith}</span>
      </div>
    </footer>
  );
}
