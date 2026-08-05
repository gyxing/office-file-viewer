import React from 'react';
import {
  GITHUB_URL,
  NPM_URL,
  SITE_ROOT,
  type WebsiteContent,
} from './home-content';

export type SiteHeaderShellProps = {
  /** 当前语言对应的站点文案。 */
  content: WebsiteContent;
  /** 当前页面的语言切换目标。 */
  languageHref: string;
  /** 桌面端主导航内容。 */
  navigation: React.ReactNode;
  /** 窄屏操作区优先展示的导航内容。 */
  mobileNavigation: React.ReactNode;
};

type HeaderActionIconProps = {
  /** 外部入口或语言切换对应的图标类型。 */
  name: 'npm' | 'github' | 'language';
};

/** 渲染站点外部入口与语言切换使用的统一矢量图标。 */
function HeaderActionIcon({ name }: HeaderActionIconProps) {
  if (name === 'npm') {
    return (
      <svg
        className="office-viewer-site-nav-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
        focusable="false"
      >
        <path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z" />
        <path d="m3.3 7 8.7 5 8.7-5" />
        <path d="M12 22V12" />
      </svg>
    );
  }

  if (name === 'github') {
    return (
      <svg
        className="office-viewer-site-nav-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
        focusable="false"
      >
        <path d="M15 22v-3.9c.04-1-.35-1.96-1.08-2.65 3.6-.4 7.38-1.77 7.38-8a6.24 6.24 0 0 0-1.67-4.34A5.8 5.8 0 0 0 19.47.8S18.16.38 15 2.46a15 15 0 0 0-8 0C3.84.38 2.53.8 2.53.8a5.8 5.8 0 0 0-.16 2.31A6.24 6.24 0 0 0 .7 7.45c0 6.22 3.78 7.6 7.38 8A4.8 4.8 0 0 0 7 18.1V22" />
        <path d="M7 19c-3 .92-3-1.5-4.2-2" />
      </svg>
    );
  }

  return (
    <svg
      className="office-viewer-site-nav-action-icon"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      <circle cx="12" cy="12" r="9.5" />
      <path d="M2.5 12h19" />
      <path d="M12 2.5c2.45 2.6 3.8 5.72 3.8 9.5S14.45 18.9 12 21.5C9.55 18.9 8.2 15.78 8.2 12S9.55 5.1 12 2.5Z" />
    </svg>
  );
}

/** 渲染首页与文档页共用的品牌导航外壳。 */
export function SiteHeaderShell({
  content,
  languageHref,
  navigation,
  mobileNavigation,
}: SiteHeaderShellProps) {
  return (
    <header className="office-viewer-site-header">
      <nav
        className="office-viewer-site-nav"
        aria-label={content.navigation.ariaLabel}
      >
        <a
          className="office-viewer-site-brand"
          href={SITE_ROOT}
          aria-label="Office File Viewer"
        >
          <span className="office-viewer-site-brand-mark" aria-hidden="true">
            <i />
            <i />
            <i />
          </span>
          <span>Office File Viewer</span>
        </a>

        <div className="office-viewer-site-nav-links">{navigation}</div>

        <div className="office-viewer-site-nav-actions">
          {mobileNavigation}
          <a
            className="office-viewer-site-npm-link"
            href={NPM_URL}
            target="_blank"
            rel="noreferrer"
          >
            <HeaderActionIcon name="npm" />
            {content.navigation.npm}
          </a>
          <a href={GITHUB_URL} target="_blank" rel="noreferrer">
            <HeaderActionIcon name="github" />
            {content.navigation.github}
          </a>
          <a className="office-viewer-site-language" href={languageHref}>
            <HeaderActionIcon name="language" />
            {content.languageLabel}
          </a>
        </div>
      </nav>
    </header>
  );
}
