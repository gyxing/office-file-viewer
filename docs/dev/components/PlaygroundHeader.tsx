import React from 'react';
import {
  getWebsiteContent,
  SITE_ROOT,
  type WebsiteLocale,
} from './home-content';
import { SiteHeaderShell } from './SiteHeaderShell';

type PlaygroundHeaderProps = {
  /** 当前在线体验页使用的语言。 */
  locale: WebsiteLocale;
};

// 页面导航始终停留在同一语言，避免切换栏目时意外改变语言环境。
const HOME_PATHS: Record<WebsiteLocale, string> = {
  'en-US': SITE_ROOT,
  'zh-CN': SITE_ROOT + 'zh-CN/',
};

const PLAYGROUND_PATHS: Record<WebsiteLocale, string> = {
  'en-US': SITE_ROOT + 'playground',
  'zh-CN': SITE_ROOT + 'zh-CN/playground',
};

/** 组合共享站点外壳与在线体验页专用导航。 */
export function PlaygroundHeader({ locale }: PlaygroundHeaderProps) {
  const content = getWebsiteContent(locale);
  const homeHref = HOME_PATHS[locale];
  const languageHref =
    locale === 'zh-CN' ? PLAYGROUND_PATHS['en-US'] : PLAYGROUND_PATHS['zh-CN'];
  const navigation = (
    <>
      <a href={homeHref + '#overview'}>{content.navigation.overview}</a>
      <a href={homeHref + '#highlights'}>{content.navigation.highlights}</a>
      <a href={homeHref + '#formats'}>{content.navigation.formats}</a>
      <a href={PLAYGROUND_PATHS[locale]} aria-current="page">
        {content.navigation.demo}
      </a>
      <a href={content.docsHref}>{content.navigation.docs}</a>
    </>
  );
  const mobileNavigation = (
    <>
      <a
        className="office-viewer-site-mobile-demo"
        href={PLAYGROUND_PATHS[locale]}
        aria-current="page"
      >
        {content.navigation.demo}
      </a>
      <a className="office-viewer-site-mobile-docs" href={content.docsHref}>
        {content.navigation.docs}
      </a>
    </>
  );

  return (
    <SiteHeaderShell
      content={content}
      languageHref={languageHref}
      navigation={navigation}
      mobileNavigation={mobileNavigation}
    />
  );
}
