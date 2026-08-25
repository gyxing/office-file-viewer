import React from 'react';
import {
  getWebsiteContent,
  SITE_ROOT,
  type WebsiteLocale,
} from './home-content';
import { SiteHeaderShell } from './SiteHeaderShell';

type DocumentationHeaderProps = {
  /** 当前文档页使用的语言。 */
  locale: WebsiteLocale;
};

// 语言切换必须停留在文档页面，不能回退到另一语言首页。
const DOCUMENTATION_PATHS: Record<WebsiteLocale, string> = {
  'en-US': SITE_ROOT + 'docs',
  'zh-CN': SITE_ROOT + 'zh-CN/docs',
};

// 文档导航中的产品区块统一返回同语言首页。
const HOME_PATHS: Record<WebsiteLocale, string> = {
  'en-US': SITE_ROOT,
  'zh-CN': SITE_ROOT + 'zh-CN/',
};

/** 组合共享站点外壳与文档页专用导航。 */
export function DocumentationHeader({ locale }: DocumentationHeaderProps) {
  const content = getWebsiteContent(locale);
  const homeHref = HOME_PATHS[locale];
  const languageHref =
    locale === 'zh-CN'
      ? DOCUMENTATION_PATHS['en-US']
      : DOCUMENTATION_PATHS['zh-CN'];
  const navigation = (
    <>
      <a href={homeHref + '#overview'}>{content.navigation.overview}</a>
      <a href={homeHref + '#highlights'}>{content.navigation.highlights}</a>
      <a href={homeHref + '#formats'}>{content.navigation.formats}</a>
      <a href={content.playgroundHref}>{content.navigation.demo}</a>
      <a href={DOCUMENTATION_PATHS[locale]} aria-current="page">
        {content.navigation.docs}
      </a>
    </>
  );
  const mobileNavigation = (
    <>
      <a
        className="office-viewer-site-mobile-demo"
        href={content.playgroundHref}
      >
        {content.navigation.demo}
      </a>
      <a
        className="office-viewer-site-mobile-docs"
        href={DOCUMENTATION_PATHS[locale]}
        aria-current="page"
      >
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
