import React from 'react';
import { type WebsiteContent } from './home-content';
import { SiteHeaderShell } from './SiteHeaderShell';

export type HomeSectionId = 'overview' | 'demo' | 'highlights' | 'formats';

type HomeHeaderProps = {
  /** 当前语言对应的站点文案。 */
  content: WebsiteContent;
  /** 当前进入阅读区域的首页区块。 */
  activeSection: HomeSectionId;
  /** 首页锚点导航的点击处理函数。 */
  onSectionLink: (event: React.MouseEvent<HTMLAnchorElement>) => void;
};

// 导航顺序与页面阅读顺序保持一致，供当前区块状态和锚点复用。
const NAVIGATION_ITEMS: Array<{
  id: HomeSectionId;
  label: keyof Pick<
    WebsiteContent['navigation'],
    'overview' | 'demo' | 'highlights' | 'formats'
  >;
}> = [
  { id: 'overview', label: 'overview' },
  { id: 'demo', label: 'demo' },
  { id: 'highlights', label: 'highlights' },
  { id: 'formats', label: 'formats' },
];

/** 组合共享站点外壳与首页区块导航。 */
export function HomeHeader({
  content,
  activeSection,
  onSectionLink,
}: HomeHeaderProps) {
  const navigation = (
    <>
      {NAVIGATION_ITEMS.map((item) => (
        <a
          key={item.id}
          href={'#' + item.id}
          aria-current={
            activeSection === item.id ? ('location' as const) : undefined
          }
          onClick={onSectionLink}
        >
          {content.navigation[item.label]}
        </a>
      ))}
      <a href={content.docsHref}>{content.navigation.docs}</a>
    </>
  );
  const mobileNavigation = (
    <>
      <a className="office-viewer-site-mobile-docs" href={content.docsHref}>
        {content.navigation.docs}
      </a>
      <a
        className="office-viewer-site-mobile-demo"
        href="#demo"
        onClick={onSectionLink}
      >
        {content.navigation.demo}
      </a>
    </>
  );

  return (
    <SiteHeaderShell
      content={content}
      languageHref={content.languageHref}
      navigation={navigation}
      mobileNavigation={mobileNavigation}
    />
  );
}
