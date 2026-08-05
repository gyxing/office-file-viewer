import { ConfigProvider } from 'antd';
import enUS from 'antd/locale/en_US';
import zhCN from 'antd/locale/zh_CN';
import { OfficeFileViewer } from 'office-file-viewer';
import React, { useEffect, useRef, useState } from 'react';
import {
  getWebsiteContent,
  GITHUB_URL,
  NPM_URL,
  type PackageManager,
  type WebsiteLocale,
} from './home-content';
import { HomeHeader, type HomeSectionId } from './HomeHeader';
import { HomeHero } from './HomeHero';
import {
  ProductHighlightsSection,
  ProductOverviewSection,
  SupportedFormatsSection,
} from './HomeProductSections';
import './OfficeViewerHomePage.less';
import { SiteFooter } from './SiteFooter';

type OfficeViewerHomePageProps = {
  /** 当前公开官网使用的语言。 */
  locale: WebsiteLocale;
};

type CopyState = 'idle' | 'copied' | 'failed';

// 包管理器顺序同时决定标签页展示顺序与左右方向键行为。
const PACKAGE_MANAGERS: PackageManager[] = ['npm', 'yarn'];
// 仅观察主导航可到达的区块，避免短区块频繁抢占当前状态。
const HOME_SECTION_IDS: HomeSectionId[] = [
  'overview',
  'demo',
  'highlights',
  'formats',
];

/** 优先使用 Clipboard API，并为权限受限或较旧的浏览器提供回退。 */
async function writeTextToClipboard(value: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }

  const textarea = document.createElement('textarea');
  textarea.value = value;
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();

  const copied = document.execCommand('copy');
  textarea.remove();
  if (!copied) {
    throw new Error('Clipboard API is unavailable');
  }
}

/** 管理复制结果反馈，并在组件卸载时清理延迟任务。 */
function useCopyFeedback() {
  const [state, setState] = useState<CopyState>('idle');
  const resetTimerRef = useRef<number>();

  useEffect(
    () => () => {
      if (resetTimerRef.current !== undefined) {
        window.clearTimeout(resetTimerRef.current);
      }
    },
    [],
  );

  const copy = async (value: string) => {
    if (resetTimerRef.current !== undefined) {
      window.clearTimeout(resetTimerRef.current);
    }

    try {
      await writeTextToClipboard(value);
      setState('copied');
    } catch {
      setState('failed');
    }

    resetTimerRef.current = window.setTimeout(() => setState('idle'), 2400);
  };

  return { copy, state };
}

/** 根据复制状态选择当前操作文案。 */
function getCopyLabel(
  state: CopyState,
  labels: { idle: string; copied: string; failed: string },
) {
  if (state === 'copied') {
    return labels.copied;
  }
  if (state === 'failed') {
    return labels.failed;
  }
  return labels.idle;
}

/** 在尊重减少动态效果偏好的前提下，为站内锚点提供平滑滚动。 */
function handleSectionLink(event: React.MouseEvent<HTMLAnchorElement>) {
  if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
    return;
  }

  const target = document.querySelector(event.currentTarget.hash);
  if (!target) {
    return;
  }

  event.preventDefault();
  target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  window.history.pushState(null, '', event.currentTarget.hash);
}

/** 根据页面中最接近视口阅读区域的区块更新导航状态。 */
function useActiveSection(): HomeSectionId {
  const [activeSection, setActiveSection] = useState<HomeSectionId>('overview');

  useEffect(() => {
    const sections = HOME_SECTION_IDS.map((id) =>
      document.getElementById(id),
    ).filter((section): section is HTMLElement => section !== null);

    const observer = new IntersectionObserver(
      (entries) => {
        const activeEntry = entries
          .filter((entry) => entry.isIntersecting)
          .sort(
            (left, right) =>
              Math.abs(left.boundingClientRect.top) -
              Math.abs(right.boundingClientRect.top),
          )[0];

        if (activeEntry) {
          setActiveSection(activeEntry.target.id as HomeSectionId);
        }
      },
      {
        rootMargin: '-20% 0px -62%',
        threshold: [0.01, 0.2, 0.5],
      },
    );

    sections.forEach((section) => observer.observe(section));
    return () => observer.disconnect();
  }, []);

  return activeSection;
}

/** 渲染默认英文、可切换中文的 Office File Viewer 产品官网。 */
export function OfficeViewerHomePage({ locale }: OfficeViewerHomePageProps) {
  const content = getWebsiteContent(locale);
  const activeSection = useActiveSection();
  const [packageManager, setPackageManager] = useState<PackageManager>('npm');
  const heroCopy = useCopyFeedback();
  const developerCopy = useCopyFeedback();
  const installCommand = content.developer.commands[packageManager];
  const antdLocale = locale === 'zh-CN' ? zhCN : enUS;

  useEffect(() => {
    // Dumi 仍会生成文档外壳，用页面标记将官网样式限制在两个公开首页。
    document.body.classList.add('office-viewer-site-page');
    return () => document.body.classList.remove('office-viewer-site-page');
  }, []);

  const handlePackageManagerKeyDown = (
    event: React.KeyboardEvent<HTMLButtonElement>,
    currentManager: PackageManager,
  ) => {
    const currentIndex = PACKAGE_MANAGERS.indexOf(currentManager);
    let nextManager: PackageManager | undefined;

    if (event.key === 'ArrowRight') {
      nextManager =
        PACKAGE_MANAGERS[(currentIndex + 1) % PACKAGE_MANAGERS.length];
    } else if (event.key === 'ArrowLeft') {
      nextManager =
        PACKAGE_MANAGERS[
          (currentIndex - 1 + PACKAGE_MANAGERS.length) % PACKAGE_MANAGERS.length
        ];
    } else if (event.key === 'Home') {
      nextManager = PACKAGE_MANAGERS[0];
    } else if (event.key === 'End') {
      nextManager = PACKAGE_MANAGERS[PACKAGE_MANAGERS.length - 1];
    }

    if (!nextManager) {
      return;
    }

    event.preventDefault();
    setPackageManager(nextManager);
    window.requestAnimationFrame(() =>
      document.getElementById(`package-tab-${nextManager}`)?.focus(),
    );
  };

  return (
    <div className="office-viewer-site" lang={content.htmlLang}>
      <a className="office-viewer-site-skip" href="#main-content">
        {content.skipToContent}
      </a>

      <HomeHeader
        content={content}
        activeSection={activeSection}
        onSectionLink={handleSectionLink}
      />

      <main id="main-content">
        <HomeHero
          content={content}
          copyLabel={getCopyLabel(heroCopy.state, {
            idle: content.hero.copy,
            copied: content.hero.copied,
            failed: content.hero.copyFailed,
          })}
          onCopy={() => heroCopy.copy(content.developer.commands.npm)}
          onSectionLink={handleSectionLink}
        />

        <ProductOverviewSection content={content} />

        <section
          className="office-viewer-site-section office-viewer-site-demo"
          id="demo"
          aria-labelledby="demo-title"
        >
          <div className="office-viewer-site-section-heading is-split">
            <div>
              <span className="office-viewer-site-eyebrow">
                {content.demo.eyebrow}
              </span>
              <h2 id="demo-title">{content.demo.title}</h2>
              <p>{content.demo.description}</p>
            </div>
            <div className="office-viewer-site-local-note">
              <strong>{content.demo.privacyTitle}</strong>
              <span>{content.demo.privacyDescription}</span>
            </div>
          </div>
          <div className="office-viewer-site-demo-frame">
            <ConfigProvider locale={antdLocale}>
              <OfficeFileViewer locale={locale} height="100%" />
            </ConfigProvider>
          </div>
        </section>

        <ProductHighlightsSection content={content} />
        <SupportedFormatsSection content={content} />

        <section
          className="office-viewer-site-section"
          aria-labelledby="developer-title"
        >
          <div className="office-viewer-site-section-heading">
            <span className="office-viewer-site-eyebrow">
              {content.developer.eyebrow}
            </span>
            <h2 id="developer-title">{content.developer.title}</h2>
            <p>{content.developer.description}</p>
          </div>

          <div className="office-viewer-site-code-grid">
            <article className="office-viewer-site-code-card">
              <div className="office-viewer-site-code-header">
                <strong>{content.developer.installLabel}</strong>
                <div
                  role="tablist"
                  aria-label={content.developer.packageManagerLabel}
                >
                  {PACKAGE_MANAGERS.map((manager) => (
                    <button
                      key={manager}
                      id={`package-tab-${manager}`}
                      type="button"
                      role="tab"
                      tabIndex={packageManager === manager ? 0 : -1}
                      aria-selected={packageManager === manager}
                      aria-controls="install-command-panel"
                      onClick={() => setPackageManager(manager)}
                      onKeyDown={(event) =>
                        handlePackageManagerKeyDown(event, manager)
                      }
                    >
                      {manager}
                    </button>
                  ))}
                </div>
              </div>

              <div
                className="office-viewer-site-code-body"
                id="install-command-panel"
                role="tabpanel"
                aria-labelledby={`package-tab-${packageManager}`}
              >
                <code>{installCommand}</code>
                <button
                  type="button"
                  onClick={() => developerCopy.copy(installCommand)}
                  aria-live="polite"
                >
                  {getCopyLabel(developerCopy.state, {
                    idle: content.developer.copy,
                    copied: content.developer.copied,
                    failed: content.developer.copyFailed,
                  })}
                </button>
              </div>

              <div className="office-viewer-site-code-links">
                <a href={GITHUB_URL} target="_blank" rel="noreferrer">
                  {content.developer.github}
                </a>
                <a href={NPM_URL} target="_blank" rel="noreferrer">
                  {content.developer.npm}
                </a>
              </div>
            </article>

            <article className="office-viewer-site-code-card is-example">
              <div className="office-viewer-site-code-header">
                <strong>{content.developer.exampleLabel}</strong>
                <span>TSX</span>
              </div>
              <pre>
                <code>{content.developer.example}</code>
              </pre>
            </article>
          </div>
        </section>

        <section
          className="office-viewer-site-section"
          id="compatibility"
          aria-labelledby="compatibility-title"
        >
          <div className="office-viewer-site-section-heading is-split">
            <div>
              <span className="office-viewer-site-eyebrow">
                {content.compatibility.eyebrow}
              </span>
              <h2 id="compatibility-title">{content.compatibility.title}</h2>
            </div>
            <p>{content.compatibility.description}</p>
          </div>

          <div className="office-viewer-site-compatibility-grid">
            {content.compatibility.items.map((item) => (
              <article key={item.name}>
                <span>{item.name}</span>
                <strong>{item.version}</strong>
                <small>{item.note}</small>
              </article>
            ))}
          </div>
        </section>

        <section
          className="office-viewer-site-final-cta"
          aria-labelledby="final-cta-title"
        >
          <div>
            <h2 id="final-cta-title">{content.finalCta.title}</h2>
            <p>{content.finalCta.description}</p>
          </div>
          <div>
            <a
              className="office-viewer-site-button is-light"
              href="#demo"
              onClick={handleSectionLink}
            >
              {content.finalCta.demo}
            </a>
            <a
              className="office-viewer-site-button is-ghost"
              href={GITHUB_URL}
              target="_blank"
              rel="noreferrer"
            >
              {content.finalCta.github}
            </a>
          </div>
        </section>
      </main>

      <SiteFooter content={content} />
    </div>
  );
}
