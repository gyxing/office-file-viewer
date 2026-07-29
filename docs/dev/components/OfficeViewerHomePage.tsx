import { ConfigProvider } from 'antd';
import enUS from 'antd/locale/en_US';
import zhCN from 'antd/locale/zh_CN';
import { OfficeFileViewer } from 'office-file-viewer';
import React, { useEffect, useRef, useState } from 'react';
import {
  getWebsiteContent,
  GITHUB_URL,
  NPM_URL,
  SITE_ROOT,
  type PackageManager,
  type WebsiteLocale,
} from './home-content';
import './OfficeViewerHomePage.less';

type OfficeViewerHomePageProps = {
  /** 当前公开官网使用的语言。 */
  locale: WebsiteLocale;
};

type CopyState = 'idle' | 'copied' | 'failed';

/** 优先使用 Clipboard API，并为权限受限或较旧的浏览器提供回退。 */
async function writeTextToClipboard(value: string) {
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

/** 渲染官网 Hero 右侧的轻量产品能力示意，不加载任何样例文档。 */
function ProductPreview({ locale }: { locale: WebsiteLocale }) {
  const content = getWebsiteContent(locale).productPreview;

  return (
    <div className="office-viewer-site-preview" aria-hidden="true">
      <div className="office-viewer-site-preview-window">
        <div className="office-viewer-site-preview-bar">
          <span className="office-viewer-site-preview-dots">
            <i />
            <i />
            <i />
          </span>
          <span>{content.windowTitle}</span>
          <span className="office-viewer-site-preview-status">
            {content.status}
          </span>
        </div>
        <div className="office-viewer-site-preview-body">
          <aside className="office-viewer-site-preview-sidebar">
            <strong>{content.fileName}</strong>
            <span>{content.fileMeta}</span>
            <div className="office-viewer-site-preview-files">
              <i className="is-active">PPTX</i>
              <i>XLSX</i>
              <i>DOCX</i>
            </div>
          </aside>
          <div className="office-viewer-site-preview-canvas">
            <div className="office-viewer-site-preview-sheet">
              <span className="office-viewer-site-preview-kicker">
                OFFICE / WEB
              </span>
              <strong>01</strong>
              <div className="office-viewer-site-preview-lines">
                <i />
                <i />
                <i />
              </div>
              <div className="office-viewer-site-preview-chart">
                <i />
                <i />
                <i />
                <i />
              </div>
            </div>
          </div>
        </div>
      </div>
      <div className="office-viewer-site-preview-progress">
        <div>
          {content.stages.map((stage, index) => (
            <span
              key={stage}
              className={
                index === content.stages.length - 1 ? 'is-active' : undefined
              }
            >
              <i />
              {stage}
            </span>
          ))}
        </div>
        <strong>{content.localOnly}</strong>
      </div>
    </div>
  );
}

/** 渲染默认英文、可切换中文的 Office File Viewer 产品官网。 */
export function OfficeViewerHomePage({ locale }: OfficeViewerHomePageProps) {
  const content = getWebsiteContent(locale);
  const [packageManager, setPackageManager] = useState<PackageManager>('npm');
  const heroCopy = useCopyFeedback();
  const developerCopy = useCopyFeedback();
  const installCommand = content.developer.commands[packageManager];
  const antdLocale = locale === 'zh-CN' ? zhCN : enUS;

  useEffect(() => {
    // Dumi 的 Markdown 路由仍会生成文档外壳，用页面标记将官网样式限制在两个公开首页。
    document.body.classList.add('office-viewer-site-page');
    return () => document.body.classList.remove('office-viewer-site-page');
  }, []);

  return (
    <div className="office-viewer-site" lang={content.htmlLang}>
      <a className="office-viewer-site-skip" href="#main-content">
        {content.skipToContent}
      </a>

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

          <div className="office-viewer-site-nav-links">
            <a href="#features" onClick={handleSectionLink}>
              {content.navigation.features}
            </a>
            <a href="#demo" onClick={handleSectionLink}>
              {content.navigation.demo}
            </a>
            <a href="#compatibility" onClick={handleSectionLink}>
              {content.navigation.compatibility}
            </a>
          </div>

          <div className="office-viewer-site-nav-actions">
            <a href={NPM_URL} target="_blank" rel="noreferrer">
              {content.navigation.npm}
            </a>
            <a href={GITHUB_URL} target="_blank" rel="noreferrer">
              {content.navigation.github}
            </a>
            <a
              className="office-viewer-site-language"
              href={content.languageHref}
            >
              {content.languageLabel}
            </a>
          </div>
        </nav>
      </header>

      <main id="main-content">
        <section
          className="office-viewer-site-hero"
          aria-labelledby="hero-title"
        >
          <div className="office-viewer-site-hero-copy">
            <span className="office-viewer-site-eyebrow">
              {content.hero.eyebrow}
            </span>
            <h1 id="hero-title">{content.hero.title}</h1>
            <p>{content.hero.description}</p>
            <div className="office-viewer-site-hero-actions">
              <a
                className="office-viewer-site-button is-primary"
                href="#demo"
                onClick={handleSectionLink}
              >
                {content.hero.primaryAction}
              </a>
              <a
                className="office-viewer-site-button is-secondary"
                href={GITHUB_URL}
                target="_blank"
                rel="noreferrer"
              >
                {content.hero.secondaryAction}
              </a>
            </div>
            <div className="office-viewer-site-command">
              <code>{content.developer.commands.npm}</code>
              <button
                type="button"
                onClick={() => heroCopy.copy(content.developer.commands.npm)}
                aria-live="polite"
              >
                {getCopyLabel(heroCopy.state, {
                  idle: content.hero.copy,
                  copied: content.hero.copied,
                  failed: content.hero.copyFailed,
                })}
              </button>
            </div>
            <span className="office-viewer-site-privacy-line">
              {content.hero.privacy}
            </span>
          </div>
          <ProductPreview locale={locale} />
        </section>

        <section
          className="office-viewer-site-capabilities"
          aria-label={content.features.eyebrow}
        >
          {content.capabilities.map((item) => (
            <div key={item.label}>
              <strong>{item.value}</strong>
              <span>{item.label}</span>
            </div>
          ))}
        </section>

        <section
          className="office-viewer-site-section office-viewer-site-features"
          id="features"
          aria-labelledby="features-title"
        >
          <div className="office-viewer-site-section-heading">
            <span className="office-viewer-site-eyebrow">
              {content.features.eyebrow}
            </span>
            <h2 id="features-title">{content.features.title}</h2>
            <p>{content.features.description}</p>
          </div>
          <div className="office-viewer-site-feature-grid">
            {content.features.items.map((feature) => (
              <article
                key={feature.title}
                className={[
                  'office-viewer-site-feature-card',
                  `is-${feature.tone}`,
                  feature.size ? `is-${feature.size}` : undefined,
                ]
                  .filter(Boolean)
                  .join(' ')}
              >
                <span>{feature.metric}</span>
                <h3>{feature.title}</h3>
                <p>{feature.description}</p>
              </article>
            ))}
          </div>
        </section>

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

        <section
          className="office-viewer-site-section office-viewer-site-developer"
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
                  {(['npm', 'yarn'] as const).map((manager) => (
                    <button
                      key={manager}
                      id={`package-tab-${manager}`}
                      type="button"
                      role="tab"
                      aria-selected={packageManager === manager}
                      aria-controls="install-command-panel"
                      onClick={() => setPackageManager(manager)}
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
          className="office-viewer-site-section office-viewer-site-compatibility"
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
    </div>
  );
}
