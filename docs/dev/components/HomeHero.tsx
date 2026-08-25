import React from 'react';
import { GITHUB_URL, type WebsiteContent } from './home-content';

type HomeHeroProps = {
  content: WebsiteContent;
  copyLabel: string;
  onCopy: () => void;
};

/** 用真实处理阶段解释文件如何在浏览器内完成预览。 */
function LocalProcessingFlow({ content }: { content: WebsiteContent }) {
  const formats = content.formats.groups.flatMap((group) => group.formats);

  return (
    <aside
      className="office-viewer-site-flow"
      aria-label={content.productFlow.title}
    >
      <div className="office-viewer-site-flow-header">
        <strong>{content.productFlow.title}</strong>
        <span>{content.productFlow.status}</span>
      </div>

      <div className="office-viewer-site-flow-file">
        <span className="office-viewer-site-flow-file-mark" aria-hidden="true">
          <i />
          <i />
        </span>
        <div>
          <strong>{content.productFlow.fileName}</strong>
          <span>{content.productFlow.fileMeta}</span>
        </div>
        <div className="office-viewer-site-flow-formats" aria-hidden="true">
          {formats.map((format) => (
            <span key={format}>{format}</span>
          ))}
        </div>
      </div>

      <ol className="office-viewer-site-flow-stages">
        {content.productFlow.stages.map((stage) => (
          <li key={stage.index}>
            <span>{stage.index}</span>
            <div>
              <strong>{stage.title}</strong>
              <p>{stage.description}</p>
            </div>
          </li>
        ))}
      </ol>

      <div className="office-viewer-site-flow-privacy">
        <span aria-hidden="true">
          <i />
        </span>
        <div>
          <strong>{content.productFlow.privacyTitle}</strong>
          <p>{content.productFlow.privacyDescription}</p>
        </div>
      </div>
    </aside>
  );
}

/** 渲染首页核心价值、主要操作、安装命令与本地处理说明。 */
export function HomeHero({ content, copyLabel, onCopy }: HomeHeroProps) {
  return (
    <section className="office-viewer-site-hero" aria-labelledby="hero-title">
      <div className="office-viewer-site-hero-copy">
        <span className="office-viewer-site-eyebrow">
          {content.hero.eyebrow}
        </span>
        <h1 id="hero-title">{content.hero.title}</h1>
        <p>{content.hero.description}</p>

        <div className="office-viewer-site-hero-actions">
          <a
            className="office-viewer-site-button is-primary"
            href={content.playgroundHref}
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
          <button type="button" onClick={onCopy} aria-live="polite">
            {copyLabel}
          </button>
        </div>

        <span className="office-viewer-site-privacy-line">
          <i aria-hidden="true" />
          {content.hero.privacy}
        </span>
      </div>

      <LocalProcessingFlow content={content} />
    </section>
  );
}
