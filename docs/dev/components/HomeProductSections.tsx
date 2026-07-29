import React from 'react';
import type { WebsiteContent } from './home-content';

type HomeProductSectionProps = {
  content: WebsiteContent;
};

/** 说明组件定位、适用场景与能够直接验证的产品事实。 */
export function ProductOverviewSection({ content }: HomeProductSectionProps) {
  return (
    <section
      className="office-viewer-site-section office-viewer-site-overview"
      id="overview"
      aria-labelledby="overview-title"
    >
      <div className="office-viewer-site-section-heading is-split">
        <div>
          <span className="office-viewer-site-eyebrow">
            {content.overview.eyebrow}
          </span>
          <h2 id="overview-title">{content.overview.title}</h2>
        </div>
        <div className="office-viewer-site-overview-copy">
          <p>{content.overview.description}</p>
          <p>{content.overview.details}</p>
        </div>
      </div>

      <div className="office-viewer-site-overview-capabilities">
        {content.overview.capabilities.map((capability) => (
          <article key={capability.index}>
            <span>{capability.index}</span>
            <h3>{capability.title}</h3>
            <p>{capability.description}</p>
          </article>
        ))}
      </div>

      <div className="office-viewer-site-overview-use-cases">
        <strong>{content.overview.useCasesLabel}</strong>
        <div>
          {content.overview.useCases.map((useCase) => (
            <span key={useCase}>{useCase}</span>
          ))}
        </div>
      </div>

      <div className="office-viewer-site-facts">
        {content.overview.facts.map((fact) => (
          <div key={fact.label}>
            <strong>{fact.value}</strong>
            <span>{fact.label}</span>
          </div>
        ))}
      </div>
    </section>
  );
}

/** 展示隐私、性能、接入与文档操作方面的核心亮点。 */
export function ProductHighlightsSection({ content }: HomeProductSectionProps) {
  return (
    <section
      className="office-viewer-site-section office-viewer-site-highlights"
      id="highlights"
      aria-labelledby="highlights-title"
    >
      <div className="office-viewer-site-section-heading">
        <span className="office-viewer-site-eyebrow">
          {content.highlights.eyebrow}
        </span>
        <h2 id="highlights-title">{content.highlights.title}</h2>
        <p>{content.highlights.description}</p>
      </div>

      <div className="office-viewer-site-highlight-grid">
        {content.highlights.items.map((highlight, index) => (
          <article
            key={highlight.title}
            className={`office-viewer-site-highlight-card is-${highlight.tone}`}
          >
            <div className="office-viewer-site-highlight-meta">
              <span>{String(index + 1).padStart(2, '0')}</span>
              <strong>{highlight.label}</strong>
            </div>
            <h3>{highlight.title}</h3>
            <p>{highlight.description}</p>
          </article>
        ))}
      </div>
    </section>
  );
}

/** 按文档家族说明七种扩展名及其对应的预览能力。 */
export function SupportedFormatsSection({ content }: HomeProductSectionProps) {
  return (
    <section
      className="office-viewer-site-section office-viewer-site-formats"
      id="formats"
      aria-labelledby="formats-title"
    >
      <div className="office-viewer-site-section-heading is-split">
        <div>
          <span className="office-viewer-site-eyebrow">
            {content.formats.eyebrow}
          </span>
          <h2 id="formats-title">{content.formats.title}</h2>
        </div>
        <p>{content.formats.description}</p>
      </div>

      <div className="office-viewer-site-format-grid">
        {content.formats.groups.map((group) => (
          <article
            key={group.category}
            className={`office-viewer-site-format-card is-${group.tone}`}
          >
            <div className="office-viewer-site-format-header">
              <span>{group.category}</span>
              <div aria-label={group.formats.join(', ')}>
                {group.formats.map((format) => (
                  <strong key={format}>{format}</strong>
                ))}
              </div>
            </div>
            <h3>{group.title}</h3>
            <p>{group.description}</p>
            <ul>
              {group.capabilities.map((capability) => (
                <li key={capability}>{capability}</li>
              ))}
            </ul>
          </article>
        ))}
      </div>

      <div className="office-viewer-site-format-roadmap">
        <strong>{content.formats.roadmapTitle}</strong>
        <p>{content.formats.roadmapDescription}</p>
      </div>
    </section>
  );
}
