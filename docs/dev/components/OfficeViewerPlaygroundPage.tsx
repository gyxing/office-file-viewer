import React, { useEffect, useState } from 'react';
import { getWebsiteContent, type WebsiteLocale } from './home-content';
import './OfficeViewerPlaygroundPage.less';
import {
  DEFAULT_PLAYGROUND_CONFIG,
  type PlaygroundConfig,
} from './playground-config';
import { getPlaygroundContent } from './playground-content';
import { PlaygroundCodePanel } from './PlaygroundCodePanel';
import { PlaygroundControlPanel } from './PlaygroundControlPanel';
import { PlaygroundHeader } from './PlaygroundHeader';
import { PlaygroundPreview } from './PlaygroundPreview';
import { SiteFooter } from './SiteFooter';

type OfficeViewerPlaygroundPageProps = {
  /** 当前公开在线体验页使用的语言。 */
  locale: WebsiteLocale;
};

/** 渲染无需内置测试资产的公开组件参数体验页。 */
export function OfficeViewerPlaygroundPage({
  locale,
}: OfficeViewerPlaygroundPageProps) {
  const websiteContent = getWebsiteContent(locale);
  const content = getPlaygroundContent(locale);
  const [config, setConfig] = useState<PlaygroundConfig>(() => ({
    ...DEFAULT_PLAYGROUND_CONFIG,
  }));
  const [selectedFile, setSelectedFile] = useState<File>();

  useEffect(() => {
    // Dumi 仍会生成文档外壳，用页面标记限制公开体验页的独立站点样式。
    document.body.classList.add('office-viewer-site-page');
    return () => document.body.classList.remove('office-viewer-site-page');
  }, []);

  const updateConfig = (patch: Partial<PlaygroundConfig>) => {
    if (patch.target && patch.target !== config.target) {
      // 两个目标组件不会共享文件解析状态，切换时清空文件名避免状态误导。
      setSelectedFile(undefined);
    }
    setConfig((current) => ({ ...current, ...patch }));
  };

  const resetConfig = () => {
    setConfig({ ...DEFAULT_PLAYGROUND_CONFIG });
    setSelectedFile(undefined);
  };

  return (
    <div className="office-viewer-site office-viewer-playground" lang={locale}>
      <a className="office-viewer-site-skip" href="#playground-main">
        {websiteContent.skipToContent}
      </a>

      <PlaygroundHeader locale={locale} />

      <main id="playground-main">
        <section
          className="office-viewer-playground-intro"
          aria-labelledby="playground-title"
        >
          <span className="office-viewer-site-eyebrow">{content.eyebrow}</span>
          <h1 id="playground-title">{content.title}</h1>
          <p>{content.description}</p>
          <div className="office-viewer-playground-privacy">
            <i aria-hidden="true" />
            {content.privacy}
          </div>
        </section>

        <div className="office-viewer-playground-workspace">
          <PlaygroundControlPanel
            config={config}
            content={content}
            onChange={updateConfig}
            onReset={resetConfig}
          />

          <div className="office-viewer-playground-result">
            <section
              className="office-viewer-playground-preview-card"
              aria-labelledby="playground-preview-title"
            >
              <div className="office-viewer-playground-preview-header">
                <div>
                  <h2 id="playground-preview-title">{content.previewTitle}</h2>
                  <p>
                    {config.target === 'viewer'
                      ? content.viewerHint
                      : content.layoutHint}
                  </p>
                </div>
                <span title={selectedFile?.name ?? content.noFile}>
                  <strong>{content.selectedFile}</strong>
                  {selectedFile?.name ?? content.noFile}
                </span>
              </div>
              <div className="office-viewer-playground-preview-surface">
                <PlaygroundPreview
                  config={config}
                  content={content}
                  locale={locale}
                  selectedFile={selectedFile}
                  onFileSelect={setSelectedFile}
                  onZoomChange={(zoom) => updateConfig({ zoom })}
                />
              </div>
            </section>

            <PlaygroundCodePanel
              config={config}
              content={content}
              locale={locale}
            />
          </div>
        </div>
      </main>

      <SiteFooter content={websiteContent} />
    </div>
  );
}
