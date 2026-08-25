import React, { useMemo, useState } from 'react';
import type { WebsiteLocale } from './home-content';
import {
  buildPlaygroundExample,
  buildPlaygroundPropsExample,
} from './playground-code';
import type { PlaygroundConfig } from './playground-config';
import type {
  PlaygroundCodeTab,
  PlaygroundContent,
} from './playground-content';
import { getSiteCopyLabel, useSiteCopyFeedback } from './site-copy';

type PlaygroundCodePanelProps = {
  /** 当前示例代码采用的组件参数。 */
  config: PlaygroundConfig;
  /** 当前语言的代码面板文案。 */
  content: PlaygroundContent;
  /** 生成组件 locale 属性使用的语言。 */
  locale: WebsiteLocale;
};

const CODE_TABS: PlaygroundCodeTab[] = ['tsx', 'props'];

/** 展示与预览参数同步的 TSX 和序列化属性，并支持一键复制。 */
export function PlaygroundCodePanel({
  config,
  content,
  locale,
}: PlaygroundCodePanelProps) {
  const [activeTab, setActiveTab] = useState<PlaygroundCodeTab>('tsx');
  const copyFeedback = useSiteCopyFeedback();
  const tsxSource = useMemo(
    () => buildPlaygroundExample(config, locale),
    [config, locale],
  );
  const propsSource = useMemo(
    () => buildPlaygroundPropsExample(config, locale),
    [config, locale],
  );
  const source = activeTab === 'tsx' ? tsxSource : propsSource;

  const handleTabKeyDown = (
    event: React.KeyboardEvent<HTMLButtonElement>,
    currentTab: PlaygroundCodeTab,
  ) => {
    const currentIndex = CODE_TABS.indexOf(currentTab);
    let nextTab: PlaygroundCodeTab | undefined;

    if (event.key === 'ArrowRight') {
      nextTab = CODE_TABS[(currentIndex + 1) % CODE_TABS.length];
    } else if (event.key === 'ArrowLeft') {
      nextTab =
        CODE_TABS[(currentIndex - 1 + CODE_TABS.length) % CODE_TABS.length];
    } else if (event.key === 'Home') {
      nextTab = CODE_TABS[0];
    } else if (event.key === 'End') {
      nextTab = CODE_TABS[CODE_TABS.length - 1];
    }

    if (!nextTab) return;
    event.preventDefault();
    setActiveTab(nextTab);
    window.requestAnimationFrame(() =>
      document.getElementById(`playground-code-tab-${nextTab}`)?.focus(),
    );
  };

  return (
    <section
      className="office-viewer-playground-code"
      aria-labelledby="playground-code-title"
    >
      <div className="office-viewer-playground-code-header">
        <div>
          <h2 id="playground-code-title">{content.exampleCode}</h2>
          <div role="tablist" aria-label={content.exampleCode}>
            {CODE_TABS.map((tab) => (
              <button
                key={tab}
                id={`playground-code-tab-${tab}`}
                type="button"
                role="tab"
                tabIndex={activeTab === tab ? 0 : -1}
                aria-selected={activeTab === tab}
                aria-controls="playground-code-panel"
                onClick={() => setActiveTab(tab)}
                onKeyDown={(event) => handleTabKeyDown(event, tab)}
              >
                {tab === 'tsx' ? content.codeTab : content.propsTab}
              </button>
            ))}
          </div>
        </div>
        <button
          className="office-viewer-playground-copy"
          type="button"
          onClick={() => copyFeedback.copy(source)}
          aria-live="polite"
        >
          {getSiteCopyLabel(copyFeedback.state, {
            idle: content.copy,
            copied: content.copied,
            failed: content.copyFailed,
          })}
        </button>
      </div>
      <pre
        id="playground-code-panel"
        role="tabpanel"
        aria-labelledby={`playground-code-tab-${activeTab}`}
      >
        <code>{source}</code>
      </pre>
    </section>
  );
}
