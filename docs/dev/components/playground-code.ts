import type { WebsiteLocale } from './home-content';
import type { PlaygroundConfig } from './playground-config';
import {
  resolvePlaygroundTheme,
  resolvePlaygroundToolbar,
  resolvePlaygroundWatermark,
} from './playground-config';

/** 将字符串转换为可直接放入 TSX 的双引号字面量。 */
function quote(value: string): string {
  return JSON.stringify(value);
}

/** 按当前参数输出可复制的主题属性片段。 */
function buildThemeProp(config: PlaygroundConfig): string {
  return `      theme={{
        mode: ${quote(config.themeMode)},
        primaryColor: ${quote(config.primaryColor)},
        tokens: { workspaceColor: ${quote(config.workspaceColor)} },
      }}`;
}

/** 只在水印启用时输出配置，关闭状态保持示例简洁。 */
function buildWatermarkProp(config: PlaygroundConfig): string | undefined {
  if (!config.watermarkEnabled) return undefined;

  return `      watermark={{
        content: ${quote(config.watermarkContent)},
        color: ${quote(config.watermarkColor)},
        opacity: ${config.watermarkOpacity},
        rotate: ${config.watermarkRotate},
      }}`;
}

/** 根据工具栏模式生成默认省略、自定义配置或关闭属性。 */
function buildToolbarProp(config: PlaygroundConfig): string | undefined {
  const toolbar = resolvePlaygroundToolbar(config);
  if (toolbar === undefined) return undefined;
  if (toolbar === false) return '      toolbar={false}';

  const options = Object.entries(toolbar)
    .map(([key, value]) => `        ${key}: ${value},`)
    .join('\n');

  return `      toolbar={{
${options}
      }}`;
}

/** 生成完整预览器的可运行 React / TSX 示例。 */
function buildViewerExample(
  config: PlaygroundConfig,
  locale: WebsiteLocale,
): string {
  const props = [
    '      uri={file}',
    `      locale=${quote(locale)}`,
    `      height={${config.previewHeight}}`,
    `      defaultViewState={{ zoom: ${config.zoom} }}`,
    buildThemeProp(config),
    buildWatermarkProp(config),
    buildToolbarProp(config),
    config.searchEnabled ? undefined : '      search={false}',
    config.reviewEnabled ? undefined : '      review={false}',
    config.imagePreviewEnabled ? undefined : '      imagePreview={false}',
  ].filter((line): line is string => Boolean(line));

  return `import { OfficeFileViewer } from 'office-file-viewer';

type PreviewProps = {
  file?: File;
};

export function Preview({ file }: PreviewProps) {
  return (
    <OfficeFileViewer
${props.join('\n')}
    />
  );
}`;
}

/** 手动缩放模式需要由宿主内容消费外壳状态并应用变换。 */
function buildLayoutContent(config: PlaygroundConfig): string {
  if (config.layoutContentScaling === 'managed') {
    return `function PreviewContent() {
  return (
    <article className="host-preview">
      <h2>Host-rendered content</h2>
      <p>Your application keeps full control of this content.</p>
    </article>
  );
}`;
  }

  return `function PreviewContent() {
  const { state } = useOfficeViewerLayout();
  const scale = state.zoom / 100;

  return (
    <article
      className="host-preview"
      style={{
        transform: \`scale(\${scale})\`,
        transformOrigin: 'top left',
        width: \`\${100 / scale}%\`,
      }}
    >
      <h2>Host-rendered content</h2>
      <p>Your application keeps full control of this content.</p>
    </article>
  );
}`;
}

/** 生成可复用预览外壳的受控或非受控接入示例。 */
function buildLayoutExample(
  config: PlaygroundConfig,
  locale: WebsiteLocale,
): string {
  const imports =
    config.layoutContentScaling === 'manual'
      ? 'OfficeViewerLayout, useOfficeViewerLayout'
      : 'OfficeViewerLayout';
  const stateLines = [
    '  const [file, setFile] = useState<File>();',
    config.layoutControlledZoom
      ? `  const [zoom, setZoom] = useState(${config.zoom});`
      : undefined,
  ].filter((line): line is string => Boolean(line));
  const zoomProps = config.layoutControlledZoom
    ? ['      zoom={zoom}', '      onZoomChange={setZoom}']
    : [`      defaultZoom={${config.zoom}}`];
  const props = [
    `      locale=${quote(locale)}`,
    '      fileName={file?.name ?? "Host content"}',
    '      onFileSelect={setFile}',
    `      height={${config.previewHeight}}`,
    ...zoomProps,
    `      contentScaling=${quote(config.layoutContentScaling)}`,
    buildThemeProp(config),
    buildWatermarkProp(config),
    buildToolbarProp(config),
  ].filter((line): line is string => Boolean(line));

  return `import { ${imports} } from 'office-file-viewer';
import { useState } from 'react';

${buildLayoutContent(config)}

export function Preview() {
${stateLines.join('\n')}

  return (
    <OfficeViewerLayout
${props.join('\n')}
    >
      <PreviewContent />
    </OfficeViewerLayout>
  );
}`;
}

/** 构造页面当前真正传入组件的参数快照。 */
function buildPropsSnapshot(
  config: PlaygroundConfig,
  locale: WebsiteLocale,
): Record<string, unknown> {
  const sharedProps: Record<string, unknown> = {
    locale,
    height: config.previewHeight,
    theme: resolvePlaygroundTheme(config),
  };
  const watermark = resolvePlaygroundWatermark(config);
  const toolbar = resolvePlaygroundToolbar(config);

  if (watermark !== false) sharedProps.watermark = watermark;
  if (toolbar !== undefined) sharedProps.toolbar = toolbar;

  if (config.target === 'viewer') {
    return {
      component: 'OfficeFileViewer',
      props: {
        uri: 'File | URL | async loader',
        ...sharedProps,
        defaultViewState: { zoom: config.zoom },
        ...(config.searchEnabled ? {} : { search: false }),
        ...(config.reviewEnabled ? {} : { review: false }),
        ...(config.imagePreviewEnabled ? {} : { imagePreview: false }),
      },
    };
  }

  return {
    component: 'OfficeViewerLayout',
    props: {
      fileName: 'Host content',
      onFileSelect: '[Function]',
      ...sharedProps,
      ...(config.layoutControlledZoom
        ? { zoom: config.zoom, onZoomChange: '[Function]' }
        : { defaultZoom: config.zoom }),
      contentScaling: config.layoutContentScaling,
      children: '[ReactNode]',
    },
  };
}

/** 按目标组件生成会随页面参数即时变化的 React / TSX 代码。 */
export function buildPlaygroundExample(
  config: PlaygroundConfig,
  locale: WebsiteLocale,
): string {
  return config.target === 'viewer'
    ? buildViewerExample(config, locale)
    : buildLayoutExample(config, locale);
}

/** 以可读 JSON 展示当前组件属性，便于核对和复制配置。 */
export function buildPlaygroundPropsExample(
  config: PlaygroundConfig,
  locale: WebsiteLocale,
): string {
  return JSON.stringify(buildPropsSnapshot(config, locale), null, 2);
}
