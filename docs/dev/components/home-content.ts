export type WebsiteLocale = 'en-US' | 'zh-CN';
export type PackageManager = 'npm' | 'yarn';

export const SITE_ROOT = '/office-file-viewer/';

type FeatureContent = {
  metric: string;
  title: string;
  description: string;
  tone: 'primary' | 'violet' | 'cyan' | 'neutral';
  size?: 'wide' | 'tall';
};

type CompatibilityContent = {
  name: string;
  version: string;
  note: string;
};

export type WebsiteContent = {
  htmlLang: WebsiteLocale;
  languageLabel: string;
  languageHref: string;
  skipToContent: string;
  navigation: {
    ariaLabel: string;
    features: string;
    demo: string;
    compatibility: string;
    github: string;
    npm: string;
  };
  hero: {
    eyebrow: string;
    title: string;
    description: string;
    primaryAction: string;
    secondaryAction: string;
    privacy: string;
    copy: string;
    copied: string;
    copyFailed: string;
  };
  productPreview: {
    windowTitle: string;
    status: string;
    fileName: string;
    fileMeta: string;
    stages: string[];
    localOnly: string;
  };
  capabilities: Array<{ value: string; label: string }>;
  features: {
    eyebrow: string;
    title: string;
    description: string;
    items: FeatureContent[];
  };
  demo: {
    eyebrow: string;
    title: string;
    description: string;
    privacyTitle: string;
    privacyDescription: string;
  };
  developer: {
    eyebrow: string;
    title: string;
    description: string;
    installLabel: string;
    exampleLabel: string;
    packageManagerLabel: string;
    commands: Record<PackageManager, string>;
    example: string;
    copy: string;
    copied: string;
    copyFailed: string;
    github: string;
    npm: string;
  };
  compatibility: {
    eyebrow: string;
    title: string;
    description: string;
    items: CompatibilityContent[];
  };
  finalCta: {
    title: string;
    description: string;
    demo: string;
    github: string;
  };
  footer: {
    description: string;
    license: string;
    builtWith: string;
  };
};

export const GITHUB_URL = 'https://github.com/gyxing/office-file-viewer';
export const NPM_URL = 'https://www.npmjs.com/package/office-file-viewer';

const WEBSITE_CONTENT: Record<WebsiteLocale, WebsiteContent> = {
  'en-US': {
    htmlLang: 'en-US',
    languageLabel: '简体中文',
    languageHref: `${SITE_ROOT}zh-CN/`,
    skipToContent: 'Skip to main content',
    navigation: {
      ariaLabel: 'Primary navigation',
      features: 'Features',
      demo: 'Live demo',
      compatibility: 'Compatibility',
      github: 'GitHub',
      npm: 'npm',
    },
    hero: {
      eyebrow: 'Open-source React component',
      title: 'Preview Office files. Keep them private.',
      description:
        'Render Word, Excel, and PowerPoint files directly in the browser—without a conversion service or an upload pipeline.',
      primaryAction: 'Try it in your browser',
      secondaryAction: 'View on GitHub',
      privacy: 'Your file stays on your device.',
      copy: 'Copy',
      copied: 'Copied',
      copyFailed: 'Copy manually',
    },
    productPreview: {
      windowTitle: 'Quarterly-review.pptx',
      status: 'Ready locally',
      fileName: 'Office File Viewer',
      fileMeta: 'Browser-only preview',
      stages: ['Read', 'Parse', 'Render'],
      localOnly: 'No server round trip',
    },
    capabilities: [
      { value: '7', label: 'Office formats' },
      { value: '100%', label: 'Browser-side' },
      { value: 'Worker', label: 'Large-file parsing' },
      { value: 'React', label: 'One component API' },
    ],
    features: {
      eyebrow: 'Designed for real applications',
      title: 'A complete viewing experience, not an iframe shortcut.',
      description:
        'Use one component across document formats while keeping file processing, progress, and resource cleanup under your control.',
      items: [
        {
          metric: 'Local',
          title: 'Privacy by architecture',
          description:
            'Local files are parsed and rendered in the browser. The viewer does not require a companion conversion server.',
          tone: 'primary',
          size: 'wide',
        },
        {
          metric: '7 formats',
          title: 'Broad Office coverage',
          description:
            'DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX share one preview surface.',
          tone: 'violet',
        },
        {
          metric: 'Worker',
          title: 'Built for larger files',
          description:
            'Worker parsing, progress events, cancellation, and progressive rendering keep the interface responsive.',
          tone: 'cyan',
          size: 'tall',
        },
        {
          metric: '1 API',
          title: 'Simple React integration',
          description:
            'Pass a File, URL, or async source and handle every supported format consistently.',
          tone: 'neutral',
        },
        {
          metric: 'Native UX',
          title: 'Document-aware controls',
          description:
            'Zoom, fullscreen, slide navigation, speaker notes, worksheet tabs, and Word outlines are included.',
          tone: 'primary',
          size: 'wide',
        },
        {
          metric: 'antd 4–6',
          title: 'Fits the host application',
          description:
            'The viewer inherits the host ConfigProvider for theme and Ant Design localization.',
          tone: 'violet',
        },
      ],
    },
    demo: {
      eyebrow: 'Live demo',
      title: 'Open a file from your device.',
      description:
        'Choose any supported Office file. Parsing and rendering happen locally in this browser tab.',
      privacyTitle: 'Local processing',
      privacyDescription:
        'This demo never uploads the selected file to a server.',
    },
    developer: {
      eyebrow: 'Developer experience',
      title: 'From install to preview in a few lines.',
      description:
        'Install the package with the peer dependencies already used by your React application.',
      installLabel: 'Install',
      exampleLabel: 'React example',
      packageManagerLabel: 'Package manager',
      commands: {
        npm: 'npm install office-file-viewer antd react react-dom',
        yarn: 'yarn add office-file-viewer antd react react-dom',
      },
      example: `import { ConfigProvider } from 'antd';
import enUS from 'antd/locale/en_US';
import { OfficeFileViewer } from 'office-file-viewer';

export function Preview({ file }: { file: File }) {
  return (
    <ConfigProvider locale={enUS}>
      <OfficeFileViewer uri={file} locale="en-US" height={640} />
    </ConfigProvider>
  );
}`,
      copy: 'Copy command',
      copied: 'Command copied',
      copyFailed: 'Select and copy manually',
      github: 'Read the source',
      npm: 'View package',
    },
    compatibility: {
      eyebrow: 'Compatibility',
      title: 'Works with the React stack you already have.',
      description:
        'Ant Design remains a peer dependency, so the viewer follows your application version and ConfigProvider.',
      items: [
        { name: 'Ant Design 4', version: '>= 4.24', note: 'React 16.9+' },
        { name: 'Ant Design 5', version: '5.x', note: 'React 16.9+' },
        { name: 'Ant Design 6', version: '6.x', note: 'React 18+' },
      ],
    },
    finalCta: {
      title: 'Give your users an Office preview without giving up their files.',
      description:
        'Start with the live demo, then add the same component to your React project.',
      demo: 'Open the demo',
      github: 'Star on GitHub',
    },
    footer: {
      description: 'Open-source Office file preview for React.',
      license: 'MIT License',
      builtWith: 'Built for the browser',
    },
  },
  'zh-CN': {
    htmlLang: 'zh-CN',
    languageLabel: 'English',
    languageHref: SITE_ROOT,
    skipToContent: '跳转到主要内容',
    navigation: {
      ariaLabel: '主导航',
      features: '核心能力',
      demo: '在线体验',
      compatibility: '兼容性',
      github: 'GitHub',
      npm: 'npm',
    },
    hero: {
      eyebrow: '开源 React 组件',
      title: 'Office 文件预览，\n无需上传。',
      description:
        '直接在浏览器中渲染 Word、Excel 和 PowerPoint 文件，不依赖文档转换服务，也不需要上传流程。',
      primaryAction: '立即在线体验',
      secondaryAction: '查看 GitHub',
      privacy: '文件始终保留在你的设备上。',
      copy: '复制',
      copied: '已复制',
      copyFailed: '请手动复制',
    },
    productPreview: {
      windowTitle: '季度复盘.pptx',
      status: '已在本地就绪',
      fileName: 'Office File Viewer',
      fileMeta: '纯浏览器预览',
      stages: ['读取', '解析', '渲染'],
      localOnly: '无需服务端往返',
    },
    capabilities: [
      { value: '7', label: '种 Office 格式' },
      { value: '100%', label: '浏览器本地处理' },
      { value: 'Worker', label: '大文件解析' },
      { value: 'React', label: '统一组件 API' },
    ],
    features: {
      eyebrow: '为真实业务而设计',
      title: '完整的预览体验，而不是简单嵌入 iframe。',
      description:
        '使用一个组件覆盖多种文档格式，同时掌控文件处理、解析进度和资源释放。',
      items: [
        {
          metric: '本地',
          title: '架构级隐私保护',
          description:
            '本地文件直接在浏览器中解析与渲染，不需要配套的文档转换服务。',
          tone: 'primary',
          size: 'wide',
        },
        {
          metric: '7 种格式',
          title: '覆盖常见 Office 文件',
          description:
            'DOC、DOCX、WPS、XLS、XLSX、PPT 和 PPTX 使用统一预览界面。',
          tone: 'violet',
        },
        {
          metric: 'Worker',
          title: '面向更大的文件',
          description: 'Worker 解析、进度事件、取消和渐进渲染让界面保持响应。',
          tone: 'cyan',
          size: 'tall',
        },
        {
          metric: '1 个 API',
          title: '简单的 React 接入',
          description:
            '传入 File、URL 或异步来源，即可用一致方式处理全部支持格式。',
          tone: 'neutral',
        },
        {
          metric: '原生体验',
          title: '理解文档的操作能力',
          description:
            '内置缩放、全屏、幻灯片导航、演讲者备注、工作表标签和 Word 大纲。',
          tone: 'primary',
          size: 'wide',
        },
        {
          metric: 'antd 4–6',
          title: '自然融入宿主应用',
          description:
            '预览器继承宿主 ConfigProvider 的主题和 Ant Design 语言配置。',
          tone: 'violet',
        },
      ],
    },
    demo: {
      eyebrow: '在线体验',
      title: '从你的设备打开一个文件。',
      description:
        '选择任意受支持的 Office 文件，解析和渲染都在当前浏览器标签页中完成。',
      privacyTitle: '本地处理',
      privacyDescription: '本示例不会将所选文件上传到服务器。',
    },
    developer: {
      eyebrow: '开发者体验',
      title: '几行代码即可完成安装与预览。',
      description: '安装组件，并复用 React 项目中已有的 peer dependencies。',
      installLabel: '安装',
      exampleLabel: 'React 示例',
      packageManagerLabel: '包管理器',
      commands: {
        npm: 'npm install office-file-viewer antd react react-dom',
        yarn: 'yarn add office-file-viewer antd react react-dom',
      },
      example: `import { ConfigProvider } from 'antd';
import zhCN from 'antd/locale/zh_CN';
import { OfficeFileViewer } from 'office-file-viewer';

export function Preview({ file }: { file: File }) {
  return (
    <ConfigProvider locale={zhCN}>
      <OfficeFileViewer uri={file} locale="zh-CN" height={640} />
    </ConfigProvider>
  );
}`,
      copy: '复制命令',
      copied: '命令已复制',
      copyFailed: '请选择后手动复制',
      github: '查看源代码',
      npm: '查看 npm 包',
    },
    compatibility: {
      eyebrow: '兼容性',
      title: '适配你已经在使用的 React 技术栈。',
      description:
        'Ant Design 保持为 peer dependency，预览器会沿用应用自身的版本与 ConfigProvider。',
      items: [
        { name: 'Ant Design 4', version: '>= 4.24', note: 'React 16.9+' },
        { name: 'Ant Design 5', version: '5.x', note: 'React 16.9+' },
        { name: 'Ant Design 6', version: '6.x', note: 'React 18+' },
      ],
    },
    finalCta: {
      title: '让用户预览 Office 文件，也让文件继续属于用户。',
      description: '先体验在线示例，再把同一个组件加入你的 React 项目。',
      demo: '打开在线体验',
      github: '在 GitHub 点星',
    },
    footer: {
      description: '面向 React 的开源 Office 文件预览组件。',
      license: 'MIT 许可证',
      builtWith: '为浏览器而构建',
    },
  },
};

/** 根据公开页面语言返回结构一致的网站文案。 */
export function getWebsiteContent(locale: WebsiteLocale): WebsiteContent {
  return WEBSITE_CONTENT[locale];
}
