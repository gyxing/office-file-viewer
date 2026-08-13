export type WebsiteLocale = 'en-US' | 'zh-CN';
export type PackageManager = 'npm' | 'yarn';

export const SITE_ROOT = '/office-file-viewer/';
export const GITHUB_URL = 'https://github.com/gyxing/office-file-viewer';
export const NPM_URL = 'https://www.npmjs.com/package/office-file-viewer';

type ProductFactContent = {
  value: string;
  label: string;
};

type ProductOverviewItemContent = {
  index: string;
  title: string;
  description: string;
};

type FlowStepContent = {
  index: string;
  title: string;
  description: string;
};

type HighlightContent = {
  label: string;
  title: string;
  description: string;
  tone: 'primary' | 'cyan' | 'violet' | 'neutral';
};

type FormatGroupContent = {
  category: string;
  title: string;
  formats: string[];
  description: string;
  capabilities: string[];
  tone: 'document' | 'spreadsheet' | 'presentation';
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
  /** 当前语言的完整文档入口。 */
  docsHref: string;
  skipToContent: string;
  navigation: {
    ariaLabel: string;
    /** 完整文档入口使用的本地化文案。 */
    docs: string;
    overview: string;
    demo: string;
    highlights: string;
    formats: string;
    github: string;
    npm: string;
  };
  documentation: {
    /** 桌面目录卡片使用的标题。 */
    tocTitle: string;
    /** 移动端打开目录的操作文案。 */
    openToc: string;
    /** 移动端关闭目录的无障碍文案。 */
    closeToc: string;
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
  productFlow: {
    title: string;
    status: string;
    fileName: string;
    fileMeta: string;
    stages: FlowStepContent[];
    privacyTitle: string;
    privacyDescription: string;
  };
  overview: {
    eyebrow: string;
    title: string;
    description: string;
    details: string;
    capabilities: ProductOverviewItemContent[];
    useCasesLabel: string;
    useCases: string[];
    facts: ProductFactContent[];
    architecture: {
      /** 当前语言对应的架构图静态资源地址。 */
      imageSrc: string;
      /** 架构图的无障碍替代文本。 */
      imageAlt: string;
    };
  };
  demo: {
    eyebrow: string;
    title: string;
    description: string;
    privacyTitle: string;
    privacyDescription: string;
  };
  highlights: {
    eyebrow: string;
    title: string;
    description: string;
    items: HighlightContent[];
  };
  formats: {
    eyebrow: string;
    title: string;
    description: string;
    groups: FormatGroupContent[];
    roadmapTitle: string;
    roadmapDescription: string;
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

// 中英文内容共用完全一致的数据结构，避免语言页面在功能区块上产生差异。
const WEBSITE_CONTENT: Record<WebsiteLocale, WebsiteContent> = {
  'en-US': {
    htmlLang: 'en-US',
    languageLabel: '简体中文',
    languageHref: `${SITE_ROOT}zh-CN/`,
    docsHref: `${SITE_ROOT}docs`,
    skipToContent: 'Skip to main content',
    navigation: {
      ariaLabel: 'Primary navigation',
      docs: 'Docs',
      overview: 'Product',
      demo: 'Live demo',
      highlights: 'Highlights',
      formats: 'Formats',
      github: 'GitHub',
      npm: 'npm',
    },
    documentation: {
      tocTitle: 'On this page',
      openToc: 'On this page',
      closeToc: 'Close table of contents',
    },
    hero: {
      eyebrow: 'Open-source React Office viewer',
      title: 'Preview Office files in the browser.',
      description:
        'Office File Viewer renders Word documents, Excel spreadsheets, and PowerPoint presentations with one React component—without a document conversion backend.',
      primaryAction: 'Try a local file',
      secondaryAction: 'View on GitHub',
      privacy: 'No upload. No conversion server. Your file stays in this tab.',
      copy: 'Copy',
      copied: 'Copied',
      copyFailed: 'Copy manually',
    },
    productFlow: {
      title: 'Local preview pipeline',
      status: 'Browser only',
      fileName: 'Your Office file',
      fileMeta: 'File API · local source',
      stages: [
        {
          index: '01',
          title: 'Select',
          description: 'Choose a file from this device.',
        },
        {
          index: '02',
          title: 'Parse',
          description: 'A Web Worker handles supported formats.',
        },
        {
          index: '03',
          title: 'Render',
          description: 'React renders the viewer in this tab.',
        },
      ],
      privacyTitle: 'The file remains local',
      privacyDescription:
        'The demo does not send the selected document to an upload service.',
    },
    overview: {
      eyebrow: 'Product',
      title: 'One React viewer for documents, spreadsheets, and slides.',
      description:
        'Office File Viewer is an open-source component for applications that need an embedded Office preview without operating a separate conversion service.',
      details:
        'It is designed for product teams building document centers, attachment previews, knowledge bases, and other workflows where Office content needs to stay inside the application experience.',
      capabilities: [
        {
          index: '01',
          title: 'Accept the source you already have',
          description:
            'Pass a local File, a URL, or an async loader without creating a separate integration for each document family.',
        },
        {
          index: '02',
          title: 'Manage the preview lifecycle',
          description:
            'Format detection, parsing progress, cancellation, errors, and resource cleanup stay within one component boundary.',
        },
        {
          index: '03',
          title: 'Adapt the interface to the file',
          description:
            'The shared viewer shell presents pages, worksheets, or slides with controls suited to the detected format.',
        },
      ],
      useCasesLabel: 'Typical use cases',
      useCases: [
        'Document management',
        'Knowledge bases',
        'Business attachments',
        'Privacy-sensitive tools',
      ],
      facts: [
        { value: '7', label: 'supported extensions' },
        { value: '3', label: 'document families' },
        { value: '1 API', label: 'shared React component' },
      ],
      architecture: {
        imageSrc: `${SITE_ROOT}assets/office-viewer-architecture-flow.en-us.svg`,
        imageAlt:
          'Architecture flow showing how Office File Viewer routes and parses Office files before rendering Word, Excel, or PowerPoint previews.',
      },
    },
    demo: {
      eyebrow: 'Live demo',
      title: 'Open a file from your device.',
      description:
        'Choose a supported Office file and use the real viewer. Parsing and rendering happen inside this browser tab.',
      privacyTitle: 'Private by default',
      privacyDescription:
        'This page does not upload the file you select to a server.',
    },
    highlights: {
      eyebrow: 'Why Office File Viewer',
      title: 'Useful capabilities for real product interfaces.',
      description:
        'The component focuses on the parts that matter when Office preview becomes part of an application rather than a standalone conversion task.',
      items: [
        {
          label: 'Privacy',
          title: 'Browser-side by architecture',
          description:
            'Local files are read, parsed, and rendered in the browser, so a companion upload or conversion service is not required.',
          tone: 'primary',
        },
        {
          label: 'Performance',
          title: 'Responsive parsing for larger files',
          description:
            'Worker parsing, progress events, cancellation, and progressive rendering help keep long-running previews responsive.',
          tone: 'cyan',
        },
        {
          label: 'Integration',
          title: 'One component across formats',
          description:
            'Use the same React API for File, URL, and async sources while keeping lifecycle and error handling in the host application.',
          tone: 'violet',
        },
        {
          label: 'Viewing experience',
          title: 'Controls that understand the document',
          description:
            'Full-document search, source hyperlinks, zoom, fullscreen, content-image preview and download, outlines, spreadsheet display modes, worksheet tabs, slide navigation, and speaker notes appear where the format supports them.',
          tone: 'neutral',
        },
      ],
    },
    formats: {
      eyebrow: 'Supported formats',
      title: 'Seven common Office extensions, grouped by how people use them.',
      description:
        'The viewer provides a shared shell while adapting navigation and controls to each document family.',
      groups: [
        {
          category: 'Documents',
          title: 'Word and WPS documents',
          formats: ['DOC', 'DOCX', 'WPS'],
          description:
            'Preview classic binary documents and modern Word packages in a paginated reading surface.',
          capabilities: [
            'Pagination',
            'Document outline',
            'Zoom and fullscreen',
          ],
          tone: 'document',
        },
        {
          category: 'Spreadsheets',
          title: 'Excel workbooks',
          formats: ['XLS', 'XLSX'],
          description:
            'Switch between a source-faithful layout and a reading mode that prioritizes complete cell text.',
          capabilities: [
            'Worksheet tabs',
            'Original and reading modes',
            'Charts and drawings',
          ],
          tone: 'spreadsheet',
        },
        {
          category: 'Presentations',
          title: 'PowerPoint presentations',
          formats: ['PPT', 'PPTX'],
          description:
            'Move through slides with controls suited to presentation review and playback.',
          capabilities: ['Slide navigation', 'Fullscreen', 'Speaker notes'],
          tone: 'presentation',
        },
      ],
      roadmapTitle: 'More formats are on the roadmap',
      roadmapDescription:
        'The current release focuses on commonly used documents, spreadsheets, and presentations. Future versions will expand Office format support based on real-world needs and community feedback.',
    },
    developer: {
      eyebrow: 'Developer experience',
      title: 'From install to preview in a few lines.',
      description: 'Install the package in your existing React application.',
      installLabel: 'Install',
      exampleLabel: 'React example',
      packageManagerLabel: 'Package manager',
      commands: {
        npm: 'npm install office-file-viewer',
        yarn: 'yarn add office-file-viewer',
      },
      example: `import { OfficeFileViewer } from 'office-file-viewer';

export function Preview({ file }: { file: File }) {
  return <OfficeFileViewer uri={file} locale="en-US" height={640} />;
}`,
      copy: 'Copy command',
      copied: 'Command copied',
      copyFailed: 'Select and copy manually',
      github: 'Read the source',
      npm: 'View package',
    },
    compatibility: {
      eyebrow: 'Compatibility',
      title: 'Fits the React stack you already use.',
      description:
        'The viewer supports React 16.9 and later, ships as ESM, and includes its compiled component styles.',
      items: [
        { name: 'React', version: '>= 16.9', note: 'Hooks support' },
        { name: 'ReactDOM', version: '>= 16.9', note: 'Match React' },
        { name: 'Module', version: 'ESM-only', note: 'Built-in CSS' },
      ],
    },
    finalCta: {
      title: 'Add Office preview without adding an upload pipeline.',
      description:
        'Try the real component with a local file, then bring the same viewer into your React project.',
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
    docsHref: `${SITE_ROOT}zh-CN/docs`,
    skipToContent: '跳转到主要内容',
    navigation: {
      ariaLabel: '主导航',
      docs: '开发文档',
      overview: '产品介绍',
      demo: '在线体验',
      highlights: '优点亮点',
      formats: '支持格式',
      github: 'GitHub',
      npm: 'npm',
    },
    documentation: {
      tocTitle: '本页目录',
      openToc: '本页目录',
      closeToc: '关闭本页目录',
    },
    hero: {
      eyebrow: '开源 React Office 预览组件',
      title: '在浏览器中预览 Office 文件',
      description:
        'Office File Viewer 使用一个 React 组件渲染 Word 文档、Excel 表格和 PowerPoint 演示文稿，无需额外部署文档转换服务。',
      primaryAction: '体验本地文件',
      secondaryAction: '查看 GitHub',
      privacy: '无需上传，无需转换服务，文件始终保留在当前标签页。',
      copy: '复制',
      copied: '已复制',
      copyFailed: '请手动复制',
    },
    productFlow: {
      title: '本地预览流程',
      status: '仅在浏览器中',
      fileName: '你的 Office 文件',
      fileMeta: 'File API · 本地来源',
      stages: [
        {
          index: '01',
          title: '选择',
          description: '从当前设备选择文件。',
        },
        {
          index: '02',
          title: '解析',
          description: '通过 Web Worker 处理支持的格式。',
        },
        {
          index: '03',
          title: '渲染',
          description: '由 React 在当前标签页呈现预览。',
        },
      ],
      privacyTitle: '文件始终保留在本地',
      privacyDescription: '在线示例不会把你选择的文档发送到上传服务。',
    },
    overview: {
      eyebrow: '产品介绍',
      title: '用一个 React 组件预览文档、表格和演示文稿。',
      description:
        'Office File Viewer 是一个开源 Office 文件预览组件，适合需要在业务系统中嵌入文档预览、又不想额外维护转换服务的 React 项目。',
      details:
        '它面向文档中心、附件预览、知识库等真实产品场景，让 Office 内容留在现有应用体验中，而不是跳转到独立转换页面。',
      capabilities: [
        {
          index: '01',
          title: '接收已有的文件来源',
          description:
            '直接传入本地 File、URL 或异步加载函数，不需要为每类文档单独设计一套接入方式。',
        },
        {
          index: '02',
          title: '管理完整预览生命周期',
          description:
            '格式识别、解析进度、取消、错误处理和资源释放都集中在一个组件边界内。',
        },
        {
          index: '03',
          title: '根据文件调整预览界面',
          description:
            '统一预览外壳会根据识别结果呈现页面、工作表或幻灯片，并提供对应操作。',
        },
      ],
      useCasesLabel: '典型使用场景',
      useCases: ['文档管理系统', '企业知识库', '业务附件预览', '隐私敏感工具'],
      facts: [
        { value: '7', label: '种受支持扩展名' },
        { value: '3', label: '类 Office 内容' },
        { value: '1 个 API', label: '统一 React 组件' },
      ],
      architecture: {
        imageSrc: `${SITE_ROOT}assets/office-viewer-architecture-flow.zh-cn.svg`,
        imageAlt:
          'Office File Viewer 从文件输入、格式路由和解析执行到 Word、Excel、PowerPoint 预览渲染的架构流程图。',
      },
    },
    demo: {
      eyebrow: '在线体验',
      title: '从你的设备打开一个文件。',
      description:
        '选择受支持的 Office 文件，直接体验真实组件。解析与渲染全部在当前浏览器标签页完成。',
      privacyTitle: '默认保护隐私',
      privacyDescription: '本页面不会把你选择的文件上传到服务器。',
    },
    highlights: {
      eyebrow: '优点亮点',
      title: '面向真实产品界面的 Office 预览能力。',
      description:
        '当 Office 预览成为业务系统的一部分时，组件不仅需要显示内容，还要兼顾隐私、性能、接入成本和文档操作体验。',
      items: [
        {
          label: '隐私',
          title: '从架构上保持浏览器本地处理',
          description:
            '本地文件直接在浏览器中读取、解析和渲染，不要求配套部署上传服务或文档转换服务。',
          tone: 'primary',
        },
        {
          label: '性能',
          title: '为较大文件保持界面响应',
          description:
            '通过 Worker 解析、进度事件、取消能力和渐进渲染，降低长时间预览对主界面交互的影响。',
          tone: 'cyan',
        },
        {
          label: '接入',
          title: '多种格式共用一个组件 API',
          description:
            '使用一致方式传入 File、URL 或异步来源，同时由宿主应用掌控生命周期、错误处理和业务状态。',
          tone: 'violet',
        },
        {
          label: '体验',
          title: '根据文档类型提供对应操作',
          description:
            '在格式支持时提供全文查找、源文档超链接、缩放、全屏、内容图片预览与下载、大纲、电子表格显示模式、工作表标签、幻灯片导航和演讲者备注等能力。',
          tone: 'neutral',
        },
      ],
    },
    formats: {
      eyebrow: '支持的文档格式',
      title: '覆盖七种常见 Office 文件扩展名。',
      description:
        '预览器复用统一外壳，同时根据文档、表格和演示文稿的特点调整导航与操作方式。',
      groups: [
        {
          category: '文本文档',
          title: 'Word 与 WPS 文档',
          formats: ['DOC', 'DOCX', 'WPS'],
          description: '在分页阅读界面中预览传统二进制文档和现代 Word 文档包。',
          capabilities: ['分页阅读', '文档大纲', '缩放与全屏'],
          tone: 'document',
        },
        {
          category: '电子表格',
          title: 'Excel 工作簿',
          formats: ['XLS', 'XLSX'],
          description:
            '可在优先还原源文件的版式与完整显示单元格文本的阅读模式之间切换。',
          capabilities: ['工作表标签', '原始版式与阅读模式', '图表与绘图对象'],
          tone: 'spreadsheet',
        },
        {
          category: '演示文稿',
          title: 'PowerPoint 演示文稿',
          formats: ['PPT', 'PPTX'],
          description: '使用适合演示文稿审阅和播放的操作方式浏览幻灯片。',
          capabilities: ['幻灯片导航', '全屏查看', '演讲者备注'],
          tone: 'presentation',
        },
      ],
      roadmapTitle: '持续扩展格式支持',
      roadmapDescription:
        '当前优先覆盖常见的文档、电子表格和演示文稿格式。后续版本将根据实际使用需求与社区反馈，继续支持更多 Office 文档格式。',
    },
    developer: {
      eyebrow: '开发者体验',
      title: '几行代码即可完成安装与预览。',
      description: '在现有 React 应用中安装组件。',
      installLabel: '安装',
      exampleLabel: 'React 示例',
      packageManagerLabel: '包管理器',
      commands: {
        npm: 'npm install office-file-viewer',
        yarn: 'yarn add office-file-viewer',
      },
      example: `import { OfficeFileViewer } from 'office-file-viewer';

export function Preview({ file }: { file: File }) {
  return <OfficeFileViewer uri={file} locale="zh-CN" height={640} />;
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
        '预览器支持 React 16.9 及以上版本，以 ESM 发布，并随包提供已编译的组件样式。',
      items: [
        { name: 'React', version: '>= 16.9', note: '支持 Hooks' },
        { name: 'ReactDOM', version: '>= 16.9', note: '与 React 匹配' },
        { name: '模块格式', version: '仅 ESM', note: '内置 CSS' },
      ],
    },
    finalCta: {
      title: '接入 Office 预览，而不是再搭建一套上传转换流程。',
      description:
        '先使用本地文件体验真实组件，再把同一个预览器接入你的 React 项目。',
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
