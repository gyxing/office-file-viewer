# Office File Viewer

[English](./README.md) | 简体中文 | [在线示例](https://gyxing.github.io/office-file-viewer/)

> A browser-based React component for offline preview of DOC/DOCX/WPS, XLS/XLSX, and PPT/PPTX files.

`office-file-viewer` 是一个面向 React 的纯浏览器 Office 文件预览组件。文件下载、解析和渲染均在浏览器内完成，不需要配套的文档转换服务，也不会主动把本地文件上传到服务器。

支持 DOC、DOCX、WPS、XLS、XLSX、PPT 和 PPTX。组件统一提供文件选择、远程加载、解析进度、缩放、全屏、幻灯片翻页和工作表切换等能力。

> 这是独立解析和渲染实现，不等同于 Microsoft Office 或 WPS 的原生排版引擎。复杂文档的显示结果可能与桌面应用存在差异，使用前请阅读[支持格式](#支持格式)和[使用边界](#使用边界)。

## 主要特性

- **纯前端解析**：本地文件无需上传，适合内网、离线环境和对数据隐私敏感的场景
- **七种文件格式**：支持 DOC/DOCX/WPS、XLS/XLSX、PPT/PPTX
- **统一 React 组件**：不同格式共用同一套加载、错误、空状态、缩放和全屏交互
- **多种文件来源**：支持 `File`、远程 URL，以及返回 `File`、`Blob`、URL 或 `Response` 的异步函数
- **Worker 解析**：DOC/WPS、XLS、PPT 可在 Web Worker 中解析，降低大文件阻塞主线程的时间
- **进度与取消**：提供解析阶段、完成度订阅和会话取消能力；切换文件时自动丢弃过期结果
- **渐进预览**：组件解析 DOC/WPS、XLS、PPT 时可逐步接收已完成内容并提前渲染
- **资源管理**：组件自动释放 Worker、订阅和 Blob URL，底层 API 也提供显式释放函数
- **宿主集成**：兼容 antd v4、v5、v6，并继承宿主项目的 `ConfigProvider`
- **图表降级**：图表或地图数据加载失败时优先使用文档内快照，无法降级时显示明确状态

## 安装

```bash
npm install office-file-viewer antd react react-dom
```

项目使用 Yarn 时：

```bash
yarn add office-file-viewer antd react react-dom
```

`react`、`react-dom` 和 `antd` 由宿主项目提供；`echarts` 与 `jszip` 是组件自身的运行时依赖。

宿主构建工具需要支持 `.less` 文件，因为组件样式会随模块一起导入。

## 版本兼容

| antd 版本           | React / ReactDOM | 支持状态 | 说明                              |
| ------------------- | ---------------- | -------- | --------------------------------- |
| `4.24.x`            | `>=16.9.0`       | 支持     | 宿主入口需要加载 antd v4 全局样式 |
| `5.x`               | `>=16.9.0`       | 支持     | 使用 antd v5 的样式机制           |
| `6.x`               | `>=18.0.0`       | 支持     | React 版本要求来自 antd v6        |
| `6.x` + React 16/17 | -                | 不支持   | 不满足 antd v6 自身要求           |

当前 peerDependencies 范围：

```text
antd: >=4.24.0 <7.0.0
react: >=16.9.0
react-dom: >=16.9.0
```

React 的 peer 下限用于兼容 antd v4/v5 宿主项目，并不表示 antd v6 可以运行在 React 16/17 上。

使用 antd v4 时，需要在宿主应用入口加载全局样式：

```tsx
import 'antd/dist/antd.css';
```

antd v5、v6 不需要导入上述文件。组件不会创建额外的根级 `ConfigProvider`，主题、语言和组件前缀由宿主配置决定。

## 快速开始

不传 `uri` 时，组件会显示文件选择入口：

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer />;
}
```

组件自身界面默认使用简体中文。切换为英文时，需要同时配置组件语言和宿主 Ant Design 的语言：

```tsx
import { ConfigProvider } from 'antd';
import antdEnUS from 'antd/locale/en_US';
import { OfficeFileViewer } from 'office-file-viewer';

export default function EnglishOfficePreview() {
  return (
    <ConfigProvider locale={antdEnUS}>
      <OfficeFileViewer locale="en-US" />
    </ConfigProvider>
  );
}
```

传入本地文件或远程地址：

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview({ file }: { file: File }) {
  return (
    <div style={{ height: 720 }}>
      <OfficeFileViewer uri={file} />
      <OfficeFileViewer
        uri="https://example.com/files/demo.pptx"
        height="80vh"
      />
    </div>
  );
}
```

未传 `height` 时，预览器使用父容器高度；因此父容器需要提供可计算的高度。也可以直接传入数字像素值、`720px`、`80vh` 或 `100%`。

使用异步文件来源、解析配置和事件回调：

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return (
    <OfficeFileViewer
      uri={async () => fetch('/files/demo.xlsx')}
      parseOptions={{ worker: 'auto' }}
      onParseProgress={(progress) => {
        const percent =
          progress.percent === undefined
            ? ''
            : ` ${Math.round(progress.percent * 100)}%`;
        console.info(progress.stage, `${progress.message}${percent}`);
      }}
      onFileParsed={(parsed, file) => {
        console.info('解析完成', parsed.kind, file.name);
      }}
      onError={(error, file) => {
        console.error('预览失败', file?.name, error);
      }}
    />
  );
}
```

## `uri` 文件来源

`uri` 支持以下形式：

```ts
type OfficeFileViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);
```

使用远程文件时需要注意：

- 跨域地址必须允许浏览器通过 CORS 访问。
- URL 最好包含受支持的文件扩展名。
- 无扩展名地址需要通过 `Content-Disposition` 文件名或响应 `Content-Type` 识别格式。
- 带有不受支持扩展名的 URL 会在下载前被拒绝，即使响应内容实际是 Office 文件。
- `uri` 变化时，旧 URL 下载会通过 `AbortController` 取消。
- 自定义异步函数本身无法被强制取消，但过期结果不会覆盖新文件。
- 用户手动选择文件时，同样会终止当前远程下载并使旧解析结果失效。

## 组件属性

| 属性              | 类型                                             | 默认值    | 说明                                                      |
| ----------------- | ------------------------------------------------ | --------- | --------------------------------------------------------- |
| `locale`          | `'zh-CN' \| 'en-US'`                             | `'zh-CN'` | 预览器界面语言；Ant Design 语言仍由 `ConfigProvider` 提供 |
| `uri`             | `OfficeFileViewerUri`                            | -         | 预加载文件来源；不传时显示文件选择入口                    |
| `defaultFileName` | `string`                                         | 语言文案  | 未加载文件时显示的名称                                    |
| `defaultZoom`     | `number`                                         | `100`     | 初始缩放百分比，最终限制在 `25` 至 `300`                  |
| `className`       | `string`                                         | -         | 根容器自定义类名                                          |
| `height`          | `CSSProperties['height']`                        | `100%`    | 预览器高度；不传时跟随父容器                              |
| `style`           | `CSSProperties`                                  | -         | 根容器自定义样式                                          |
| `parseOptions`    | `OfficeParseOptions`                             | `{}`      | Worker 模式与自定义 Worker 工厂                           |
| `onParseProgress` | `(progress: ParseProgress) => void`              | -         | 解析阶段或完成度变化时触发                                |
| `onFileParsed`    | `(parsed: ParsedOfficeFile, file: File) => void` | -         | 完整文件解析成功后触发一次，不接收内部渐进结果            |
| `onError`         | `(error: Error, file?: File) => void`            | -         | 文件下载、解析或全屏操作失败时触发                        |

`PreviewKind` 的可选值为 `'docx' | 'doc' | 'xlsx' | 'xls' | 'pptx' | 'ppt'`；WPS 文件复用 `'doc'` 预览模型。

传入 `height` 时优先于 `style.height`。百分比高度依赖父级高度；如果页面没有固定高度，建议使用数字、`px` 或 `vh`。

常用公开类型：

```ts
import type {
  OfficeFileViewerProps,
  OfficeFileViewerLocale,
  OfficeFileViewerUri,
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParsedOfficeFile,
  ParseProgress,
  ParseStage,
  PreviewKind,
  WorkerMode,
} from 'office-file-viewer';
```

## 支持格式

| 文档类型           | 扩展名  | 当前能力                                                               |
| ------------------ | ------- | ---------------------------------------------------------------------- |
| Word OOXML         | `.docx` | 富文本段落、列表、表格、图片、图表、VML/WPG 形状、超链接、样式与主题色 |
| Word 97–2003       | `.doc`  | CFB、FIB、Piece Table、FKP、SPRM、正文结构、表格、列表及图片提取       |
| WPS 文字           | `.wps`  | 复用 DOC 二进制解析链路，以内容可读和资源提取为目标                    |
| Excel OOXML        | `.xlsx` | 多工作表、单元格值与样式、合并单元格、行列尺寸、浮动图片和图表         |
| Excel 97–2003      | `.xls`  | BIFF8 工作簿、单元格、格式、合并区域、行列尺寸、OfficeArt 图片及图表   |
| PowerPoint XML     | `.pptx` | 母版与布局继承、文本、形状、图片、表格、背景、渐变、阴影及常见图表     |
| PowerPoint 97–2003 | `.ppt`  | 二进制记录、母版、文本、形状、图片、嵌入图表及无法解析内容的静态预览   |

Office/WPS 图表当前覆盖线图、柱图、饼图、环形图、面积图、散点图、气泡图、雷达图和地图等常见类型。非标准扩展或损坏数据会尽量使用文档内快照，无法降级时显示明确状态。

格式表中的能力表示解析器已经覆盖的主要路径，并不保证所有 Office 版本、厂商扩展、宏、嵌入对象、动画或复杂排版都能完整还原。

## Web Worker 与性能

`parseOptions.worker` 控制解析任务的执行位置：

| 模式       | DOC/WPS、XLS、PPT                                         | DOCX、XLSX、PPTX                         |
| ---------- | --------------------------------------------------------- | ---------------------------------------- |
| `'auto'`   | 默认优先 Worker；环境不支持或 Worker 启动失败时回退主线程 | 当前直接使用主线程                       |
| `'always'` | 强制使用 Worker；无法创建时抛出配置错误                   | 当前尚未完成 Worker 迁移，会抛出配置错误 |
| `'never'`  | 始终使用主线程                                            | 始终使用主线程                           |

```tsx
<OfficeFileViewer parseOptions={{ worker: 'auto' }} />
```

每个解析会话拥有独立 Worker。文件缓冲区会通过 transferable object 移交给 Worker，解析过程通过有序消息、ACK 背压和取消消息协调。组件切换文件或卸载时会终止当前 Worker。

`workerFactory` 允许宿主接管 Worker 创建方式，主要用于有特殊资源路径或 CSP 配置的构建环境。工厂必须返回与当前解析协议兼容的 Worker；一般项目保持默认即可。

大文件注意事项：

- Worker 能降低旧二进制格式解析对主线程的阻塞，但不会减少文件本身和解析模型占用的内存。
- DOCX、XLSX、PPTX 当前仍在主线程解析，复杂或超大文件可能造成界面短暂卡顿。
- 组件没有设置文件大小、ZIP 条目数量、单条目大小或累计解压大小限制，宿主应根据业务场景预先校验。

## 底层解析会话

需要自行管理解析生命周期时，可以直接使用 `createOfficeParseSession`：

```ts
import {
  createOfficeParseSession,
  disposeDocDocument,
  disposePresentationDocument,
  disposeSpreadsheetWorkbook,
  type ParsedOfficeFile,
} from 'office-file-viewer';

export async function parseFile(file: File) {
  const session = createOfficeParseSession(file, { worker: 'auto' });
  const unsubscribe = session.subscribe((progress) => {
    console.info(progress.stage, progress.percent, progress.message);
  });

  try {
    const parsed = await session.result;
    console.info('解析完成', parsed.kind);
    return parsed;
  } finally {
    unsubscribe();
    session.dispose();
  }
}

export function disposeParsedFile(parsed: ParsedOfficeFile) {
  if (parsed.kind === 'doc') disposeDocDocument(parsed.document);
  if (parsed.kind === 'ppt') disposePresentationDocument(parsed.document);
  if (parsed.kind === 'xls' || parsed.kind === 'xlsx') {
    disposeSpreadsheetWorkbook(parsed.workbook);
  }
}
```

`OfficeParseSession` 提供：

- `result`：最终解析结果 Promise
- `status`：`starting`、`running`、`completed`、`cancelled` 或 `failed`
- `subscribe(listener)`：订阅解析进度，返回取消订阅函数
- `cancel()`：请求取消当前任务
- `dispose()`：释放 Worker、订阅和未移交给结果的临时资源

进度阶段包括 `reading`、`container`、`structure`、`content`、`resources` 和 `assembling`。`percent` 存在时取值范围为 `0` 到 `1`。

`OfficeFileViewer` 会自动完成取消和资源释放。直接使用底层会话时，应在结束后调用 `session.dispose()`；长期保存的解析结果不再使用时，再调用对应的文档释放函数清理 Blob URL。

## 交互说明

- 缩放范围为 `25%` 至 `300%`，工具栏快捷档位为 `50%`、`75%`、`100%`、`125%`、`150%`、`200%`。
- PPT/PPTX 支持幻灯片翻页和缩略图导航。
- XLS/XLSX 支持工作表标签切换。
- DOC/DOCX/WPS 不显示额外文档标题栏；旧格式解析警告会显示在原标题栏位置。
- 全屏依赖浏览器 Fullscreen API；不支持时按钮会禁用，按 `Esc` 退出后状态会自动同步。
- 地图图表可能需要加载外部 GeoJSON；网络失败时优先显示文档快照，否则显示加载失败状态。

## 使用边界

- 组件面向现代浏览器，依赖 `File`、`fetch`、`DOMParser`、`AbortController`、`IntersectionObserver`、`ResizeObserver`、Blob URL、Canvas、Web Worker 和 Fullscreen API 等能力。
- 本地文件解析和常规渲染不依赖服务端；远程 `uri`、外链图片或动态地图数据仍可能需要网络访问。
- 远程文件必须满足浏览器 CORS 策略。组件不会代理下载，也不会绕过鉴权和跨域限制。
- 面向不可信文件时，建议宿主在解析前限制文件大小、扩展名、MIME 类型和来源，并在服务端补充病毒扫描等安全策略。
- DOC/WPS、XLS、PPT 属于旧二进制格式，目前优先保证内容可读，不保证复杂分页、动画、锚点、图文环绕和排版与桌面 Office 完全一致。
- OOXML 文档可能包含未覆盖的厂商扩展、宏、ActiveX、OLE 对象、SmartArt 或复杂动画；这些内容可能降级、忽略或显示静态预览。
- 预览器只负责读取与显示，不提供编辑、保存、格式转换、打印排版或 PDF 导出能力。

## 项目结构

```text
src/
├── index.ts
└── office-file-viewer/
    ├── OfficeFileViewer.tsx  # 对外主组件与文件加载编排
    ├── shell/                # 工具栏、预览分发和通用状态
    ├── services/
    │   ├── parsing/          # 解析会话、Worker 协议、运行时与结果组装
    │   ├── doc/ docx/        # DOC/WPS 与 DOCX 解析
    │   ├── xls/ xlsx/        # XLS 与 XLSX 解析
    │   └── ppt/ pptx/        # PPT 与 PPTX 解析
    ├── formats/              # 各文档格式的 React 渲染器
    └── shared/
        ├── binary/           # CFB 等二进制基础能力
        ├── officeart/        # OfficeArt 图形记录
        ├── ooxml/            # ZIP、XML、关系、主题、媒体和图表适配
        └── chart/            # ECharts 渲染和失败降级
```

核心数据流：

```text
File / URL / async loader
  → 文件名与 MIME 类型识别
  → 主线程或 Web Worker 解析
  → 标准化 TypeScript 文档模型
  → React 格式渲染器
  → 浏览器预览
```

## 本地开发

```bash
yarn
yarn start
```

构建组件库和 Dumi 文档：

```bash
yarn build
yarn docs:build
```

项目使用 TypeScript、ESLint、Stylelint 和 Prettier 进行静态检查。发布前建议在目标 React/antd 组合中验证本地文件、远程 URI、Worker、全屏和各格式示例文档。

## License

[MIT](./LICENSE)
