---
title: 使用文档
siteLayout: docs
toc: content
---

# Office File Viewer 使用文档

`office-file-viewer` 是一个面向 React 的纯浏览器 Office 文件预览组件，支持 DOC、DOCX、WPS、XLS、XLSX、PPT 和 PPTX。文件下载、解析与渲染均在浏览器中完成，不需要配套的文档转换服务。

<nav className="office-viewer-docs__quick-links" aria-label="文档快捷入口">
  <a href="#quick-start">快速接入</a>
  <a href="#component-api">组件 API</a>
  <a href="#advanced-api">高级 API</a>
  <a href="#limitations">限制说明</a>
</nav>

## 安装

```bash
npm install office-file-viewer
```

也可以使用 Yarn：

```bash
yarn add office-file-viewer
```

`react` 和 `react-dom` 由宿主项目提供。当前包仅发布 ESM，因此宿主构建工具必须支持 ESM。组件构建产物已包含 CSS，不要求宿主配置 Less loader。

公共 API 只能从 `office-file-viewer` 根入口导入。未文档化的 `dist` 或源码深层路径不属于兼容性承诺。

## 版本兼容

| 项目     | 要求       | 说明                          |
| -------- | ---------- | ----------------------------- |
| React    | `>=16.9.0` | 支持 Hooks 的 React 版本      |
| ReactDOM | `>=16.9.0` | 建议与 React 保持相同主版本   |
| 模块格式 | 仅 ESM     | 使用支持 ESM 的浏览器构建工具 |

当前 peer dependency 范围：

```text
react: >=16.9.0
react-dom: >=16.9.0
```

<a id="quick-start"></a>

## 快速接入

### 默认文件选择器

不传 `uri` 时，预览器会显示内置文件选择入口：

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer height="80vh" />;
}
```

不传 `height` 时，预览器跟随父容器高度。请确保父容器高度可计算，或者传入数字、`px`、`%`、`vh` 等值。

### 本地文件与远程 URL

```tsx | pure
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

### 异步来源与回调

异步来源可以返回 `File`、`Blob`、URL 字符串或 `Response`：

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return (
    <OfficeFileViewer
      uri={async (signal) => fetch('/files/demo.xlsx', { signal })}
      parseOptions={{ worker: 'auto' }}
      onParseProgress={(progress) => {
        const percent =
          progress.percent === undefined
            ? undefined
            : Math.round(progress.percent * 100);
        console.info(progress.stage, percent, progress.message);
      }}
      onPreviewReady={(info, file) => {
        console.info('Preview ready', info.previewKind, info.mode, file.name);
      }}
      onFileParsed={(parsed, file) => {
        console.info('Parsing completed', parsed.kind, file.name);
      }}
      onWarning={(warning, file) => {
        console.warn('Preview warning', file.name, warning.code);
      }}
      onError={(error, file) => {
        console.error('Preview failed', file?.name, error);
      }}
    />
  );
}
```

### 英文界面

预览器界面默认使用简体中文。将 `locale` 设为英文即可使用内置英文文案：

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function EnglishOfficePreview() {
  return <OfficeFileViewer locale="en-US" height="80vh" />;
}
```

## `uri` 文件来源

```ts | pure
type OfficeFileViewerUriLoader = (
  signal?: AbortSignal,
) => Promise<File | Blob | string | Response>;

type OfficeFileViewerUri = File | string | OfficeFileViewerUriLoader;
```

远程来源规则：

- 跨域 URL 必须允许浏览器通过 CORS 访问。
- URL 最好以受支持的 Office 扩展名结尾。
- 无扩展名 URL 必须通过 `Content-Disposition` 提供文件名，或通过 `Content-Type` 提供受支持的 MIME 类型。
- URL 已显式包含不支持的扩展名时，会在下载前拒绝，即使响应体实际包含 Office 文件。
- `uri` 变化时，之前的 URL 下载和解析会被取消。
- 异步加载函数会收到可选 `AbortSignal`；应把它传给 `fetch` 等支持取消的 API。原有无参数函数仍兼容，且过期结果不会覆盖最新来源。
- 手动选择文件也会使进行中的远程加载或解析结果失效。
- 加载或解析失败后，内置错误状态可以重新加载最近一次文件来源。

<a id="component-api"></a>

## `OfficeFileViewer` API

| 属性                             | 类型                                                 | 默认值     | 说明                                   |
| -------------------------------- | ---------------------------------------------------- | ---------- | -------------------------------------- |
| `locale`                         | `'zh-CN' \| 'en-US'`                                 | `'zh-CN'`  | 预览器内置界面语言                     |
| `uri`                            | `OfficeFileViewerUri`                                | -          | 预加载文件来源；未传时显示文件选择器   |
| `defaultFileName`                | `string`                                             | 本地化文案 | 来源无法提供有效文件名时使用的备用名称 |
| `defaultZoom`                    | `number`                                             | `100`      | 初始缩放百分比，限制在 `25` 到 `300`   |
| `defaultViewState`               | `Partial<OfficeFileViewerViewState>`                 | -          | 非受控视图字段的统一初始值             |
| `viewState`                      | `Partial<OfficeFileViewerViewState>`                 | -          | 按字段控制缩放、页面、侧栏等视图状态   |
| `onViewStateChange`              | `(state, change) => void`                            | -          | 用户请求改变任一视图字段时触发         |
| `defaultShowSpeakerNotes`        | `boolean`                                            | `false`    | 非受控模式下演讲者备注的初始状态       |
| `showSpeakerNotes`               | `boolean`                                            | -          | 受控模式下演讲者备注是否显示           |
| `onSpeakerNotesVisibilityChange` | `(visible: boolean) => void`                         | -          | 演讲者备注显示状态变化时触发           |
| `className`                      | `string`                                             | -          | 预览器根节点附加类名                   |
| `height`                         | `CSSProperties['height']`                            | 跟随父容器 | 预览器高度；优先级高于 `style.height`  |
| `style`                          | `CSSProperties`                                      | -          | 预览器根节点内联样式                   |
| `onFileParsed`                   | `(parsed: ParsedOfficeFile, file: File) => void`     | -          | 完整实体化解析结果可用时触发一次       |
| `onPreviewReady`                 | `(info: OfficePreviewReadyInfo, file: File) => void` | -          | 首个可用预览就绪时触发一次             |
| `onError`                        | `(error: Error, file?: File) => void`                | -          | 加载、解析或预览器操作失败时触发       |
| `onWarning`                      | `(warning, file) => void`                            | -          | 非致命解析降级或部分预览保留时触发     |
| `parseOptions`                   | `OfficeParseOptions`                                 | `{}`       | Worker 策略与可选资源限制              |
| `imagePreview`                   | `boolean \| OfficeFileViewerImagePreviewOptions`     | `true`     | 内容图片预览、下载与右键菜单配置       |
| `hyperlink`                      | `boolean`                                            | `true`     | 是否启用源文档明确声明的超链接         |
| `onHyperlinkActivate`            | `(event: OfficeHyperlinkActivateEvent) => void`      | -          | 链接有效激活时触发，可阻止默认导航     |
| `onParseProgress`                | `(progress: ParseProgress) => void`                  | -          | 解析阶段或完成度变化时触发             |

### 内容图片预览

DOC、DOCX、WPS、XLS 和 XLSX 中的可见内容图片默认支持双击、`Enter` 或空格键打开预览层。预览层支持适应窗口、`10%` 到 `500%` 缩放、拖拽、顺时针旋转、复位和下载；图片右键菜单只提供“预览”和“下载”。

```ts | pure
type OfficeFileViewerImagePreviewOptions = {
  download?: boolean;
  contextMenu?: boolean;
};

type OfficeFileViewerImagePreviewConfig =
  | boolean
  | OfficeFileViewerImagePreviewOptions;
```

不传或传入 `true` 时启用全部能力；传入 `false` 时关闭图片附加交互。对象配置可分别关闭下载或自定义右键菜单，例如：

```tsx | pure
<OfficeFileViewer
  uri={file}
  imagePreview={{ download: false, contextMenu: true }}
/>
```

关闭 `contextMenu` 后保留双击和键盘预览，并恢复浏览器原生右键菜单。页眉装饰图、背景、水印、页面绘图层、图表以及 PPT/PPTX 图片不属于该交互范围。

### Office 超链接

预览器只处理源 Office 文件明确声明的文字、单元格、图片、形状和按钮链接，不会把正文中的普通 URL 文本自动转成链接。为避免浏览文档时误触，Windows 和 Linux 使用 `Ctrl + 单击` 激活，macOS 使用 `Command + 单击` 激活；聚焦链接后可按 `Enter` 激活，触屏设备需要在提示有效期内再次点击确认。文字、单元格、形状和按钮链接支持通过右键菜单直接打开或跳转，外部安全地址还可以复制；本地路径、受限地址和图片继续沿用各自现有的右键行为。

```tsx | pure
import {
  OfficeFileViewer,
  type OfficeHyperlinkActivateEvent,
} from 'office-file-viewer';

export default function OfficePreview({ file }: { file: File }) {
  const handleHyperlinkActivate = (event: OfficeHyperlinkActivateEvent) => {
    console.info(event.sourceType, event.sourceId, event.hyperlink);

    // 需要交给应用路由、权限弹窗或审计流程时，阻止组件默认导航。
    if (event.hyperlink.kind === 'external') {
      event.preventDefault();
    }
  };

  return (
    <OfficeFileViewer
      uri={file}
      hyperlink
      onHyperlinkActivate={handleHyperlinkActivate}
    />
  );
}
```

`OfficeHyperlinkActivateEvent` 提供 `hyperlink`、`file`、`previewKind`、`sourceType`、`sourceId`、`defaultPrevented` 和 `preventDefault()`。组件的默认处理规则如下：

- `http` 和 `https` 在新标签页打开；`mailto` 和 `tel` 交给浏览器或系统处理。
- Word 书签、Excel 工作表/单元格/定义名称和 PowerPoint 幻灯片动作在当前预览器中定位。
- 本地路径、UNC 路径以及没有可靠远程来源 URL 的相对地址只触发回调，默认不打开。
- `javascript`、`data`、`vbscript` 和未知可执行协议始终阻止默认执行；宿主回调不能解除该安全限制。
- 图片的修饰键单击用于链接；普通双击和右键菜单仍按内容图片预览配置处理。

设置 `hyperlink={false}` 会关闭链接焦点、提示和激活行为，但不会改变源文档保存的颜色、下划线或版式。

### 统一视图状态

`defaultViewState` 用于设置非受控初始值；`viewState` 只控制实际传入的字段，未传字段仍由组件内部管理。可控制字段如下：

```ts | pure
type OfficeFileViewerViewState = {
  zoom: number;
  activeSlideIndex: number;
  activeSheetId?: string;
  wordOutlineVisible: boolean;
  speakerNotesVisible: boolean;
  spreadsheetViewMode: 'source' | 'reading';
};
```

`onViewStateChange` 的第一个参数是应用本次请求后的完整视图状态，第二个参数是 `{ key, value }` 形式的单字段变化。无效缩放值会回退并限制在 `25` 到 `300`，页面索引和工作表标识会按当前文件校验。

旧版 `defaultZoom`、`defaultShowSpeakerNotes`、`showSpeakerNotes` 和 `onSpeakerNotesVisibilityChange` 继续兼容。若同时传入，`defaultViewState` 中对应字段优先于旧版默认值，`viewState.speakerNotesVisible` 优先于 `showSpeakerNotes`。

### 回调时机

- `onPreviewReady` 在完整模型（`mode: 'materialized'`）或按需数据源（`mode: 'source'`）能够渲染首个可用预览时触发。
- `onFileParsed` 在完整实体化结果就绪后触发；渐进解析的中间结果不会触发它。
- `onParseProgress` 可以多次触发，且可能不包含精确数量或 `percent`。
- `onWarning` 用稳定的 `code`、`previewKind` 和 `source` 区分解析器警告与保留部分预览；它不会代替 `onError`。
- 错误发生在有效文件解析出来之前时，`onError` 可能收不到 `file`。资源策略错误可以通过 `OfficeResourceLimitError.code` 进一步判断。

## 解析配置与进度

```ts | pure
type WorkerMode = 'auto' | 'always' | 'never';

type OfficeParseResourcePolicy = {
  maxFileBytes?: number;
  maxArchiveEntries?: number;
  maxArchiveEntryBytes?: number;
  maxArchiveInflatedBytes?: number;
  maxCompressionRatio?: number;
  timeoutMs?: number;
};

type OfficeParseOptions = {
  worker?: WorkerMode;
  workerFactory?: () => Worker;
  resourcePolicy?: OfficeParseResourcePolicy;
};
```

| 模式       | DOC/WPS、XLS、PPT                       | DOCX、XLSX、PPTX                    |
| ---------- | --------------------------------------- | ----------------------------------- |
| `'auto'`   | 优先使用 Worker，必要时回退到主线程     | 当前在主线程解析                    |
| `'always'` | 必须创建兼容 Worker，无法创建时直接失败 | Worker 迁移尚未完成，会抛出配置错误 |
| `'never'`  | 始终在主线程解析                        | 始终在主线程解析                    |

每个正在运行的旧格式解析会话拥有独立 Worker。来源切换或组件卸载时，预览器会取消并释放当前会话。`workerFactory` 主要用于需要特殊资源路径或内容安全策略的宿主；大多数应用应使用内置工厂。

### 资源限制

默认不限制文件大小、解析时长或 OOXML 归档规模，避免擅自拒绝业务文件。处理不可信文件时，可以按宿主设备和业务边界配置限制：

```tsx | pure
<OfficeFileViewer
  parseOptions={{
    resourcePolicy: {
      maxFileBytes: 100 * 1024 * 1024,
      maxArchiveEntries: 10_000,
      maxArchiveEntryBytes: 64 * 1024 * 1024,
      maxArchiveInflatedBytes: 512 * 1024 * 1024,
      maxCompressionRatio: 500,
      timeoutMs: 120_000,
    },
  }}
/>
```

`maxFileBytes` 和 `timeoutMs` 适用于全部格式；归档条目、解压体积和压缩比限制适用于 DOCX、XLSX、PPTX。所有已配置值都必须是大于零的有限数。触发限制时，`onError` 收到 `OfficeResourceLimitError`，可以读取 `code`、`limit`、`actual` 和可选 `path`。`timeoutMs` 会发出取消请求；主线程中的同步解析片段需要等到下一次让出执行权后才能响应。

`percent` 存在时使用 `0` 到 `1` 的归一化进度值：

```ts | pure
type ParseStage =
  | 'reading'
  | 'container'
  | 'structure'
  | 'content'
  | 'resources'
  | 'assembling';

type ParseProgress = {
  stage: ParseStage;
  completed?: number;
  total?: number;
  percent?: number;
  message: string;
};

type OfficePreviewReadyInfo = {
  previewKind: PreviewKind;
  mode: 'materialized' | 'source';
};
```

`completed`、`total` 和 `percent` 都是可选字段。使用方应支持无法确定精确进度的状态，不能假设每种解析器都能返回完整进度。

<a id="advanced-api"></a>

## 高级解析 API

只有需要脱离 `OfficeFileViewer` 自行管理解析流程时，才使用底层会话：

```ts | pure
type OfficeParseSessionStatus =
  | 'starting'
  | 'running'
  | 'completed'
  | 'cancelled'
  | 'failed';

type OfficeParseSession<TParsed> = {
  readonly result: Promise<TParsed>;
  readonly status: OfficeParseSessionStatus;
  subscribe(listener: (progress: ParseProgress) => void): () => void;
  cancel(): void;
  dispose(): void;
};
```

`createOfficeParseSession(file, options)` 接收受支持的 `File`，返回 `OfficeParseSession<ParsedOfficeFile>`。取消操作会使 `result` 被拒绝；调用方等待该 Promise 时，应像处理其他解析失败一样处理取消错误。

### 会话生命周期

```ts | pure
import {
  createOfficeParseSession,
  type ParsedOfficeFile,
} from 'office-file-viewer';

export async function parseFile(file: File): Promise<ParsedOfficeFile> {
  const session = createOfficeParseSession(file, { worker: 'auto' });
  const unsubscribe = session.subscribe((progress) => {
    console.info(progress.stage, progress.percent, progress.message);
  });

  try {
    return await session.result;
  } finally {
    unsubscribe();
    session.dispose();
  }
}
```

始终在 `finally` 中取消订阅并调用 `session.dispose()`。会话仍在运行时调用 `dispose()` 也会请求取消。它会清理会话运行时与监听器，但成功返回的解析结果拥有独立的资源生命周期。

### 解析结果资源

宿主不再使用解析结果后，调用统一的异步释放函数：

```ts | pure
import { disposeParsedOfficeFile } from 'office-file-viewer';
import { parseFile } from './parse-file';

const parsed = await parseFile(file);

try {
  console.info(parsed.kind);
  // 在这里读取或渲染解析模型。
} finally {
  await disposeParsedOfficeFile(parsed);
}
```

`OfficeFileViewer` 会自动管理自身会话和解析资源。使用底层 API 时，这两个生命周期都由调用方负责。专用集成也可以使用 `disposeDocDocument`、`disposePresentationDocument` 和 `disposeSpreadsheetWorkbook`，但优先使用 `disposeParsedOfficeFile`，因为它还会释放完整结果附带的文档会话。

## 支持格式与交互

| 文档类型           | 扩展名  | 主要解析范围                                                                                        |
| ------------------ | ------- | --------------------------------------------------------------------------------------------------- |
| Word OOXML         | `.docx` | 段落、列表、表格、图片、图表、形状、链接、样式和主题颜色                                            |
| Word 97-2003       | `.doc`  | 二进制文档结构、文本运行、表格、列表、格式、图片、字段链接、正文书签和基础页面级 OfficeArt 浮动分层 |
| WPS Writer         | `.wps`  | 复用 DOC 二进制管线，优先还原可读内容、字段链接、正文书签和文档资源                                 |
| Excel OOXML        | `.xlsx` | 工作表、值、样式、合并单元格、尺寸、浮动图片、图表和链接                                            |
| Excel 97-2003      | `.xls`  | BIFF8 单元格、格式、合并区域、尺寸、OfficeArt 图片、图表和链接                                      |
| PowerPoint OOXML   | `.pptx` | 母版/版式继承、文本、形状、图片、表格、链接、背景、效果、备注、动态字段和常用图表                   |
| PowerPoint 97-2003 | `.ppt`  | 二进制记录、母版、文本、形状、图片、链接、嵌入图表、兼容文本、演讲者备注和静态回退                  |

常见图表覆盖折线图、柱状图、饼图、圆环图、面积图、散点图、气泡图、雷达图和地图。无法解析或损坏的图表内容会尽可能使用文档内嵌快照。

预览器交互能力包括：

- 支持 `25%` 到 `300%` 缩放、手动输入数值、常用档位，以及每次 `10%` 的放大/缩小操作。
- DOC/DOCX/WPS 与 XLS/XLSX 的可见内容图片支持双击预览、缩放、旋转、下载和自定义右键菜单。
- DOC/DOCX/WPS 大纲默认隐藏，仅在存在可用标题时提供开关，展开后可以左右调整宽度。
- XLS/XLSX 支持工作表标签，以及原始版式与阅读模式切换。
- PPT/PPTX 支持幻灯片与缩略图导航；只有一页时隐藏上一页/下一页。
- 演示文档存在备注时，可以开关演讲者备注并上下调整备注区域高度。
- 支持浏览器全屏，按 `Esc` 退出后会自动同步状态。

### 电子表格显示模式

显示模式选择器仅在 XLS/XLSX 预览中出现：

- **原始版式**为默认模式，保留源文件的行高、列宽、换行、缩小字体填充、合并单元格和内容裁切规则。
- **阅读模式**保留列宽，对长文本自动换行并按需增大行高，以便完整阅读单元格内容。页面版式可能与源文件不同，但不会修改值、公式、工作簿结构或源文件。
- 在同一文件内切换工作表时保留当前模式；打开其他文件时重置为原始版式。
- 对于大型工作表，阅读模式只计算已加载区域的布局调整，不扫描整个工作簿。

<a id="limitations"></a>

## 限制、性能与安全

`office-file-viewer` 是独立实现的解析与渲染引擎，并非 Microsoft Office 或 WPS Office 的原生排版引擎，因此不能保证所有 Office 版本、厂商扩展或复杂文档都达到像素级一致。

### 渲染边界

- 旧格式 DOC/WPS、XLS 和 PPT 优先保证内容可读。复杂分页、锚点、文字环绕、OfficeArt 效果和动画可能与桌面应用不同。
- DOC/WPS 的页面级 OfficeArt 画布当前以单张 SVG 图片显示，无法精确保留画布内每个子形状的独立链接点击区；这类链接会产生非阻断降级提示，字段链接和正文书签不受影响。
- OOXML 文件可能包含尚未支持的宏、ActiveX 控件、OLE 对象、SmartArt、厂商扩展或复杂动画。这些内容可能被忽略、降级，或使用内嵌静态快照显示。
- 预览器为只读组件，不提供编辑、保存、格式转换、打印排版，也不把 Office 文件导出为 PDF 或图片。
- 外部链接图片和动态地图数据仍可能需要网络。地图数据失败时会优先使用内嵌快照；没有快照时显示明确失败状态。

### 浏览器与性能边界

组件面向现代浏览器，依赖 `File`、`fetch`、`DOMParser`、`AbortController`、`IntersectionObserver`、`ResizeObserver`、Blob URL、Canvas、Web Worker 和 Fullscreen API 等能力。

Worker 可以减少受支持旧格式在主线程上的解析工作，但不会消除源文件或解析模型占用的内存。DOCX、XLSX 和 PPTX 当前在主线程解析，因此超大或复杂文件可能短暂降低界面响应速度。

组件默认不限制源文件大小、ZIP 条目数量、单条目大小或解压后总量。处理不可信文件时，宿主应通过 `parseOptions.resourcePolicy` 按目标设备和安全模型设置限制。

### 不可信文件与远程访问

- 解析不可信输入前，校验文件大小、扩展名、MIME 类型和来源。
- 业务需要时增加服务端恶意文件扫描或内容策略检查。
- 远程文件仍受浏览器 CORS、身份认证和内容安全策略约束；预览器不会代理下载或绕过这些控制。
- 本地文件解析在浏览器内完成且不会主动上传，但宿主应用仍需对自己的日志、遥测和数据处理策略负责。
