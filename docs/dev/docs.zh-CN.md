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
      uri={async () => fetch('/files/demo.xlsx')}
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
type OfficeFileViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);
```

远程来源规则：

- 跨域 URL 必须允许浏览器通过 CORS 访问。
- URL 最好以受支持的 Office 扩展名结尾。
- 无扩展名 URL 必须通过 `Content-Disposition` 提供文件名，或通过 `Content-Type` 提供受支持的 MIME 类型。
- URL 已显式包含不支持的扩展名时，会在下载前拒绝，即使响应体实际包含 Office 文件。
- `uri` 变化时，之前的 URL 下载会通过 `AbortController` 取消。
- 自定义异步加载函数无法被强制中止，但过期结果不会覆盖最新来源。
- 手动选择文件也会使进行中的远程加载或解析结果失效。

<a id="component-api"></a>

## `OfficeFileViewer` API

| 属性                             | 类型                                                 | 默认值     | 说明                                   |
| -------------------------------- | ---------------------------------------------------- | ---------- | -------------------------------------- |
| `locale`                         | `'zh-CN' \| 'en-US'`                                 | `'zh-CN'`  | 预览器内置界面语言                     |
| `uri`                            | `OfficeFileViewerUri`                                | -          | 预加载文件来源；未传时显示文件选择器   |
| `defaultFileName`                | `string`                                             | 本地化文案 | 来源无法提供有效文件名时使用的备用名称 |
| `defaultZoom`                    | `number`                                             | `100`      | 初始缩放百分比，限制在 `25` 到 `300`   |
| `defaultShowSpeakerNotes`        | `boolean`                                            | `false`    | 非受控模式下演讲者备注的初始状态       |
| `showSpeakerNotes`               | `boolean`                                            | -          | 受控模式下演讲者备注是否显示           |
| `onSpeakerNotesVisibilityChange` | `(visible: boolean) => void`                         | -          | 演讲者备注显示状态变化时触发           |
| `className`                      | `string`                                             | -          | 预览器根节点附加类名                   |
| `height`                         | `CSSProperties['height']`                            | 跟随父容器 | 预览器高度；优先级高于 `style.height`  |
| `style`                          | `CSSProperties`                                      | -          | 预览器根节点内联样式                   |
| `onFileParsed`                   | `(parsed: ParsedOfficeFile, file: File) => void`     | -          | 完整实体化解析结果可用时触发一次       |
| `onPreviewReady`                 | `(info: OfficePreviewReadyInfo, file: File) => void` | -          | 首个可用预览就绪时触发一次             |
| `onError`                        | `(error: Error, file?: File) => void`                | -          | 加载、解析或预览器操作失败时触发       |
| `parseOptions`                   | `OfficeParseOptions`                                 | `{}`       | Worker 策略与可选 Worker 工厂          |
| `onParseProgress`                | `(progress: ParseProgress) => void`                  | -          | 解析阶段或完成度变化时触发             |

### 受控演讲者备注

受控模式使用 `showSpeakerNotes` 和 `onSpeakerNotesVisibilityChange`；非受控模式只用 `defaultShowSpeakerNotes` 指定初始状态。该控制仅在 PPT/PPTX 文件存在演讲者备注时有实际内容可显示。

### 回调时机

- `onPreviewReady` 在完整模型（`mode: 'materialized'`）或按需数据源（`mode: 'source'`）能够渲染首个可用预览时触发。
- `onFileParsed` 在完整实体化结果就绪后触发；渐进解析的中间结果不会触发它。
- `onParseProgress` 可以多次触发，且可能不包含精确数量或 `percent`。
- 错误发生在有效文件解析出来之前时，`onError` 可能收不到 `file`。

## 解析配置与进度

```ts | pure
type WorkerMode = 'auto' | 'always' | 'never';

type OfficeParseOptions = {
  worker?: WorkerMode;
  workerFactory?: () => Worker;
};
```

| 模式       | DOC/WPS、XLS、PPT                       | DOCX、XLSX、PPTX                    |
| ---------- | --------------------------------------- | ----------------------------------- |
| `'auto'`   | 优先使用 Worker，必要时回退到主线程     | 当前在主线程解析                    |
| `'always'` | 必须创建兼容 Worker，无法创建时直接失败 | Worker 迁移尚未完成，会抛出配置错误 |
| `'never'`  | 始终在主线程解析                        | 始终在主线程解析                    |

每个正在运行的旧格式解析会话拥有独立 Worker。来源切换或组件卸载时，预览器会取消并释放当前会话。`workerFactory` 主要用于需要特殊资源路径或内容安全策略的宿主；大多数应用应使用内置工厂。

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

| 文档类型           | 扩展名  | 主要解析范围                                                                                     |
| ------------------ | ------- | ------------------------------------------------------------------------------------------------ |
| Word OOXML         | `.docx` | 段落、列表、表格、图片、图表、形状、链接、样式和主题颜色                                         |
| Word 97-2003       | `.doc`  | 二进制文档结构、文本运行、表格、列表、格式、图片和基础页面级 OfficeArt 浮动分层                  |
| WPS Writer         | `.wps`  | 复用 DOC 二进制管线，优先还原可读内容和文档资源                                                  |
| Excel OOXML        | `.xlsx` | 工作表、值、样式、合并单元格、尺寸、浮动图片和图表                                               |
| Excel 97-2003      | `.xls`  | BIFF8 单元格、格式、合并区域、尺寸、OfficeArt 图片和图表                                         |
| PowerPoint OOXML   | `.pptx` | 母版/版式继承、文本、形状、图片、表格、背景、效果、演讲者备注、幻灯片编号/日期时间字段和常用图表 |
| PowerPoint 97-2003 | `.ppt`  | 二进制记录、母版、文本、形状、图片、嵌入图表、内嵌 DrawingML 兼容文本、演讲者备注和静态回退      |

常见图表覆盖折线图、柱状图、饼图、圆环图、面积图、散点图、气泡图、雷达图和地图。无法解析或损坏的图表内容会尽可能使用文档内嵌快照。

预览器交互能力包括：

- 支持 `25%` 到 `300%` 缩放、手动输入数值、常用档位，以及每次 `10%` 的放大/缩小操作。
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
- OOXML 文件可能包含尚未支持的宏、ActiveX 控件、OLE 对象、SmartArt、厂商扩展或复杂动画。这些内容可能被忽略、降级，或使用内嵌静态快照显示。
- 预览器为只读组件，不提供编辑、保存、格式转换、打印排版，也不把 Office 文件导出为 PDF 或图片。
- 外部链接图片和动态地图数据仍可能需要网络。地图数据失败时会优先使用内嵌快照；没有快照时显示明确失败状态。

### 浏览器与性能边界

组件面向现代浏览器，依赖 `File`、`fetch`、`DOMParser`、`AbortController`、`IntersectionObserver`、`ResizeObserver`、Blob URL、Canvas、Web Worker 和 Fullscreen API 等能力。

Worker 可以减少受支持旧格式在主线程上的解析工作，但不会消除源文件或解析模型占用的内存。DOCX、XLSX 和 PPTX 当前在主线程解析，因此超大或复杂文件可能短暂降低界面响应速度。

组件本身不限制源文件大小、ZIP 条目数量、单条目大小或解压后总量。宿主应根据目标设备和安全模型设置合适限制。

### 不可信文件与远程访问

- 解析不可信输入前，校验文件大小、扩展名、MIME 类型和来源。
- 业务需要时增加服务端恶意文件扫描或内容策略检查。
- 远程文件仍受浏览器 CORS、身份认证和内容安全策略约束；预览器不会代理下载或绕过这些控制。
- 本地文件解析在浏览器内完成且不会主动上传，但宿主应用仍需对自己的日志、遥测和数据处理策略负责。
