# 单 npm 包的 Office Core、组合式外壳与插件扩展设计

状态：已确认

日期：2026-09-07

## 1. 背景

当前 `office-file-viewer` 已提供 DOC、DOCX、WPS、XLS、XLSX、PPT、PPTX、PPTM、POTX 等格式的浏览器预览能力，并具备 Worker、渐进解析、按需数据源、资源生命周期、主题、水印、工具栏扩展和结构化错误等能力。

现有包以 `OfficeFileViewer` 为主要入口，解析器、资源仓库、格式模型、布局外壳和格式渲染器已经形成可复用基础，但部分能力仍位于内部 `services`、`formats` 和 `shell` 目录。未来还计划增加独立的 Office Web Editor，因此需要在不立即拆分多个 npm 包的前提下，建立稳定的共享边界。

## 2. 设计目标

- 继续提供简单、兼容的 `OfficeFileViewer` 高层入口。
- 让复杂宿主可以组合工具栏、视口、侧栏、状态栏和自定义内容。
- 让宿主可以通过稳定插件协议增加格式、解析器、渲染器和导出器。
- 对外共享解析、模型、资源、生命周期、布局契约和导出契约。
- 未来 `office-web-editor` 可以复用共享能力，而不依赖 Viewer 内部实现。
- 保持当前单一 npm 包、ESM-only、React `>=16.9` 和按需加载策略。
- 能力对象沿用 `previewKind` 命名，文档快照使用 `format`，由适配层统一映射。
- 通过稳定子路径隔离公共 API，避免宿主依赖未文档化的深层路径。

## 3. 非目标

- 本阶段不实现完整 Office Web Editor。
- 本阶段不立即拆分多个 npm 包或多个仓库。
- 本阶段不承诺 DOC、XLS、PPT、WPS 的浏览器端完整原格式写回。
- 本阶段不开放内部 Worker 消息、DOM 结构或格式解析私有类。
- 本阶段不引入实时协作、CRDT 或服务端编辑引擎。

## 4. 总体决策

采用单包、多层、稳定子路径和内部插件化的结构：

```text
office-file-viewer
├── 高层 Viewer 入口
├── core       解析、模型、资源、生命周期和能力
├── layout     React 组合式预览外壳
├── plugins    插件协议和注册表
└── export     原始文件及模型导出协议
```

内置格式也通过内部插件注册，但默认由 `OfficeFileViewer` 自动装配，普通宿主不需要理解插件机制。

未来编辑器依赖 `core`、`layout`、`plugins` 和 `export`，不依赖 Viewer 的格式渲染实现。

## 5. 依赖方向

```text
core
 ↑
 ├── layout
 ├── plugins
 ├── viewer
 └── future office-web-editor
```

- `core` 不依赖 React、Viewer UI 或具体编辑器。
- `layout` 依赖 React 和 Core 的布局契约，不依赖具体 Office 格式。
- `plugins` 只定义和管理稳定协议，不直接操作 DOM。
- `viewer` 组合 Core、Layout、内置插件和格式渲染器。
- `office-web-editor` 复用 Core、Layout、Plugins、Export，并自行实现编辑交互。
- 禁止 Core 反向依赖 Viewer，禁止 Editor 依赖 Viewer 内部模块。

## 6. 稳定公共入口

第一阶段只承诺以下入口：

```text
office-file-viewer
office-file-viewer/core
office-file-viewer/layout
office-file-viewer/plugins
office-file-viewer/export
```

根入口继续导出 `OfficeFileViewer` 及常用类型。`layout`、`core`、`plugins` 和 `export` 入口必须分别生成完整的 ESM 文件和 TypeScript 声明。

Core 初期从一个总入口导出模型、解析、资源、生命周期、错误和能力类型；只有在包体积或编辑器按需加载确实需要时，才增加 `core/model`、`core/parsing`、`core/resources` 和 `core/export` 子路径。

任何 `services`、`formats`、`shared` 或 `dist` 深层路径都不属于兼容性承诺。

## 7. Core 模型分层

不把当前渲染器模型直接当作编辑器模型，采用三层结构。

### 7.1 Source Model

保留源文件相关信息：

- 原始格式和部件路径；
- 源对象与节点映射；
- 资源引用；
- 未知部件的保留句柄；
- 页面、工作表和幻灯片按需读取信息。

### 7.2 Document Snapshot

面向 Viewer 和未来 Editor 的只读语义快照，要求：

- 节点 ID 稳定；
- 可跨 Worker 序列化；
- 文本范围和几何区域可定位；
- 样式、链接、批注、备注和资源使用引用；
- 保留格式扩展数据；
- 不把 Blob URL 直接写入模型。

建议提供统一外壳：

```ts
type OfficeDocumentSnapshot = {
  id: string;
  format: PreviewKind;
  family: PreviewFamily;
  metadata: OfficeDocumentMetadata;
  capabilities: OfficeCapabilities;
  resources: readonly OfficeResourceRef[];
  warnings?: readonly OfficeWarning[];
  sourceMap?: OfficeSourceMap;
  extensionData?: Readonly<Record<string, OfficeJsonValue>>;
};
```

`extensionData` 和变更操作中的值限定为 JSON-like 数据，禁止函数、DOM 节点、Blob 和循环引用，以保证快照与变更集可以结构化序列化。

解析会话在需要携带格式专属物化模型或按需 Source 时返回运行时封装：`{ snapshot: OfficeDocumentSnapshot, mode: 'materialized', model }` 或 `{ snapshot: OfficeDocumentSnapshot, mode: 'source', source }`。`OfficeDocumentSnapshot` 本身保持可跨 Worker 序列化，运行时封装只在所属会话内存活，资源由会话释放。

```ts
type OfficeDocumentRuntime =
  | {
      snapshot: OfficeDocumentSnapshot;
      mode: 'materialized';
      model: OfficeDocumentModel;
    }
  | {
      snapshot: OfficeDocumentSnapshot;
      mode: 'source';
      source: OfficeDocumentSource;
    };
```

Core 同时提供 `createOfficeSourceRuntime(snapshot, source)`，让外部 Source 插件不需要构造未公开的联合字段。

Word、Excel、PowerPoint 继续保留格式专属模型，不强行抹平格式差异。

### 7.3 Editable Model

未来编辑器再增加：

- `OfficeChangeSet`；
- `OfficeTransaction`；
- `OfficeCommand`；
- `OfficeHistory`；
- `OfficeSelection`；
- `OfficeSourceMap`。

`OfficeCommandDefinition` 的 `execute` 只在 Editor 运行时使用，不能放入可序列化的 `OfficeChangeSet`。

编辑状态不写回只读快照，导出器消费快照和变更集，避免 Viewer 承担编辑状态。

## 8. 共享解析能力

`office-file-viewer/core` 对外共享以下稳定契约：

- `createOfficeParseSession`；
- `OfficeParseSession`；
- `OfficeDocumentSource`；
- `OfficeParser`、`OfficeParserInput` 和最小 `OfficePluginResolver`；
- `OfficeFormatPlugin`；
- `OfficeParseOptions`、`WorkerMode` 和资源限制；
- `ParseProgress`；
- 结构化错误、警告和错误码；
- 取消、重试和释放协议。

插件解析器接收不含 UI 和内部 Worker 消息的输入：

```ts
type OfficeParserInput = {
  file: File;
  sessionId: string;
  signal: AbortSignal;
  options: Omit<OfficeParseOptions, 'pluginRegistry'>;
  resources: OfficeResourceSession;
  documentSession: OfficeDocumentSession;
};

type OfficeWorkerParserInput = {
  file: File;
  sessionId: string;
  options: Omit<OfficeParseOptions, 'pluginRegistry' | 'workerFactory'>;
};
```

内置格式继续按需加载：

```text
DOC / DOCX / WPS
XLS / XLSX
PPT / PPTX / PPTM / POTX
```

Worker 消息结构、内部适配器和具体解析类保持私有。公开的是解析会话和插件契约，避免宿主绑定内部协议。

大文件策略保持现有方向：

- `auto / always / never` Worker 模式；
- 渐进式解析；
- 页面、工作表和幻灯片按需数据源；
- `AbortSignal`；
- 资源和时间限制；
- 解析会话统一释放。

## 9. 共享资源能力

`core` 提供与 React 无关的资源层：

- `OfficeResourceSource`；
- `OfficeResourceStore`；
- `OfficeResourceSession`；
- `OfficeResourceRef`；
- Blob、Object URL、图片、媒体和字体生命周期；
- 缓存和引用计数；
- 资源限制；
- 宿主资源解析器；
- Worker 与主线程资源转移。

资源加载必须支持取消、错误转换、MIME 校验和释放。React Provider 只放在 Viewer 或 Layout 层，Core 不依赖 React。
`OfficeResourceSession.resolve(source, signal)` 是宿主资源解析的公共入口；`acquire()`/`release()` 负责可见资源的引用计数，默认复用内置加载逻辑。宿主 resolver 只替换加载来源，不改变引用计数和释放责任。
快照中的 `OfficeResourceRef` 只保存可序列化定位信息（资源 ID、MIME、大小和 URL/惰性键），不携带 `load` 函数、Blob 或长期 Object URL；实际资源对象由 `OfficeResourceSession` 持有。

## 10. 共享导出能力

导出分为原始文件导出和修改后文件导出。

### 10.1 原始文件导出

```ts
exportOriginalOfficeFile(source): Promise<OfficeExportResult>
```

适用于所有当前支持格式，不修改源文件内容。

### 10.2 模型导出

```ts
type OfficeExportRequest = {
  document: OfficeDocumentSnapshot;
  changes?: OfficeChangeSet;
  format?: string;
  signal?: AbortSignal;
  preserveUnknownParts?: boolean;
};

type OfficeExportResult = {
  blob: Blob;
  fileName: string;
  format: string;
  warnings?: readonly OfficeWarning[];
};
```

导出器按能力声明是否支持修改后写回，不支持的特性必须返回结构化警告或错误。

首批目标格式为 XLSX、DOCX、PPTX。DOC、XLS、PPT 和 WPS 初期保持只读、转换导出或交给服务端处理。

宏格式不执行宏代码，尽量保留未修改的宏部件，并提示重新打包可能导致数字签名失效。

## 11. 兼容式 Viewer 入口

`OfficeFileViewer` 继续是默认入口：

```tsx
<OfficeFileViewer uri={file} />
```

现有 Props 保持兼容，包括 `uri`、`toolbar`、`theme`、`watermark`、`imagePreview`、`search`、`review`、`parseOptions`、`viewState`、`onWarning` 和 `onError`。

新增控制能力：

```ts
type OfficeViewerHandle = {
  openFile(file: File): void;
  reload(): void;
  setZoom(value: number): void;
  zoomIn(): void;
  zoomOut(): void;
  fitWidth(): void;
  fitPage(): void;
  goToPage(index: number): void;
  goToSheet(id: string): void;
  goToSlide(index: number): void;
  toggleFullscreen(): void;
  openSearch(): void;
  openReview(): void;
  getViewState(): OfficeFileViewerViewState;
  getCapabilities(): OfficeCapabilities | undefined;
};
```

受控模式下，句柄动作必须通过 `onViewStateChange` 与宿主状态协作，不能绕过受控值。

## 12. 能力描述

Viewer、Plugin、Layout 和 Editor 共享能力类型：

```ts
type OfficeCapabilities = {
  previewKind: PreviewKind;
  family: PreviewFamily;
  canSearch: boolean;
  canReview: boolean;
  canNavigatePages: boolean;
  canNavigateSheets: boolean;
  canNavigateSlides: boolean;
  canShowSpeakerNotes: boolean;
  canFitPage: boolean;
  canFitWidth: boolean;
  canFullscreen: boolean;
  canExportOriginal: boolean;
  exportFormats: readonly string[];
  features: Readonly<Record<string, boolean>>;
};
```

能力应附加到首屏就绪信息和 Layout Context，宿主不根据格式字符串自行推断功能。

## 13. 组合式 Layout

保留当前 `OfficeViewerLayout` 作为默认封装，并增加 `OfficeViewerShell`：

```tsx
<OfficeViewerShell.Root>
  <OfficeViewerShell.Toolbar>
    <OfficeViewerShell.FileInfo />
    <OfficeViewerShell.Zoom />
    <HostToolbarAction />
  </OfficeViewerShell.Toolbar>

  <OfficeViewerShell.Sidebar>
    <HostSidebar />
  </OfficeViewerShell.Sidebar>

  <OfficeViewerShell.Viewport>{children}</OfficeViewerShell.Viewport>
  <OfficeViewerShell.StatusBar>
    <HostStatus />
  </OfficeViewerShell.StatusBar>
</OfficeViewerShell.Root>
```

外壳只管理缩放、全屏、视口、工具栏、侧栏和状态栏，不识别具体 Office 格式。

现有 `useOfficeViewerLayout()` 扩展为稳定的 `state / actions / meta` Context。组合式 API 使用 children 和 Provider，不继续增加大量 `showXxx` 或 `renderXxx` Props。

## 14. 插件体系

内置格式也统一适配为插件，但由 Viewer 自动注册。外部插件采用实例级或 Provider 级注册，不使用全局可变注册表。
同一插件 ID 总是拒绝重复注册；同一扩展名只有在显式优先级不同且最高者唯一时才允许覆盖，注册表释放只清理自身监听，不处置文档资源。

```ts
type OfficeCorePlugin = {
  id: string;
  extensions: readonly string[];
  mimeTypes?: readonly string[];
  priority?: number;
  detect(
    file: File,
    context: { signal: AbortSignal },
  ): boolean | Promise<boolean>;
  parse(input: OfficeParserInput): Promise<OfficeDocumentRuntime>;
  createSource?(
    input: OfficeParserInput & { snapshot?: OfficeDocumentSnapshot },
  ): Promise<OfficeDocumentRuntime | undefined>;
  export?(request: OfficeExportRequest): Promise<OfficeExportResult>;
  capabilities: OfficeCapabilities;
  workerSupport: 'none' | 'main-thread' | 'worker';
  workerFactory?: () => Worker;
  parseInWorker?(
    input: OfficeWorkerParserInput,
    worker: Worker,
  ): Promise<OfficeDocumentRuntime>;
  dispose?(runtime: OfficeDocumentRuntime): void | Promise<void>;
};

type OfficeFormatPlugin = OfficeCorePlugin;
```

Core 插件不依赖 React。Viewer 渲染和 Editor 编辑分别使用适配器：

```ts
type OfficeViewerPluginAdapter = {
  pluginId: string;
  renderDocument: React.ComponentType<OfficeViewerDocumentProps>;
  renderUnsupported?: React.ComponentType<{ fileName: string }>;
};

type OfficeEditorPluginAdapter = {
  pluginId: string;
  createEditorModel(document: OfficeDocumentRuntime): unknown;
  commands: readonly OfficeCommandDefinition[];
};
```

`OfficeViewerDocumentProps` 只包含公开运行时封装、文件名和错误回调；`OfficeCommandDefinition` 属于未来 Editor 的运行时命令描述，不进入 Core 快照。

```ts
type OfficeViewerDocumentProps = {
  runtime: OfficeDocumentRuntime;
  fileName: string;
  onError?: (error: unknown) => void;
};

type OfficeCommandDefinition = {
  id: string;
  label: string;
  execute(): OfficeCommand;
};

type OfficeCommand = {
  id: string;
  label: string;
  changes: OfficeChangeSet;
};
```

插件不能直接修改 Viewer DOM，也不能从远程地址自动执行代码。第一版插件接口标记为实验性，稳定后再承诺长期兼容。

## 15. 未来 Office Web Editor

Editor 作为独立组件库实现，但当前不拆 npm 包。它只依赖公开子路径：

```text
office-file-viewer/core
office-file-viewer/layout
office-file-viewer/plugins
office-file-viewer/export
```

Editor 自己负责光标、选区、编辑命令、撤销重做、修订、协作和编辑 UI；解析、模型、资源、导出和外壳由共享层提供。

Viewer 不引入 Editor 代码，避免 Viewer 主包体积和状态复杂度增长。

## 16. 单包目录规划

先建立公共门面层，不立即搬迁所有现有文件：

```text
src/office-file-viewer/
├── core/
│   ├── index.ts
│   ├── model.ts
│   ├── parsing.ts
│   ├── parsingContracts.ts
│   ├── resources.ts
│   ├── export.ts
│   ├── editing.ts
│   ├── capabilities.ts（未来按需拆分）
│   └── errors.ts（未来按需拆分）
├── shell/layout/（Layout 与 Shell 实现）
├── plugins/
├── export/
├── formats/
├── services/
└── shared/
```

`core/*` 初期可以从现有 `services/*` 稳定 re-export，Viewer 再逐步切换到公共契约。内部实现目录继续保留，不让宿主依赖。

## 17. Package Exports

第一阶段稳定导出：

```json
{
  ".": "./dist/index.js",
  "./core": "./dist/core.js",
  "./layout": "./dist/layout.js",
  "./plugins": "./dist/plugins.js",
  "./export": "./dist/export.js",
  "./styles.css": "./dist/styles.css"
}
```

上面的 JSON 只表达路径映射；实际 `package.json` 对代码入口同时提供 `types`、`import` 和 `default` 条件，并保持 ESM-only，不提供 `require` 条件。CSS 入口只提供 `import`/`default` 可解析的样式文件。

每个入口都提供 ESM 文件和 TypeScript 声明。Core 入口不加载 Viewer UI，Layout 入口不加载全部格式解析器，Plugins 和 Export 按需加载。

只有在包体积或编辑器按需导入有明确需要时，才追加 `core/model`、`core/parsing`、`core/resources` 和 `core/export` 子路径。

## 18. 迁移策略

### 阶段 0：契约设计

冻结模型、资源所有权、解析会话、Worker、插件、导出、Layout Context、能力和错误协议。

### 阶段 1：Core 门面

建立 `core` 公共入口和兼容 re-export，不改变 Viewer 的渲染行为。确认 Core 不依赖 React UI。

### 阶段 2：A 和 B

增加 `OfficeViewerHandle`、`OfficeCapabilities`、统一生命周期、`OfficeViewerShell`、Provider、Toolbar、Viewport、Sidebar、StatusBar 和可替换状态内容。

### 阶段 3：C 和 Export

建立插件注册表、内置格式插件适配、外部 Parser/Source/Renderer 协议、原始文件导出和首批 OOXML 导出契约。

### 阶段 4：Editor 基础

增加 Change Set、Transaction、Command、History、Selection 和 Editor Adapter，为未来 `office-web-editor` 做准备，但不把编辑器实现打入 Viewer。

## 19. 兼容性、性能与安全

- 保持 React `>=16.9`，使用 `forwardRef`、Context 和 `useImperativeHandle` 等兼容能力。
- 保持 ESM-only，不提供 CommonJS 构建。
- 不删除现有 Props 和根入口；新增 API 采用向后兼容方式。
- 内置解析器、插件和导出器继续动态加载，避免所有能力进入首屏。
- Core 不引入 React 或 UI 依赖，ECharts 和大型格式模块继续按需加载。
- Layout Context 的 `state / actions / meta` 引用保持稳定，避免宿主内容无关重渲染。
- 插件必须由宿主明确提供，不允许从远程地址自动执行插件代码。
- 宏、ActiveX、OLE 和脚本不执行；导出时提示宏部件和数字签名风险。
- 外部资源必须经过宿主资源策略，遵守取消、超时、MIME 和大小限制。
- 未知格式或未知包部件采用安全降级，并通过结构化警告说明。

## 20. 验证方案

### 公共入口

- 根入口、`/core`、`/layout`、`/plugins`、`/export` 均可导入。
- 各入口具有完整类型声明和 ESM 构建产物。
- 不支持未文档化深层路径。

### 向后兼容

- 现有 `OfficeFileViewer` 和 `OfficeViewerLayout` 示例无需修改即可构建。
- 旧 Props、回调、受控状态和资源释放语义不变。
- React 16.9、17、18 消费端类型检查和浏览器打包通过。

### 生命周期与资源

- 解析取消、Worker 释放、Blob URL 回收、缓存淘汰和文档切换无泄漏。
- 插件 `dispose` 和导出取消语义可重复调用。

### 插件与导出

- 插件格式识别、解析、按需 Source、Renderer Adapter、Exporter 和错误边界。
- 原始文件导出和能力查询。
- XLSX、DOCX、PPTX 的导出回读验证在对应实现阶段完成。

### 性能

- 根入口、Core、Layout、Plugins 和 Export 分别统计包体积。
- 首屏不加载全部格式、Editor 或导出器。
- 大文件仍使用 Worker、按需读取和资源预算。

### 最终全量示例文件烟测门禁

全部实施任务完成、准备交付前，才使用现有 `docs/dev/smoke-test.md` 路由和现有示例文件进行一次完整的逐项验证。中间实施阶段只执行与当前改动直接相关的类型、Lint、构建或局部验证，不要求逐个示例文件烟测。

烟测顺序和要求：

1. 从当前示例文件清单动态读取全部文件，逐个加载并确认文件名、格式识别、加载状态、解析警告和首屏状态。
2. Word、DOCX、WPS：逐页检查文字、字体、字号、行高、段间距、分页、表格、图片、目录、超链接、批注/修订和滚动定位。
3. XLS、XLSX：逐个工作表检查表头、单元格文字、换行、行高、列宽、合并单元格、冻结区域、图片、图表、批注、滚动条和大范围滚动稳定性。
4. PPT、PPTX、PPTM、POTX：逐张幻灯片检查文字、字体、形状、图片、母版、背景、缩略图、导航、备注、超链接、切换和缩放行为。
5. 每个文件记录可复现的问题位置，至少精确到页码、工作表/区域或幻灯片，不用 `看起来正常` 作为验收结论。
6. 发现问题后必须先修复通用解析、模型、渲染或布局逻辑，禁止按文件名、文件 ID 或单个示例添加特例。
7. 修复后先重新检查受影响文件的全部页面/工作表/幻灯片，再从示例清单第一项开始进行全量回归。
8. 只有全部示例逐项通过、没有新增回归问题，并完成类型、Lint、构建和文档检查后，整体任务才可标记完成。

烟测过程只使用现有示例和现有路由，不把私有烟测说明、临时截图或调试资产放入发布包、公共 API 或对外 Release 文案。

## 21. 版本策略

- 当前 `office-file-viewer` 继续使用单一版本号。
- 新增 Core、Layout、Plugin 和 Export 能力计划以次版本发布。
- 插件和导出协议初版标记为实验性，稳定后再承诺长期兼容。
- 将来即使拆成多个 npm 包，当前子路径仍保留兼容转发层。

## 22. 验收标准

1. 普通宿主仍可只使用 `<OfficeFileViewer />`。
2. 高级宿主可以组合工具栏、视口、侧栏和状态栏。
3. 宿主可以通过 Handle 和 Context 控制通用操作。
4. 宿主可以只使用 Core 的解析、模型、资源和生命周期能力。
5. 宿主可以注册自定义格式、Renderer 和 Exporter 插件。
6. 未来 Editor 不依赖 Viewer 内部实现。
7. 现有 API、React 兼容范围、ESM-only 和性能策略保持不变。
8. 全部实施任务完成后，执行一次现有示例文件的逐文件、逐页/逐表/逐幻灯片详细烟测；发现问题修复后重新执行全量回归。

## 23. 明确不做的事情

- 不将 Viewer 和 Editor 合并成一个巨型组件。
- 不让宿主修改 Viewer 内部 DOM 或依赖内部 CSS 类名。
- 不继续为每项能力增加顶层布尔 Props。
- 不直接开放 `services`、`formats`、`shared` 的内部文件路径。
- 不立即实现全部旧版二进制格式和 WPS 的浏览器端原格式写回。
- 不在 Core 中引入 UI 依赖或编辑器代码。

## 24. 规格自检

- 目标、非目标和依赖方向一致；当前仍是单 npm 包。
- A、B、C 三条路线均有独立职责，且由 Core 连接。
- 解析、模型、资源、布局和导出都有明确共享边界。
- 未来 Editor 不依赖 Viewer 内部实现。
- 没有要求宿主使用未文档化深层导入。
- 没有引入 CommonJS、远程插件执行或宏执行要求。
- 导出能力按格式分阶段，不承诺无法保证的全格式写回。
- 验证范围包含入口、兼容性、生命周期、插件、导出和性能。

本规格不包含具体实现代码。规格复核通过后，再使用 `superpowers:writing-plans` 拆分实施任务。
