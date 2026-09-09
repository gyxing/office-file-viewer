# 单包 Office Core 与可扩展 Viewer 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:executing-plans` to implement this plan task-by-task in the current session. Steps use checkbox syntax for tracking.

**Goal:** 在保持单一 `office-file-viewer` npm 包和现有接入方式兼容的前提下，公开可复用的 Core、组合式 Layout、插件和导出契约，为未来 `office-web-editor` 提供稳定基础。

**Architecture:** 以现有 `services`、`shell` 和 `formats` 实现为基础增加公共门面层，不做一次性目录搬迁。根入口继续提供开箱即用的 `OfficeFileViewer`，`core`、`layout`、`plugins` 和 `export` 通过稳定子路径对外提供能力，内置格式通过内部注册表统一装配，未来 Editor 只依赖这些公开契约。

**Tech Stack:** TypeScript、React `>=16.9.0`、ReactDOM `>=16.9.0`、Less、Yarn 1.22、Father、Dumi、Web Worker、zip.js、CFB/OfficeArt 解析器、现有 ESM 构建链。

**Spec:** `docs/superpowers/specs/2026-09-07-single-package-core-extensibility-design.md`

## Global Constraints

- 当前只维护一个 `office-file-viewer` npm 包，不创建或发布多包 workspace。
- 保持 `office-file-viewer`、`office-file-viewer/layout` 和现有公共类型的兼容语义。
- 公共能力对象沿用现有 `previewKind` 命名；文档快照遵循语义模型使用 `format`，适配层负责一次映射，禁止两者在同一对象中漂移。
- 新增 `office-file-viewer/core`、`office-file-viewer/plugins` 和 `office-file-viewer/export` 稳定入口；禁止依赖未文档化的深层路径。
- Core 不依赖 React、Viewer UI 或编辑器；Layout 不依赖具体 Office 格式。
- 保持 React `>=16.9.0` 和 ESM-only；不使用 React 19 专属 API，不添加 CommonJS 构建。
- 不新增运行时依赖；继续使用现有解析器、资源 Store、Worker、zip.js、ECharts 按需加载策略。
- 不改变现有 Word、Excel、PowerPoint 的解析和视觉还原算法，除非公共契约适配确实需要且经过单独验证。
- 不为形式上的分层拆分现有各格式解析器；格式实现继续按现有族和动态 loader 组织，只新增必要的适配门面。
- 只有文件同时承担两个以上可独立验证的职责且体量明显超过当前模块平均值时才拆分；本计划新增文件每个只承担一个公共边界，不回迁或复制既有解析逻辑。
- 不继续增加孤立的 `showXxx`、`enableXxx` 或 `renderXxx` 顶层 Props；优先使用 Context、Provider、children 和 Compound Components。
- 插件使用实例级或 Provider 级注册，不使用全局可变注册表，不执行远程插件代码。
- `plugins`、`export` 和 Editor 预留类型首版标记为 `@experimental`，Core 基础解析/资源/生命周期和根 Viewer API 才作为稳定兼容面。
- 解析、资源、导出和插件生命周期必须支持取消、错误转换、释放和重复调用安全。
- 新增公开类型、属性、枚举映射和业务常量必须补简短中文 TSDoc，说明用途或约束，不在注释中重复属性名。
- 中间任务只做直接相关的类型、Lint、构建或局部验证；全部实施任务完成后才进行一次现有示例文件全量详细烟测。
- 最终烟测必须逐文件、逐页/逐表/逐幻灯片执行；发现问题先修复通用逻辑，再重新全量回归。
- 不新增独立烟测文档、私有截图或发布测试资产；现有 `docs/dev/smoke-test.md` 和示例清单只读使用。
- 不执行 Git commit、push、pull、merge、rebase 或分支操作，除非用户另行明确授权。
- 不启用子智能体；按任务顺序在当前会话执行，并在每项任务后重新读取修改文件。
- 本计划不自动修改版本号、不发布 npm 或 GitHub Release；全部实现和最终烟测通过后，再由用户单独决定下一次次版本号。

---

## 执行记录（2026-09-08）

本计划列出的 Task 1-14 已在当前工作区完成。已落地 Core 模型/解析/资源/生命周期门面、导出契约与原始文件导出、作用域级插件注册表、Viewer Provider 接入、控制句柄、组合式 Shell、状态插槽、Editor 类型契约，以及 `core`、`layout`、`plugins`、`export` 和 `styles.css` 公共入口。执行过程中发现并修正了 Core 类型触发 lint、未知扩展名无法命中插件索引、Provider 解析结果未进入 PreviewStage、内置插件嵌套会话重复持有资源、稳定 ESM 相对路径和聚合样式重复包含等问题。

已完成项包括：`goToPage` 接入现有分页导航控制器并增加不支持/越界错误；内置格式统一适配；中英文 README 与开发文档补齐稳定入口、扩展边界和迁移说明；公共入口浏览器打包、React 16.9/17/18 消费者和包体积探针；发布脚本补齐严格 ESM 的相对 `.js` 路径和 `styles.css` 聚合处理。Task 14 使用 Codex CUA In-app Browser 在 `http://localhost:8001/office-file-viewer/dev/smoke-test` 完成 32/32 示例加载回归：17 个 Word 文件连续滚动与专项交互、8 个 Excel 文件共 59 个工作表逐表四向滚动、7 个 PowerPoint 文件共 121 张幻灯片逐张切换；同时验证 Worker 三种模式、Shell 缩放/全屏、主题/水印、插件内容外壳、超链接、图片菜单和日志。详细记录保存在系统临时文件 `office-viewer-final-qa-20260908.md`；缺少 PPTM、POTX 等本地样本及外部 Office/PDF 像素基准，已在记录中注明。

失败探针和临时消费者仅用于验证错误分支，验证后均已清理；复选框按实际完成状态更新。未执行 Git 提交、发布或其他历史写操作。

增量复核（2026-09-09）：针对最终审计又补齐了物化/按需模型的资源引用与警告快照、资源会话的会话级取消信号、Shell 的侧栏网格布局、卸载后 Handle 无副作用、插件/导出器输入校验和自定义状态插槽的懒加载回退；同时补充了布局类型导出、双语 Handle/slots/扩展入口文档，并重新通过类型、Lint、样式和 Father 声明生成检查。

最终收口复核（2026-09-09）：修正 `slots.statusBar` 与文档约定不一致的并行渲染问题，补充 Provider 在 Registry 切换和 StrictMode effect 重放下的释放保护；插件解析避免二次 `detect()`，补齐同优先级歧义检测、无点扩展名规范化、强制 Worker 能力边界和外层取消到资源会话的传播；Core 快照与能力对象完成运行时只读冻结。最新构建冷启动的 Codex CUA 标签页 error/warn 为 0，代表性 DOC/DOCX/XLS/PPT 均通过。

最终门禁后复跑（2026-09-09）：完整 `yarn run check` 触发开发页热更新后，在同一 Codex CUA 标签页按顺序重新加载 32 个现有 Office 样本，四批均为 8/8 就绪，最终 PPTX 状态正常且控制台 error/warn 为 0。

## 文件结构与职责

### 新增公共门面

- Create: `src/core.ts` - `office-file-viewer/core` 的顶层入口。
- Create: `src/plugins.ts` - `office-file-viewer/plugins` 的顶层入口。
- Create: `src/export.ts` - `office-file-viewer/export` 的顶层入口。
- Modify: `src/layout.ts` - 扩展 Layout 稳定导出，保留现有导出。
- Modify: `src/office-file-viewer/index.ts` - 根入口补充公共类型和兼容导出。

### Core 契约

- Create: `src/office-file-viewer/core/types.ts` - 文档、能力、资源、解析和导出公共类型。
- Create: `src/office-file-viewer/core/model.ts` - 现有格式模型的公共适配和快照封装。
- Create: `src/office-file-viewer/core/resources.ts` - 资源 Store、资源会话和生命周期门面。
- Create: `src/office-file-viewer/core/parsing.ts` - 解析会话、Source 和解析器契约门面。
- Create: `src/office-file-viewer/core/parsingContracts.ts` - 不依赖解析实现的 Parser/Plugin 低层契约，避免服务层反向导入门面。
- Create: `src/office-file-viewer/core/parsingSession.ts` - 将现有解析会话适配为公开运行时快照。
- Create: `src/office-file-viewer/core/export.ts` - 导出请求、结果和导出能力门面。
- Create: `src/office-file-viewer/core/editing.ts` - 未来 Editor 使用的变更、命令、事务、历史和选区契约。
- Create: `src/office-file-viewer/core/index.ts` - Core 统一导出。

### 插件与导出实现

- Create: `src/office-file-viewer/plugins/types.ts` - Core 插件和注册表契约。
- Create: `src/office-file-viewer/plugins/viewerTypes.ts` - Viewer Renderer Adapter 契约。
- Create: `src/office-file-viewer/plugins/OfficePluginRegistry.ts` - 作用域级插件注册表。
- Create: `src/office-file-viewer/plugins/index.ts` - 插件公共入口。
- Create: `src/office-file-viewer/export/types.ts` - 导出器契约和导出错误类型。
- Create: `src/office-file-viewer/export/OfficeExportService.ts` - 原始文件导出和导出器调度。
- Create: `src/office-file-viewer/export/index.ts` - 导出公共入口。

### Viewer 与 Layout 扩展

- Create: `src/office-file-viewer/shell/controller/OfficeViewerHandle.ts` - Viewer 控制句柄类型。
- Create: `src/office-file-viewer/shell/controller/officeViewerCapabilities.ts` - 格式能力计算。
- Create: `src/office-file-viewer/shell/layout/OfficeViewerShell.tsx` - 组合式外壳组件。
- Modify: `src/office-file-viewer/shell/layout/types.ts` - 通用 Shell Context 和组合组件类型。
- Modify: `src/office-file-viewer/shell/layout/OfficeViewerLayoutContext.tsx` - Provider 和 Context 扩展。
- Modify: `src/office-file-viewer/shell/layout/OfficeViewerLayout.tsx` - 兼容封装和公共 Shell 复用。
- Modify: `src/office-file-viewer/shell/layout/index.ts` - Layout 公共导出。
- Modify: `src/office-file-viewer/OfficeFileViewer.tsx` - Handle、能力、插件注册和状态插槽接线。

### 构建、文档和验证

- Modify: `package.json` - 单包 export map 和稳定子路径。
- Modify: `.fatherrc.ts` 或现有构建脚本 - 新入口和样式产物校验。
- Modify: `README.md`、`README.zh-CN.md`、`docs/dev/docs.md`、`docs/dev/docs.zh-CN.md` - 公共接入和扩展文档。
- Read-only verify: `docs/dev/smoke-test.md`、`docs/dev/smoke-test.samples.ts` - 最终全量烟测。

## 实施阶段

1. **共享 Core 基础（Task 1-5）**：先冻结类型、模型、解析、资源、生命周期和导出边界，确保未来组件库有可直接复用的无 UI 能力。
2. **C 路线插件（Task 6-7）**：建立实例级注册表和 Viewer Adapter，内置格式保持按需加载，外部格式可在不改 Viewer 主组件的情况下接入。
3. **A/B 路线 Viewer 与 Layout（Task 8-10）**：提供 React 16.9 兼容的 Handle、能力查询、组合式 Shell、状态插槽和样式入口。
4. **Editor 预留与发布边界（Task 11-13）**：公开可序列化变更契约，完成单包 export map、文档、包体积和兼容性检查，不实现编辑器 UI。
5. **最终质量门禁（Task 14）**：全部实现完成后只执行一次现有示例文件的逐文件详细烟测；修复后从第一项重新全量回归。

## Task 1: 建立公共入口和 Core 类型契约

**Files:**

- Create: `src/office-file-viewer/core/types.ts`
- Create: `src/office-file-viewer/core/index.ts`
- Create: `src/core.ts`
- Modify: `src/office-file-viewer/index.ts`
- Modify: `package.json`

**Interfaces:**

- Consumes: 现有 `PreviewKind`、`PreviewFamily`、`ParsedOfficeFile`、`OfficeResourceSource`、`OfficeParseSession` 和格式专属模型。
- Produces: `OfficeCapabilities`、可序列化的 `OfficeDocumentSnapshot`、会话运行时 `OfficeDocumentRuntime`、`OfficeResourceRef`、`OfficeSourceMap` 以及可导入的 `office-file-viewer/core`。

- [x] **Step 1: 写公共入口失败探针**

在系统临时目录创建 `$env:TEMP\office-core-contract-probe.ts`，内容至少包含：

```ts
import type {
  OfficeCapabilities,
  OfficeDocumentSnapshot,
  OfficeResourceRef,
} from 'D:/workspace/github/office-x-viewer/src/core';

const capabilities: OfficeCapabilities = {
  previewKind: 'xlsx',
  family: 'spreadsheet',
  canSearch: true,
  canReview: false,
  canNavigatePages: false,
  canNavigateSheets: true,
  canNavigateSlides: false,
  canShowSpeakerNotes: false,
  canFitPage: false,
  canFitWidth: true,
  canFullscreen: true,
  canExportOriginal: true,
  exportFormats: ['xlsx'],
  features: { cellOverflow: true },
};
void (capabilities satisfies OfficeCapabilities);
void (undefined as OfficeDocumentSnapshot | undefined);
void (undefined as OfficeResourceRef | undefined);
```

- [x] **Step 2: 运行失败探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-core-contract-probe.ts --bundle --platform=node --format=esm --outfile=$env:TEMP\office-core-contract-probe.mjs
```

Expected: 在新增入口和类型前，出现无法解析 `src/core` 或缺少类型导出的失败。

- [x] **Step 3: 定义最小稳定类型**

在 `core/types.ts` 定义以下契约，字段全部使用明确类型：

```ts
import type { DocDocument } from '../services/doc/types';
import type { DocxDocument } from '../services/docx/types';
import type { DocWordPageSource } from '../services/doc/DocWordPageSource';
import type { DocxPagePreviewSource } from '../services/docx/DocxWordPageSource';
import type { PresentationSource } from '../services/presentation/PresentationSource';
import type { PresentationDocument } from '../services/presentation/types';
import type { PptxDocument } from '../services/pptx/types';
import type { SpreadsheetSource } from '../services/spreadsheet/SpreadsheetSource';
import type { SpreadsheetWorkbook } from '../services/spreadsheet/types';
import type { PreviewFamily, PreviewKind } from '../services/preview';

export type OfficeCapabilities = {
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
  /** 可扩展能力键，新增格式可声明额外功能而不改动基础字段。 */
  features: Readonly<Record<string, boolean>>;
};

export type OfficeResourceRef = {
  id: string;
  kind: 'image' | 'media' | 'font' | 'other';
  /** 可跨 Worker 传输的资源定位信息，不包含 load 函数或 Blob。 */
  locator?: OfficeResourceLocator;
  mimeType?: string;
  size?: number;
};

export type OfficeResourceLocator =
  | Readonly<{ kind: 'url'; url: string }>
  | Readonly<{ kind: 'lazy'; id: string; mimeType: string; size: number }>;

// `url` 只允许源文件声明的稳定地址；渲染层临时生成的 blob URL 不进入快照。

export type OfficeSourceMap = Readonly<Record<string, OfficeSourceLocation>>;

export type OfficeSourceLocation = {
  part?: string;
  objectId?: string;
  offset?: number;
};

export type OfficeDocumentMetadata = {
  fileName?: string;
  fileSize?: number;
  title?: string;
  author?: string;
};

export type OfficeNodeId = string;

export type OfficeJsonValue =
  | string
  | number
  | boolean
  | null
  | readonly OfficeJsonValue[]
  | Readonly<{ [key: string]: OfficeJsonValue }>;

export type OfficeDocumentSnapshotBase = {
  id: string;
  /** 对外稳定的格式字段；内部 ParsedOfficeFile.kind 在适配层映射到这里。 */
  format: PreviewKind;
  family: PreviewFamily;
  metadata: OfficeDocumentMetadata;
  capabilities: OfficeCapabilities;
  resources: readonly OfficeResourceRef[];
  warnings?: readonly OfficeWarning[];
  sourceMap?: OfficeSourceMap;
  extensionData?: Readonly<Record<string, OfficeJsonValue>>;
};

/** 可跨 Worker 传输的只读文档语义快照，不携带运行时对象或释放方法。 */
export type OfficeDocumentSnapshot = Readonly<OfficeDocumentSnapshotBase>;

/** 会话内部把快照与物化模型或按需 Source 绑定，资源由会话统一释放。 */
export type OfficeDocumentRuntime<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> =
  | Readonly<{
      snapshot: OfficeDocumentSnapshot;
      mode: 'materialized';
      model: TModel;
      source?: never;
    }>
  | Readonly<{
      snapshot: OfficeDocumentSnapshot;
      mode: 'source';
      model?: never;
      source: TSource;
    }>;

export type OfficeDocumentModel =
  | DocDocument
  | DocxDocument
  | SpreadsheetWorkbook
  | PresentationDocument
  | PptxDocument;

export type OfficeDocumentSource =
  | DocWordPageSource
  | DocxPagePreviewSource
  | SpreadsheetSource
  | PresentationSource;

export type OfficeWarning = {
  code: string;
  message: string;
  source?: string;
  recoverable?: boolean;
};

export type OfficeChangeSet = Readonly<{
  id: string;
  baseDocumentId: string;
  operations: readonly OfficeChangeOperation[];
}>;

export type OfficeChangeOperation =
  | Readonly<{
      type: 'insert-text';
      nodeId: OfficeNodeId;
      offset: number;
      text: string;
    }>
  | Readonly<{
      type: 'delete-text';
      nodeId: OfficeNodeId;
      offset: number;
      length: number;
    }>
  | Readonly<{
      type: 'replace-node';
      nodeId: OfficeNodeId;
      value: OfficeJsonValue;
    }>;
```

`OfficeDocumentModel` 和 `OfficeDocumentSource` 的导入必须使用 `import type`，确保 Core 入口不因类型引用加载格式运行时代码。运行时模型或按需 Source 只由会话内部持有，不能把带方法的对象放进可跨 Worker 传输的快照；快照的 `format`、元数据、能力、引用和 `extensionData` 必须保持可结构化序列化，扩展数据只允许 JSON-like 值，遇到函数、DOM 节点或循环引用时在适配层丢弃并生成警告。资源释放由 `OfficeDocumentSession` 或解析会话负责，快照本身不增加 `dispose()` 方法。

- [x] **Step 4: 建立 Core 门面**

在 `core/index.ts` 只导出文档化契约和经过审查的现有实现：

```ts
export type {
  OfficeCapabilities,
  OfficeChangeOperation,
  OfficeChangeSet,
  OfficeDocumentMetadata,
  OfficeDocumentModel,
  OfficeDocumentRuntime,
  OfficeDocumentSnapshot,
  OfficeDocumentSource,
  OfficeNodeId,
  OfficeJsonValue,
  OfficeResourceRef,
  OfficeResourceLocator,
  OfficeSourceSnapshotInput,
  OfficeSourceLocation,
  OfficeSourceMap,
  OfficeWarning,
} from './types';
export {
  createOfficeDocumentRuntime,
  createOfficeDocumentSnapshot,
  createOfficeSourceRuntime,
  createOfficeSourceSnapshot,
  getOfficeCapabilities,
  getOfficeDocumentModel,
} from './model';
export type {
  OfficeParser,
  OfficeFormatPlugin,
  OfficeParserInput,
  OfficeParserPlugin,
  OfficePluginResolver,
  OfficeSourceFactory,
  OfficeWorkerParser,
  OfficeWorkerParserInput,
} from './parsing';
export type {
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  OfficeParseResourcePolicy,
  ParseProgress,
  ParseStage,
  WorkerMode,
  OfficeFileViewerErrorCode,
  OfficeFileViewerErrorContext,
  OfficeFileViewerErrorStage,
} from '../services/parsing';
export { createOfficeDocumentParseSession } from './parsingSession';
export {
  createOfficeParseSession,
  isOfficeFileViewerError,
  OfficeFileViewerError,
  OfficeResourceLimitError,
} from '../services/parsing';
export type {
  OfficeResourceSource,
  OfficeResourceStore,
  OfficeResourceStoreOptions,
} from '../services/resource-store/types';
export type {
  OfficeResourceResolver,
  OfficeResourceSession,
} from './resources';
export { createOfficeResourceSession } from './resources';
export { createOfficeResourceStore } from '../services/resource-store/OfficeResourceStore';
export {
  createOfficeDocumentSession,
  disposeDocumentSession,
} from './lifecycle';
export type { OfficeDocumentSession, OfficeSessionResource } from './lifecycle';
```

`src/core.ts` 只 re-export `./office-file-viewer/core`，不得导入 `OfficeFileViewer`、格式 React 组件或 ECharts。

- [x] **Step 5: 补充根入口和 export map**

根入口继续导出现有 API，并以类型导出方式补充 Core 常用类型。`package.json` 增加 `./core`，保留 `./`、`./layout` 和 `./package.json` 的现有行为。

- [x] **Step 6: 运行探针和入口检查**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-core-contract-probe.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-core-contract-probe.mjs
yarn run check:types
```

Expected: 探针通过，Core 入口不携带 Viewer UI；临时文件完成后删除，不放入仓库。

## Task 2: 公共模型快照和格式适配

**Files:**

- Create: `src/office-file-viewer/core/model.ts`
- Modify: `src/office-file-viewer/services/preview.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Test: 系统临时目录中的 `office-core-model-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `ParsedOfficeFile`、`PresentationDocument`、`SpreadsheetWorkbook`、`DocDocument`、`DocxDocument` 以及现有资源释放函数。
- Produces: `OfficeDocumentModel`、`OfficeDocumentSnapshot` 的物化适配函数、Source 快照构造函数和格式能力计算入口。

- [x] **Step 1: 写模型适配失败探针**

在临时探针中构造最小 `ParsedOfficeFile` 联合值，并调用计划中的函数：

```ts
import { createOfficeDocumentRuntime } from 'D:/workspace/github/office-x-viewer/src/core';

const runtime = createOfficeDocumentRuntime({
  sessionId: 'probe-session',
  file: new File([], 'probe.xlsx'),
  parsed: {
    kind: 'xlsx',
    workbook: {
      sheets: [],
      resources: { objectUrls: [] },
    },
  } as never,
});

if (
  runtime.snapshot.format !== 'xlsx' ||
  runtime.mode !== 'materialized' ||
  !runtime.model
) {
  throw new Error('模型快照没有保留格式和物化模式');
}
```

- [x] **Step 2: 运行探针确认缺少适配函数**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-core-model-probe.ts --bundle --platform=node --format=esm --outfile=$env:TEMP\office-core-model-probe.mjs
```

Expected: 出现 `createOfficeDocumentRuntime` 未导出的失败。

- [x] **Step 3: 定义格式映射类型**

在 `core/model.ts` 定义按 `PreviewKind` 收窄的输入和输出：

```ts
export type OfficeDocumentSnapshotInput = {
  sessionId: string;
  file: File;
  parsed: ParsedOfficeFile;
};

export function createOfficeDocumentSnapshot(
  input: OfficeDocumentSnapshotInput,
): OfficeDocumentSnapshot;

export function createOfficeDocumentRuntime<TModel = OfficeDocumentModel>(
  input: OfficeDocumentSnapshotInput,
): OfficeDocumentRuntime<TModel, never>;

export type OfficeSourceSnapshotInput = Readonly<{
  sessionId: string;
  fileName?: string;
  fileSize?: number;
  format: PreviewKind;
  family: PreviewFamily;
  capabilities: OfficeCapabilities;
  resources?: readonly OfficeResourceRef[];
  warnings?: readonly OfficeWarning[];
  sourceMap?: OfficeSourceMap;
  extensionData?: Readonly<Record<string, OfficeJsonValue>>;
}>;

export function createOfficeSourceSnapshot(
  input: OfficeSourceSnapshotInput,
): OfficeDocumentSnapshot;

export function createOfficeSourceRuntime<TSource = OfficeDocumentSource>(
  snapshot: OfficeDocumentSnapshot,
  source: TSource,
): OfficeDocumentRuntime<never, TSource>;

export function getOfficeCapabilities(
  format: PreviewKind,
  available?: Partial<OfficeCapabilities>,
): OfficeCapabilities;

export function getOfficeDocumentModel<
  TModel = OfficeDocumentModel,
  TSource = never,
>(runtime: OfficeDocumentRuntime<TModel, TSource>): TModel | undefined;
```

实现规则：

- `parsed.kind`、`snapshot.format` 和 `snapshot.family` 必须一致；
- `sessionId` 作为快照 ID 的组成部分，不使用文件名生成唯一标识；
- 物化模型转移现有资源所有权，不复制 Blob URL；`OfficeResourceRef` 只提取 ID、MIME、大小和可序列化 locator；
- Source 模式只在 `OfficeDocumentRuntime` 保存 Source 引用和摘要，不把 Source 放入可序列化快照；
- 格式专属字段通过 `extensionData` 或格式模型保留，不使用 `any`。

模型适配器只能从 `formatDefinitions` 和格式类型文件引入必要内容，不能从 `formats/*` React 渲染目录引入运行时代码；公共快照生成必须是纯同步、无 DOM 副作用的步骤。
Source 快照统一通过 `createOfficeSourceSnapshot` 构造，文件名、大小、能力和资源引用缺失时使用明确的 `undefined`，不从文件内容重复读取或猜测。

物化映射使用完整判别分支而不是类型断言：`doc/docx/ppt/pptx` 读取 `parsed.document`，`xls/xlsx` 读取 `parsed.workbook`；未知 `kind` 直接抛出 `INVALID_FILE`，不把字段降级为 `any`。
`getOfficeCapabilities()` 以 `formatDefinitions` 的 family/Worker 元数据为默认值，再合并解析结果明确声明的基础字段和 `features`；缺少实际 Renderer 或导出器时必须返回 `false`，宿主不得通过格式字符串自行猜测。

- [x] **Step 4: 接入现有 preview 类型**

从 `services/preview.ts` 复用 `ParsedOfficeFile`、`PreviewKind` 和 `PreviewFamily`，不复制一份格式联合。`disposeParsedOfficeFile` 的释放语义保持不变；运行时对象的释放必须交给 `OfficeDocumentSession`，不能在快照上增加第二个释放入口。

- [x] **Step 5: 运行模型探针和类型检查**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-core-model-probe.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-core-model-probe.mjs
yarn run check:types
```

Expected: 探针和类型检查通过；删除临时探针。

## Task 3: 对外共享资源和生命周期门面

**Files:**

- Create: `src/office-file-viewer/core/resources.ts`
- Create: `src/office-file-viewer/core/lifecycle.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Read/verify: `src/office-file-viewer/services/resource-store/OfficeResourceStore.ts`
- Read/verify: `src/office-file-viewer/services/session/OfficeDocumentSession.ts`
- Test: 系统临时目录中的 `office-core-resource-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `createOfficeResourceStore`、`OfficeResourceStore`、`createOfficeDocumentSession` 和现有 `dispose` 逻辑。
- Produces: `OfficeResourceSession`、`createOfficeResourceSession`、`OfficeDocumentSession` 的稳定 Core 导出。

- [x] **Step 1: 写资源会话失败探针**

探针先从 `src/core` 导入 `createOfficeResourceSession`，创建会话，调用 `resolve()` 读取一个惰性资源，再调用两次 `dispose()`。当前 Core 入口没有该导出时，构建应失败。

- [x] **Step 2: 定义轻量资源会话**

在 `core/resources.ts` 定义：

```ts
import type {
  OfficeResourceSource,
  OfficeResourceStore,
  OfficeResourceStoreOptions,
} from '../services/resource-store/types';

export type OfficeResourceSession = {
  readonly store: OfficeResourceStore;
  resolve(source: OfficeResourceSource, signal?: AbortSignal): Promise<Blob>;
  acquire(source: OfficeResourceSource, signal?: AbortSignal): Promise<string>;
  release(source: OfficeResourceSource): void;
  dispose(): Promise<void>;
};

export function createOfficeResourceSession(
  options?: OfficeResourceStoreOptions & {
    resolver?: OfficeResourceResolver;
  },
): OfficeResourceSession;
```

实现只包装现有 Store，不再复制引用计数或 LRU；`resolve()` 在有宿主 resolver 时委托它，否则复用现有 `portableResourceToBlob`，并由 Store 负责 Object URL 引用；传入 `resolver` 时仅为 `lazy` Source 建立受控代理，`url` Source 仍直接交给 Store。`dispose()` 必须直接复用 Store 的幂等释放逻辑。

- [x] **Step 3: 暴露文档会话契约**

`core/lifecycle.ts` 只 re-export `OfficeDocumentSession`、`OfficeSessionResource`、`createOfficeDocumentSession` 和幂等的 `disposeDocumentSession`。不把 React Provider、DOM 节点或内部解析句柄暴露到 Core。

- [x] **Step 4: 补充资源解析器接口**

定义宿主可选的资源解析契约：

```ts
export type OfficeResourceResolver = (
  source: OfficeResourceSource,
  context: { signal: AbortSignal; sessionId: string },
) => Promise<Blob>;
```

该接口只作为可选契约，不改变默认 Store 的 `source.load()` 行为。代理加载前检查 `signal.aborted`，加载后校验声明的 MIME 和大小，失败统一转换为 `OfficeFileViewerError` 的资源阶段错误；resolver 抛错时不增加引用计数，也不留下 Object URL。

- [x] **Step 5: 运行资源探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-core-resource-probe.ts --bundle --platform=node --format=esm --outfile=$env:TEMP\office-core-resource-probe.mjs
node $env:TEMP\office-core-resource-probe.mjs
yarn run check:types
```

Expected: Store 可以创建并幂等释放；临时探针删除。

## Task 4: 抽出 Core 解析入口并保持按需加载

**Files:**

- Create: `src/office-file-viewer/core/parsing.ts`
- Create: `src/office-file-viewer/core/parsingContracts.ts`
- Create: `src/office-file-viewer/core/parsingSession.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Modify: `src/office-file-viewer/services/parsing/types.ts`
- Modify: `src/office-file-viewer/services/parsing/createParseSession.ts`
- Modify: `src/office-file-viewer/services/parsing/formatParserRegistry.ts`
- Modify: `src/office-file-viewer/services/errors/OfficeFileViewerError.ts`
- Test: 系统临时目录中的 `office-core-parsing-probe.ts` 和 `office-core-runtime-session-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: 现有 `createOfficeParseSession`、`OfficeParseOptions`、`OfficeParseSession`、`OfficeFormatDefinition`、Worker 运行时和 Source 工厂。
- Produces: Core 公开解析契约、`OfficeParserInput`/`OfficeParser`/`OfficeSourceFactory` 类型，以及可选 `OfficePluginRegistry` 注入点的类型预留。

- [x] **Step 1: 写解析入口失败探针**

探针从 `src/core` 导入 `createOfficeParseSession`、`OfficeParseOptions` 和 `ParseProgress`，确认 Core 门面没有复制或缺失现有定义。

- [x] **Step 2: 创建解析门面**

`core/parsing.ts` 只 re-export 文档化的解析会话和 Source 类型：

```ts
export { createOfficeParseSession } from '../services/parsing/createParseSession';
export { parseOfficeFile } from '../preview';
export { createOfficeDocumentParseSession } from './parsingSession';
export type {
  OfficeParseOptions,
  OfficeParseSession,
  ParseProgress,
  ParseStage,
  WorkerMode,
} from '../services/parsing';
```

`parseOfficeFile` 仅作为完整物化模型的便捷函数；不得从该入口导出 `createOfficeFileViewerParseSession`、`RuntimeSink`、内部预览句柄或 Worker 消息联合。

`core/parsing.ts` 追加：

```ts
export type {
  OfficeParser,
  OfficeParserInput,
  OfficeParserPlugin,
  OfficePluginResolver,
  OfficeSourceFactory,
} from './parsingContracts';
```

`core/parsingSession.ts` 定义：

```ts
import type { OfficeDocumentRuntime } from './types';
import type {
  OfficeParseOptions,
  OfficeParseSession,
} from '../services/parsing';

export function createOfficeDocumentParseSession(
  file: File,
  options?: OfficeParseOptions,
): OfficeParseSession<OfficeDocumentRuntime>;
```

该适配会把完整 `ParsedOfficeFile` 转为 `OfficeDocumentRuntime` 的 `mode: 'materialized'` 分支；满足现有大文件 Source 条件时通过公开 Source 快照返回 `mode: 'source'`。原有 `createOfficeParseSession` 仍返回 `ParsedOfficeFile`，不改变旧调用方的类型。

- [x] **Step 3: 预留解析器注册接口**

在 `core/parsingContracts.ts` 定义供插件实现使用的公开输入合同，不暴露内部 Runtime；`core/parsing.ts` 只 re-export 这些类型和解析会话，服务层仅以 `import type` 依赖 `parsingContracts.ts`，避免反向导入门面：

```ts
import type {
  OfficeCapabilities,
  OfficeDocumentModel,
  OfficeDocumentRuntime,
  OfficeDocumentSnapshot,
  OfficeDocumentSource,
} from './types';
import type { OfficeResourceSession } from './resources';
import type { OfficeDocumentSession } from './lifecycle';
import type { OfficeParseOptions } from '../services/parsing';

export type OfficeParserInput = Readonly<{
  file: File;
  sessionId: string;
  signal: AbortSignal;
  options: Omit<OfficeParseOptions, 'pluginRegistry'>;
  resources: OfficeResourceSession;
  documentSession: OfficeDocumentSession;
}>;

export type OfficeParser<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeParserInput,
) => Promise<OfficeDocumentRuntime<TModel, TSource>>;

/** Worker 适配器接收的可序列化字段，不包含 Store、Session 或函数。 */
export type OfficeWorkerParserInput = Readonly<{
  file: File;
  sessionId: string;
  options: Omit<OfficeParseOptions, 'pluginRegistry' | 'workerFactory'>;
}>;

export type OfficeWorkerParser<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeWorkerParserInput,
  worker: Worker,
) => Promise<OfficeDocumentRuntime<TModel, TSource>>;

export type OfficeSourceFactory<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeParserInput & { snapshot?: OfficeDocumentSnapshot },
) => Promise<OfficeDocumentRuntime<TModel, TSource> | undefined>;

/** Core 只依赖解析所需的最小解析器查询能力，避免反向依赖 plugins 入口。 */
export type OfficeParserPlugin<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = Readonly<{
  id: string;
  capabilities: OfficeCapabilities;
  parse: OfficeParser<TModel, TSource>;
  createSource?: OfficeSourceFactory<TModel, TSource>;
  dispose?(
    runtime: OfficeDocumentRuntime<TModel, TSource>,
  ): void | Promise<void>;
}>;

/** 格式扩展的低层兼容别名，完整注册元数据由 plugins 入口补充。 */
export type OfficeFormatPlugin = OfficeParserPlugin;

export type OfficePluginResolver = Readonly<{
  resolve(
    file: File,
    signal?: AbortSignal,
  ): Promise<OfficeParserPlugin | undefined>;
}>;
```

`OfficeParseOptions` 增加可选的 `pluginRegistry?: OfficePluginResolver` 时，必须标记该字段不可结构化克隆；Worker 只接收插件 ID 和可序列化配置。插件注册表的完整公开接口在 Task 6 定义，Core 只依赖这里的最小结构。

在 `services/parsing/types.ts` 仅使用 `import type` 从 `../../core/parsingContracts` 引入 `OfficePluginResolver` 并增加该可选字段；在 `createParseSession.ts` 把它作为主线程解析选择器使用。该类型导入在编译后被擦除，不能形成 Core 与 UI 的运行时循环。

字段形态固定为：

```ts
export type OfficeParseOptions = {
  // 保留现有 worker、workerFactory 和 resourcePolicy 字段。
  worker?: WorkerMode;
  workerFactory?: () => Worker;
  resourcePolicy?: OfficeParseResourcePolicy;
  /** 当前线程的格式插件解析器查询入口，不会传入 Worker。 */
  pluginRegistry?: OfficePluginResolver;
};
```

- [x] **Step 4: 保持内置解析路径不变**

`createParseSession()` 仍然先校验格式、资源限制、加密文件和取消信号，再选择 Source 或完整解析。未传插件注册表时，调用结果、错误码、进度顺序和动态 import 路径保持原样。

在结构化错误类型中追加 `plugin`、`export`、`navigation` 阶段及其代码，但不改写已有代码的含义：`PLUGIN_CONFLICT`、`PLUGIN_WORKER_UNSUPPORTED`、`PLUGIN_RENDERER_MISSING`、`EXPORT_UNSUPPORTED`、`EXPORT_CANCELLED`、`EXPORT_FILENAME_REQUIRED`、`EXPORT_RESOURCE_FAILED`、`EXPORT_WRITE_FAILED`、`NAVIGATION_UNSUPPORTED` 和 `NAVIGATION_OUT_OF_RANGE`。同时为错误上下文增加可选 `format` 字段，归一化函数保留原始 `cause`，错误消息不包含本地路径或文件字节。

错误归一化规则固定为：插件异常使用 `stage: 'plugin'`，导出异常使用 `stage: 'export'`，导航边界使用 `stage: 'navigation'`；调用方仍可通过 `isOfficeFileViewerError(error)` 和 `error.code` 做分支，不需要解析错误消息。

- [x] **Step 5: 检查 Core 是否误载入 Viewer**

运行：

```powershell
yarn --silent esbuild src/core.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-core-entry.mjs
rg -n "OfficeFileViewer|formats/.+\.tsx|antd|echarts" $env:TEMP\office-core-entry.mjs
yarn run check:types
```

Expected: Core 入口不包含 Viewer JSX、Ant Design 或 ECharts UI；临时产物删除。

## Task 5: 建立导出契约和原始文件导出

**Files:**

- Create: `src/office-file-viewer/export/types.ts`
- Create: `src/office-file-viewer/export/OfficeExportService.ts`
- Create: `src/office-file-viewer/export/index.ts`
- Create: `src/office-file-viewer/core/export.ts`
- Modify: `src/office-file-viewer/core/types.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Create: `src/export.ts`
- Test: 系统临时目录中的 `office-export-contract-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `File`、`Blob`、`OfficeDocumentSnapshot`、未来 `OfficeChangeSet` 和导出器注册接口。
- Produces: `OfficeExportRequest`、`OfficeExportResult`、`OfficeExportError`、`exportOriginalOfficeFile` 和 `OfficeExporter` 契约。

- [x] **Step 1: 写导出契约失败探针**

探针从 `src/export` 导入 `exportOriginalOfficeFile` 和 `OfficeExportResult`，传入一个 `Blob` 并检查文件名、MIME 和字节内容。当前入口没有实现时应失败。

同时用已取消的 `AbortController` 调用一次，断言拒绝值可被 `isOfficeFileViewerError` 识别且代码为 `EXPORT_CANCELLED`；探针不触发浏览器下载。

- [x] **Step 2: 定义导出类型**

在 `export/types.ts` 定义：

```ts
import type {
  OfficeChangeSet,
  OfficeDocumentSnapshot,
  OfficeWarning,
} from '../core/types';
import type { OfficeFileViewerError } from '../services/errors/OfficeFileViewerError';

export type OfficeExportRequest = {
  document: OfficeDocumentSnapshot;
  changes?: OfficeChangeSet;
  format?: string;
  signal?: AbortSignal;
  preserveUnknownParts?: boolean;
};

export type OfficeExportResult = {
  blob: Blob;
  fileName: string;
  format: string;
  warnings?: readonly OfficeWarning[];
};

export type OfficeExporter = {
  id: string;
  formats: readonly string[];
  export(request: OfficeExportRequest): Promise<OfficeExportResult>;
};

export type OfficeExportErrorCode =
  | 'EXPORT_UNSUPPORTED'
  | 'EXPORT_CANCELLED'
  | 'EXPORT_FILENAME_REQUIRED'
  | 'EXPORT_RESOURCE_FAILED'
  | 'EXPORT_WRITE_FAILED';

// 这些代码必须同时加入 OfficeFileViewerErrorCode，保证统一错误判断。

export type OfficeExportError = OfficeFileViewerError;

export type OfficeExporterRegistry = {
  list(): readonly OfficeExporter[];
  getByFormat(format: string): OfficeExporter | undefined;
  register(exporter: OfficeExporter): () => void;
  dispose(): void;
};

export function createOfficeExporterRegistry(
  exporters?: readonly OfficeExporter[],
): OfficeExporterRegistry;
```

`OfficeExportError` 使用现有 `OfficeFileViewerError` 结构化错误基类，`stage` 固定为 `export`，至少区分格式不支持、导出取消、文件名缺失、资源失败和写回失败；不另造一套无法被 `isOfficeFileViewerError` 识别的错误类。

- [x] **Step 3: 实现原始文件导出**

实现 `exportOriginalOfficeFile(source, options)`：

```ts
export type OfficeOriginalSource = File | Blob | ArrayBuffer | Uint8Array;

export function exportOriginalOfficeFile(
  source: OfficeOriginalSource,
  options?: { fileName?: string; format?: string; signal?: AbortSignal },
): Promise<OfficeExportResult>;
```

函数只复制或包装原始字节，不解析、不重写、不执行宏。`signal` 已取消时返回标准 AbortError，不能生成半截 Blob。
`File` 默认沿用原文件名和 MIME；`Blob`、`ArrayBuffer` 和 `Uint8Array` 必须要求调用方提供 `fileName` 或可推断的 `format`，无法推断时使用明确的 `EXPORT_FILENAME_REQUIRED` 分支（不静默写成 `.bin`）。函数不创建长期 Object URL，下载由宿主决定。

- [x] **Step 4: 暴露未实现格式的明确能力**

`exportDocument(request)` 通过导出器注册表调度；没有对应 Writer 时抛出结构化 `EXPORT_UNSUPPORTED`，不伪造成功结果。`OfficeCapabilities.canExportOriginal` 和 `exportFormats` 必须反映真实能力。

`OfficeExportService.ts` 提供 `createOfficeExportService(options?: { exporters?: readonly OfficeExporter[] })`，返回包含 `exportDocument(request)`、`exportOriginalOfficeFile(source, options)` 和 `dispose()` 的实例；注册表采用实例级快照，导出期间不读取或修改 Viewer 状态。内置首版只实现原始文件导出，XLSX/DOCX/PPTX Writer 仅在确有实现和回读验证后注册。

`core/export.ts` 只转发导出契约和工厂，`src/export.ts` 只转发 `core/export`，禁止从 Export 入口引入 Viewer、React 或格式渲染器；Exporter 的动态实现只能在 `exportDocument()` 真正选中目标格式后加载。

- [x] **Step 5: 运行导出探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-export-contract-probe.ts --bundle --platform=node --format=esm --outfile=$env:TEMP\office-export-contract-probe.mjs
node $env:TEMP\office-export-contract-probe.mjs
yarn run check:types
```

Expected: 原始 Blob 内容和文件名保持一致，取消和不支持导出分支有明确错误；注册表重复格式被拒绝；删除临时探针。

## Task 6: 建立作用域级插件注册表

**Files:**

- Create: `src/office-file-viewer/plugins/types.ts`
- Create: `src/office-file-viewer/plugins/OfficePluginRegistry.ts`
- Create: `src/office-file-viewer/plugins/index.ts`
- Create: `src/plugins.ts`
- Modify: `src/office-file-viewer/core/types.ts`
- Modify: `src/office-file-viewer/core/parsing.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Read/verify: `src/office-file-viewer/services/parsing/formatDefinitions.ts`
- Modify: `src/office-file-viewer/services/parsing/formatParserRegistry.ts`
- Test: 系统临时目录中的 `office-plugin-registry-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `OfficeCapabilities`、`OfficeDocumentRuntime`、`OfficeDocumentSnapshot`、`OfficeParser`、`OfficeSourceFactory`、`OfficeExportRequest`、现有 `OfficeFormatDefinition` 和动态 loader。
- Produces: 不依赖 React 的 `OfficeCorePlugin`、`OfficePluginRegistry` 和 built-in registry；Viewer/Editor Adapter 分别在 UI 任务中定义。

- [x] **Step 1: 写注册表失败探针**

探针创建两个同名扩展插件，期望第二次注册被拒绝；再创建同一扩展名但相同优先级的插件，期望被拒绝，显式更高优先级时验证 `resolve()` 选择高优先级插件，并验证 `list()` 返回不可变数组。当前注册表不存在时，构建应失败。

- [x] **Step 2: 定义插件契约**

在 `plugins/types.ts` 定义：

```ts
import type { OfficeDocumentModel, OfficeDocumentSource } from '../core/types';
import type {
  OfficeParserPlugin,
  OfficeWorkerParser,
  OfficePluginResolver,
} from '../core/parsingContracts';
import type { OfficeExportRequest, OfficeExportResult } from '../export/types';
import type { PreviewKind } from '../services/preview';

export type OfficePluginWorkerSupport = 'none' | 'main-thread' | 'worker';

export type OfficeCorePlugin<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = OfficeParserPlugin<TModel, TSource> & {
  extensions: readonly string[];
  mimeTypes?: readonly string[];
  /** 同一扩展名存在多个插件时的确定性选择顺序，数值越大越优先。 */
  priority?: number;
  workerSupport: OfficePluginWorkerSupport;
  detect(
    file: File,
    context: { signal: AbortSignal },
  ): boolean | Promise<boolean>;
  /** worker 模式由宿主在当前线程创建 Worker，函数不会被结构化克隆。 */
  workerFactory?: () => Worker;
  /** 自定义 Worker 协议适配器；缺失时只能走主线程。 */
  parseInWorker?: OfficeWorkerParser<TModel, TSource>;
  export?: (request: OfficeExportRequest) => Promise<OfficeExportResult>;
};

/** OfficeFormatPlugin 是面向格式扩展的兼容别名，首版与 Core 插件契约相同。 */
export type OfficeFormatPlugin = OfficeCorePlugin;

export type OfficePluginRegistry = OfficePluginResolver & {
  list(): readonly OfficeCorePlugin[];
  getById(id: string): OfficeCorePlugin | undefined;
  getByKind(kind: PreviewKind): OfficeCorePlugin | undefined;
  resolve(
    file: File,
    signal?: AbortSignal,
  ): Promise<OfficeCorePlugin | undefined>;
  register(plugin: OfficeCorePlugin): () => void;
  dispose(): void;
};
```

Viewer 和 Editor Adapter 单独定义，Core 类型不能引用 `React.ComponentType`。
注册表内部以 `OfficeCorePlugin<unknown, unknown>` 保存异构格式，Renderer Adapter 通过 `pluginId` 约束自己的运行时模型类型；内置格式仍提供具体类型别名，新增格式不需要修改现有联合。

- [x] **Step 3: 实现实例级注册表**

实现 `createOfficePluginRegistry(options)`，内部使用拷贝后的 Map：

```ts
export function createOfficePluginRegistry(options?: {
  includeBuiltIns?: boolean;
  plugins?: readonly OfficeCorePlugin[];
}): OfficePluginRegistry;
```

注册时校验 `id`、扩展名、能力对象和 `parse` 函数；重复 ID 总是拒绝，同一扩展名只有在显式 `priority` 不同且最高者唯一时才允许共存，否则抛出明确错误。`register()` 返回幂等注销函数，不影响其他 Registry 实例；`dispose()` 只清理注册表监听，不调用插件文档资源释放。`list()` 返回冻结副本，`resolve()` 按显式优先级、扩展名、MIME 和异步 `detect()` 顺序选择唯一插件。
`resolve()` 先用小写扩展名和 MIME 索引筛选候选，不对全部插件读取文件；只有没有索引命中时才按注册顺序调用 `detect()`，每次调用都传入同一个取消信号，并在首个唯一命中后停止。
注册时统一把扩展名规范化为带点小写形式、MIME 规范化为小写并冻结数组；空扩展名、路径片段和包含通配符的值直接以 `PLUGIN_CONFLICT` 拒绝。

- [x] **Step 4: 适配内置格式定义**

在 `formatParserRegistry.ts` 增加只读转换函数，把现有六个解析定义（`OfficeFormatDefinition`）包装为 `OfficeParserPlugin`。内置插件的 `parse` 闭包只在第一次调用时执行现有 `loadParser()`，再由会话适配器转换为 `OfficeDocumentRuntime`；`createSource` 闭包同样延迟加载现有 Source 工厂。所有 `.doc/.wps/.docx/.docm/.dotx/.xls/.xlsx/.xlsm/.xltx/.ppt/.pptx/.pptm/.potx` 别名继续使用当前识别规则。
该转换函数只依赖 Core 的 `OfficeParserPlugin` 结构类型，不反向导入 `plugins` 运行时代码；完整的 `OfficeCorePlugin` 元数据和注册逻辑由 `OfficePluginRegistry.ts` 包装完成。
内置插件固定 `priority: 0`；外部插件若要接管已有扩展名必须显式设置更高优先级，并在开发环境发出覆盖警告，默认不改变现有格式行为。

- [x] **Step 5: 明确 Worker 策略**

外部插件没有完整 Worker 适配器时标记 `workerSupport: 'main-thread'`：

- `worker: 'auto'` 自动改用主线程；
- `worker: 'never'` 正常运行；
- `worker: 'always'` 返回结构化 `PLUGIN_WORKER_UNSUPPORTED`。

声明 `workerSupport: 'worker'` 的插件必须同时提供 `workerFactory` 和 `parseInWorker`；两者只在宿主线程调用，传给 Worker 的消息只包含文件字节、插件 ID 和可序列化配置，不能把插件函数或资源 Store 直接结构化克隆给 Worker。缺少任一项时注册失败并指出缺失字段。

- [x] **Step 6: 接入解析会话**

在 `OfficeParseOptions` 增加可选 `pluginRegistry`，格式解析先由 Registry 按扩展名、MIME 和 `detect()` 优先级解析文件，再回退到内置格式判定。命中插件时将当前会话的 `file`、`sessionId`、`AbortSignal`、资源会话和只读选项组成 `OfficeParserInput`，调用 `parse` 或满足 Worker 条件的 `parseInWorker`；插件失败统一经过现有结构化错误归一化。未传 Registry 时，结果与当前实现一致；`pluginRegistry` 只在主线程使用，进入 Worker 前转换为插件 ID 和可序列化配置。

- [x] **Step 7: 运行注册表探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-plugin-registry-probe.ts --bundle --platform=node --format=esm --outfile=$env:TEMP\office-plugin-registry-probe.mjs
node $env:TEMP\office-plugin-registry-probe.mjs
yarn run check:types
```

`plugins/index.ts` 和 `src/plugins.ts` 只导出注册表工厂、Core/Viewer 类型和 Provider，不在模块求值阶段加载任一格式解析器或 ECharts；内置插件的动态 loader 只在 `resolve()` 后调用。

Expected: 作用域隔离、重复 ID/扩展名优先级校验、内置格式回退和 Worker 能力判断通过；删除临时探针。

## Task 7: 将插件注册接入 Viewer 且保持默认接入不变

**Files:**

- Create: `src/office-file-viewer/plugins/OfficeViewerPluginContext.tsx`
- Create: `src/office-file-viewer/plugins/OfficeViewerProvider.tsx`
- Create: `src/office-file-viewer/plugins/viewerTypes.ts`
- Modify: `src/office-file-viewer/OfficeFileViewer.tsx`
- Modify: `src/office-file-viewer/shell/controller/useOfficeViewerController.ts`
- Modify: `src/office-file-viewer/shell/PreviewStage.tsx`
- Modify: `src/office-file-viewer/services/parsing/createParseSession.ts`
- Modify: `src/office-file-viewer/index.ts`
- Modify: `src/plugins.ts`
- Test: 系统临时目录中的 `office-viewer-plugin-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `OfficePluginRegistry`、`OfficeCorePlugin`、现有 Viewer 控制器和 `OfficeParseOptions`。
- Produces: 作用域插件 Provider、Viewer Renderer Adapter，以及 `OfficeFileViewer` 对自定义 Registry 的兼容接入。

- [x] **Step 1: 写 Viewer 插件接入失败探针**

探针使用 `OfficeViewerProvider` 包裹 `OfficeFileViewer`，并传入一个只支持主线程的插件；在 Provider 尚不存在时，构建应失败。

- [x] **Step 2: 定义 Provider 边界**

Provider 接收：

```ts
type OfficeViewerProviderProps = {
  registry?: OfficePluginRegistry;
  plugins?: readonly OfficeCorePlugin[];
  adapters?: readonly OfficeViewerPluginAdapter[];
  children: ReactNode;
};
```

在 `viewerTypes.ts` 定义：

```ts
import type { ComponentType } from 'react';
import type { OfficeDocumentRuntime } from '../core/types';

export type OfficeViewerDocumentProps<TModel = unknown, TSource = unknown> = {
  runtime: OfficeDocumentRuntime<TModel, TSource>;
  fileName: string;
  onError?: (error: unknown) => void;
};

export type OfficeViewerPluginAdapter<TModel = unknown, TSource = unknown> = {
  pluginId: string;
  renderDocument: ComponentType<OfficeViewerDocumentProps<TModel, TSource>>;
  renderUnsupported?: ComponentType<{ fileName: string }>;
};
```

如果同时传 `registry` 和 `plugins`，使用显式 Registry 并报告开发期警告；Provider 内部只创建一次 Registry，初始插件在创建时注入，后续数组变化在 effect 中按 ID 增量更新，Adapter 按 `pluginId` 替换而不覆盖其他插件。Provider 不负责 dispose 宿主传入的 Registry，只释放自己创建的注册表监听。

- [x] **Step 3: 接入 Viewer 控制器**

`OfficeFileViewer` 从最近的 Provider 读取 Registry 和 Adapter。未提供 Provider 时创建默认 built-in Registry 及内置 Adapter。解析会话在选择格式前使用该 Registry；匹配到外部插件时，只有存在对应 Adapter 才渲染自定义内容。
默认 Registry 和 Adapter 必须用 `useMemo`/模块级只读定义保持引用稳定，不能在每次渲染时重新注册或重新加载格式模块。

内置 Adapter 继续消费当前 `OfficeFileViewerPreviewHandle`，只在边界处转换为 `OfficeDocumentRuntime` 的 `snapshot`/`model`/`source` 视图，不改动各格式 Viewer 的 Props；外部 Adapter 只接收公开运行时合同。转换失败必须回到现有错误态并释放当前会话。

- [x] **Step 4: 处理自定义插件的渲染边界**

Core 插件没有 React 渲染器时，Viewer 使用现有 unsupported 状态并提供 `PLUGIN_RENDERER_MISSING` 警告；不把任意插件对象直接渲染为 JSX。Adapter 只接收 `OfficeDocumentRuntime` 和文件名，不得读取内部 DOM 或 RuntimeSink。
插件不得自行持有全局缓存或创建未登记的 Object URL；解析产生的资源必须通过 `OfficeParserInput.resources` 或所属 `OfficeDocumentSession` 登记，返回 runtime 前完成快照和资源引用配对。

- [x] **Step 5: 保持受控状态和资源生命周期**

插件切换、文件切换和 Provider 卸载必须复用现有 `OfficeDocumentSession`、`OfficeResourceStore` 和 `dispose` 路径；插件的可选 `dispose(runtime)` 在会话资源释放前后只调用一次，并在异常时继续清理其余资源。插件 Registry 只管理定义，不拥有文档资源。

- [x] **Step 6: 运行 Viewer 插件探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-viewer-plugin-probe.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-viewer-plugin-probe.mjs
yarn run check:types
```

Expected: 未配置插件时旧示例仍可构建，配置主线程插件和对应 Adapter 时能进入插件解析路径；缺少 Adapter 时显示结构化 unsupported 状态；删除临时探针。

## Task 8: 增加 A 路线的控制句柄和能力查询

**Files:**

- Create: `src/office-file-viewer/shell/controller/OfficeViewerHandle.ts`
- Create: `src/office-file-viewer/shell/controller/officeViewerCapabilities.ts`
- Modify: `src/office-file-viewer/shell/controller/useOfficeViewerController.ts`
- Modify: `src/office-file-viewer/OfficeFileViewer.tsx`
- Modify: `src/office-file-viewer/shell/PreviewStage.tsx`
- Modify: `src/office-file-viewer/formats/doc/DocViewer.tsx`
- Modify: `src/office-file-viewer/formats/docx/DocxViewer.tsx`
- Modify: `src/office-file-viewer/formats/word-pages/types.ts`
- Modify: `src/office-file-viewer/services/parsing/types.ts`
- Modify: `src/office-file-viewer/index.ts`
- Test: 系统临时目录中的 `office-viewer-handle-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: 当前 `OfficeViewerActions`、`OfficeViewerMeta`、`OfficeFileViewerViewState`、Word `WordPageNavigationController`、格式导航控制器和全屏实现。
- Produces: `OfficeViewerHandle`、`OfficeCapabilities` 和扩展后的 `OfficePreviewReadyInfo`。

- [x] **Step 1: 写句柄失败探针**

使用 React 16.9 兼容写法声明：

```tsx
const ref = useRef<OfficeViewerHandle>(null);
<OfficeFileViewer ref={ref} />;
```

在 `forwardRef` 接入前运行类型检查，Expected: `ref` 类型或组件 ref 类型不匹配。

- [x] **Step 2: 定义 Handle 和 Capabilities**

Handle 方法必须对应已有 Controller Action：`openFile` 映射 `selectFile`，`reload` 映射 `retry`，缩放/导航/搜索/审阅/全屏映射同名动作；没有现成动作的方法先返回明确的能力错误，不在本任务中复制控制器状态。`OfficeCapabilities` 至少包含 `previewKind`、`family`、查找、审阅、页面/工作表/幻灯片导航、备注、缩放、全屏和导出能力。

在 `OfficeViewerHandle.ts` 固定以下签名；异步虚拟页定位由内部排队处理，公开句柄保持同步调用兼容：

```ts
export type OfficeViewerHandle = {
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

- [x] **Step 3: 将 Viewer 改为 React 16.9 兼容的 forwardRef**

保留现有 Locale Provider、资源 Provider、Hyperlink Provider、图片预览 Provider、字体 Provider、搜索边界和审阅边界的嵌套顺序。只在最外层增加 `forwardRef`，通过 `useImperativeHandle` 映射 `OfficeViewerActions`。
实现形态固定为 `memo(forwardRef<OfficeViewerHandle, OfficeFileViewerProps>(...))`，不要把 `ref` 加进普通 Props，也不要改变现有默认导出名称。
`getViewState()` 和 `getCapabilities()` 读取最新 state/meta 的 ref，避免宿主在同一渲染周期中拿到闭包旧值；卸载后句柄方法保持无副作用。

- [x] **Step 4: 收敛受控模式语义**

句柄触发缩放、页面、工作表、幻灯片、备注、搜索或审阅动作时，统一走现有 `onViewStateChange`。当宿主传入受控值时，不直接修改内部受控字段。无文档或能力不存在时保持无副作用。

为补齐 Word 页导航，在 `OfficePreviewStage` 创建稳定的 `WordPageNavigationController` ref，并传给 `DocViewer`、`DocxViewer`；两者继续把 ref 交给 `VirtualWordPageList`。Controller 新增 `selectPage(index)`，只调用 `scrollToPage(index)` 并在未挂载时等待 `ensurePageMounted(index, signal)`，不复制虚拟列表的测量和定位算法。`goToPage` 对 DOC、DOCX、WPS 统一走该动作，索引越界时转换为可识别的导航错误。

- [x] **Step 5: 扩展首屏就绪信息**

在 `OfficePreviewReadyInfo` 增加 `capabilities`，并确保物化模式与 Source 模式都生成同一结构。能力计算只读取快照和 Controller Meta，不触发额外解析或资源加载；未加载文档时 `getCapabilities()` 返回 `undefined`，不伪造格式能力。
`canExportOriginal` 仅在当前会话仍持有可重新读取的原始 `File`/Blob 时为 `true`；远程 URL 被释放或不可重读时返回 `false`，避免 Handle 暴露虚假能力。

- [x] **Step 6: 运行 Handle 探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-viewer-handle-probe.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-viewer-handle-probe.mjs
yarn run check:types
```

Expected: Handle 类型、受控状态和能力字段通过；删除临时探针。

## Task 9: 实现 B 路线的组合式 Layout / Shell

**Files:**

- Create: `src/office-file-viewer/shell/layout/OfficeViewerShell.tsx`
- Modify: `src/office-file-viewer/shell/layout/types.ts`
- Modify: `src/office-file-viewer/shell/layout/OfficeViewerLayoutContext.tsx`
- Modify: `src/office-file-viewer/shell/layout/OfficeViewerLayout.tsx`
- Modify: `src/office-file-viewer/shell/layout/index.ts`
- Modify: `src/layout.ts`
- Modify: `src/office-file-viewer/shell/layout/index.less`
- Test: 系统临时目录中的 `office-shell-composition-probe.tsx`，完成后删除。

**Interfaces:**

- Consumes: 现有 `OfficeViewerLayoutContextValue`、`useOfficeViewerLayoutController`、`OfficeToolbar`、`OfficeViewerFrame` 和水印视口。
- Produces: `OfficeViewerShell` Compound Components、通用 Shell Context 和兼容的 `OfficeViewerLayout`。

- [x] **Step 1: 写组合式 API 失败探针**

探针使用以下结构：

```tsx
<OfficeViewerShell.Provider value={value}>
  <OfficeViewerShell.Root>
    <OfficeViewerShell.Toolbar>
      <CustomAction />
    </OfficeViewerShell.Toolbar>
    <OfficeViewerShell.Viewport>
      <CustomDocument />
    </OfficeViewerShell.Viewport>
  </OfficeViewerShell.Root>
</OfficeViewerShell.Provider>
```

在组件不存在时运行 TypeScript/ESBuild，Expected: 对应导出缺失。

- [x] **Step 2: 扩展 Shell Context 类型**

保留 `state / actions / meta` 三段结构，并补充可复用字段：

```ts
import type { OfficeCapabilities } from '../../core/types';

type OfficeViewerShellState = {
  zoom: number;
  isFullscreen: boolean;
};

type OfficeViewerShellActions = {
  changeZoom(value: number): void;
  zoomIn(): void;
  zoomOut(): void;
  toggleFullscreen(): void | Promise<void>;
};

type OfficeViewerShellMeta = {
  viewerRef: RefObject<HTMLDivElement>;
  fullscreenSupported: boolean;
  contentScaling: OfficeViewerLayoutContentScaling;
  capabilities?: OfficeCapabilities;
};

export type OfficeViewerShellContextValue = Readonly<{
  state: OfficeViewerShellState;
  actions: OfficeViewerShellActions;
  meta: OfficeViewerShellMeta;
}>;
```

`OfficeViewerShell.Provider` 支持两种互斥用法：传入完整 `value` 时由宿主或未来 Editor 控制；不传 `value` 时复用现有 `useOfficeViewerLayoutController` 创建非受控状态，并接受 `defaultZoom`、`zoom`、`onZoomChange`、`contentScaling` 等现有布局参数。两种模式都只创建一个 Context，不能同时传 `value` 与状态控制参数。Editor 或宿主可以提供同一契约的不同 Provider。
`OfficeViewerLayoutContextValue` 作为兼容别名指向同一 `OfficeViewerShellContextValue` 结构，旧的 `useOfficeViewerLayout()` 返回值和类型名称继续保留；无 Provider 时的错误提示改为说明必须位于 `OfficeViewerShell.Provider` 或兼容 `OfficeViewerLayout` 内。

- [x] **Step 3: 创建最小 Compound Components**

`OfficeViewerShell` 至少提供 `Provider`、`Root`、`Toolbar`、`Viewport`、`Sidebar` 和 `StatusBar`。结构组件只负责布局和 children，不读取文件格式，不自行创建解析会话；`Root` 使用 `display: grid` 或现有布局变量协调区域，避免通过固定宽度挤压宿主内容。

静态属性使用 React 16.9 可表达的接口类型，不依赖 React 19 的 `ref` 作为普通属性：

```ts
import type { ComponentType, ReactNode } from 'react';
import type { OfficeViewerThemeOptions } from '../../shared/theme';
import type { OfficeViewerWatermark } from '../../shared/watermark';

export type OfficeViewerShellProviderProps = {
  value?: OfficeViewerShellContextValue;
  defaultZoom?: number;
  zoom?: number;
  onZoomChange?: (zoom: number) => void;
  contentScaling?: OfficeViewerLayoutContentScaling;
  theme?: OfficeViewerThemeOptions;
  watermark?: OfficeViewerWatermark;
  children: ReactNode;
};

export type OfficeViewerShellComponent = {
  Provider: ComponentType<OfficeViewerShellProviderProps>;
  Root: ComponentType<{ className?: string; children: ReactNode }>;
  Toolbar: ComponentType<{ children?: ReactNode }>;
  Viewport: ComponentType<{ children: ReactNode }>;
  Sidebar: ComponentType<{ children?: ReactNode }>;
  StatusBar: ComponentType<{ children?: ReactNode }>;
  FileInfo: ComponentType<{ children?: ReactNode }>;
  Zoom: ComponentType;
  Fullscreen: ComponentType;
};
```

- [x] **Step 4: 复用现有控件**

`FileInfo`、`Zoom` 和 `Fullscreen` 作为可选子组件复用现有 `OfficeToolbar` 的控制实现。不要复制缩放归一化、全屏错误处理或主题变量；子组件未挂载时不渲染空占位。
`Provider` 的 `theme` 和 `watermark` 继续复用现有主题/水印实现，Shell 只负责注入上下文；宿主传入的文档 children 不需要知道 CSS 变量或水印实现细节。

- [x] **Step 5: 保持 OfficeViewerLayout 兼容**

现有 `OfficeViewerLayout` 继续使用默认 Provider 和默认子组件组合。原有 Props、`contentScaling`、`toolbarExtra`、主题、水印和文件选择行为保持不变。

- [x] **Step 6: 补齐布局样式和无障碍**

新增类名使用现有 `office-viewer-layout` 前缀；Toolbar、Viewport、Sidebar 和 StatusBar 具备最小尺寸、滚动、焦点和 `prefers-reduced-motion` 兼容，不使用固定的超大 z-index。

- [x] **Step 7: 运行组合式探针**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-shell-composition-probe.tsx --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-shell-composition-probe.mjs
yarn run check:types
yarn run check:styles
```

Expected: 自定义内容可以复用缩放和全屏 Context，旧 `OfficeViewerLayout` 示例仍可构建；删除临时探针。

## Task 10: 增加宿主状态插槽和稳定样式入口

**Files:**

- Modify: `src/office-file-viewer/OfficeFileViewer.tsx`
- Modify: `src/office-file-viewer/shell/Loading.tsx`
- Modify: `src/office-file-viewer/shell/Error.tsx`
- Modify: `src/office-file-viewer/shared/ui/OfficeEmptyState.tsx`
- Modify: `scripts/build-package-styles.mjs`
- Modify: `package.json`
- Modify: `README.md`
- Modify: `README.zh-CN.md`
- Modify: `docs/dev/docs.md`
- Modify: `docs/dev/docs.zh-CN.md`
- Test: 系统临时目录中的 `office-viewer-slots-probe.tsx`，完成后删除。

**Interfaces:**

- Consumes: 当前 Loading、Error、Empty、ParseStatus、主题变量和 `OfficeViewerFrame`。
- Produces: `OfficeViewerSlots`、加载/空/错误状态替换能力和由构建脚本生成的 `office-file-viewer/styles.css` 稳定入口。

- [x] **Step 1: 定义状态插槽类型**

新增单一对象配置，避免增加多个顶层布尔属性：

```ts
type OfficeViewerSlots = {
  loading?: ReactNode;
  empty?: ReactNode;
  error?: ReactNode;
  statusBar?: ReactNode;
};
```

在 `OfficeFileViewerProps` 增加单一 `slots?: OfficeViewerSlots` 字段，其他既有 Props 不移动、不改名。错误详情继续通过 `onError` 提供，静态 `error` 节点不能吞掉重试能力；需要动作时使用 Shell Context。

- [x] **Step 2: 写插槽失败探针**

在临时 JSX 中传入 `slots={{ loading: <Loading /> }}`，运行类型检查，Expected: 当前 Props 不接受该配置。

- [x] **Step 3: 接入状态渲染**

在 `OfficeFileViewerContent` 中仅在对应状态存在时使用插槽；没有插槽时继续走现有默认组件。不得改变解析状态机、错误码或资源释放。

- [x] **Step 4: 生成稳定总样式**

不新增未被入口引用的 Less 文件。修改 `scripts/build-package-styles.mjs`，在现有 Less 编译和相对引用改写完成后，按路径排序收集 `dist` 内除 `styles.css` 外的全部 CSS 文件，按完整文本去重后写入 `dist/styles.css`；重复执行构建时先排除旧的 `styles.css`，避免自包含。保留逐文件 CSS 产物，确保根入口的按需样式引用不变，并验证总样式不包含 `.less` 引用。

聚合逻辑固定为：`files.filter(file => file.endsWith('.css') && basename(file) !== 'styles.css').sort()`，逐文件读取后用 `Set<string>` 按完整 CSS 文本去重，使用 `writeFile(resolve(distRoot, 'styles.css'), chunks.join('\n'), 'utf8')` 写出；不要用 glob 回读 `styles.css`，不要修改 JS 中已有相对 CSS 引用。

- [x] **Step 5: 更新 export map 和文档**

`package.json` 增加 `./styles.css`，README 和开发文档说明自动样式与显式样式入口的关系、CSS 变量优先级和稳定类名边界。
`OfficeViewerSlots` 从根入口以类型方式导出；不把内部 Loading/Error/Empty 组件路径写入对外文档，也不要求宿主依赖内部类名。

- [x] **Step 6: 运行插槽和样式检查**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-viewer-slots-probe.tsx --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-viewer-slots-probe.mjs
yarn run check:styles
yarn run build:styles
```

Expected: 默认状态和自定义状态都可渲染，`dist/styles.css` 存在且不包含 Less 引用；删除临时探针。

## Task 11: 公开未来 Editor 使用的变更契约

**Files:**

- Create: `src/office-file-viewer/core/editing.ts`
- Modify: `src/office-file-viewer/core/types.ts`
- Modify: `src/office-file-viewer/core/index.ts`
- Modify: `src/core.ts`
- Test: 系统临时目录中的 `office-editing-contract-probe.ts`，完成后删除。

**Interfaces:**

- Consumes: `OfficeDocumentSnapshot`、稳定节点 ID、`OfficeSourceMap` 和导出请求。
- Produces: 不包含编辑器 UI 的 `OfficeChangeSet`、`OfficeTransaction`、`OfficeCommand`、`OfficeCommandDefinition`、`OfficeHistory`、`OfficeSelection` 和 `OfficeEditorPluginAdapter` 类型。

- [x] **Step 1: 写编辑契约失败探针**

临时探针构造一个只读变更集，并从 `src/core` 导入类型。当前没有编辑契约时，TypeScript 应报告缺少导出。

- [x] **Step 2: 定义可序列化事务和选区类型**

`OfficeChangeOperation` 和 `OfficeChangeSet` 沿用 Task 1 已冻结的判别联合；本任务新增：

```ts
export type OfficeCommand = Readonly<{
  id: string;
  label: string;
  changes: OfficeChangeSet;
}>;

export type OfficeCommandDefinition = Readonly<{
  id: string;
  label: string;
  execute(): OfficeCommand;
}>;

export type OfficeTransaction = Readonly<{
  id: string;
  label: string;
  changes: OfficeChangeSet;
}>;

export type OfficeSelection = Readonly<{
  nodeId: OfficeNodeId;
  start: number;
  end: number;
}>;

export type OfficeHistory = {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  undo(): void;
  redo(): void;
  push(transaction: OfficeTransaction): void;
  clear(): void;
};

export type OfficeEditorPluginAdapter<TModel = unknown, TSource = unknown> = {
  pluginId: string;
  createEditorModel(document: OfficeDocumentRuntime<TModel, TSource>): unknown;
  commands: readonly OfficeCommandDefinition[];
};
```

`replace-node.value` 在第一阶段只能作为 `OfficeJsonValue` 受控序列化字段使用，不对外承诺任意对象可写入；编辑器实现前需要为各格式补充具体操作联合。
`OfficeCommand`、`OfficeTransaction` 和 `OfficeSelection` 可结构化序列化；`OfficeHistory` 是编辑器运行时接口，不进入 Worker 消息或只读 Viewer 快照。

- [x] **Step 3: 明确只读 Viewer 边界**

Viewer 不创建或持有 `OfficeHistory`，只允许把可选 `changes` 传给导出器。Core 类型文件不得导入 React、DOM 或编辑器组件。

将 `OfficeCommand`、`OfficeCommandDefinition`、`OfficeEditorPluginAdapter`、`OfficeHistory`、`OfficeSelection` 和 `OfficeTransaction` 从 `core/index.ts` 以 `export type` 方式公开；运行时代码只保留纯数据校验，避免把编辑器实现打入 Viewer。

- [x] **Step 4: 运行编辑契约检查**

运行：

```powershell
yarn --silent esbuild $env:TEMP\office-editing-contract-probe.ts --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-editing-contract-probe.mjs
yarn run check:types
```

Expected: 变更集可结构化序列化，Core 产物不包含 Editor UI；删除临时探针。

## Task 12: 完成单包入口、类型声明和包体积边界

**Files:**

- Modify: `package.json`
- Modify: `.fatherrc.ts`（仅在当前 Father 配置无法识别新增入口时）
- Modify: `scripts/check-package-budget.mjs`
- Modify: `README.md`
- Modify: `README.zh-CN.md`
- Modify: `docs/dev/docs.md`
- Modify: `docs/dev/docs.zh-CN.md`
- Test: 系统临时目录中的 `office-public-exports-probe.tsx`，完成后删除。

**Interfaces:**

- Consumes: `src/core.ts`、`src/layout.ts`、`src/plugins.ts`、`src/export.ts`、现有根入口和样式产物。
- Produces: 可安装包中的稳定 export map、类型声明、安装说明、扩展示例和包体积预算。

- [x] **Step 1: 写安装消费者失败探针**

临时消费者分别导入：

```ts
import { OfficeFileViewer } from 'office-file-viewer';
import {
  OfficeViewerLayout,
  OfficeViewerShell,
} from 'office-file-viewer/layout';
import { createOfficeParseSession } from 'office-file-viewer/core';
import { createOfficePluginRegistry } from 'office-file-viewer/plugins';
import { exportOriginalOfficeFile } from 'office-file-viewer/export';
```

在 export map 和构建入口完成前运行消费者打包，Expected: 子路径解析失败。

- [x] **Step 2: 配置 Father 入口**

确认 Father 会为 `src/core.ts`、`src/plugins.ts`、`src/export.ts` 生成对应 ESM 和 `.d.ts`。只有实际构建未发现入口时才修改 `.fatherrc.ts`，不重排现有构建配置。

- [x] **Step 3: 更新 package exports**

在 `package.json` 增加 `./core`、`./plugins`、`./export` 和 `./styles.css`，保留 `./`、`./layout` 和 `./package.json`。代码入口每个条件同时提供 `types`、`import` 和现有指向 ESM 的 `default`；不添加 `require`，CSS 入口只提供 `default`/`import` 可解析的 CSS 文件。

- [x] **Step 4: 维护包体积预算**

`check-package-budget.mjs` 分别记录根入口、Core、Layout、Plugins、Export 和 Worker 产物大小；新增入口只统计自身入口文件及其独立 chunk，公共 chunk 只计入一次。阈值沿用现有预算，只在新增入口有明确体积证据时调整；脚本输出每个分组的文件数、原始字节数和超限路径，便于定位首屏膨胀。

- [x] **Step 5: 更新对外文档**

中英文 README 和开发文档新增：

- 普通 Viewer 接入；
- Core 解析/模型/资源/生命周期；
- Layout 组合式外壳；
- Plugin 注册和 Worker 限制；
- 原始文件导出与格式能力；
- 未来 Editor 复用边界；
- 稳定入口与禁止深层导入说明。

至少提供一组可复制的 ESM 示例（中英文只翻译说明，不改变 API）：

```tsx
import { OfficeFileViewer } from 'office-file-viewer';
import { OfficeViewerProvider } from 'office-file-viewer/plugins';
import { createOfficePluginRegistry } from 'office-file-viewer/plugins';

const registry = createOfficePluginRegistry({ includeBuiltIns: true });

<OfficeViewerProvider registry={registry}>
  <OfficeFileViewer uri={file} />
</OfficeViewerProvider>;
```

同时展示 `office-file-viewer/core` 的解析会话、`office-file-viewer/layout` 的 Shell 组合和 `office-file-viewer/export` 的原始文件导出；示例只使用稳定子路径，不引用 `services`、`formats` 或 `shared` 深层文件。
样式说明必须给出 `import 'office-file-viewer/styles.css';` 的显式用法，并说明如果宿主构建器已处理根入口的 CSS 副作用则不要再次导入，避免重复规则。
迁移说明明确：未文档化的 `services/*`、`formats/*`、`shared/*` 和 `dist/*` 深层导入不属于兼容承诺；新增功能必须改用五个稳定入口，不能为旧深层路径增加永久转发层。

- [x] **Step 6: 运行安装消费者检查**

运行：

```powershell
yarn build
yarn --silent esbuild $env:TEMP\office-public-exports-probe.tsx --bundle --platform=browser --format=esm --outfile=$env:TEMP\office-public-exports-probe.mjs
yarn run check:package-budget
```

Expected: 五个公共入口和类型声明均可解析，包体积预算通过；删除临时探针。

## Task 13: 完成公共 API 兼容和工程质量验证

**Files:**

- Read/verify: `src/office-file-viewer/index.ts`
- Read/verify: `src/layout.ts`
- Read/verify: `src/core.ts`
- Read/verify: `src/plugins.ts`
- Read/verify: `src/export.ts`
- Read/verify: `package.json`
- Read/verify: `scripts/check-react-compatibility.mjs`
- Test: 系统临时目录中的 React 消费者和 API 类型探针，完成后删除。

**Interfaces:**

- Consumes: 全部新增公共入口、旧入口、类型声明、Controller Handle 和 Shell Context。
- Produces: React 16.9、17、18 的安装/类型/浏览器打包证据，以及无深层导入依赖的 API 清单。

- [x] **Step 1: 检查旧 API 没有被删除**

使用 `rg` 对 `OfficeFileViewerProps`、`OfficeViewerLayoutProps`、`OfficeFileViewerUri`、`OfficeParseOptions`、`OfficeFileViewerImagePreviewOptions` 和现有回调进行导出检查。发现旧字段被删除时先恢复兼容导出，不在本任务中改名。
同时从构建后的 `dist/index.d.ts` 和 `dist/layout.d.ts` 生成旧导出名单，与当前 master 的只读基线逐项比较；差异只能是新增符号或明确的兼容别名，不能因重排入口而删除旧符号。

- [x] **Step 2: 运行 React 消费者验证**

运行现有脚本：

```powershell
yarn run check:react-compatibility
```

Expected: React 16.9、17、18 消费端类型检查和 ESM 浏览器打包通过。若临时 Yarn 安装因网络或注册表无响应，记录环境阻塞，不修改 peer 版本或构建配置掩盖问题。

- [x] **Step 3: 检查 React 兼容语法**

运行：

```powershell
rg -n "from 'react'.*\b(use|forwardRef)|use\(|ref=\{[A-Za-z]+\}" src/office-file-viewer/core src/office-file-viewer/shell/layout src/office-file-viewer/plugins
```

确认没有 React 19 `use()`、ref-as-prop 或 Context 新语法；`forwardRef` 和 `useImperativeHandle` 只在 Viewer UI 层使用。

- [x] **Step 4: 中文注释自检**

检查新增公共类型、业务常量、插件注册、资源释放、Worker 分支和导出错误分支是否有简短中文注释，注释说明业务原因，不重复属性名称。

- [x] **Step 5: 运行静态检查**

运行：

```powershell
yarn run check:types
yarn run check:lint
yarn --silent eslint src/core.ts src/plugins.ts src/export.ts src/layout.ts --ext .ts,.tsx
yarn run check:styles
yarn --silent prettier --check src/core.ts src/plugins.ts src/export.ts src/layout.ts src/office-file-viewer/core src/office-file-viewer/plugins src/office-file-viewer/export src/office-file-viewer/shell/layout
```

Expected: 类型、Lint、样式和目标范围格式检查通过。

- [x] **Step 6: 验证图表依赖按需加载**

分别对 `src/core.ts`、`src/layout.ts`、`src/plugins.ts` 和 `src/export.ts` 做浏览器 ESM bundle，使用 `rg -n "echarts|register.*chart"` 确认公共入口不包含 ECharts；再在临时浏览器探针中加载带图表的现有示例，确认首次渲染前才触发图表模块，切换到无图表文件后不重复注册。
Expected: 公共入口保持无 ECharts 的初始依赖，图表样本只加载一次图表模块；探针完成后删除临时产物。

## Task 14: 全部任务完成后的逐文件详细烟测与回归

**Files:**

- Read-only: `docs/dev/smoke-test.md`
- Read-only: `docs/dev/smoke-test.samples.ts`
- Read-only: `docs/files/`
- Verify: `http://localhost:8000/office-file-viewer/dev/smoke-test`
- Record outside repository: `$env:TEMP\office-viewer-final-qa-YYYYMMDD.md`

**Interfaces:**

- Consumes: 已完成的 Core、Viewer、Layout、Plugin、Export 公共入口和全部现有格式渲染器。
- Produces: 每个示例文件的逐项检查记录、问题修复闭环和最终无回归证据。

- [x] **Step 1: 启动并锁定测试环境**

运行：

```powershell
yarn start
```

打开 `http://localhost:8000/office-file-viewer/dev/smoke-test`，记录浏览器版本、视口尺寸、设备像素比、当前包版本和示例清单长度。测试期间固定这些环境参数，示例数量从 `OFFICE_VISUAL_SAMPLES` 动态读取，不手工写死。
视觉检查使用固定缩放和滚动位置截取稳定画面；有可用的 Office/PDF 源画面时做并排或半透明叠加比对，重点核对字体度量、基线、边框、页边距、裁切和颜色，不以可读到文字代替像素级检查。截图和比对产物只放系统临时目录。

- [x] **Step 2: 逐个加载示例并记录基础结果**

按左侧示例列表顺序逐个选择文件。每个文件都要记录：

- 文件名、扩展名和预期格式；
- 实际 `previewKind`、Worker/主线程模式和首屏状态；
- 完整解析是否完成；
- 警告和错误数量；
- 文件是否能重新选择、切换和释放。

任一文件未就绪、格式识别不符或出现未解释警告，都视为失败，不能跳过继续标记通过。

- [x] **Step 3: Word、DOCX、WPS 逐页检查**

对每个 Word 示例读取实际页数并逐页进入可视范围，不能只检查初始化、中间和末页。每页检查：

- 文字内容、字体、字号、字重、颜色、行高和段间距；
- 段落换行、列表编号、表格边框和单元格内容；
- 图片、浮动对象、图表、页眉页脚、脚注和尾注；
- 大纲显隐、当前项跟随、点击定位和左侧滚动；
- 超链接 Ctrl/Command + 单击、图片双击/右键和键盘入口；
- 缩放、垂直滚动、分页边界和页面空白。

每页至少保留一张稳定截图或明确的人工检查记录；长文档要覆盖每一页，而不是仅依赖默认检查点。

- [x] **Step 4: XLS、XLSX 逐工作表和可视区域检查**

对每个 Excel 示例逐一打开每个工作表，并沿已使用范围从左上、顶部中段、右侧、底部中段到右下移动视口。每个工作表检查：

- 表头、行号、列号、列宽、行高、字体、对齐和边框；
- 长文本、换行、缩进、缩小字体、合并单元格和公式显示；
- 冻结行列、工作表切换、横向/纵向滚动和滚动条拖动稳定性；
- 浮动图片、图表、批注、条件格式和超链接；
- 原始版式与阅读模式的差异，以及放大/缩小后的内容边界。

大表不能只看首屏；至少要按虚拟窗口连续滚动覆盖实际使用区域，并记录未加载区域的占位和加载完成结果。

- [x] **Step 5: PPT、PPTX、PPTM、POTX 逐张幻灯片检查**

对每个演示文稿示例逐张切换全部幻灯片，逐页检查：

- 文字、字体、字号、字重、颜色、行距和文本框边界；
- 形状、层叠、旋转、裁剪、图片和媒体占位；
- 母版、版式、背景和图片填充是否正确排除或显示；
- 缩略图与主页面一致性；
- 上一页/下一页边界、单页隐藏规则和页码；
- 演讲者备注、超链接、图片双击/右键、旋转和下载；
- 缩放、拖拽、全屏和页级切换。

PPTM 和 POTX 使用现有格式路径验证，不新增独立示例文件。

- [x] **Step 6: 问题修复闭环**

发现问题时，在 QA 记录中写明文件、页/表/幻灯片、复现操作、预期效果和实际效果。修复要求：

- 先定位公共解析、模型、资源、渲染或布局根因；
- 不按文件名、样本 ID、页码或单份文档添加特例；
- 修复后先重新检查受影响文件的全部页面/工作表/幻灯片；
- 再从示例清单第一项开始重新执行全量检查；
- 新问题继续按同一闭环处理，直到全量清单无失败项。

- [x] **Step 7: 完成最终工程检查**

在全量烟测和所有修复完成后运行：

```powershell
yarn run check
yarn run check:react-compatibility
git diff --check
```

只有示例逐项通过、无未解释警告、无视觉回归，且上述工程检查通过，才可以将整体任务标记完成。QA 记录保存在系统临时目录，不提交到 npm 包或公共文档。

## 执行顺序和检查点

按以下顺序执行，前一项的公共契约和局部验证通过后才进入后一项：

1. Task 1：公共类型和 `core` 入口。
2. Task 2：模型快照与格式适配。
3. Task 3：资源和生命周期门面。
4. Task 4：解析门面和插件注入边界。
5. Task 5：导出契约和原始文件导出。
6. Task 6：作用域级插件注册表。
7. Task 7：插件接入 Viewer。
8. Task 8：高层 Handle 与能力查询。
9. Task 9：组合式 Layout / Shell。
10. Task 10：状态插槽和样式入口。
11. Task 11：Editor 变更契约。
12. Task 12：单包导出、文档和包体积边界。
13. Task 13：公共 API、React 和工程质量检查。
14. Task 14：全部任务完成后的唯一一次全量详细烟测。

每项中间任务完成后：

- 重新读取本任务修改的文件；
- 检查中文注释、类型和导出是否完整；
- 运行本任务列出的局部验证；
- 不执行全量示例烟测；
- 不执行 Git 写操作。

Task 14 结束前不标记整体任务完成，也不发布版本。

## 最终交付物

- 单一 `office-file-viewer` npm 包；
- 根 Viewer 入口的向后兼容 API；
- `core`、`layout`、`plugins` 和 `export` 稳定子路径；
- Core 解析、模型、资源和生命周期契约；
- 组合式 `OfficeViewerShell`；
- 作用域级插件注册表；
- 原始文件导出和导出器协议；
- 面向未来 Editor 的变更、事务和选区类型；
- 中英文接入、扩展和迁移文档；
- 最终全量示例文件烟测与修复记录（仅保存在系统临时目录）。

## 计划自检

- Spec coverage：Task 1-5 覆盖 Core、模型、解析、资源和导出；Task 6-7 覆盖插件；Task 8-10 覆盖 A/B；Task 11 覆盖未来 Editor；Task 12-14 覆盖单包、兼容、性能、安全、文档和最终烟测。
- Scope：没有创建多 npm 包，没有实现完整编辑器，也没有把旧版格式写回承诺提前到当前阶段。
- Type consistency：`OfficeCapabilities`、`OfficeDocumentSnapshot`、`OfficeCorePlugin`、`OfficeViewerHandle`、`OfficeChangeSet` 和 `OfficeExportRequest` 在前置任务定义后由后续任务复用。
- Validation：中间任务只做局部验证；Task 14 才做现有示例逐文件、逐页/逐表/逐幻灯片全量烟测，并要求修复后从头回归。
- Placeholder scan：已检查所有步骤，未发现未完成占位语句或模糊的实施要求。

## Git 约束

本计划不包含 commit、push、pull、merge、rebase 或分支切换。只有用户明确授权后，才在最终验证通过的工作区执行 Git 操作。
