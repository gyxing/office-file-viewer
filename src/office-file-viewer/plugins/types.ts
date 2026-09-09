import type {
  OfficeParserPlugin,
  OfficePluginResolver,
  OfficeWorkerParser,
} from '../core/parsingContracts';
import type { OfficeDocumentModel, OfficeDocumentSource } from '../core/types';
import type { OfficeExportRequest, OfficeExportResult } from '../export/types';
import type { PreviewKind } from '../services/parsing/formatDefinitions';

/** 声明插件能够在哪个线程执行解析。 */
export type OfficePluginWorkerSupport = 'none' | 'main-thread' | 'worker';

/**
 * 带格式识别和 Worker 元数据的 Core 插件合同。
 * @experimental 插件协议首版，后续版本可能调整字段。
 */
export type OfficeCorePlugin<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = OfficeParserPlugin<TModel, TSource> & {
  /** 插件支持的文件扩展名，注册时会规范化为小写带点形式。 */
  extensions: readonly string[];
  /** 插件声明的 MIME 类型。 */
  mimeTypes?: readonly string[];
  /** 同一扩展名存在多个插件时的确定性优先级。 */
  priority?: number;
  /** 当前插件的 Worker 支持级别。 */
  workerSupport: OfficePluginWorkerSupport;
  /** 对候选文件进行额外格式识别。 */
  detect(
    file: File,
    context: { signal: AbortSignal },
  ): boolean | Promise<boolean>;
  /** 在宿主线程创建插件专属 Worker。 */
  workerFactory?: () => Worker;
  /** 插件专属 Worker 协议适配器。 */
  parseInWorker?: OfficeWorkerParser<TModel, TSource>;
  /** 可选的修改后导出函数。 */
  export?: (request: OfficeExportRequest) => Promise<OfficeExportResult>;
};

/**
 * 格式插件的兼容名称。
 * @experimental 插件协议首版兼容别名。
 */
export type OfficeFormatPlugin<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = OfficeCorePlugin<TModel, TSource>;

/**
 * 插件定义的实例级注册表。
 * @experimental 插件注册表首版。
 */
export type OfficePluginRegistry = OfficePluginResolver & {
  /** 返回当前注册表的冻结插件列表。 */
  list(): readonly OfficeCorePlugin[];
  /** 按插件 ID 查询。 */
  getById(id: string): OfficeCorePlugin | undefined;
  /** 按内置预览格式查询。 */
  getByKind(kind: PreviewKind): OfficeCorePlugin | undefined;
  /** 注册插件并返回幂等注销函数。 */
  register(plugin: OfficeCorePlugin): () => void;
  /** 按插件 ID 注销并返回是否成功。 */
  unregister(id: string): boolean;
  /** 释放注册表索引和监听，不释放文档资源。 */
  dispose(): void;
};

/** 创建插件注册表的选项。 */
export type OfficePluginRegistryOptions = {
  /** 是否装配当前包内置格式，默认 true。 */
  includeBuiltIns?: boolean;
  /** 初始化时注册的外部插件。 */
  plugins?: readonly OfficeCorePlugin[];
};
