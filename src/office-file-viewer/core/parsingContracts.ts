import type { OfficeParseOptions } from '../services/parsing/types';
import type { OfficeDocumentSession } from './lifecycle';
import type { OfficeResourceSession } from './resources';
import type {
  OfficeCapabilities,
  OfficeDocumentModel,
  OfficeDocumentRuntime,
  OfficeDocumentSnapshot,
  OfficeDocumentSource,
} from './types';

/** 插件解析器可接收的主线程输入，不包含 Viewer 或内部 Runtime。 */
export type OfficeParserInput = Readonly<{
  /** 待读取的原始文件。 */
  file: File;
  /** 当前解析会话的稳定标识。 */
  sessionId: string;
  /** 取消当前解析和资源读取。 */
  signal: AbortSignal;
  /** 已剔除插件注册表的只读选项；Worker 工厂仅在主线程解析中可见。 */
  options: Omit<OfficeParseOptions, 'pluginRegistry'>;
  /** 当前解析会话的资源门面。 */
  resources: OfficeResourceSession;
  /** 当前解析会话的生命周期门面。 */
  documentSession: OfficeDocumentSession;
}>;

/** 可结构化传输给 Worker 的解析输入。 */
export type OfficeWorkerParserInput = Readonly<{
  /** 待读取的原始文件。 */
  file: File;
  /** 当前解析会话的稳定标识。 */
  sessionId: string;
  /** 不含函数和注册表的解析选项。 */
  options: Omit<OfficeParseOptions, 'pluginRegistry' | 'workerFactory'>;
}>;

/** 插件主线程解析函数。 */
export type OfficeParser<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeParserInput,
) => Promise<OfficeDocumentRuntime<TModel, TSource>>;

/** 插件按需 Source 创建函数。 */
export type OfficeSourceFactory<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeParserInput & { snapshot?: OfficeDocumentSnapshot },
) => Promise<OfficeDocumentRuntime<TModel, TSource> | undefined>;

/** 插件自定义 Worker 解析函数，由宿主线程负责创建和销毁 Worker。 */
export type OfficeWorkerParser<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = (
  input: OfficeWorkerParserInput,
  worker: Worker,
) => Promise<OfficeDocumentRuntime<TModel, TSource>>;

/** Core 只依赖的最小插件解析契约。 */
export type OfficeParserPlugin<
  TModel = OfficeDocumentModel,
  TSource = OfficeDocumentSource,
> = Readonly<{
  /** 插件稳定标识。 */
  id: string;
  /** 插件声明的基础能力。 */
  capabilities: OfficeCapabilities;
  /** 主线程解析入口。 */
  parse: OfficeParser<TModel, TSource>;
  /** 可选的按需 Source 入口。 */
  createSource?: OfficeSourceFactory<TModel, TSource>;
  /** 释放插件运行时对象和其私有资源。 */
  dispose?(
    runtime: OfficeDocumentRuntime<TModel, TSource>,
  ): void | Promise<void>;
}>;

/** Core 查询插件的最小结构，避免直接依赖完整注册表实现。 */
export type OfficePluginResolver = Readonly<{
  resolve(
    file: File,
    signal?: AbortSignal,
  ): Promise<OfficeParserPlugin | undefined>;
}>;

/** 低层格式插件契约的兼容别名。 */
export type OfficeFormatPlugin = OfficeParserPlugin;
