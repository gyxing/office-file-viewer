import type { OfficeParseResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import type { PreviewKind } from './formatDefinitions';

export type { OfficeParseResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';

/** 控制解析任务是否使用 Web Worker：自动选择、强制使用或禁用。 */
export type WorkerMode = 'auto' | 'always' | 'never';

/** 枚举文件读取、容器解析、结构提取、内容解析、资源处理和组装等阶段。 */
export type ParseStage =
  | 'reading'
  | 'container'
  | 'structure'
  | 'content'
  | 'resources'
  | 'assembling';

/** 描述解析会话任务的当前进度。 */
export type ParseProgress = {
  /** 解析任务当前所处的处理阶段。 */
  stage: ParseStage;
  /** 当前阶段已经完成的工作量。 */
  completed?: number;
  /** 当前阶段预计需要完成的总工作量。 */
  total?: number;
  /** 面向界面展示的解析完成百分比。 */
  percent?: number;
  /** 当前阶段的可读说明，可直接用于进度界面。 */
  message: string;
};

/** 描述预览首屏已经可以使用时的数据形态。 */
export type OfficePreviewReadyInfo = {
  /** 当前文件对应的预览格式类别。 */
  previewKind: PreviewKind;
  /** 当前预览使用完整模型或按需数据源。 */
  mode: 'materialized' | 'source';
};

/** 定义解析会话的可选配置。 */
export type OfficeParseOptions = {
  /** 控制解析任务使用 Worker、主线程或自动降级的策略。 */
  worker?: WorkerMode;
  /** 创建解析 Worker 的自定义工厂，便于宿主接管 Worker 加载方式。 */
  workerFactory?: () => Worker;
  /** 宿主可选配置的资源上限；未提供时不限制文件大小或解析时长。 */
  resourcePolicy?: OfficeParseResourcePolicy;
};

/** 枚举解析会话从启动到完成、取消或失败的生命周期状态。 */
export type OfficeParseSessionStatus =
  | 'starting'
  | 'running'
  | 'completed'
  | 'cancelled'
  | 'failed';

/** 定义解析会话任务的会话控制接口。 */
export type OfficeParseSession<TParsed> = {
  /** 解析任务的最终结果 Promise；任务失败或取消时会被拒绝。 */
  readonly result: Promise<TParsed>;
  /** 解析会话当前状态，只读且随任务生命周期变化。 */
  readonly status: OfficeParseSessionStatus;
  /** 订阅解析进度，并返回用于取消订阅的函数。 */
  subscribe(listener: (progress: ParseProgress) => void): () => void;
  /** 请求取消当前解析任务。 */
  cancel(): void;
  /** 释放运行时和订阅；完整结果的资源由 disposeParsedOfficeFile 释放。 */
  dispose(): void;
};

/** 表示预览器内部使用的完整模型或按需内容源状态。 */
export type OfficePreviewState<
  TModel extends object,
  TSource extends object,
  TSummary extends object,
> =
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: PreviewKind;
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'materialized';
      /** 模型。 */
      model: TModel;
      /** 当前预览使用的按需加载数据源。 */
      source?: undefined;
      /** 当前预览内容的摘要信息。 */
      summary?: undefined;
    }
  | {
      /** 当前解析或预览会话的标识。 */
      sessionId: string;
      /** 当前文件使用的预览格式类别。 */
      previewKind: PreviewKind;
      /** 当前数据源或渲染器采用的工作模式。 */
      mode: 'source';
      /** 模型。 */
      model?: undefined;
      /** 当前预览使用的按需加载数据源。 */
      source: TSource;
      /** 当前预览内容的摘要信息。 */
      summary: TSummary;
    };

/** 为内部预览状态附加统一的异步资源释放入口。 */
export type OfficePreviewHandle<
  TModel extends object,
  TSource extends object,
  TSummary extends object,
> = OfficePreviewState<TModel, TSource, TSummary> & {
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
};
