import type { PreviewKind } from '../../preview';
import type { ParseStage } from '../types';

/** 描述跨线程解析协议失败时使用的结构化错误。 */
export type SerializedParseError = {
  /** SerializedParseError 的稳定代码，用于程序化识别具体情况。 */
  code: string;
  /** SerializedParseError 面向调用方或用户展示的具体警告、错误说明。 */
  message: string;
  /** SerializedParseError 的文件格式或协议类型标识；未提供时使用来源格式或渲染器的默认行为。 */
  format?: PreviewKind;
  /** 解析任务当前所处的处理阶段。 */
  stage?: ParseStage;
  /** SerializedParseError 在所属数据范围中的偏移位置。 */
  offset?: number;
  /** 表示该错误是否允许自动降级到其他解析方式。 */
  recoverable: boolean;
};

/** 汇总跨线程解析协议当前步骤需要共享的上下文。 */
type ErrorWithContext = Error & {
  /** ErrorWithContext 的稳定代码，用于程序化识别具体情况。 */
  code?: unknown;
  /** ErrorWithContext 在所属数据范围中的偏移位置。 */
  offset?: unknown;
};

/** 将运行时异常收敛为可安全跨线程传输的错误信息。 */
export function serializeParseError(
  error: unknown,
  context: {
    /** 当前局部结构 的文件格式或协议类型标识；未提供时使用来源格式或渲染器的默认行为。 */
    format?: PreviewKind;
    /** 解析任务当前所处的处理阶段。 */
    stage?: ParseStage;
    /** 表示该错误是否允许自动降级到其他解析方式。 */
    recoverable?: boolean;
  } = {},
): SerializedParseError {
  const normalized =
    error instanceof Error ? (error as ErrorWithContext) : undefined;
  return {
    code:
      typeof normalized?.code === 'string'
        ? normalized.code
        : 'WORKER_PARSE_FAILED',
    message: normalized?.message ?? '文件解析失败',
    format: context.format,
    stage: context.stage,
    offset:
      typeof normalized?.offset === 'number' ? normalized.offset : undefined,
    recoverable: context.recoverable ?? false,
  };
}

/** 将跨线程错误恢复为 Error，并保留稳定错误码供调用方判断。 */
export function deserializeParseError(source: SerializedParseError): Error {
  const error = new Error(source.message) as Error & {
    /** 序列化错误对象的稳定代码，用于程序化识别具体情况。 */
    code: string;
    /** 当前局部结构 的文件格式或协议类型标识；未提供时使用来源格式或渲染器的默认行为。 */
    format?: PreviewKind;
    /** 解析任务当前所处的处理阶段。 */
    stage?: ParseStage;
    /** 当前结构 在所属数据范围中的偏移位置。 */
    offset?: number;
    /** 表示该错误是否允许自动降级到其他解析方式。 */
    recoverable: boolean;
  };
  error.name = 'OfficeParseError';
  error.code = source.code;
  error.format = source.format;
  error.stage = source.stage;
  error.offset = source.offset;
  error.recoverable = source.recoverable;
  return error;
}
