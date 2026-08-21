import type { OfficeResourceLimitCode } from '../../shared/resource/OfficeResourceLimitError';
import type { PreviewKind } from '../parsing/formatDefinitions';
import type { ParseStage } from '../parsing/types';

/** OfficeFileViewer 对外稳定暴露的错误阶段。 */
export type OfficeFileViewerErrorStage =
  | 'input'
  | 'download'
  | 'format'
  | 'parsing'
  | 'worker'
  | 'resource'
  | 'fullscreen';

/** OfficeFileViewer 对外稳定暴露的错误代码。 */
export type OfficeFileViewerErrorCode =
  | 'INVALID_INPUT'
  | 'UNSUPPORTED_FILE_TYPE'
  | 'FILE_NAME_UNRECOGNIZED'
  | 'FILE_LOAD_FAILED'
  | 'FILE_DOWNLOAD_FAILED'
  | 'ENCRYPTED_FILE'
  | 'INVALID_FILE'
  | 'WORKER_FAILED'
  | 'PARSE_FAILED'
  | 'FULLSCREEN_FAILED'
  | OfficeResourceLimitCode;

/** 创建结构化错误时可以附加的安全上下文。 */
export type OfficeFileViewerErrorContext = {
  /** 错误发生的稳定阶段。 */
  stage: OfficeFileViewerErrorStage;
  /** 已识别时记录内部预览格式。 */
  previewKind?: PreviewKind;
  /** 已知时记录源文件名，不包含本地路径。 */
  fileName?: string;
  /** 底层解析器或运行时提供的原始错误码。 */
  originalCode?: string;
  /** 解析协议已知时记录当前解析阶段。 */
  parseStage?: ParseStage;
  /** 错误是否可以由调用方重试。 */
  recoverable?: boolean;
  /** 保留底层异常供调用方诊断，不参与界面展示。 */
  cause?: unknown;
};

/** OfficeFileViewer 公共错误，统一输入、解析、资源与交互失败语义。 */
export class OfficeFileViewerError extends Error {
  readonly code: OfficeFileViewerErrorCode;
  readonly stage: OfficeFileViewerErrorStage;
  readonly previewKind?: PreviewKind;
  readonly fileName?: string;
  readonly originalCode?: string;
  readonly parseStage?: ParseStage;
  readonly recoverable: boolean;
  readonly cause?: unknown;

  constructor(
    code: OfficeFileViewerErrorCode,
    message: string,
    context: OfficeFileViewerErrorContext,
  ) {
    super(message);
    this.name = 'OfficeFileViewerError';
    this.code = code;
    this.stage = context.stage;
    this.previewKind = context.previewKind;
    this.fileName = context.fileName;
    this.originalCode = context.originalCode;
    this.parseStage = context.parseStage;
    this.recoverable = context.recoverable ?? false;
    this.cause = context.cause;
  }
}

/** 判断未知异常是否已经使用公共结构化错误模型。 */
export function isOfficeFileViewerError(
  error: unknown,
): error is OfficeFileViewerError {
  return error instanceof OfficeFileViewerError;
}

type ErrorWithRuntimeContext = Error & {
  code?: unknown;
  stage?: unknown;
  recoverable?: unknown;
};

/** 将格式解析器、Worker 和普通异常统一映射为公共错误。 */
export function normalizeOfficeFileViewerError(
  error: unknown,
  context: Omit<OfficeFileViewerErrorContext, 'cause' | 'stage'> & {
    /** 无法从底层错误推断时使用的阶段。 */
    stage?: OfficeFileViewerErrorStage;
  } = {},
): OfficeFileViewerError {
  if (isOfficeFileViewerError(error)) return error;

  const source =
    error instanceof Error ? (error as ErrorWithRuntimeContext) : undefined;
  const originalCode =
    typeof source?.code === 'string' ? source.code : context.originalCode;
  const encrypted =
    originalCode === 'ENCRYPTED_FILE' || originalCode === 'PPT_ENCRYPTED';
  const worker = Boolean(originalCode?.startsWith('WORKER_'));
  const invalid =
    originalCode === 'INVALID_CFB' || originalCode === 'INVALID_SIGNATURE';
  const code: OfficeFileViewerErrorCode = encrypted
    ? 'ENCRYPTED_FILE'
    : worker
    ? 'WORKER_FAILED'
    : invalid
    ? 'INVALID_FILE'
    : 'PARSE_FAILED';
  const stage: OfficeFileViewerErrorStage = worker
    ? 'worker'
    : context.stage ?? 'parsing';

  return new OfficeFileViewerError(code, source?.message ?? '文件解析失败', {
    ...context,
    stage,
    originalCode,
    parseStage:
      typeof source?.stage === 'string'
        ? (source.stage as ParseStage)
        : context.parseStage,
    recoverable:
      typeof source?.recoverable === 'boolean'
        ? source.recoverable
        : context.recoverable,
    cause: error,
  });
}
