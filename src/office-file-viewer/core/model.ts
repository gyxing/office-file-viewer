import { OfficeFileViewerError } from '../services/errors/OfficeFileViewerError';
import type { PreviewFamily } from '../services/parsing/formatDefinitions';
import {
  getOfficeFormatMetadata,
  getPreviewFamily,
  OFFICE_FORMAT_METADATA,
  type PreviewKind,
} from '../services/parsing/formatDefinitions';
import type { ParsedOfficeFile } from '../services/preview';
import type { OfficeResourceDescriptor } from '../services/resource-store/types';
import type {
  OfficeCapabilities,
  OfficeDocumentModel,
  OfficeDocumentRuntime,
  OfficeDocumentSnapshot,
  OfficeDocumentSource,
  OfficeFormatId,
  OfficeJsonValue,
  OfficeResourceRef,
  OfficeSourceMap,
  OfficeWarning,
} from './types';

/** 由解析会话提供的物化模型适配输入。 */
export type OfficeDocumentSnapshotInput = Readonly<{
  /** 当前解析会话标识。 */
  sessionId: string;
  /** 原始文件对象。 */
  file: File;
  /** 现有格式解析器产出的完整模型。 */
  parsed: ParsedOfficeFile;
}>;

/** 构造 Source 模式快照所需的可序列化描述。 */
export type OfficeSourceSnapshotInput = Readonly<{
  /** 当前解析会话标识。 */
  sessionId: string;
  /** 原始文件名。 */
  fileName?: string;
  /** 原始文件大小。 */
  fileSize?: number;
  /** 当前 Source 对应的预览格式。 */
  format: OfficeFormatId;
  /** Source 所属的预览界面族。 */
  family: PreviewFamily;
  /** Source 当前实际可用的能力。 */
  capabilities: OfficeCapabilities;
  /** Source 使用的资源引用。 */
  resources?: readonly OfficeResourceRef[];
  /** Source 解析阶段产生的可恢复警告。 */
  warnings?: readonly OfficeWarning[];
  /** 语义节点到源文件位置的映射。 */
  sourceMap?: OfficeSourceMap;
  /** 格式插件提供的可序列化扩展数据。 */
  extensionData?: Readonly<Record<string, OfficeJsonValue>>;
}>;

function sanitizeJsonValue(
  value: unknown,
  seen: Set<object>,
): OfficeJsonValue | undefined {
  if (value === null || typeof value === 'string' || typeof value === 'boolean')
    return value;
  if (typeof value === 'number')
    return Number.isFinite(value) ? value : undefined;
  if (typeof value !== 'object' || seen.has(value)) return undefined;
  seen.add(value);
  try {
    if (Array.isArray(value)) {
      const result: OfficeJsonValue[] = [];
      for (const item of value) {
        const sanitized = sanitizeJsonValue(item, seen);
        if (sanitized === undefined && item !== null) return undefined;
        result.push(sanitized ?? null);
      }
      return Object.freeze(result);
    }
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return undefined;
    const result = Object.create(null) as Record<string, OfficeJsonValue>;
    for (const [key, item] of Object.entries(value)) {
      const sanitized = sanitizeJsonValue(item, seen);
      if (sanitized === undefined && item !== null) return undefined;
      result[key] = sanitized ?? null;
    }
    return Object.freeze(result);
  } finally {
    seen.delete(value);
  }
}

function sanitizeExtensionData(
  extensionData: Readonly<Record<string, OfficeJsonValue>> | undefined,
) {
  if (!extensionData) return { value: undefined, warning: undefined };
  const value = sanitizeJsonValue(extensionData, new Set());
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {
      value: undefined,
      warning: {
        code: 'INVALID_EXTENSION_DATA',
        message: '插件扩展数据包含不可序列化值，已忽略。',
        source: 'core-model',
        recoverable: true,
      } as const,
    };
  }
  return { value, warning: undefined };
}

function sanitizeSourceMap(sourceMap: OfficeSourceMap | undefined) {
  if (!sourceMap || typeof sourceMap !== 'object') return undefined;
  const result = Object.create(null) as Record<string, OfficeSourceMap[string]>;
  Object.entries(sourceMap).forEach(([nodeId, location]) => {
    if (!location || typeof location !== 'object') return;
    const safeLocation = {
      ...(typeof location.part === 'string' ? { part: location.part } : {}),
      ...(typeof location.objectId === 'string'
        ? { objectId: location.objectId }
        : {}),
      ...(typeof location.offset === 'number' &&
      Number.isFinite(location.offset)
        ? { offset: location.offset }
        : {}),
    };
    result[nodeId] = Object.freeze(safeLocation);
  });
  return Object.freeze(result);
}

function freezeCapabilities(
  capabilities: OfficeCapabilities,
): OfficeCapabilities {
  return Object.freeze({
    ...capabilities,
    exportFormats: Object.freeze([...capabilities.exportFormats]),
    features: Object.freeze({ ...capabilities.features }),
  });
}

/** 统一冻结对外快照的嵌套容器，避免只读类型在运行时被意外改写。 */
function freezeOfficeSnapshot(
  snapshot: OfficeDocumentSnapshot,
): OfficeDocumentSnapshot {
  return Object.freeze({
    ...snapshot,
    metadata: Object.freeze({ ...snapshot.metadata }),
    capabilities: freezeCapabilities(snapshot.capabilities),
    resources: Object.freeze([...snapshot.resources]),
    warnings: snapshot.warnings
      ? Object.freeze(
          snapshot.warnings.map((warning) => Object.freeze({ ...warning })),
        )
      : undefined,
  });
}

function createSnapshotInputError(message: string, format?: string) {
  return new OfficeFileViewerError('INVALID_FILE', message, {
    stage: 'format',
    format,
    recoverable: false,
  });
}

function assertSourceSnapshotInput(input: OfficeSourceSnapshotInput) {
  if (!input || typeof input !== 'object') {
    throw createSnapshotInputError('Source 快照输入无效');
  }
  if (typeof input.sessionId !== 'string' || !input.sessionId.trim()) {
    throw createSnapshotInputError('Source 快照缺少会话标识');
  }
  if (typeof input.format !== 'string' || !input.format.trim()) {
    throw createSnapshotInputError('Source 快照缺少格式标识');
  }
  if (
    input.family !== 'word' &&
    input.family !== 'spreadsheet' &&
    input.family !== 'presentation'
  ) {
    throw createSnapshotInputError('Source 快照的界面族无效', input.format);
  }
  if (
    !input.capabilities ||
    input.capabilities.previewKind !== input.format ||
    input.capabilities.family !== input.family
  ) {
    throw createSnapshotInputError(
      'Source 快照的能力与格式或界面族不一致',
      input.format,
    );
  }
  if (
    input.fileSize !== undefined &&
    (!Number.isFinite(input.fileSize) || input.fileSize < 0)
  ) {
    throw createSnapshotInputError('Source 快照文件大小无效', input.format);
  }
}
/** 依据格式和当前渲染能力生成统一能力描述。 */
export function getOfficeCapabilities(
  format: PreviewKind,
  available: Partial<OfficeCapabilities> = {},
): OfficeCapabilities {
  if (!Object.prototype.hasOwnProperty.call(OFFICE_FORMAT_METADATA, format)) {
    throw new OfficeFileViewerError(
      'INVALID_FILE',
      `无法生成未知格式的能力描述：${String(format)}`,
      { stage: 'format', format: String(format), recoverable: false },
    );
  }
  const family = getPreviewFamily(format);
  const metadata = getOfficeFormatMetadata(format);
  const {
    features: availableFeatures,
    // 格式、界面族和导出后缀由唯一元数据源决定，避免调用方传入不一致值。
    previewKind: _previewKind,
    family: _family,
    exportFormats: _exportFormats,
    ...availableFields
  } = available;
  void _previewKind;
  void _family;
  void _exportFormats;
  return Object.freeze({
    previewKind: format,
    family,
    canSearch: family !== undefined,
    // 三类内置格式都支持只读批注或修订展示；是否显示入口仍由实际内容决定。
    canReview: true,
    canNavigatePages: family === 'word',
    canNavigateSheets: family === 'spreadsheet',
    canNavigateSlides: family === 'presentation',
    canShowSpeakerNotes: family === 'presentation',
    canFitPage: true,
    canFitWidth: true,
    canFullscreen: true,
    canExportOriginal: true,
    exportFormats: Object.freeze([...metadata.extensions]),
    ...availableFields,
    features: Object.freeze({
      ...(availableFeatures ?? {}),
    }),
  });
}

function getParsedModel(parsed: ParsedOfficeFile): OfficeDocumentModel {
  if (!parsed || typeof parsed !== 'object') {
    throw createSnapshotInputError('解析结果无效');
  }
  let model: OfficeDocumentModel;
  switch (parsed.kind) {
    case 'doc':
    case 'docx':
    case 'ppt':
    case 'pptx':
      model = parsed.document;
      break;
    case 'xls':
    case 'xlsx':
      model = parsed.workbook;
      break;
    default:
      throw new OfficeFileViewerError(
        'INVALID_FILE',
        '解析结果缺少有效的 Office 格式',
        { stage: 'format', recoverable: false },
      );
  }
  if (!model || typeof model !== 'object') {
    throw createSnapshotInputError('解析结果缺少文档模型', parsed.kind);
  }
  return model;
}

type ResourceSourceLike =
  | { kind: 'url'; url: string }
  | { kind: 'lazy'; id: string; mimeType: string; size: number };

const RESOURCE_FIELD_NAMES = new Set([
  'src',
  'imageRef',
  'imageSrc',
  'posterSrc',
  'snapshotSource',
  'snapshotSrc',
  'source',
  'resource',
  'resources',
]);

/** 模型中可能继续包含资源引用的容器字段，避免快照生成遍历全部单元格样式。 */
const RESOURCE_CONTAINER_FIELD_NAMES = new Set([
  ...RESOURCE_FIELD_NAMES,
  'resourceRefs',
  'images',
  'image',
  'slides',
  'elements',
  'children',
  'background',
  'chart',
  'shape',
  'table',
  'media',
  'pages',
  'blocks',
  'rows',
  'cells',
  'inlines',
  'items',
  'headers',
  'footers',
  'default',
  'first',
  'even',
]);

function isResourceSource(value: unknown): value is ResourceSourceLike {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Record<string, unknown>;
  if (candidate.kind === 'url') {
    return typeof candidate.url === 'string' && candidate.url.length > 0;
  }
  return (
    candidate.kind === 'lazy' &&
    typeof candidate.id === 'string' &&
    typeof candidate.mimeType === 'string' &&
    typeof candidate.size === 'number'
  );
}

function inferResourceKind(
  fieldName: string | undefined,
  mimeType?: string,
): OfficeResourceRef['kind'] {
  const field = fieldName?.toLowerCase() ?? '';
  const mime = mimeType?.toLowerCase() ?? '';
  if (
    field.includes('font') ||
    mime.startsWith('font/') ||
    /(?:ttf|otf|woff2?|eot)(?:$|[+;])/i.test(mime)
  ) {
    return 'font';
  }
  if (
    field.includes('media') ||
    field.includes('poster') ||
    mime.startsWith('audio/') ||
    mime.startsWith('video/') ||
    mime.includes('media')
  ) {
    return 'media';
  }
  if (
    field.includes('image') ||
    field.includes('picture') ||
    field === 'src' ||
    mime.startsWith('image/')
  ) {
    return 'image';
  }
  return 'other';
}

function hashResourceUrl(value: string) {
  // 只用哈希生成没有显式 ID 的临时引用，避免把 Blob URL 写进快照。
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16);
}

function mergeResourceRef(
  refs: Map<string, OfficeResourceRef>,
  ref: OfficeResourceRef,
) {
  const previous = refs.get(ref.id);
  if (!previous) {
    refs.set(ref.id, Object.freeze({ ...ref }));
    return;
  }
  refs.set(
    ref.id,
    Object.freeze({
      ...previous,
      ...ref,
      locator: ref.locator ?? previous.locator,
      mimeType: ref.mimeType ?? previous.mimeType,
      size: ref.size ?? previous.size,
    }),
  );
}

function descriptorToResourceRef(
  descriptor: OfficeResourceDescriptor,
): OfficeResourceRef {
  return Object.freeze({
    id: descriptor.id,
    kind: descriptor.kind,
    mimeType: descriptor.mimeType,
    size: descriptor.size,
    locator:
      descriptor.mimeType && descriptor.size !== undefined
        ? {
            kind: 'lazy' as const,
            id: descriptor.id,
            mimeType: descriptor.mimeType,
            size: descriptor.size,
          }
        : undefined,
  });
}

function sanitizeResourceRefs(
  resources: readonly OfficeResourceRef[] | undefined,
): readonly OfficeResourceRef[] {
  if (!resources?.length) return Object.freeze([]);
  const refs = new Map<string, OfficeResourceRef>();
  resources.forEach((resource) => {
    if (!resource || typeof resource.id !== 'string' || !resource.id) return;
    const kind =
      resource.kind === 'image' ||
      resource.kind === 'media' ||
      resource.kind === 'font' ||
      resource.kind === 'other'
        ? resource.kind
        : 'other';
    const size =
      typeof resource.size === 'number' &&
      Number.isFinite(resource.size) &&
      resource.size >= 0
        ? resource.size
        : undefined;
    const locator = resource.locator;
    const safeLocator =
      locator?.kind === 'lazy' &&
      typeof locator.id === 'string' &&
      typeof locator.mimeType === 'string' &&
      Number.isFinite(locator.size) &&
      locator.size >= 0
        ? {
            kind: 'lazy' as const,
            id: locator.id,
            mimeType: locator.mimeType,
            size: locator.size,
          }
        : locator?.kind === 'url' &&
          typeof locator.url === 'string' &&
          !/^blob:/i.test(locator.url)
        ? { kind: 'url' as const, url: locator.url }
        : undefined;
    mergeResourceRef(
      refs,
      Object.freeze({
        id: resource.id,
        kind,
        mimeType:
          typeof resource.mimeType === 'string' ? resource.mimeType : undefined,
        size,
        locator: safeLocator,
      }),
    );
  });
  return Object.freeze([...refs.values()]);
}

/** 从模型或 Source 摘要提取不含运行时对象的资源引用。 */
export function collectOfficeResourceRefs(
  value: unknown,
): readonly OfficeResourceRef[] {
  const refs = new Map<string, OfficeResourceRef>();
  const descriptorIds = new Set<string>();
  const visited = new Set<object>();

  const visit = (
    current: unknown,
    fieldName?: string,
    ownerId?: string,
    depth = 0,
    inResourceRefs = false,
  ): void => {
    if (depth > 16 || current === null || current === undefined) return;
    if (isResourceSource(current)) {
      const kind = inferResourceKind(
        fieldName,
        current.kind === 'lazy' ? current.mimeType : undefined,
      );
      if (current.kind === 'lazy') {
        mergeResourceRef(refs, {
          id: current.id,
          kind,
          mimeType: current.mimeType,
          size: current.size,
          locator: {
            kind: 'lazy' as const,
            id: current.id,
            mimeType: current.mimeType,
            size: current.size,
          },
        });
      } else {
        const id = ownerId ?? `url:${hashResourceUrl(current.url)}`;
        mergeResourceRef(refs, {
          id,
          kind,
          locator: /^blob:/i.test(current.url)
            ? undefined
            : { kind: 'url' as const, url: current.url },
        });
      }
      return;
    }
    if (typeof current === 'string') {
      if (current.startsWith('office-resource:')) {
        const id = current.slice('office-resource:'.length);
        if (id) {
          mergeResourceRef(refs, {
            id,
            kind: inferResourceKind(fieldName),
          });
        }
      } else if (
        RESOURCE_FIELD_NAMES.has(fieldName ?? '') &&
        /^(?:blob:|https?:|data:)/i.test(current)
      ) {
        const id = ownerId ?? `url:${hashResourceUrl(current)}`;
        mergeResourceRef(refs, {
          id,
          kind: inferResourceKind(fieldName),
          locator: /^blob:/i.test(current)
            ? undefined
            : { kind: 'url' as const, url: current },
        });
      }
      return;
    }
    if (typeof current !== 'object' || visited.has(current)) return;
    if (
      current instanceof ArrayBuffer ||
      ArrayBuffer.isView(current) ||
      (typeof Blob !== 'undefined' && current instanceof Blob)
    ) {
      return;
    }
    if (Array.isArray(current)) {
      visited.add(current);
      try {
        current.forEach((item) =>
          visit(item, fieldName, ownerId, depth + 1, inResourceRefs),
        );
      } finally {
        visited.delete(current);
      }
      return;
    }
    visited.add(current);
    try {
      const record = current as Record<string, unknown>;
      const currentId =
        typeof record.id === 'string' && record.id.length > 0
          ? record.id
          : ownerId;
      if (inResourceRefs && typeof record.id === 'string') {
        const descriptor = record as unknown as OfficeResourceDescriptor;
        if (
          descriptor.kind === 'image' ||
          descriptor.kind === 'media' ||
          descriptor.kind === 'font' ||
          descriptor.kind === 'other'
        ) {
          descriptorIds.add(descriptor.id);
          mergeResourceRef(refs, descriptorToResourceRef(descriptor));
        }
      }
      Object.entries(record).forEach(([key, child]) => {
        // 函数、对象 URL 列表和底层二进制不属于可序列化快照内容。
        if (
          key === 'load' ||
          key === 'buffer' ||
          key === 'objectUrls' ||
          !RESOURCE_CONTAINER_FIELD_NAMES.has(key)
        ) {
          return;
        }
        visit(
          child,
          key,
          currentId,
          depth + 1,
          inResourceRefs || key === 'resourceRefs',
        );
      });
    } finally {
      visited.delete(current);
    }
  };

  visit(value);
  if (descriptorIds.size) {
    [...refs.keys()].forEach((id) => {
      const resource = refs.get(id)!;
      if (!descriptorIds.has(id) && resource.locator === undefined) {
        refs.delete(id);
      }
    });
  }
  return Object.freeze([...refs.values()]);
}

/** 将各格式的警告对象收敛为可序列化的 Core 警告。 */
export function normalizeOfficeWarnings(
  warnings: unknown,
): readonly OfficeWarning[] {
  if (!Array.isArray(warnings)) return [];
  return Object.freeze(
    warnings.flatMap((warning): OfficeWarning[] => {
      if (typeof warning === 'string') {
        return [
          Object.freeze({
            code: 'PARSE_WARNING',
            message: warning,
            source: 'parser',
            recoverable: true,
          }),
        ];
      }
      if (!warning || typeof warning !== 'object') return [];
      const record = warning as Record<string, unknown>;
      if (typeof record.message !== 'string') return [];
      return [
        Object.freeze({
          code: typeof record.code === 'string' ? record.code : 'PARSE_WARNING',
          message: record.message,
          source: typeof record.source === 'string' ? record.source : 'parser',
          recoverable:
            typeof record.recoverable === 'boolean' ? record.recoverable : true,
        }),
      ];
    }),
  );
}

function getParsedWarnings(parsed: ParsedOfficeFile): readonly OfficeWarning[] {
  const model =
    parsed.kind === 'xls' || parsed.kind === 'xlsx'
      ? parsed.workbook
      : parsed.document;
  return normalizeOfficeWarnings((model as { warnings?: unknown }).warnings);
}

function getParsedResources(
  parsed: ParsedOfficeFile,
): readonly OfficeResourceRef[] {
  const model =
    parsed.kind === 'xls' || parsed.kind === 'xlsx'
      ? parsed.workbook
      : parsed.document;
  const refs = new Map(
    collectOfficeResourceRefs(model).map((resource) => [resource.id, resource]),
  );
  const descriptors = (
    model as {
      resources?: { resourceRefs?: readonly OfficeResourceDescriptor[] };
    }
  ).resources?.resourceRefs;
  if (descriptors?.length) {
    const descriptorIds = new Set(
      descriptors.map((descriptor) => descriptor.id),
    );
    [...refs.keys()].forEach((id) => {
      const resource = refs.get(id)!;
      // 有完整描述符时丢弃仅由临时 Blob URL 推导出的模型对象 ID，避免同一资源出现两份引用。
      if (!descriptorIds.has(id) && resource.locator === undefined) {
        refs.delete(id);
      }
    });
  }
  descriptors?.forEach((descriptor) => {
    mergeResourceRef(refs, descriptorToResourceRef(descriptor));
  });
  return Object.freeze([...refs.values()]);
}

/** 从现有物化解析结果生成可跨 Worker 传输的语义快照。 */
export function createOfficeDocumentSnapshot(
  input: OfficeDocumentSnapshotInput,
): OfficeDocumentSnapshot {
  if (!input || typeof input !== 'object' || !input.file) {
    throw createSnapshotInputError('物化快照输入无效');
  }
  if (typeof input.sessionId !== 'string' || !input.sessionId.trim()) {
    throw createSnapshotInputError('物化快照缺少会话标识');
  }
  const { parsed, file, sessionId } = input;
  const model = getParsedModel(parsed);
  const capabilities = getOfficeCapabilities(parsed.kind);
  const metadata = {
    fileName: file.name,
    fileSize: file.size,
    ...(typeof (model as { title?: unknown }).title === 'string'
      ? { title: (model as { title: string }).title }
      : {}),
    ...('author' in model && typeof model.author === 'string'
      ? { author: model.author }
      : {}),
  };
  const resources = getParsedResources(parsed);
  const warnings = getParsedWarnings(parsed);
  return freezeOfficeSnapshot({
    id: `${sessionId}:${parsed.kind}`,
    format: parsed.kind,
    family: capabilities.family,
    metadata,
    capabilities,
    resources,
    warnings: warnings.length ? warnings : undefined,
  });
}

/** 将物化模型和语义快照绑定为会话运行时对象。 */
export function createOfficeDocumentRuntime<TModel = OfficeDocumentModel>(
  input: OfficeDocumentSnapshotInput,
): OfficeDocumentRuntime<TModel, never> {
  return Object.freeze({
    snapshot: createOfficeDocumentSnapshot(input),
    mode: 'materialized',
    model: getParsedModel(input.parsed) as TModel,
  });
}

/** 从 Source 描述生成不携带运行时对象的语义快照。 */
export function createOfficeSourceSnapshot(
  input: OfficeSourceSnapshotInput,
): OfficeDocumentSnapshot {
  assertSourceSnapshotInput(input);
  const sanitizedExtensionData = sanitizeExtensionData(input.extensionData);
  const sourceWarnings = normalizeOfficeWarnings(input.warnings);
  return freezeOfficeSnapshot({
    id: `${input.sessionId}:${input.format}`,
    format: input.format,
    family: input.family,
    metadata: {
      fileName: input.fileName,
      fileSize: input.fileSize,
    },
    capabilities: input.capabilities,
    resources: sanitizeResourceRefs(input.resources),
    sourceMap: sanitizeSourceMap(input.sourceMap),
    extensionData: sanitizedExtensionData.value as
      | Readonly<Record<string, OfficeJsonValue>>
      | undefined,
    warnings:
      sourceWarnings.length || sanitizedExtensionData.warning
        ? Object.freeze([
            ...sourceWarnings,
            ...(sanitizedExtensionData.warning
              ? [sanitizedExtensionData.warning]
              : []),
          ])
        : undefined,
  });
}

/** 将公开快照与外部 Source 绑定为运行时对象。 */
export function createOfficeSourceRuntime<TSource = OfficeDocumentSource>(
  snapshot: OfficeDocumentSnapshot,
  source: TSource,
): OfficeDocumentRuntime<never, TSource> {
  return Object.freeze({ snapshot, mode: 'source', source });
}

/** 从运行时对象读取物化模型；Source 模式返回 undefined。 */
export function getOfficeDocumentModel<
  TModel = OfficeDocumentModel,
  TSource = never,
>(runtime: OfficeDocumentRuntime<TModel, TSource>): TModel | undefined {
  return runtime.mode === 'materialized' ? runtime.model : undefined;
}
