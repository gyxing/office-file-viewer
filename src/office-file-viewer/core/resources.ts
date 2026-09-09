import { normalizeOfficeFileViewerError } from '../services/errors/OfficeFileViewerError';
import { createOfficeResourceStore } from '../services/resource-store/OfficeResourceStore';
import type {
  OfficeResourceSource,
  OfficeResourceStore,
  OfficeResourceStoreOptions,
} from '../services/resource-store/types';
import { createOfficeDocumentSession } from '../services/session/OfficeDocumentSession';

/** 宿主可替换的资源读取函数，不能创建全局缓存或长期 Object URL。 */
export type OfficeResourceResolver = (
  source: OfficeResourceSource,
  context: { signal: AbortSignal; sessionId: string },
) => Promise<Blob>;

/** 供解析器和未来编辑器复用的资源会话。 */
export type OfficeResourceSession = {
  readonly store: OfficeResourceStore;
  resolve(source: OfficeResourceSource, signal?: AbortSignal): Promise<Blob>;
  acquire(source: OfficeResourceSource, signal?: AbortSignal): Promise<string>;
  release(source: OfficeResourceSource): void;
  dispose(): Promise<void>;
};

type ResourceSessionOptions = OfficeResourceStoreOptions & {
  resolver?: OfficeResourceResolver;
};

function createAbortError() {
  const error = new Error('资源加载已取消');
  error.name = 'AbortError';
  return error;
}

function validateResourceBlob(source: OfficeResourceSource, blob: Blob) {
  if (
    !blob ||
    typeof blob !== 'object' ||
    typeof blob.size !== 'number' ||
    !Number.isFinite(blob.size) ||
    typeof blob.type !== 'string'
  ) {
    throw new Error('资源解析器必须返回有效 Blob');
  }
  if (source.kind === 'lazy' && source.size > 0 && blob.size !== source.size) {
    throw new Error('资源大小与声明不一致');
  }
  if (
    source.kind === 'lazy' &&
    blob.type &&
    source.mimeType &&
    blob.type.toLowerCase() !== source.mimeType.toLowerCase()
  ) {
    throw new Error('资源 MIME 类型与声明不一致');
  }
  return blob;
}

async function resolveDefaultResource(
  source: OfficeResourceSource,
  signal: AbortSignal | undefined,
) {
  if (signal?.aborted) throw createAbortError();
  const blob =
    source.kind === 'lazy'
      ? await source.load(signal)
      : await (async () => {
          const response = await fetch(source.url, { signal });
          if (!response.ok) throw new Error(`资源请求失败：${response.status}`);
          return response.blob();
        })();
  return validateResourceBlob(source, blob);
}

/** 创建绑定文档会话的资源门面，Store 统一管理引用计数和 Object URL。 */
export function createOfficeResourceSession(
  options: ResourceSessionOptions = {},
): OfficeResourceSession {
  const { resolver, ...storeOptions } = options;
  const store = createOfficeResourceStore(storeOptions);
  const documentSession = createOfficeDocumentSession();
  const sourceAdapters = new Map<string, OfficeResourceSource>();
  documentSession.register({ dispose: () => store.dispose() });

  const adaptSource = (source: OfficeResourceSource) => {
    if (!resolver || source.kind !== 'lazy') return source;
    const cached = sourceAdapters.get(source.id);
    if (cached) return cached;
    const adapted: OfficeResourceSource = {
      ...source,
      load: async (signal) => {
        const blob = await resolver(source, {
          signal: signal ?? documentSession.signal,
          sessionId: documentSession.id,
        });
        return validateResourceBlob(source, blob);
      },
    };
    sourceAdapters.set(source.id, adapted);
    return adapted;
  };

  const resolve = async (
    source: OfficeResourceSource,
    signal?: AbortSignal,
  ) => {
    const effectiveSignal = signal ?? documentSession.signal;
    try {
      const blob =
        resolver && source.kind === 'lazy'
          ? await resolver(source, {
              signal: effectiveSignal,
              sessionId: documentSession.id,
            })
          : await resolveDefaultResource(source, effectiveSignal);
      if (effectiveSignal.aborted) throw createAbortError();
      return validateResourceBlob(source, blob);
    } catch (error) {
      if (error instanceof Error && error.name === 'AbortError') throw error;
      throw normalizeOfficeFileViewerError(error, { stage: 'resource' });
    }
  };

  return {
    store,
    resolve,
    async acquire(source, signal) {
      try {
        return await store.acquire(
          adaptSource(source),
          signal ?? documentSession.signal,
        );
      } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') throw error;
        throw normalizeOfficeFileViewerError(error, { stage: 'resource' });
      }
    },
    release(source) {
      store.release(adaptSource(source));
    },
    dispose() {
      sourceAdapters.clear();
      return documentSession.dispose();
    },
  };
}
