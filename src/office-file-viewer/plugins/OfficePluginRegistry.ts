import { getOfficeCapabilities } from '../core/model';
import { OfficeFileViewerError } from '../services/errors/OfficeFileViewerError';
import {
  getOfficeFormatMetadata,
  getPreviewFamily,
  OFFICE_FORMAT_METADATA,
  tryDetectPreviewKind,
  type PreviewKind,
} from '../services/parsing/formatDefinitions';
import { createOfficeBuiltInParserPlugin } from '../services/parsing/formatParserRegistry';
import type {
  OfficeCorePlugin,
  OfficePluginRegistry,
  OfficePluginRegistryOptions,
} from './types';

function normalizeExtension(extension: string) {
  const rawValue =
    typeof extension === 'string' ? extension.trim().toLowerCase() : '';
  // 对外允许写 `docx` 或 `.docx`，注册表统一保存带点形式；路径、通配符和
  // 多段扩展名会破坏索引确定性，因此仍然直接拒绝。
  const value =
    rawValue && rawValue.startsWith('.') ? rawValue : `.${rawValue}`;
  if (!rawValue || !/^\.?[a-z0-9]+$/.test(rawValue)) {
    throw new OfficeFileViewerError('PLUGIN_CONFLICT', '插件扩展名格式无效', {
      stage: 'plugin',
      recoverable: false,
    });
  }
  return value;
}

function normalizeMimeType(mimeType: string) {
  const value =
    typeof mimeType === 'string' ? mimeType.trim().toLowerCase() : '';
  if (!value || value.includes(' ')) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件 MIME 类型格式无效',
      { stage: 'plugin', recoverable: false },
    );
  }
  return value;
}

function validatePlugin(plugin: OfficeCorePlugin) {
  if (
    !plugin ||
    typeof plugin.id !== 'string' ||
    !plugin.id.trim() ||
    typeof plugin.parse !== 'function' ||
    typeof plugin.detect !== 'function'
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件必须提供唯一 ID 和 parse 函数',
      { stage: 'plugin', recoverable: false },
    );
  }
  if (
    !Array.isArray(plugin.extensions) ||
    !plugin.extensions.length ||
    plugin.extensions.some((extension) => typeof extension !== 'string')
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件至少需要声明一个文件扩展名',
      { stage: 'plugin', recoverable: false },
    );
  }
  if (
    plugin.mimeTypes !== undefined &&
    (!Array.isArray(plugin.mimeTypes) ||
      plugin.mimeTypes.some((mimeType) => typeof mimeType !== 'string'))
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件 MIME 类型定义无效',
      { stage: 'plugin', recoverable: false },
    );
  }
  if (plugin.export !== undefined && typeof plugin.export !== 'function') {
    throw new OfficeFileViewerError('PLUGIN_CONFLICT', '插件导出函数定义无效', {
      stage: 'plugin',
      recoverable: false,
    });
  }
  if (
    plugin.workerSupport !== 'none' &&
    plugin.workerSupport !== 'main-thread' &&
    plugin.workerSupport !== 'worker'
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件 Worker 支持级别无效',
      { stage: 'plugin', recoverable: false },
    );
  }
  if (plugin.workerSupport === 'worker') {
    if (!plugin.workerFactory || !plugin.parseInWorker) {
      throw new OfficeFileViewerError(
        'PLUGIN_WORKER_UNSUPPORTED',
        'Worker 插件缺少 workerFactory 或 parseInWorker',
        { stage: 'plugin', recoverable: false },
      );
    }
  }
  const capabilities = plugin.capabilities;
  const capabilityBooleans = [
    'canSearch',
    'canReview',
    'canNavigatePages',
    'canNavigateSheets',
    'canNavigateSlides',
    'canShowSpeakerNotes',
    'canFitPage',
    'canFitWidth',
    'canFullscreen',
    'canExportOriginal',
  ] as const;
  if (
    !capabilities ||
    typeof capabilities !== 'object' ||
    typeof capabilities.previewKind !== 'string' ||
    !capabilities.previewKind ||
    (capabilities.family !== 'word' &&
      capabilities.family !== 'spreadsheet' &&
      capabilities.family !== 'presentation') ||
    !Array.isArray(capabilities.exportFormats) ||
    capabilities.exportFormats.some((format) => typeof format !== 'string') ||
    !capabilities.features ||
    typeof capabilities.features !== 'object' ||
    Array.isArray(capabilities.features) ||
    Object.values(capabilities.features).some(
      (feature) => typeof feature !== 'boolean',
    ) ||
    capabilityBooleans.some((key) => typeof capabilities[key] !== 'boolean')
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件必须提供完整 capabilities',
      { stage: 'plugin', recoverable: false },
    );
  }
  if (
    plugin.priority !== undefined &&
    (!Number.isFinite(plugin.priority) || !Number.isInteger(plugin.priority))
  ) {
    throw new OfficeFileViewerError(
      'PLUGIN_CONFLICT',
      '插件优先级必须是有限整数',
      { stage: 'plugin', recoverable: false },
    );
  }
}

function createBuiltInPlugin(kind: PreviewKind): OfficeCorePlugin {
  const metadata = getOfficeFormatMetadata(kind);
  const parserPlugin = createOfficeBuiltInParserPlugin(kind);
  return {
    ...parserPlugin,
    extensions: metadata.extensions,
    mimeTypes: metadata.mimeTypes.map(({ mimeType }) => mimeType),
    priority: 0,
    workerSupport: 'main-thread',
    detect: (file) => tryDetectPreviewKind(file.name) === kind,
    capabilities: getOfficeCapabilities(kind, {
      family: getPreviewFamily(kind),
    }),
  };
}

function createBuiltInPlugins() {
  return (Object.keys(OFFICE_FORMAT_METADATA) as PreviewKind[]).map(
    createBuiltInPlugin,
  );
}

function getCandidatePriority(plugin: OfficeCorePlugin) {
  return plugin.priority ?? 0;
}

/** 创建实例级插件注册表，不修改全局内置定义。 */
export function createOfficePluginRegistry(
  options: OfficePluginRegistryOptions = {},
): OfficePluginRegistry {
  const byId = new Map<string, OfficeCorePlugin>();
  const byExtension = new Map<string, OfficeCorePlugin[]>();
  const byMimeType = new Map<string, OfficeCorePlugin[]>();
  let disposed = false;

  const removeFromIndex = (plugin: OfficeCorePlugin) => {
    plugin.extensions.forEach((extension) => {
      const candidates = byExtension.get(extension);
      if (!candidates) return;
      const next = candidates.filter((candidate) => candidate !== plugin);
      if (next.length) byExtension.set(extension, next);
      else byExtension.delete(extension);
    });
    plugin.mimeTypes?.forEach((mimeType) => {
      const candidates = byMimeType.get(mimeType);
      if (!candidates) return;
      const next = candidates.filter((candidate) => candidate !== plugin);
      if (next.length) byMimeType.set(mimeType, next);
      else byMimeType.delete(mimeType);
    });
  };

  const unregisterById = (id: string) => {
    const plugin = byId.get(id);
    if (!plugin) return false;
    byId.delete(id);
    removeFromIndex(plugin);
    return true;
  };

  const register = (input: OfficeCorePlugin) => {
    if (disposed) {
      throw new OfficeFileViewerError('PLUGIN_CONFLICT', '插件注册表已释放', {
        stage: 'plugin',
        recoverable: false,
      });
    }
    validatePlugin(input);
    if (byId.has(input.id)) {
      throw new OfficeFileViewerError(
        'PLUGIN_CONFLICT',
        `插件 ID 已注册：${input.id}`,
        { stage: 'plugin', recoverable: false },
      );
    }
    const plugin = Object.freeze({
      ...input,
      extensions: Object.freeze(input.extensions.map(normalizeExtension)),
      mimeTypes: input.mimeTypes
        ? Object.freeze(input.mimeTypes.map(normalizeMimeType))
        : undefined,
      capabilities: Object.freeze({
        ...input.capabilities,
        exportFormats: Object.freeze([...input.capabilities.exportFormats]),
        features: Object.freeze({ ...input.capabilities.features }),
      }),
    }) as OfficeCorePlugin;
    const normalizedExtensions = plugin.extensions;
    const normalizedMimeTypes = plugin.mimeTypes ?? [];
    if (
      new Set(normalizedExtensions).size !== normalizedExtensions.length ||
      new Set(normalizedMimeTypes).size !== normalizedMimeTypes.length
    ) {
      throw new OfficeFileViewerError(
        'PLUGIN_CONFLICT',
        '插件声明了重复的扩展名或 MIME 类型',
        { stage: 'plugin', recoverable: false },
      );
    }
    const assertConflicts = (
      values: readonly string[],
      index: Map<string, OfficeCorePlugin[]>,
      label: string,
    ) => {
      const conflicts = values.flatMap((value) => index.get(value) ?? []);
      conflicts.forEach((candidate) => {
        const samePriority =
          getCandidatePriority(candidate) === getCandidatePriority(plugin);
        const hasExplicitPriority =
          candidate.priority !== undefined && plugin.priority !== undefined;
        if (!hasExplicitPriority || samePriority) {
          throw new OfficeFileViewerError(
            'PLUGIN_CONFLICT',
            `插件${label}冲突：${values.join(', ')}`,
            { stage: 'plugin', recoverable: false },
          );
        }
      });
    };
    assertConflicts(normalizedExtensions, byExtension, '扩展名');
    assertConflicts(normalizedMimeTypes, byMimeType, 'MIME 类型');
    const overriddenBuiltIns = [
      ...normalizedExtensions.flatMap(
        (extension) => byExtension.get(extension) ?? [],
      ),
      ...normalizedMimeTypes.flatMap(
        (mimeType) => byMimeType.get(mimeType) ?? [],
      ),
    ].filter(
      (candidate, index, candidates) =>
        candidate.id.startsWith('builtin:') &&
        getCandidatePriority(plugin) > getCandidatePriority(candidate) &&
        candidates.indexOf(candidate) === index,
    );
    if (
      overriddenBuiltIns.length > 0 &&
      typeof process !== 'undefined' &&
      process.env.NODE_ENV !== 'production'
    ) {
      console.warn(
        `插件 ${plugin.id} 将以更高优先级覆盖内置格式：${overriddenBuiltIns
          .map((candidate) => candidate.id)
          .join(', ')}`,
      );
    }
    byId.set(plugin.id, plugin);
    plugin.extensions.forEach((extension) => {
      const candidates = byExtension.get(extension) ?? [];
      byExtension.set(extension, [...candidates, plugin]);
    });
    plugin.mimeTypes?.forEach((mimeType) => {
      const candidates = byMimeType.get(mimeType) ?? [];
      byMimeType.set(mimeType, [...candidates, plugin]);
    });
    let active = true;
    return () => {
      if (!active) return;
      active = false;
      unregisterById(plugin.id);
    };
  };

  const initialPlugins = [
    ...(options.includeBuiltIns === false ? [] : createBuiltInPlugins()),
    ...(options.plugins ?? []),
  ];
  const unregisters: Array<() => void> = [];
  try {
    initialPlugins.forEach((plugin) => {
      unregisters.push(register(plugin));
    });
  } catch (error) {
    unregisters
      .splice(0)
      .reverse()
      .forEach((unregister) => unregister());
    byId.clear();
    byExtension.clear();
    byMimeType.clear();
    throw error;
  }

  const resolve = async (file: File, signal?: AbortSignal) => {
    if (disposed) return undefined;
    const resolveSignal = signal ?? new AbortController().signal;
    if (resolveSignal.aborted) return undefined;
    const extensionMatch = /\.[a-z0-9]+$/i.exec(file.name);
    const extension = extensionMatch?.[0]?.toLowerCase();
    const indexed = [
      ...(extension ? byExtension.get(extension) ?? [] : []),
      ...(file.type ? byMimeType.get(file.type.toLowerCase()) ?? [] : []),
    ];
    const candidates = [
      ...new Set(indexed.length ? indexed : byId.values()),
    ].sort(
      (left, right) => getCandidatePriority(right) - getCandidatePriority(left),
    );
    let matchedPlugin: OfficeCorePlugin | undefined;
    let matchedPriority: number | undefined;
    for (const [candidateIndex, plugin] of candidates.entries()) {
      if (disposed || resolveSignal.aborted) return undefined;
      const priority = getCandidatePriority(plugin);
      if (matchedPlugin && priority < matchedPriority!) return matchedPlugin;
      try {
        const detected = await plugin.detect(file, { signal: resolveSignal });
        if (disposed || resolveSignal.aborted) return undefined;
        if (!detected) continue;
        if (matchedPlugin && priority === matchedPriority) {
          throw new OfficeFileViewerError(
            'PLUGIN_CONFLICT',
            `插件格式识别结果不唯一：${matchedPlugin.id}、${plugin.id}`,
            { stage: 'plugin', recoverable: false },
          );
        }
        matchedPlugin = plugin;
        matchedPriority = priority;
        // 候选已按优先级降序排列；后续优先级更低时不可能覆盖当前命中。
        const nextCandidate = candidates[candidateIndex + 1];
        if (!nextCandidate || getCandidatePriority(nextCandidate) < priority) {
          return matchedPlugin;
        }
      } catch (error) {
        if (
          resolveSignal.aborted ||
          (error instanceof Error && error.name === 'AbortError')
        ) {
          return undefined;
        }
        if (error instanceof OfficeFileViewerError) throw error;
        throw new OfficeFileViewerError(
          'PLUGIN_CONFLICT',
          `插件格式识别失败：${plugin.id}`,
          { stage: 'plugin', recoverable: false, cause: error },
        );
      }
    }
    return matchedPlugin;
  };

  return {
    list: () => Object.freeze([...byId.values()]),
    getById: (id) => byId.get(id),
    getByKind: (kind) =>
      [...byId.values()]
        .filter((plugin) => plugin.capabilities.previewKind === kind)
        .sort(
          (left, right) =>
            getCandidatePriority(right) - getCandidatePriority(left),
        )[0],
    resolve,
    register,
    unregister: unregisterById,
    dispose() {
      if (disposed) return;
      disposed = true;
      unregisters.splice(0).forEach((unregister) => unregister());
      byId.clear();
      byExtension.clear();
      byMimeType.clear();
    },
  };
}
