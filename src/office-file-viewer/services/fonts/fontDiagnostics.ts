import type { PreviewKind } from '../preview';
import type { OfficeFontFallbackWarning } from '../previewWarnings';
import type { OfficeFontResolution } from './types';

/** 浏览器 Font Loading API 中诊断所需的最小接口。 */
export type OfficeFontSet = {
  /** 判断指定 CSS 字体声明是否可用。 */
  check: (font: string) => boolean;
  /** 浏览器已完成当前字体加载时兑现。 */
  ready?: PromiseLike<unknown>;
};

/** 字体诊断器运行参数。 */
export type OfficeFontDiagnosticsOptions = {
  /** 当前解析会话标识，用于隔离重复警告。 */
  sessionKey?: string;
  /** 当前文件格式。 */
  previewKind: PreviewKind;
  /** 可选的浏览器字体集合；不存在时静默跳过诊断。 */
  fontSet?: OfficeFontSet;
  /** 是否报告缺失字体，默认开启。 */
  warnOnMissing?: boolean;
  /** 缺失字体警告接收器。 */
  onWarning?: (warning: OfficeFontFallbackWarning) => void;
};

/** 收集渲染字体，并在首屏就绪后按批执行诊断。 */
export type OfficeFontDiagnostics = {
  /** 登记渲染阶段遇到的源字体。 */
  register: (resolution: OfficeFontResolution) => void;
  /** 标记首屏已就绪，并为当前及后续登记字体开启批量检查。 */
  activate: () => void;
  /** 检查尚未报告的已登记字体。 */
  flush: () => Promise<void>;
  /** 清理当前会话尚未处理的字体。 */
  dispose: () => void;
};

/** 创建与单个文档会话绑定的字体诊断器。 */
export function createOfficeFontDiagnostics(
  options: OfficeFontDiagnosticsOptions,
): OfficeFontDiagnostics {
  const pending = new Map<string, OfficeFontResolution>();
  const checked = new Set<string>();
  let disposed = false;
  let active = false;
  let flushTimer: ReturnType<typeof setTimeout> | undefined;
  let flushPromise: Promise<void> | undefined;
  const enabled = Boolean(
    options.warnOnMissing !== false && options.fontSet && options.onWarning,
  );

  /** Font Loading API 接受 CSS font 简写，字体名始终加引号可避免空格歧义。 */
  const buildCheckQuery = (fontFamily: string) =>
    `12px "${fontFamily.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;

  const runFlush = async () => {
    if (disposed || !enabled || !options.fontSet || !options.onWarning) {
      return;
    }
    if (options.fontSet.ready) {
      try {
        await options.fontSet.ready;
      } catch {
        // 字体集合就绪失败不应阻断文档预览，继续以当前可用状态检查。
      }
    }
    if (disposed) return;
    pending.forEach((resolution, identity) => {
      if (checked.has(identity) || !resolution.requestedFamily) return;
      checked.add(identity);
      pending.delete(identity);
      let available = true;
      try {
        available =
          options.fontSet?.check(buildCheckQuery(resolution.requestedFamily)) ??
          true;
      } catch {
        // 非法或浏览器不支持的字体声明只跳过诊断，不能影响文档渲染。
        return;
      }
      if (available) {
        return;
      }
      options.onWarning?.({
        code: 'FONT_FALLBACK_APPLIED',
        message: `字体“${
          resolution.requestedFamily
        }”在当前浏览器中不可用，已应用回退字体链：${resolution.candidates.join(
          '、',
        )}`,
        previewKind: options.previewKind,
        source: 'font',
        requestedFamily: resolution.requestedFamily,
        candidates: resolution.candidates,
      });
    });
  };

  const flush = () => {
    if (!flushPromise) {
      flushPromise = runFlush().finally(() => {
        flushPromise = undefined;
      });
    }
    return flushPromise;
  };

  const scheduleFlush = () => {
    if (!enabled || disposed || !active || flushTimer !== undefined) return;
    flushTimer = setTimeout(() => {
      flushTimer = undefined;
      void flush();
    }, 0);
  };

  return {
    register(resolution) {
      if (!enabled || disposed || !resolution.requestedFamily) return;
      const identity = `${
        options.sessionKey ?? ''
      }\u0000${resolution.requestedFamily.toLocaleLowerCase()}`;
      if (!checked.has(identity)) pending.set(identity, resolution);
      scheduleFlush();
    },
    activate() {
      if (disposed) return;
      active = true;
      scheduleFlush();
    },
    flush,
    dispose() {
      disposed = true;
      if (flushTimer !== undefined) clearTimeout(flushTimer);
      pending.clear();
    },
  };
}
