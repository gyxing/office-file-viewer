import type { ReactNode, RefObject } from 'react';
import React, {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from 'react';
import type { DocxEmbeddedFontSession } from '../../services/docx/loadDocxEmbeddedFonts';
import {
  resolveOfficeFont,
  resolveOfficeFontFamily,
} from '../../services/fonts/OfficeFontResolver';
import {
  createOfficeFontDiagnostics,
  type OfficeFontSet,
} from '../../services/fonts/fontDiagnostics';
import type {
  OfficeFileViewerFontOptions,
  OfficeFontFamilyResolver,
} from '../../services/fonts/types';
import type { PreviewKind } from '../../services/preview';
import type { OfficeFileViewerWarning } from '../../services/previewWarnings';

/** 未处于预览器 Provider 内时仍保留基础字体链解析能力。 */
const OfficeFontResolverContext = createContext<OfficeFontFamilyResolver>(
  resolveOfficeFontFamily,
);

/** 非 DOCX 或不支持内嵌字体的环境无需等待额外字体资源。 */
const OfficeEmbeddedFontsReadyContext = createContext(true);

/** 隔离相邻文档会话的字体加载状态，避免新文档复用旧会话的就绪值。 */
type OfficeEmbeddedFontLoadState = {
  /** 当前状态所属的文件对象。 */
  file?: File;
  /** 当前状态所属的文档会话标识。 */
  sessionKey?: string;
  /** 当前会话的内嵌字体是否已完成加载或安全降级。 */
  ready: boolean;
  /** 源字体名到当前文档内嵌字体族的映射。 */
  aliases?: Readonly<Record<string, string>>;
};

/** 字体 Provider 与当前文档会话的连接参数。 */
type OfficeFontProviderProps = {
  /** 当前预览器包含的格式渲染内容。 */
  children: ReactNode;
  /** 宿主提供的字体别名、全局回退和诊断配置。 */
  options?: OfficeFileViewerFontOptions;
  /** 当前预览器根节点，用于读取其所属文档的字体集合。 */
  containerRef: RefObject<HTMLDivElement>;
  /** 当前文档会话标识。 */
  sessionKey?: string;
  /** 首屏是否已经具备可渲染内容。 */
  ready: boolean;
  /** 当前正在预览的文件。 */
  file?: File;
  /** 当前文件格式。 */
  previewKind?: PreviewKind;
  /** 字体缺失时通知宿主。 */
  onWarning?: (warning: OfficeFileViewerWarning, file: File) => void;
};

/** 读取当前预览器所属文档的 Font Loading API，旧浏览器中返回空。 */
function getOwnerFontSet(containerRef: RefObject<HTMLDivElement>) {
  const ownerDocument =
    containerRef.current?.ownerDocument ??
    (typeof document === 'undefined' ? undefined : document);
  return (ownerDocument as (Document & { fonts?: OfficeFontSet }) | undefined)
    ?.fonts;
}

/** 为一个文档会话统一解析字体链，并在首屏提交后批量诊断缺失字体。 */
export function OfficeFontProvider({
  children,
  options,
  containerRef,
  sessionKey,
  ready,
  file,
  previewKind,
  onWarning,
}: OfficeFontProviderProps) {
  const shouldLoadEmbeddedFonts = previewKind === 'docx' && Boolean(file);
  const [embeddedFontState, setEmbeddedFontState] =
    useState<OfficeEmbeddedFontLoadState>({ ready: true });

  useEffect(() => {
    if (!shouldLoadEmbeddedFonts || !file) {
      setEmbeddedFontState({ ready: true });
      return undefined;
    }
    const ownerDocument =
      containerRef.current?.ownerDocument ??
      (typeof document === 'undefined' ? undefined : document);
    if (!ownerDocument) {
      setEmbeddedFontState({ file, sessionKey, ready: true });
      return undefined;
    }

    const controller = new AbortController();
    let fontSession: DocxEmbeddedFontSession | undefined;
    setEmbeddedFontState({ file, sessionKey, ready: false });
    void import('../../services/docx/loadDocxEmbeddedFonts')
      .then((module) =>
        module.loadDocxEmbeddedFonts(file, ownerDocument, controller.signal),
      )
      .then(
        (session) => {
          if (controller.signal.aborted) {
            session.dispose();
            return;
          }
          fontSession = session;
          setEmbeddedFontState({
            file,
            sessionKey,
            ready: true,
            aliases: session.aliases,
          });
        },
        () => {
          if (!controller.signal.aborted) {
            // 字体包异常时继续使用既有回退链，不能阻断文档正文预览。
            setEmbeddedFontState({ file, sessionKey, ready: true });
          }
        },
      );

    return () => {
      controller.abort();
      fontSession?.dispose();
    };
  }, [containerRef, file, previewKind, sessionKey, shouldLoadEmbeddedFonts]);

  const embeddedFontsReady =
    !shouldLoadEmbeddedFonts ||
    (embeddedFontState.file === file &&
      embeddedFontState.sessionKey === sessionKey &&
      embeddedFontState.ready);
  const embeddedFontAliases = embeddedFontsReady
    ? embeddedFontState.aliases
    : undefined;
  const diagnostics = useMemo(() => {
    if (!file || !previewKind) return undefined;
    return createOfficeFontDiagnostics({
      sessionKey,
      previewKind,
      fontSet: getOwnerFontSet(containerRef),
      warnOnMissing: options?.warnOnMissing,
      onWarning: onWarning
        ? (warning) => {
            try {
              onWarning(warning, file);
            } catch (error) {
              // 宿主观察回调异常不应反向中断预览渲染。
              setTimeout(() => {
                throw error;
              }, 0);
            }
          }
        : undefined,
    });
  }, [
    containerRef,
    file,
    onWarning,
    options?.warnOnMissing,
    previewKind,
    sessionKey,
  ]);
  useEffect(() => () => diagnostics?.dispose(), [diagnostics]);
  useEffect(() => {
    if (!ready || !embeddedFontsReady || !diagnostics) return undefined;
    // 激活后由诊断器合并当前字体及异步内容后续登记的字体，避免逐节点同步检查。
    diagnostics.activate();
    return undefined;
  }, [diagnostics, embeddedFontsReady, ready]);

  const resolutionCache = useMemo(
    () => new Map<string, ReturnType<typeof resolveOfficeFont>>(),
    [options],
  );
  const resolveFontFamily = useCallback<OfficeFontFamilyResolver>(
    (requestedFamily) => {
      const cacheKey = requestedFamily ?? '';
      const resolution =
        resolutionCache.get(cacheKey) ??
        resolveOfficeFont(requestedFamily, options);
      if (!resolutionCache.has(cacheKey)) {
        resolutionCache.set(cacheKey, resolution);
      }
      const embeddedFamily = resolution.requestedFamily
        ? embeddedFontAliases?.[
            resolution.requestedFamily.toLocaleLowerCase()
          ]
        : undefined;
      if (!embeddedFamily) diagnostics?.register(resolution);
      return [
        embeddedFamily ? `"${embeddedFamily}"` : undefined,
        resolution.cssFamily,
      ]
        .filter(Boolean)
        .join(', ');
    },
    [diagnostics, embeddedFontAliases, options, resolutionCache],
  );

  return (
    <OfficeFontResolverContext.Provider value={resolveFontFamily}>
      <OfficeEmbeddedFontsReadyContext.Provider value={embeddedFontsReady}>
        {children}
      </OfficeEmbeddedFontsReadyContext.Provider>
    </OfficeFontResolverContext.Provider>
  );
}

/** 获取当前文档会话统一的字体链解析函数。 */
export function useOfficeFontResolver() {
  return useContext(OfficeFontResolverContext);
}

/** 判断当前文档内嵌字体是否已经完成注册，供分页测量避开回退字体。 */
export function useOfficeEmbeddedFontsReady() {
  return useContext(OfficeEmbeddedFontsReadyContext);
}
