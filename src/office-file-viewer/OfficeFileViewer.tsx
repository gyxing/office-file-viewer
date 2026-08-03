// OfficeFileViewer 是组件库对外主入口，负责文件上传解析、格式状态和全局工具栏交互。
import { Layout } from 'antd';
import type { CSSProperties, ReactElement } from 'react';
import React, {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import './index.less';
import {
  OfficeFileViewerLocaleProvider,
  useOfficeFileViewerMessages,
  type OfficeFileViewerLocale,
  type OfficeFileViewerMessages,
} from './locale';
import type {
  DocWordPageSource,
  DocWordPreviewSummary,
} from './services/doc/DocWordPageSource';
import { disposeDocDocument, type DocDocument } from './services/doc/types';
import type {
  DocxWordPageSource,
  DocxWordPreviewSummary,
} from './services/docx/DocxWordPageSource';
import type { DocxDocument } from './services/docx/types';
import {
  type OfficeParseOptions,
  type ParseProgress,
} from './services/parsing';
import { createOfficeFileViewerParseSession } from './services/parsing/createParseSession';
import type {
  OfficeFileViewerParseSession,
  OfficeFileViewerPreviewHandle,
  OfficeFileViewerPreviewState,
} from './services/parsing/internalTypes';
import type { PptxDocument } from './services/pptx/types';
import { disposePresentationDocument } from './services/presentation/dispose';
import type { PresentationSource } from './services/presentation/types';
import {
  detectPreviewKind,
  disposeParsedOfficeFile,
  getPreviewFamily,
  isPresentationPreviewKind,
  isSpreadsheetPreviewKind,
  isSupportedOfficeFileName,
  type ParsedOfficeFile,
  type PreviewKind,
} from './services/preview';
import {
  createOfficeResourceStore,
  OfficeResourceStoreProvider,
} from './services/resource-store';
import {
  createOfficeDocumentSession,
  disposeDocumentSession,
  type OfficeDocumentSession,
} from './services/session';
import type { SpreadsheetSource } from './services/spreadsheet/SpreadsheetSource';
import {
  disposeSpreadsheetWorkbook,
  type SpreadsheetWorkbook,
} from './services/spreadsheet/types';
import { useExternalStoreSnapshot } from './shared/react/useExternalStoreSnapshot';
import { OfficeParseStatus } from './shell/ParseStatus';
import { OfficePreviewStage } from './shell/PreviewStage';
import { OfficeToolbar } from './shell/Toolbar';
import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_STEP,
} from './shell/constants';

const { Content } = Layout;

/** 保存当前加载代次收到的渐进解析结果，用于过滤过期任务输出。 */
type PendingPartialResult = {
  /** 本次加载的递增代次，用于丢弃上一次异步解析产生的过期结果。 */
  loadGeneration: number;
  /** 当前加载代次产生的部分或完整内部预览状态。 */
  preview: OfficeFileViewerPreviewState;
};

/** 定义预览文件来源，可直接传入 File、URL，或返回文件数据的异步加载函数。 */
export type OfficeFileViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);

/** Office文件预览器组件属性。 */
export type OfficeFileViewerProps = {
  /** 预览器使用的界面语言，默认使用简体中文。 */
  locale?: OfficeFileViewerLocale;
  /** 待预览文件的来源，可为本地文件、URL 或异步加载函数。 */
  uri?: OfficeFileViewerUri;
  /** 无法从文件来源推断名称时使用的默认文件名。 */
  defaultFileName?: string;

  /** 组件首次渲染时采用的缩放比例。 */
  defaultZoom?: number;
  /** 非受控模式下演讲者备注是否默认展开。 */
  defaultShowSpeakerNotes?: boolean;
  /** 受控模式下演讲者备注是否展开。 */
  showSpeakerNotes?: boolean;
  /** 演讲者备注展开状态变化时触发。 */
  onSpeakerNotesVisibilityChange?: (visible: boolean) => void;
  /** 附加到预览器根元素的自定义类名。 */
  className?: string;
  /** 预览区域高度，支持任意 CSS 高度值；未提供时使用父容器高度。 */
  height?: CSSProperties['height'];
  /** 传递给预览器根元素的内联样式。 */
  style?: CSSProperties;
  /** 完整文件解析成功后触发一次；渐进解析结果不会触发该回调。 */
  onFileParsed?: (parsed: ParsedOfficeFile, file: File) => void;
  /** 文件加载或解析失败时触发的回调。 */
  onError?: (error: Error, file?: File) => void;
  /** 传递给底层解析会话的运行配置。 */
  parseOptions?: OfficeParseOptions;
  /** 解析阶段或完成度变化时触发的进度回调。 */
  onParseProgress?: (progress: ParseProgress) => void;
};

/** Office MIME 类型到文件扩展名的映射。 */
const OFFICE_MIME_EXTENSION_MAP: Record<string, string> = {
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':
    '.pptx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
  'application/vnd.ms-excel': '.xls',
  'application/vnd.ms-powerpoint': '.ppt',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
    '.docx',
  'application/msword': '.doc',
  'application/wps-office.wps': '.wps',
};

/** 从 URL 路径末段提取并解码文件名；无法解析时返回 undefined。 */
function getFileNameFromUrl(url: string) {
  try {
    const parsedUrl = new URL(url, window.location.href);
    const lastSegment = parsedUrl.pathname.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  } catch {
    const path = url.split(/[?#]/)[0];
    const lastSegment = path.split('/').filter(Boolean).pop();
    return lastSegment ? decodeURIComponent(lastSegment) : undefined;
  }
}

/** 从 Content-Disposition 响应头提取 UTF-8 或普通文件名。 */
function getFileNameFromContentDisposition(contentDisposition: string | null) {
  if (!contentDisposition) return undefined;

  const encodedMatch = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch?.[1]) return decodeURIComponent(encodedMatch[1]);

  const plainMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
  return plainMatch?.[1] ? decodeURIComponent(plainMatch[1]) : undefined;
}

/** 根据响应 MIME 类型推断受支持的 Office 文件扩展名。 */
function getExtensionFromMimeType(mimeType: string) {
  return OFFICE_MIME_EXTENSION_MAP[
    mimeType.split(';')[0]?.trim().toLowerCase()
  ];
}

/** 判断文件名末尾是否包含扩展名。 */
function hasFileExtension(fileName: string) {
  return /\.[^./\\]+$/.test(fileName);
}

/** 标记由预览器输入校验产生、可安全直接展示给用户的错误。 */
class OfficeFileViewerInputError extends Error {}

/** 校验文件扩展名是否属于当前组件支持的 Office 格式。 */
function ensureSupportedOfficeFile(
  file: File,
  messages: OfficeFileViewerMessages,
) {
  if (!isSupportedOfficeFileName(file.name)) {
    throw new OfficeFileViewerInputError(messages.file.unsupported);
  }
}

/** 根据 Blob、显式文件名或 MIME 类型创建可识别格式的 File。 */
function createFileFromBlob(
  blob: Blob,
  messages: OfficeFileViewerMessages,
  fileName?: string,
) {
  if (
    fileName &&
    hasFileExtension(fileName) &&
    !isSupportedOfficeFileName(fileName)
  ) {
    throw new OfficeFileViewerInputError(messages.file.unrecognized);
  }

  const extension = getExtensionFromMimeType(blob.type);
  const inferredFileName =
    fileName && isSupportedOfficeFileName(fileName)
      ? fileName
      : extension
      ? `office-file${extension}`
      : undefined;

  if (!inferredFileName) {
    throw new OfficeFileViewerInputError(messages.file.unrecognized);
  }

  return new File([blob], inferredFileName, { type: blob.type });
}

/** 校验下载响应，并将响应体及响应头文件名转换为 File。 */
async function createFileFromResponse(
  response: Response,
  messages: OfficeFileViewerMessages,
  fallbackFileName?: string,
) {
  if (!response.ok) {
    throw new OfficeFileViewerInputError(
      messages.file.downloadFailed(response.status, response.statusText),
    );
  }

  const blob = await response.blob();
  const fileName =
    getFileNameFromContentDisposition(
      response.headers.get('Content-Disposition'),
    ) || fallbackFileName;
  return createFileFromBlob(blob, messages, fileName);
}

/** 下载远程 Office 文件，并沿用 URL 或响应头中可识别的文件名。 */
async function downloadOfficeFile(
  url: string,
  messages: OfficeFileViewerMessages,
  signal?: AbortSignal,
) {
  const urlFileName = getFileNameFromUrl(url);
  if (
    urlFileName &&
    hasFileExtension(urlFileName) &&
    !isSupportedOfficeFileName(urlFileName)
  ) {
    throw new OfficeFileViewerInputError(messages.file.unsupported);
  }

  const response = await fetch(url, { signal });
  return createFileFromResponse(response, messages, urlFileName);
}

/** 将组件支持的各种 URI 输入统一解析为可交给解析层的 File。 */
async function normalizeOfficeFileUri(
  uri: OfficeFileViewerUri,
  messages: OfficeFileViewerMessages,
  signal?: AbortSignal,
) {
  const resolvedUri = typeof uri === 'function' ? await uri() : uri;

  if (resolvedUri instanceof File) return resolvedUri;
  if (resolvedUri instanceof Response)
    return createFileFromResponse(resolvedUri, messages);
  if (resolvedUri instanceof Blob)
    return createFileFromBlob(resolvedUri, messages);
  if (typeof resolvedUri === 'string')
    return downloadOfficeFile(resolvedUri, messages, signal);

  throw new OfficeFileViewerInputError(messages.file.invalidUri);
}

/** 实现 Office 文件加载、解析会话、渐进结果和工具栏状态。 */
function OfficeFileViewerContent({
  uri,
  defaultFileName,
  defaultZoom = OFFICE_DEFAULT_ZOOM,
  defaultShowSpeakerNotes = false,
  showSpeakerNotes,
  onSpeakerNotesVisibilityChange,
  className,
  height,
  style,
  onFileParsed,
  onError,
  parseOptions,
  onParseProgress,
}: Omit<OfficeFileViewerProps, 'locale'>) {
  const messages = useOfficeFileViewerMessages();
  // OfficeFileViewer 是公共组件入口，集中管理“文件状态”和“格式私有状态”，避免使用者再组合多个子组件。
  const [fileName, setFileName] = useState<string>();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [parseProgress, setParseProgress] = useState<ParseProgress>();
  const [partialWarning, setPartialWarning] = useState(false);
  const [previewKind, setPreviewKind] = useState<PreviewKind>();
  const [pptxDocument, setPptxDocument] = useState<PptxDocument>();
  const presentationDocumentRef = useRef<PptxDocument>();
  const [presentationPreviewSource, setPresentationPreviewSource] =
    useState<PresentationSource>();
  const presentationPreviewSummary = useExternalStoreSnapshot(
    presentationPreviewSource,
  );
  const presentationPreviewSourceRef = useRef<PresentationSource>();
  const [spreadsheetWorkbook, setSpreadsheetWorkbook] =
    useState<SpreadsheetWorkbook>();
  const spreadsheetWorkbookRef = useRef<SpreadsheetWorkbook>();
  const [spreadsheetPreviewSource, setSpreadsheetPreviewSource] =
    useState<SpreadsheetSource>();
  const spreadsheetPreviewSummary = useExternalStoreSnapshot(
    spreadsheetPreviewSource,
  );
  const spreadsheetPreviewSourceRef = useRef<SpreadsheetSource>();
  const [docxDocument, setDocxDocument] = useState<DocxDocument>();
  const docxDocumentRef = useRef<DocxDocument>();
  const [docxPreviewSource, setDocxPreviewSource] =
    useState<DocxWordPageSource>();
  const [docxPreviewSummary, setDocxPreviewSummary] =
    useState<DocxWordPreviewSummary>();
  const docxPreviewSourceRef = useRef<DocxWordPageSource>();
  const [docDocument, setDocDocument] = useState<DocDocument>();
  const docDocumentRef = useRef<DocDocument>();
  const [docPreviewSource, setDocPreviewSource] = useState<DocWordPageSource>();
  const [docPreviewSummary, setDocPreviewSummary] =
    useState<DocWordPreviewSummary>();
  const docPreviewSourceRef = useRef<DocWordPageSource>();
  const [activeIndex, setActiveIndex] = useState(0);
  const [activeSheetId, setActiveSheetId] = useState<string>();
  const [zoom, setZoom] = useState(defaultZoom);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [internalShowSpeakerNotes, setInternalShowSpeakerNotes] = useState(
    defaultShowSpeakerNotes,
  );
  const [showWordOutline, setShowWordOutline] = useState(false);
  const previewFamily = previewKind ? getPreviewFamily(previewKind) : undefined;
  // 工具栏只订阅当前 Word 数据源的大纲计数，正文增量不会重复维护布尔状态。
  const isWordPreview = previewFamily === 'word';
  const wordOutlineProvider = isWordPreview
    ? previewKind === 'docx'
      ? docxPreviewSource?.outline
      : docPreviewSource?.outline
    : undefined;
  const wordOutlineSnapshot = useExternalStoreSnapshot(wordOutlineProvider);
  const materializedWordOutlineCount = isWordPreview
    ? previewKind === 'docx'
      ? docxDocument?.outline?.length ?? 0
      : docDocument?.outline?.length ?? 0
    : 0;
  const hasWordOutline =
    (wordOutlineSnapshot?.count ?? materializedWordOutlineCount) > 0;
  const viewerRef = useRef<HTMLDivElement | null>(null);
  const resourceStore = useMemo(() => createOfficeResourceStore(), []);
  const loadGenerationRef = useRef(0);
  const documentSessionIdRef = useRef<string>();
  const requestControllerRef = useRef<AbortController>();
  const parseSessionRef = useRef<OfficeFileViewerParseSession>();
  const pendingPartialRef = useRef<PendingPartialResult>();
  const partialFrameRef = useRef<number>();
  const defaultZoomRef = useRef(defaultZoom);
  const defaultShowSpeakerNotesRef = useRef(defaultShowSpeakerNotes);
  const showSpeakerNotesRef = useRef(showSpeakerNotes);
  const onFileParsedRef = useRef(onFileParsed);
  const onErrorRef = useRef(onError);
  const parseOptionsRef = useRef(parseOptions);
  const onParseProgressRef = useRef(onParseProgress);
  const messagesRef = useRef(messages);

  defaultZoomRef.current = defaultZoom;
  defaultShowSpeakerNotesRef.current = defaultShowSpeakerNotes;
  showSpeakerNotesRef.current = showSpeakerNotes;
  onFileParsedRef.current = onFileParsed;
  onErrorRef.current = onError;
  parseOptionsRef.current = parseOptions;
  onParseProgressRef.current = onParseProgress;
  messagesRef.current = messages;

  const cancelPartialFrame = useCallback(() => {
    if (
      partialFrameRef.current !== undefined &&
      typeof window !== 'undefined'
    ) {
      window.cancelAnimationFrame(partialFrameRef.current);
    }
    partialFrameRef.current = undefined;
    pendingPartialRef.current = undefined;
  }, []);

  const clearPreviewDocuments = useCallback(() => {
    // 先移除 React 模型，再释放 refs 中唯一拥有资源的完整或失败冻结模型。
    setPptxDocument(undefined);
    setPresentationPreviewSource(undefined);
    setSpreadsheetWorkbook(undefined);
    setSpreadsheetPreviewSource(undefined);
    setDocxDocument(undefined);
    setDocxPreviewSource(undefined);
    setDocxPreviewSummary(undefined);
    setDocDocument(undefined);
    setDocPreviewSource(undefined);
    setDocPreviewSummary(undefined);
    setPreviewKind(undefined);
    setActiveIndex(0);
    setActiveSheetId(undefined);
    setShowWordOutline(false);

    const spreadsheet = spreadsheetWorkbookRef.current;
    spreadsheetWorkbookRef.current = undefined;
    disposeSpreadsheetWorkbook(spreadsheet);
    void disposeDocumentSession(spreadsheet);
    const spreadsheetSource = spreadsheetPreviewSourceRef.current;
    spreadsheetPreviewSourceRef.current = undefined;
    void disposeDocumentSession(spreadsheetSource);
    const presentation = presentationDocumentRef.current;
    presentationDocumentRef.current = undefined;
    disposePresentationDocument(presentation);
    void disposeDocumentSession(presentation);
    const presentationSource = presentationPreviewSourceRef.current;
    presentationPreviewSourceRef.current = undefined;
    void disposeDocumentSession(presentationSource);
    const docx = docxDocumentRef.current;
    docxDocumentRef.current = undefined;
    void disposeDocumentSession(docx);
    const docxSource = docxPreviewSourceRef.current;
    docxPreviewSourceRef.current = undefined;
    void disposeDocumentSession(docxSource);
    const documentModel = docDocumentRef.current;
    docDocumentRef.current = undefined;
    disposeDocDocument(documentModel);
    void disposeDocumentSession(documentModel);
    const documentSource = docPreviewSourceRef.current;
    docPreviewSourceRef.current = undefined;
    void disposeDocumentSession(documentSource);
  }, []);

  const installPartialSnapshot = useCallback(
    (preview: OfficeFileViewerPreviewState) => {
      if (preview.mode === 'source') {
        if (preview.previewKind === 'doc') {
          docPreviewSourceRef.current = preview.source;
          setDocPreviewSource(preview.source);
          setDocPreviewSummary(preview.summary);
        } else if (preview.previewKind === 'docx') {
          docxPreviewSourceRef.current = preview.source;
          setDocxPreviewSource(preview.source);
          setDocxPreviewSummary(preview.summary);
        } else if (isSpreadsheetPreviewKind(preview.previewKind)) {
          spreadsheetPreviewSourceRef.current = preview.source;
          setSpreadsheetPreviewSource(preview.source);
          setActiveSheetId((current) =>
            current &&
            preview.summary.sheets.some((sheet) => sheet.id === current)
              ? current
              : preview.summary.sheets[0]?.id,
          );
        } else {
          presentationPreviewSourceRef.current = preview.source;
          setPresentationPreviewSource(preview.source);
          setActiveIndex((current) =>
            Math.min(current, Math.max(0, preview.summary.slideCount - 1)),
          );
        }
        return;
      }
      const parsed = preview.model;
      if (parsed.kind === 'xls') {
        setSpreadsheetWorkbook(parsed.workbook);
        setActiveSheetId((current) =>
          current &&
          parsed.workbook.sheets.some((sheet) => sheet.id === current)
            ? current
            : parsed.workbook.sheets[0]?.id,
        );
        return;
      }
      if (parsed.kind === 'ppt') {
        setPptxDocument(parsed.document);
        setActiveIndex((current) =>
          Math.min(current, Math.max(0, parsed.document.slides.length - 1)),
        );
        return;
      }
      if (parsed.kind === 'doc') {
        setDocDocument(parsed.document);
      }
    },
    [],
  );

  const schedulePartialSnapshot = useCallback(
    (preview: OfficeFileViewerPreviewState, loadGeneration: number) => {
      pendingPartialRef.current = { preview, loadGeneration };
      if (partialFrameRef.current !== undefined) return;
      if (typeof window === 'undefined') {
        if (
          loadGeneration === loadGenerationRef.current &&
          preview.sessionId === documentSessionIdRef.current
        ) {
          installPartialSnapshot(preview);
        }
        pendingPartialRef.current = undefined;
        return;
      }
      partialFrameRef.current = window.requestAnimationFrame(() => {
        partialFrameRef.current = undefined;
        const pending = pendingPartialRef.current;
        pendingPartialRef.current = undefined;
        if (
          pending &&
          pending.loadGeneration === loadGenerationRef.current &&
          pending.preview.sessionId === documentSessionIdRef.current
        ) {
          installPartialSnapshot(pending.preview);
        }
      });
    },
    [installPartialSnapshot],
  );

  const loadFile = useCallback(
    async (file: File, loadGeneration: number) => {
      parseSessionRef.current?.cancel();
      parseSessionRef.current?.dispose();
      parseSessionRef.current = undefined;
      documentSessionIdRef.current = undefined;
      cancelPartialFrame();
      clearPreviewDocuments();
      setLoading(true);
      setError(undefined);
      setParseProgress(undefined);
      setPartialWarning(false);
      let retainedPartial = false;
      let documentSession: OfficeDocumentSession | undefined;

      try {
        ensureSupportedOfficeFile(file, messagesRef.current);
        if (loadGeneration !== loadGenerationRef.current) return;

        // 上传新文件时同步重置所有格式相关状态，防止上一份文档的页码/缩放/工作表残留到新文档。
        const fileKind = detectPreviewKind(file.name);
        setPreviewKind(fileKind);
        setFileName(file.name);
        setActiveIndex(0);
        setZoom(defaultZoomRef.current);
        if (showSpeakerNotesRef.current === undefined) {
          setInternalShowSpeakerNotes(defaultShowSpeakerNotesRef.current);
        }

        documentSession = createOfficeDocumentSession();
        documentSessionIdRef.current = documentSession.id;
        const isCurrentSession = () =>
          loadGeneration === loadGenerationRef.current &&
          documentSession?.id === documentSessionIdRef.current;
        const parseSession = createOfficeFileViewerParseSession(
          file,
          parseOptionsRef.current,
          documentSession,
        );
        parseSessionRef.current = parseSession;
        const unsubscribeProgress = parseSession.subscribe((progress) => {
          if (!isCurrentSession()) return;
          setParseProgress(progress);
          onParseProgressRef.current?.(progress);
        });
        const unsubscribePartial = parseSession.subscribePartial((partial) => {
          if (
            !isCurrentSession() ||
            partial.sessionId !== documentSession?.id
          ) {
            return;
          }
          schedulePartialSnapshot(partial, loadGeneration);
        });
        let preview: OfficeFileViewerPreviewHandle;
        try {
          preview = await parseSession.result;
        } catch (nextError) {
          cancelPartialFrame();
          const partial = parseSession.partialResult;
          if (partial) {
            if (
              !isCurrentSession() ||
              partial.sessionId !== documentSession.id
            ) {
              if (partial.mode === 'materialized') {
                await disposeParsedOfficeFile(partial.model);
              } else {
                await documentSession.dispose();
              }
            } else if (
              partial.mode === 'materialized' &&
              partial.model.kind === 'xls'
            ) {
              spreadsheetWorkbookRef.current = partial.model.workbook;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'materialized' &&
              partial.model.kind === 'ppt'
            ) {
              presentationDocumentRef.current = partial.model.document;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'materialized' &&
              partial.model.kind === 'doc'
            ) {
              docDocumentRef.current = partial.model.document;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'source' &&
              partial.previewKind === 'doc'
            ) {
              docPreviewSourceRef.current = partial.source;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'source' &&
              partial.previewKind === 'docx'
            ) {
              docxPreviewSourceRef.current = partial.source;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'source' &&
              isPresentationPreviewKind(partial.previewKind)
            ) {
              presentationPreviewSourceRef.current = partial.source;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else if (
              partial.mode === 'source' &&
              isSpreadsheetPreviewKind(partial.previewKind)
            ) {
              spreadsheetPreviewSourceRef.current = partial.source;
              installPartialSnapshot(partial);
              setPartialWarning(true);
              setError(undefined);
              retainedPartial = true;
            } else {
              await documentSession.dispose();
            }
          }
          throw nextError;
        } finally {
          unsubscribeProgress();
          unsubscribePartial();
          if (parseSessionRef.current === parseSession) {
            parseSession.dispose();
            parseSessionRef.current = undefined;
          }
        }
        if (!isCurrentSession() || preview.sessionId !== documentSession.id) {
          if (preview.mode === 'materialized') {
            await disposeParsedOfficeFile(preview.model);
          } else {
            await preview.dispose();
          }
          return;
        }
        if (preview.mode === 'source') {
          cancelPartialFrame();
          setParseProgress(undefined);
          setPartialWarning(false);
          if (preview.previewKind === 'doc') {
            docPreviewSourceRef.current = preview.source;
            setDocPreviewSource(preview.source);
            setDocPreviewSummary(preview.summary);
            setDocDocument(undefined);
            docDocumentRef.current = undefined;
          } else if (preview.previewKind === 'docx') {
            docxPreviewSourceRef.current = preview.source;
            setDocxPreviewSource(preview.source);
            setDocxPreviewSummary(preview.summary);
            setDocxDocument(undefined);
            docxDocumentRef.current = undefined;
          } else if (isSpreadsheetPreviewKind(preview.previewKind)) {
            spreadsheetPreviewSourceRef.current = preview.source;
            setSpreadsheetPreviewSource(preview.source);
            setSpreadsheetWorkbook(undefined);
            spreadsheetWorkbookRef.current = undefined;
            setActiveSheetId((current) =>
              current &&
              preview.summary.sheets.some((sheet) => sheet.id === current)
                ? current
                : preview.summary.sheets[0]?.id,
            );
          } else {
            presentationPreviewSourceRef.current = preview.source;
            setPresentationPreviewSource(preview.source);
            setPptxDocument(undefined);
            presentationDocumentRef.current = undefined;
          }
          return;
        }
        const parsed = preview.model;

        cancelPartialFrame();
        setParseProgress(undefined);
        setPartialWarning(false);
        const nextPresentationDocument = isPresentationPreviewKind(parsed.kind)
          ? parsed.document
          : undefined;
        disposePresentationDocument(presentationDocumentRef.current);
        void disposeDocumentSession(presentationDocumentRef.current);
        presentationDocumentRef.current = nextPresentationDocument;
        setPptxDocument(nextPresentationDocument);
        presentationPreviewSourceRef.current = undefined;
        setPresentationPreviewSource(undefined);
        const nextSpreadsheetWorkbook = isSpreadsheetPreviewKind(parsed.kind)
          ? parsed.workbook
          : undefined;
        disposeSpreadsheetWorkbook(spreadsheetWorkbookRef.current);
        void disposeDocumentSession(spreadsheetWorkbookRef.current);
        spreadsheetWorkbookRef.current = nextSpreadsheetWorkbook;
        setSpreadsheetWorkbook(nextSpreadsheetWorkbook);
        spreadsheetPreviewSourceRef.current = undefined;
        setSpreadsheetPreviewSource(undefined);
        const nextDocxDocument =
          parsed.kind === 'docx' ? parsed.document : undefined;
        void disposeDocumentSession(docxDocumentRef.current);
        docxDocumentRef.current = nextDocxDocument;
        setDocxDocument(nextDocxDocument);
        const nextDocDocument =
          parsed.kind === 'doc' ? parsed.document : undefined;
        disposeDocDocument(docDocumentRef.current);
        void disposeDocumentSession(docDocumentRef.current);
        docDocumentRef.current = nextDocDocument;
        setDocDocument(nextDocDocument);
        setActiveSheetId((current) => {
          if (!isSpreadsheetPreviewKind(parsed.kind)) return undefined;
          return current &&
            parsed.workbook.sheets.some((sheet) => sheet.id === current)
            ? current
            : parsed.workbook.sheets[0]?.id;
        });
        onFileParsedRef.current?.(parsed, file);
      } catch (nextError) {
        if (
          loadGeneration !== loadGenerationRef.current ||
          (documentSession &&
            documentSession.id !== documentSessionIdRef.current)
        ) {
          return;
        }

        // 界面只展示可本地化的概述，原始解析错误仍通过回调交给调用方诊断。
        const normalizedError =
          nextError instanceof Error
            ? nextError
            : new Error(messagesRef.current.file.parseFailed);
        setParseProgress(undefined);
        if (!retainedPartial) {
          if (documentSessionIdRef.current === documentSession?.id) {
            documentSessionIdRef.current = undefined;
          }
          setError(
            normalizedError instanceof OfficeFileViewerInputError
              ? normalizedError.message
              : messagesRef.current.file.parseFailed,
          );
        }
        onErrorRef.current?.(normalizedError, file);
      } finally {
        if (loadGeneration === loadGenerationRef.current) {
          setLoading(false);
        }
      }
    },
    [
      cancelPartialFrame,
      clearPreviewDocuments,
      installPartialSnapshot,
      schedulePartialSnapshot,
    ],
  );

  const handleSelectFile = useCallback(
    async (file: File) => {
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      const loadGeneration = ++loadGenerationRef.current;
      await loadFile(file, loadGeneration);
    },
    [loadFile],
  );

  useEffect(() => {
    if (!uri) return;

    parseSessionRef.current?.cancel();
    parseSessionRef.current?.dispose();
    parseSessionRef.current = undefined;
    documentSessionIdRef.current = undefined;
    cancelPartialFrame();
    clearPreviewDocuments();
    setParseProgress(undefined);
    setPartialWarning(false);
    // 固化本次 effect 的文件来源，避免异步闭包丢失类型收窄。
    const uriToLoad = uri;
    const loadGeneration = ++loadGenerationRef.current;
    const requestController =
      typeof AbortController === 'undefined'
        ? undefined
        : new AbortController();
    requestControllerRef.current?.abort();
    requestControllerRef.current = requestController;

    async function loadUri() {
      setLoading(true);
      setError(undefined);

      let file: File | undefined;

      try {
        file = await normalizeOfficeFileUri(
          uriToLoad,
          messagesRef.current,
          requestController?.signal,
        );
        if (loadGeneration !== loadGenerationRef.current) return;
        await loadFile(file, loadGeneration);
      } catch (nextError) {
        if (
          loadGeneration !== loadGenerationRef.current ||
          requestController?.signal.aborted
        )
          return;

        const normalizedError =
          nextError instanceof Error
            ? nextError
            : new Error(messagesRef.current.file.loadFailed);
        setError(
          normalizedError instanceof OfficeFileViewerInputError
            ? normalizedError.message
            : messagesRef.current.file.loadFailed,
        );
        onErrorRef.current?.(normalizedError, file);
        setLoading(false);
      } finally {
        if (requestControllerRef.current === requestController) {
          requestControllerRef.current = undefined;
        }
      }
    }

    void loadUri();

    return () => {
      requestController?.abort();
    };
  }, [cancelPartialFrame, clearPreviewDocuments, loadFile, uri]);

  useEffect(
    () => () => {
      // 组件卸载时让所有不可取消的本地解析任务失效，避免异步结果继续写入状态。
      loadGenerationRef.current += 1;
      requestControllerRef.current?.abort();
      parseSessionRef.current?.cancel();
      parseSessionRef.current?.dispose();
      parseSessionRef.current = undefined;
      documentSessionIdRef.current = undefined;
      cancelPartialFrame();
      disposeSpreadsheetWorkbook(spreadsheetWorkbookRef.current);
      void disposeDocumentSession(spreadsheetWorkbookRef.current);
      spreadsheetWorkbookRef.current = undefined;
      void disposeDocumentSession(spreadsheetPreviewSourceRef.current);
      spreadsheetPreviewSourceRef.current = undefined;
      disposePresentationDocument(presentationDocumentRef.current);
      void disposeDocumentSession(presentationDocumentRef.current);
      presentationDocumentRef.current = undefined;
      void disposeDocumentSession(presentationPreviewSourceRef.current);
      presentationPreviewSourceRef.current = undefined;
      void disposeDocumentSession(docxDocumentRef.current);
      docxDocumentRef.current = undefined;
      void disposeDocumentSession(docxPreviewSourceRef.current);
      docxPreviewSourceRef.current = undefined;
      disposeDocDocument(docDocumentRef.current);
      void disposeDocumentSession(docDocumentRef.current);
      docDocumentRef.current = undefined;
      void disposeDocumentSession(docPreviewSourceRef.current);
      docPreviewSourceRef.current = undefined;
    },
    [cancelPartialFrame],
  );

  useEffect(() => () => void resourceStore.dispose(), [resourceStore]);

  useEffect(() => {
    // 切换文件或渐进解析确认无大纲时，收起已打开的空侧栏。
    if (!hasWordOutline) setShowWordOutline(false);
  }, [hasWordOutline]);

  const presentationSlideCount =
    presentationPreviewSummary?.slideCount ?? pptxDocument?.slides.length ?? 0;
  const hasDocument = useMemo(() => {
    // 工具栏的格式操作只依赖当前已识别文件的可渲染内容。
    if (!previewKind || !previewFamily) return false;
    if (previewFamily === 'presentation') {
      return presentationSlideCount > 0;
    }
    if (previewFamily === 'spreadsheet') {
      return Boolean(
        spreadsheetWorkbook?.sheets.length ||
          spreadsheetPreviewSummary?.sheets.length,
      );
    }
    if (previewKind === 'docx') {
      return Boolean(docxDocument?.blocks.length || docxPreviewSummary);
    }
    return Boolean(
      docDocument?.paragraphs.length ||
        docPreviewSource?.getSnapshot().pages.length,
    );
  }, [
    docDocument,
    docPreviewSource,
    docxDocument,
    docxPreviewSummary,
    pptxDocument,
    presentationSlideCount,
    previewFamily,
    previewKind,
    spreadsheetWorkbook,
    spreadsheetPreviewSummary,
  ]);

  const canGoPreviousSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    activeIndex > 0;
  const canGoNextSlide =
    previewFamily === 'presentation' &&
    presentationSlideCount > 1 &&
    activeIndex < presentationSlideCount - 1;

  const handlePreviousSlide = useCallback(() => {
    setActiveIndex((value) => Math.max(value - 1, 0));
  }, []);

  const handleNextSlide = useCallback(() => {
    setActiveIndex((value) =>
      Math.min(value + 1, Math.max(0, presentationSlideCount - 1)),
    );
  }, [presentationSlideCount]);

  const handleZoomOut = useCallback(() => {
    setZoom((value) => Math.max(OFFICE_MIN_ZOOM, value - OFFICE_ZOOM_STEP));
  }, []);

  const handleZoomIn = useCallback(() => {
    setZoom((value) => Math.min(OFFICE_MAX_ZOOM, value + OFFICE_ZOOM_STEP));
  }, []);

  const handleResetZoom = useCallback(() => {
    setZoom(defaultZoom);
  }, [defaultZoom]);

  const speakerNotesVisible = showSpeakerNotes ?? internalShowSpeakerNotes;
  const handleToggleSpeakerNotes = useCallback(() => {
    const nextVisible = !(showSpeakerNotes ?? internalShowSpeakerNotes);
    if (showSpeakerNotes === undefined) {
      setInternalShowSpeakerNotes(nextVisible);
    }
    onSpeakerNotesVisibilityChange?.(nextVisible);
  }, [
    internalShowSpeakerNotes,
    onSpeakerNotesVisibilityChange,
    showSpeakerNotes,
  ]);
  const handleToggleWordOutline = useCallback(() => {
    if (!hasWordOutline) return;
    setShowWordOutline((visible) => !visible);
  }, [hasWordOutline]);
  const handleCloseWordOutline = useCallback(() => {
    setShowWordOutline(false);
  }, []);

  const fullscreenSupported =
    typeof document !== 'undefined' &&
    typeof document.documentElement.requestFullscreen === 'function';

  useEffect(() => {
    if (typeof document === 'undefined') return;

    // 浏览器和 ESC 键都可能改变全屏状态，因此以 fullscreenchange 作为唯一状态来源。
    const handleFullscreenChange = () => {
      setIsFullscreen(document.fullscreenElement === viewerRef.current);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
    };
  }, []);

  const handleFullscreen = useCallback(async () => {
    const viewer = viewerRef.current;
    if (
      !viewer ||
      typeof document === 'undefined' ||
      typeof viewer.requestFullscreen !== 'function'
    )
      return;

    try {
      if (document.fullscreenElement === viewer) {
        await document.exitFullscreen();
      } else {
        await viewer.requestFullscreen();
      }
    } catch (nextError) {
      const reason =
        nextError instanceof Error
          ? nextError.message
          : messagesRef.current.file.fullscreenRejected;
      onErrorRef.current?.(
        new Error(messagesRef.current.file.fullscreenFailed(reason)),
      );
    }
  }, []);

  // 专用 height 配置优先于 style.height，避免两个入口同时传值时结果不确定。
  const viewerStyle = height === undefined ? style : { ...style, height };
  const displayedFileName =
    fileName ?? defaultFileName ?? messages.file.unloaded;
  const loadingTip = parseProgress
    ? messages.progress.stages[parseProgress.stage]
    : undefined;

  return (
    <div
      ref={viewerRef}
      className={['office-file-viewer', className].filter(Boolean).join(' ')}
      style={viewerStyle}
    >
      <Layout className="office-file-viewer__layout">
        <OfficeToolbar
          fileName={displayedFileName}
          previewKind={previewKind}
          zoom={zoom}
          hasDocument={hasDocument}
          canGoPreviousSlide={canGoPreviousSlide}
          canGoNextSlide={canGoNextSlide}
          showSpeakerNotes={speakerNotesVisible}
          onToggleSpeakerNotes={handleToggleSpeakerNotes}
          hasWordOutline={hasWordOutline}
          showWordOutline={showWordOutline}
          onToggleWordOutline={handleToggleWordOutline}
          onSelectFile={handleSelectFile}
          onPreviousSlide={handlePreviousSlide}
          onNextSlide={handleNextSlide}
          onZoomOut={handleZoomOut}
          onZoomIn={handleZoomIn}
          onZoomChange={setZoom}
          onResetZoom={handleResetZoom}
          isFullscreen={isFullscreen}
          fullscreenSupported={fullscreenSupported}
          onFullscreen={handleFullscreen}
        />
        <Content className="office-file-viewer__content">
          <OfficeResourceStoreProvider store={resourceStore}>
            <OfficePreviewStage
              loading={loading}
              loadingTip={loadingTip}
              hasRenderableContent={hasDocument}
              error={error}
              previewKind={previewKind}
              documentSessionId={documentSessionIdRef.current}
              pptxDocument={pptxDocument}
              presentationPreviewSource={presentationPreviewSource}
              spreadsheetWorkbook={spreadsheetWorkbook}
              spreadsheetPreviewSource={spreadsheetPreviewSource}
              docxDocument={docxDocument}
              docxPreviewSource={docxPreviewSource}
              docxPreviewSummary={docxPreviewSummary}
              docDocument={docDocument}
              docPreviewSource={docPreviewSource}
              docPreviewSummary={docPreviewSummary}
              activeIndex={activeIndex}
              activeSheetId={activeSheetId}
              zoom={zoom}
              showSpeakerNotes={speakerNotesVisible}
              showWordOutline={showWordOutline && hasWordOutline}
              onCloseWordOutline={handleCloseWordOutline}
              onSelectSlide={setActiveIndex}
              onSelectSheet={setActiveSheetId}
            />
          </OfficeResourceStoreProvider>
          <OfficeParseStatus
            progress={loading && hasDocument ? parseProgress : undefined}
            warning={partialWarning}
          />
        </Content>
      </Layout>
    </div>
  );
}

/** 渲染 Office 文件预览器，并为当前实例提供独立的界面语言。 */
function OfficeFileViewerComponent({
  locale = 'zh-CN',
  ...props
}: OfficeFileViewerProps): ReactElement {
  return (
    <OfficeFileViewerLocaleProvider locale={locale}>
      <OfficeFileViewerContent {...props} />
    </OfficeFileViewerLocaleProvider>
  );
}

export const OfficeFileViewer = memo(OfficeFileViewerComponent);
