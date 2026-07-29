import type { ParseStage } from '../services/parsing';
import type { PreviewKind } from '../services/preview';

/** OfficeFileViewer 内置支持的界面语言。 */
export type OfficeFileViewerLocale = 'zh-CN' | 'en-US';

/** 定义预览器内部全部用户可见文案，保证两种语言结构一致。 */
export type OfficeFileViewerMessages = {
  file: {
    unloaded: string;
    unsupported: string;
    unrecognized: string;
    invalidUri: string;
    parseFailed: string;
    loadFailed: string;
    downloadFailed: (status: number, statusText: string) => string;
    fullscreenRejected: string;
    fullscreenFailed: (reason: string) => string;
  };
  toolbar: {
    selectFile: string;
    previousSlide: string;
    nextSlide: string;
    showSpeakerNotes: string;
    hideSpeakerNotes: string;
    speakerNotes: string;
    zoomOut: string;
    zoomIn: string;
    fullscreen: string;
    exitFullscreen: string;
  };
  empty: Record<PreviewKind, string>;
  loading: { parsing: string };
  error: { previewFailed: string };
  progress: {
    stages: Record<ParseStage, string>;
    partialTitle: string;
    partialDescription: string;
  };
  outline: {
    region: string;
    expand: string;
    collapse: string;
    title: string;
    tree: string;
  };
  spreadsheet: {
    dimensions: (rows: number, columns: number) => string;
    imageLoadFailed: (alt?: string) => string;
  };
  presentation: {
    slide: (index: number) => string;
    slideCount: (count: number) => string;
    notesRegion: string;
    resizeNotes: string;
    emptyNotes: string;
  };
  document: { images: string };
  chart: {
    invalidSize: string;
    staticAlt: string;
    renderFailed: string;
    mapLoadFailed: string;
  };
};
