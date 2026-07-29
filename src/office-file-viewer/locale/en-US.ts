import type { OfficeFileViewerMessages } from './types';

/** OfficeFileViewer 内置的英文界面文案。 */
export const enUS: OfficeFileViewerMessages = {
  file: {
    unloaded: 'No file loaded',
    unsupported:
      'This file type is not supported. Select a PPTX, PPT, XLSX, XLS, DOCX, DOC, or WPS file.',
    unrecognized:
      'The Office file type could not be identified. Provide a PPTX, PPT, XLSX, XLS, DOCX, DOC, or WPS file.',
    invalidUri:
      'uri must be a File, URL string, or an async function that returns a File, Blob, URL, or Response.',
    parseFailed: 'Failed to parse the file.',
    loadFailed: 'Failed to load the file.',
    downloadFailed: (status, statusText) =>
      `File download failed: ${status} ${statusText}`,
    fullscreenRejected: 'The browser rejected the fullscreen request.',
    fullscreenFailed: (reason) => `Fullscreen failed: ${reason}`,
  },
  toolbar: {
    selectFile: 'Select file',
    previousSlide: 'Previous slide',
    nextSlide: 'Next slide',
    showSpeakerNotes: 'Show speaker notes',
    hideSpeakerNotes: 'Hide speaker notes',
    speakerNotes: 'Notes',
    zoomOut: 'Zoom out',
    zoomIn: 'Zoom in',
    fullscreen: 'Fullscreen',
    exitFullscreen: 'Exit fullscreen',
  },
  empty: {
    pptx: 'Upload a PPTX file to start previewing',
    ppt: 'Upload a PPT file to start previewing',
    xlsx: 'Upload an XLSX file to start previewing',
    xls: 'Upload an XLS file to start previewing',
    docx: 'Upload a DOCX file to start previewing',
    doc: 'Upload a DOC or WPS file to start previewing',
  },
  loading: { parsing: 'Parsing file' },
  error: { previewFailed: 'Preview failed' },
  progress: {
    stages: {
      reading: 'Reading file',
      container: 'Opening file container',
      structure: 'Reading document structure',
      content: 'Parsing document content',
      resources: 'Processing document resources',
      assembling: 'Building preview',
    },
    partialTitle: 'Document parsing is incomplete',
    partialDescription: 'Only successfully parsed content is currently shown.',
  },
  outline: {
    region: 'Document outline',
    expand: 'Expand document outline',
    collapse: 'Collapse document outline',
    title: 'Outline',
    tree: 'Outline navigation',
  },
  spreadsheet: {
    dimensions: (rows, columns) => `${rows} rows × ${columns} columns`,
    imageLoadFailed: (alt) =>
      alt ? `${alt} (image failed to load)` : 'Image failed to load',
  },
  presentation: {
    slide: (index) => `Slide ${index}`,
    slideCount: (count) => `${count} slides`,
    notesRegion: 'Speaker notes',
    resizeNotes: 'Resize speaker notes',
    emptyNotes: 'No speaker notes for this slide',
  },
  document: { images: 'Document images' },
  chart: {
    invalidSize: 'Invalid chart size',
    staticAlt: 'Static chart',
    renderFailed: 'Chart rendering failed',
    mapLoadFailed: 'Map data failed to load',
  },
};
