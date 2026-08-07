---
title: Documentation
siteLayout: docs
toc: content
---

# Office File Viewer Documentation

`office-file-viewer` is a browser-only React component for previewing DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX files. Downloading, parsing, and rendering happen in the browser, without a companion document-conversion service.

<nav className="office-viewer-docs__quick-links" aria-label="Documentation shortcuts">
  <a href="#quick-start">Quick start</a>
  <a href="#component-api">Component API</a>
  <a href="#advanced-api">Advanced API</a>
  <a href="#limitations">Limitations</a>
</nav>

## Installation

```bash
npm install office-file-viewer
```

Or with Yarn:

```bash
yarn add office-file-viewer
```

`react` and `react-dom` are peer dependencies supplied by the host. The package is ESM-only, so the host build tool must support ESM. Component CSS is included by the package build and does not require a Less loader.

Import public APIs only from `office-file-viewer`. Undocumented `dist` or source-file imports are not part of the compatibility contract.

## Version compatibility

| Item          | Requirement | Notes                                 |
| ------------- | ----------- | ------------------------------------- |
| React         | `>=16.9.0`  | Hooks-capable React versions          |
| ReactDOM      | `>=16.9.0`  | Keep the same major version as React  |
| Module format | ESM-only    | Use an ESM-capable browser build tool |

Current peer dependency ranges:

```text
react: >=16.9.0
react-dom: >=16.9.0
```

<a id="quick-start"></a>

## Quick start

### Default file picker

When `uri` is omitted, the viewer displays its own file picker:

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer height="80vh" />;
}
```

When `height` is omitted, the viewer follows its parent height. Ensure the parent has a computable height, or pass a number, `px`, `%`, or `vh` value.

### Local files and remote URLs

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview({ file }: { file: File }) {
  return (
    <div style={{ height: 720 }}>
      <OfficeFileViewer uri={file} />
      <OfficeFileViewer
        uri="https://example.com/files/demo.pptx"
        height="80vh"
      />
    </div>
  );
}
```

### Async sources and callbacks

An async source can return a `File`, `Blob`, URL string, or `Response`:

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return (
    <OfficeFileViewer
      uri={async () => fetch('/files/demo.xlsx')}
      parseOptions={{ worker: 'auto' }}
      onParseProgress={(progress) => {
        const percent =
          progress.percent === undefined
            ? undefined
            : Math.round(progress.percent * 100);
        console.info(progress.stage, percent, progress.message);
      }}
      onPreviewReady={(info, file) => {
        console.info('Preview ready', info.previewKind, info.mode, file.name);
      }}
      onFileParsed={(parsed, file) => {
        console.info('Parsing completed', parsed.kind, file.name);
      }}
      onError={(error, file) => {
        console.error('Preview failed', file?.name, error);
      }}
    />
  );
}
```

### English UI

The viewer UI defaults to Simplified Chinese. Set `locale` to use the built-in English messages:

```tsx | pure
import { OfficeFileViewer } from 'office-file-viewer';

export default function EnglishOfficePreview() {
  return <OfficeFileViewer locale="en-US" height="80vh" />;
}
```

## `uri` file sources

```ts | pure
type OfficeFileViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);
```

Remote source rules:

- Cross-origin URLs must allow browser access through CORS.
- URLs should preferably end with a supported Office extension.
- An extensionless URL must expose a filename through `Content-Disposition` or a supported MIME type through `Content-Type`.
- An explicitly unsupported URL extension is rejected before download, even if its response body contains an Office file.
- Changing `uri` cancels the previous URL download through `AbortController`.
- A custom async loader cannot be forcibly cancelled, but a stale result cannot replace the latest source.
- Manually selecting a file also invalidates an in-flight remote or parsing result.

<a id="component-api"></a>

## `OfficeFileViewer` API

| Prop                             | Type                                                 | Default           | Description                                                      |
| -------------------------------- | ---------------------------------------------------- | ----------------- | ---------------------------------------------------------------- |
| `locale`                         | `'zh-CN' \| 'en-US'`                                 | `'zh-CN'`         | Built-in viewer UI language                                      |
| `uri`                            | `OfficeFileViewerUri`                                | -                 | File source to preload; the file picker is shown when omitted    |
| `defaultFileName`                | `string`                                             | Localized message | Fallback name when the source does not provide a usable filename |
| `defaultZoom`                    | `number`                                             | `100`             | Initial zoom percentage, clamped from `25` to `300`              |
| `defaultShowSpeakerNotes`        | `boolean`                                            | `false`           | Initial speaker-notes state in uncontrolled mode                 |
| `showSpeakerNotes`               | `boolean`                                            | -                 | Controlled speaker-notes visibility                              |
| `onSpeakerNotesVisibilityChange` | `(visible: boolean) => void`                         | -                 | Called when the presentation notes visibility changes            |
| `className`                      | `string`                                             | -                 | Additional class name for the viewer root                        |
| `height`                         | `CSSProperties['height']`                            | Parent height     | Viewer height; takes precedence over `style.height`              |
| `style`                          | `CSSProperties`                                      | -                 | Inline styles for the viewer root                                |
| `onFileParsed`                   | `(parsed: ParsedOfficeFile, file: File) => void`     | -                 | Called once when the complete materialized result is available   |
| `onPreviewReady`                 | `(info: OfficePreviewReadyInfo, file: File) => void` | -                 | Called once when the first usable preview is ready               |
| `onError`                        | `(error: Error, file?: File) => void`                | -                 | Called when loading, parsing, or a viewer operation fails        |
| `parseOptions`                   | `OfficeParseOptions`                                 | `{}`              | Worker strategy and optional Worker factory                      |
| `onParseProgress`                | `(progress: ParseProgress) => void`                  | -                 | Called when the current parse stage or progress changes          |

### Controlled speaker notes

Use `showSpeakerNotes` with `onSpeakerNotesVisibilityChange` for controlled state. Use `defaultShowSpeakerNotes` only for the initial state in uncontrolled mode. The control is relevant to PPT/PPTX files when speaker notes are available.

### Callback timing

- `onPreviewReady` fires when either a complete model (`mode: 'materialized'`) or an on-demand source (`mode: 'source'`) can render its first usable preview.
- `onFileParsed` fires after the complete materialized result is ready; progressive intermediate results do not trigger it.
- `onParseProgress` can fire many times and may omit exact counts or `percent`.
- `onError` can receive no `file` when an error occurs before a usable file has been resolved.

## Parsing configuration and progress

```ts | pure
type WorkerMode = 'auto' | 'always' | 'never';

type OfficeParseOptions = {
  worker?: WorkerMode;
  workerFactory?: () => Worker;
};
```

| Mode       | DOC/WPS, XLS, PPT                                              | DOCX, XLSX, PPTX                                                   |
| ---------- | -------------------------------------------------------------- | ------------------------------------------------------------------ |
| `'auto'`   | Prefers a Worker and falls back to the main thread if needed   | Currently parses on the main thread                                |
| `'always'` | Requires a compatible Worker and fails if it cannot be created | Worker migration is incomplete and a configuration error is thrown |
| `'never'`  | Always parses on the main thread                               | Always parses on the main thread                                   |

Each active legacy-format session owns its Worker. The viewer cancels and disposes the current session when the source changes or the component unmounts. `workerFactory` is intended for hosts with special asset paths or Content Security Policy requirements; most applications should use the built-in factory.

Progress uses normalized values from `0` to `1` when `percent` is available:

```ts | pure
type ParseStage =
  | 'reading'
  | 'container'
  | 'structure'
  | 'content'
  | 'resources'
  | 'assembling';

type ParseProgress = {
  stage: ParseStage;
  completed?: number;
  total?: number;
  percent?: number;
  message: string;
};

type OfficePreviewReadyInfo = {
  previewKind: PreviewKind;
  mode: 'materialized' | 'source';
};
```

`completed`, `total`, and `percent` are optional. Consumers should support an indeterminate state instead of assuming every parser can report exact progress.

<a id="advanced-api"></a>

## Advanced parsing API

Use a low-level session only when the host must manage parsing outside `OfficeFileViewer`:

```ts | pure
type OfficeParseSessionStatus =
  | 'starting'
  | 'running'
  | 'completed'
  | 'cancelled'
  | 'failed';

type OfficeParseSession<TParsed> = {
  readonly result: Promise<TParsed>;
  readonly status: OfficeParseSessionStatus;
  subscribe(listener: (progress: ParseProgress) => void): () => void;
  cancel(): void;
  dispose(): void;
};
```

`createOfficeParseSession(file, options)` accepts a supported `File` and returns `OfficeParseSession<ParsedOfficeFile>`. Cancellation rejects `result`; callers should handle it in the same way as other parse failures when they await the promise.

### Session lifecycle

```ts | pure
import {
  createOfficeParseSession,
  type ParsedOfficeFile,
} from 'office-file-viewer';

export async function parseFile(file: File): Promise<ParsedOfficeFile> {
  const session = createOfficeParseSession(file, { worker: 'auto' });
  const unsubscribe = session.subscribe((progress) => {
    console.info(progress.stage, progress.percent, progress.message);
  });

  try {
    return await session.result;
  } finally {
    unsubscribe();
    session.dispose();
  }
}
```

Always unsubscribe and call `session.dispose()` in `finally`. Calling `dispose()` while a session is running also requests cancellation. It clears session runtime and listeners, but a successfully returned parsed result has its own resource lifetime.

### Parsed-result resources

Use the unified asynchronous disposer after the host no longer needs a parsed result:

```ts | pure
import { disposeParsedOfficeFile } from 'office-file-viewer';
import { parseFile } from './parse-file';

const parsed = await parseFile(file);

try {
  console.info(parsed.kind);
  // Read or render the parsed model here.
} finally {
  await disposeParsedOfficeFile(parsed);
}
```

`OfficeFileViewer` handles its own sessions and parsed resources automatically. Low-level consumers own both lifecycles. `disposeDocDocument`, `disposePresentationDocument`, and `disposeSpreadsheetWorkbook` are available for specialized integrations, but `disposeParsedOfficeFile` is preferred because it also releases the document session attached to a complete result.

## Supported formats and interactions

| Document type      | Extension | Main parser coverage                                                                                                                           |
| ------------------ | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| Word OOXML         | `.docx`   | Paragraphs, lists, tables, images, charts, shapes, links, styles, and theme colors                                                             |
| Word 97-2003       | `.doc`    | Binary document structure, text runs, tables, lists, formatting, images, and basic page-level OfficeArt floating layers                        |
| WPS Writer         | `.wps`    | Reuses the DOC binary pipeline, prioritizing readable content and resource extraction                                                          |
| Excel OOXML        | `.xlsx`   | Worksheets, values, styles, merged cells, dimensions, floating images, and charts                                                              |
| Excel 97-2003      | `.xls`    | BIFF8 cells, formatting, merged ranges, dimensions, OfficeArt images, and charts                                                               |
| PowerPoint OOXML   | `.pptx`   | Master/layout inheritance, text, shapes, images, tables, backgrounds, effects, speaker notes, slide-number/date-time fields, and common charts |
| PowerPoint 97-2003 | `.ppt`    | Binary records, masters, text, shapes, images, embedded charts, embedded DrawingML-compatible text, speaker notes, and static fallbacks        |

Common chart coverage includes line, column, pie, doughnut, area, scatter, bubble, radar, and map charts. Embedded document snapshots are used when possible for unsupported or damaged chart content.

Viewer interactions include:

- Zoom from `25%` to `300%` using manual numeric input, common presets, and `10%` step controls.
- The DOC/DOCX/WPS outline is hidden by default, appears only when usable headings exist, and can be resized horizontally when opened.
- Worksheet tabs plus Original layout and Reading mode for XLS/XLSX.
- Slide and thumbnail navigation for PPT/PPTX; previous/next controls are hidden for a single-slide deck.
- Toggleable and vertically resizable speaker notes when a presentation contains notes.
- Browser fullscreen with automatic state synchronization after leaving through `Esc`.

### Spreadsheet display modes

The display-mode selector appears only for XLS/XLSX previews:

- **Original layout** is the default. It preserves source row heights, column widths, wrapping, shrink-to-fit, merged cells, and clipping behavior.
- **Reading mode** keeps column widths but wraps long text and expands row heights as needed so cell content is easier to read. The resulting layout can differ from the source, but values, formulas, workbook structure, and the source file remain unchanged.
- The selected mode is preserved when switching worksheets in the same file. Opening another file resets the viewer to Original layout.
- For large worksheets, reading-layout adjustments are calculated only for loaded ranges instead of scanning the entire workbook.

<a id="limitations"></a>

## Limitations, performance, and security

`office-file-viewer` is an independent parser and renderer. It is not the native layout engine used by Microsoft Office or WPS Office, so it cannot guarantee pixel-identical output for every Office version, vendor extension, or complex document.

### Rendering boundaries

- Legacy DOC/WPS, XLS, and PPT parsing prioritizes readable content. Complex pagination, anchors, text wrapping, OfficeArt effects, and animations may differ from desktop applications.
- OOXML files can contain unsupported macros, ActiveX controls, OLE objects, SmartArt, vendor extensions, or complex animations. Such content may be ignored, degraded, or represented by an embedded static snapshot.
- The viewer is read-only. It does not edit, save, convert, print-layout, or export Office files to PDF or images.
- Externally linked images and dynamic map data can still require network access. If map data fails, an embedded snapshot is used when available; otherwise the viewer shows an explicit failure state.

### Browser and performance boundaries

The component targets modern browsers and depends on APIs such as `File`, `fetch`, `DOMParser`, `AbortController`, `IntersectionObserver`, `ResizeObserver`, Blob URLs, Canvas, Web Workers, and the Fullscreen API.

Workers reduce main-thread parsing work for supported legacy formats, but they do not eliminate the memory cost of the source file or parsed model. DOCX, XLSX, and PPTX currently parse on the main thread, so very large or complex files can briefly make the UI less responsive.

The library does not impose a maximum source size, ZIP entry count, individual entry size, or total decompressed size. Hosts should set limits appropriate to their devices and threat model.

### Untrusted files and remote access

- Validate file size, extension, MIME type, and source before parsing untrusted input.
- Add server-side malware scanning or content policy checks when required by the application.
- Remote files remain subject to browser CORS, authentication, and Content Security Policy rules; the viewer does not proxy downloads or bypass those controls.
- Parsing local files happens in the browser and does not actively upload them, but the host application remains responsible for its own logging, telemetry, and data-handling policy.
