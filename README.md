# Office File Viewer

English | [简体中文](./README.zh-CN.md) | <a href="https://gyxing.github.io/office-file-viewer/" target="_blank" rel="noopener noreferrer">Live Demo</a>

> A browser-based React component for offline preview of DOC/DOCX/WPS, XLS/XLSX, and PPT/PPTX files.

`office-file-viewer` is a browser-only Office file viewer component for React. File downloading, parsing, and rendering all take place in the browser. It requires no companion document conversion service and does not actively upload local files to a server.

It supports DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX files. The component provides a consistent experience across formats, including file selection, remote loading, parsing progress, zoom, fullscreen mode, slide navigation, and worksheet switching.

> This is an independent parsing and rendering implementation. It is not the native layout engine used by Microsoft Office or WPS Office. Complex documents may render differently from desktop applications. Read [Supported formats](#supported-formats) and [Limitations](#limitations) before use.

## Features

- **Browser-only parsing**: Local files do not need to be uploaded, making the component suitable for intranets, offline environments, and privacy-sensitive scenarios
- **Seven file formats**: Supports DOC/DOCX/WPS, XLS/XLSX, and PPT/PPTX
- **Unified React component**: All formats share the same loading, error, empty-state, zoom, and fullscreen interactions
- **Multiple file sources**: Accepts a `File`, a remote URL, or an async function that returns a `File`, `Blob`, URL, or `Response`
- **Worker-based parsing**: DOC/WPS, XLS, and PPT files can be parsed in a Web Worker to reduce main-thread blocking for large files
- **Progress and cancellation**: Exposes parsing stages, progress subscriptions, and session cancellation; stale results are automatically discarded when switching files
- **Progressive preview**: The component can receive completed content incrementally and render it early while parsing DOC/WPS, XLS, and PPT files
- **Resource management**: The component automatically releases Workers, subscriptions, and Blob URLs; the low-level API also provides explicit disposal functions
- **Host integration**: Compatible with antd v4, v5, and v6, and inherits the host application's `ConfigProvider`
- **Chart fallback**: Uses snapshots embedded in the document when chart or map data fails to load, and shows an explicit state when no fallback is available

## Installation

```bash
npm install office-file-viewer antd react react-dom
```

For projects using Yarn:

```bash
yarn add office-file-viewer antd react react-dom
```

`react`, `react-dom`, and `antd` are provided by the host project. `echarts`, `@zip.js/zip.js`, `emf-converter`, `saxes`, and `@babel/runtime` are installed as runtime dependencies of the component itself.

The package is ESM-only, so the host must use a browser build tool that supports ESM. Component styles are published as CSS and loaded automatically; the host does not need a Less loader. Import public APIs only from `office-file-viewer` rather than undocumented `dist` paths.

## Version compatibility

| antd version        | React / ReactDOM | Status      | Notes                                          |
| ------------------- | ---------------- | ----------- | ---------------------------------------------- |
| `4.24.x`            | `>=16.9.0`       | Supported   | The host entry must load antd v4 global styles |
| `5.x`               | `>=16.9.0`       | Supported   | Uses the antd v5 styling system                |
| `6.x`               | `>=18.0.0`       | Supported   | The React requirement comes from antd v6       |
| `6.x` + React 16/17 | -                | Unsupported | Does not meet antd v6's own requirements       |

Current peer dependency ranges:

```text
antd: >=4.24.0 <7.0.0
react: >=16.9.0
react-dom: >=16.9.0
```

The React peer dependency floor supports host projects using antd v4 or v5. It does not mean antd v6 can run on React 16 or 17.

When using antd v4, load the global stylesheet in the host application entry:

```tsx
import 'antd/dist/antd.css';
```

antd v5 and v6 do not require the stylesheet above. The component does not create an additional root-level `ConfigProvider`; theme, locale, and component prefix settings come from the host configuration.

## Quick start

When `uri` is omitted, the component displays a file picker:

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer />;
}
```

The viewer's own UI defaults to Simplified Chinese. To use English, configure both the viewer and the host Ant Design locale:

```tsx
import { ConfigProvider } from 'antd';
import antdEnUS from 'antd/locale/en_US';
import { OfficeFileViewer } from 'office-file-viewer';

export default function EnglishOfficePreview() {
  return (
    <ConfigProvider locale={antdEnUS}>
      <OfficeFileViewer locale="en-US" />
    </ConfigProvider>
  );
}
```

Pass a local file or remote URL:

```tsx
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

When `height` is omitted, the viewer uses the height of its parent container, so the parent must have a computable height. You can also pass a numeric pixel value, `720px`, `80vh`, or `100%` directly.

Use an async file source, parsing options, and event callbacks:

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return (
    <OfficeFileViewer
      uri={async () => fetch('/files/demo.xlsx')}
      parseOptions={{ worker: 'auto' }}
      onParseProgress={(progress) => {
        const percent =
          progress.percent === undefined
            ? ''
            : ` ${Math.round(progress.percent * 100)}%`;
        console.info(progress.stage, `${progress.message}${percent}`);
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

## `uri` file sources

`uri` accepts the following forms:

```ts
type OfficeFileViewerUri =
  | File
  | string
  | (() => Promise<File | Blob | string | Response>);
```

Consider the following when using remote files:

- Cross-origin URLs must allow browser access through CORS.
- URLs should preferably include a supported file extension.
- URLs without an extension must expose the format through the filename in `Content-Disposition` or the response `Content-Type`.
- A URL with an unsupported extension is rejected before downloading, even if the response actually contains an Office file.
- When `uri` changes, the previous URL download is cancelled through `AbortController`.
- A custom async function cannot be forcibly cancelled, but its stale result will not replace a newer file.
- Manually selecting a file also stops the current remote download and invalidates stale parsing results.

## Component props

| Prop              | Type                                             | Default        | Description                                                             |
| ----------------- | ------------------------------------------------ | -------------- | ----------------------------------------------------------------------- |
| `locale`          | `'zh-CN' \| 'en-US'`                             | `'zh-CN'`      | Viewer UI language; Ant Design locale still comes from `ConfigProvider` |
| `uri`             | `OfficeFileViewerUri`                            | -              | File source to preload; displays the file picker when omitted           |
| `defaultFileName` | `string`                                         | Locale message | Name displayed when no file is loaded                                   |
| `defaultZoom`     | `number`                                         | `100`          | Initial zoom percentage, clamped to the range from `25` to `300`        |
| `className`       | `string`                                         | -              | Custom class name for the root container                                |
| `height`          | `CSSProperties['height']`                        | `100%`         | Viewer height; follows the parent container when omitted                |
| `style`           | `CSSProperties`                                  | -              | Custom styles for the root container                                    |
| `parseOptions`    | `OfficeParseOptions`                             | `{}`           | Worker mode and custom Worker factory                                   |
| `onParseProgress` | `(progress: ParseProgress) => void`              | -              | Called when the parsing stage or progress changes                       |
| `onFileParsed`    | `(parsed: ParsedOfficeFile, file: File) => void` | -              | Called once after full parsing; does not receive progressive internals  |
| `onError`         | `(error: Error, file?: File) => void`            | -              | Called when file downloading, parsing, or fullscreen operation fails    |

The available `PreviewKind` values are `'docx' | 'doc' | 'xlsx' | 'xls' | 'pptx' | 'ppt'`. WPS files reuse the `'doc'` preview model.

When provided, `height` takes precedence over `style.height`. Percentage heights depend on the parent height; if the page has no fixed height, use a number, `px`, or `vh` value.

Common public types:

```ts
import type {
  OfficeFileViewerProps,
  OfficeFileViewerLocale,
  OfficeFileViewerUri,
  OfficeParseOptions,
  OfficeParseSession,
  OfficeParseSessionStatus,
  ParsedOfficeFile,
  ParseProgress,
  ParseStage,
  PreviewKind,
  WorkerMode,
} from 'office-file-viewer';
```

## Supported formats

| Document type      | Extension | Current capabilities                                                                                                  |
| ------------------ | --------- | --------------------------------------------------------------------------------------------------------------------- |
| Word OOXML         | `.docx`   | Rich-text paragraphs, lists, tables, images, charts, VML/WPG shapes, hyperlinks, styles, and theme colors             |
| Word 97-2003       | `.doc`    | CFB, FIB, Piece Table, FKP, SPRM, body structure, tables, lists, and image extraction                                 |
| WPS Writer         | `.wps`    | Reuses the DOC binary parsing pipeline, prioritizing readable content and resource extraction                         |
| Excel OOXML        | `.xlsx`   | Multiple worksheets, cell values and styles, merged cells, row and column dimensions, floating images, and charts     |
| Excel 97-2003      | `.xls`    | BIFF8 workbooks, cells, formatting, merged ranges, row and column dimensions, OfficeArt images, and charts            |
| PowerPoint XML     | `.pptx`   | Master and layout inheritance, text, shapes, images, tables, backgrounds, gradients, shadows, and common chart types  |
| PowerPoint 97-2003 | `.ppt`    | Binary records, masters, text, shapes, images, embedded charts, and static previews for content that cannot be parsed |

Office and WPS charts currently cover common types such as line, column, pie, doughnut, area, scatter, bubble, radar, and map charts. Non-standard extensions or corrupted data use snapshots embedded in the document when possible; otherwise, the viewer displays an explicit state.

The capabilities in the table describe the main parser paths currently covered. They do not guarantee complete restoration of every Office version, vendor extension, macro, embedded object, animation, or complex layout.

## Web Worker and performance

`parseOptions.worker` controls where parsing runs:

| Mode       | DOC/WPS, XLS, PPT                                              | DOCX, XLSX, PPTX                                            |
| ---------- | -------------------------------------------------------------- | ----------------------------------------------------------- |
| `'auto'`   | Prefers a Worker; falls back to the main thread when necessary | Currently runs directly on the main thread                  |
| `'always'` | Requires a Worker; throws a configuration error if unavailable | Worker migration is not yet complete; throws a config error |
| `'never'`  | Always runs on the main thread                                 | Always runs on the main thread                              |

```tsx
<OfficeFileViewer parseOptions={{ worker: 'auto' }} />
```

Each parsing session owns an independent Worker. File buffers are transferred to the Worker as transferable objects. Ordered messages, ACK backpressure, and cancellation messages coordinate parsing. The component terminates the current Worker when switching files or unmounting.

`workerFactory` allows the host to control Worker creation, mainly for build environments with special asset paths or CSP settings. The factory must return a Worker compatible with the current parsing protocol; most projects should keep the default.

Notes for large files:

- Workers reduce the time spent blocking the main thread while parsing legacy binary formats, but do not reduce the memory used by the file itself or the parsed model.
- DOCX, XLSX, and PPTX files are currently parsed on the main thread, so complex or very large files may briefly make the interface unresponsive.
- The component does not limit file size, ZIP entry count, individual entry size, or total decompressed size. The host should validate these according to its use case.

## Low-level parsing sessions

Use `createOfficeParseSession` when you need to manage the parsing lifecycle directly:

```ts
import {
  createOfficeParseSession,
  disposeDocDocument,
  disposePresentationDocument,
  disposeSpreadsheetWorkbook,
  type ParsedOfficeFile,
} from 'office-file-viewer';

export async function parseFile(file: File) {
  const session = createOfficeParseSession(file, { worker: 'auto' });
  const unsubscribe = session.subscribe((progress) => {
    console.info(progress.stage, progress.percent, progress.message);
  });

  try {
    const parsed = await session.result;
    console.info('Parsing completed', parsed.kind);
    return parsed;
  } finally {
    unsubscribe();
    session.dispose();
  }
}

export function disposeParsedFile(parsed: ParsedOfficeFile) {
  if (parsed.kind === 'doc') disposeDocDocument(parsed.document);
  if (parsed.kind === 'ppt') disposePresentationDocument(parsed.document);
  if (parsed.kind === 'xls' || parsed.kind === 'xlsx') {
    disposeSpreadsheetWorkbook(parsed.workbook);
  }
}
```

`OfficeParseSession` provides:

- `result`: A Promise for the final parsing result
- `status`: `starting`, `running`, `completed`, `cancelled`, or `failed`
- `subscribe(listener)`: Subscribes to parsing progress and returns an unsubscribe function
- `cancel()`: Requests cancellation of the current task
- `dispose()`: Releases the Worker, subscriptions, and temporary resources not transferred to the result

Progress stages include `reading`, `container`, `structure`, `content`, `resources`, and `assembling`. When present, `percent` ranges from `0` to `1`.

`OfficeFileViewer` automatically handles cancellation and resource cleanup. When using the low-level session directly, call `session.dispose()` after completion. When a long-lived parsing result is no longer needed, call the corresponding document disposal function to release Blob URLs.

## Interactions

- The zoom range is `25%` to `300%`, with toolbar presets at `50%`, `75%`, `100%`, `125%`, `150%`, and `200%`.
- PPT/PPTX supports slide navigation and thumbnail navigation.
- XLS/XLSX supports switching between worksheet tabs.
- DOC/DOCX/WPS does not display an additional document title bar, leaving more vertical space for document content.
- Fullscreen mode depends on the browser Fullscreen API. The button is disabled when unsupported, and the state synchronizes automatically after exiting with `Esc`.
- Map charts may need to load external GeoJSON data. If the network request fails, the viewer uses a document snapshot when available; otherwise, it displays a failure state.

## Limitations

- The component targets modern browsers and depends on APIs including `File`, `fetch`, `DOMParser`, `AbortController`, `IntersectionObserver`, `ResizeObserver`, Blob URLs, Canvas, Web Workers, and the Fullscreen API.
- Local file parsing and normal rendering do not require a server. A remote `uri`, externally linked images, or dynamic map data may still require network access.
- Remote files must comply with browser CORS policies. The component does not proxy downloads or bypass authentication and cross-origin restrictions.
- For untrusted files, hosts should validate file size, extension, MIME type, and source before parsing, and add server-side protections such as malware scanning where appropriate.
- DOC/WPS, XLS, and PPT are legacy binary formats. The current implementation prioritizes readable content and does not guarantee that complex pagination, animation, anchoring, text wrapping, or layout will match desktop Office applications exactly.
- OOXML documents may contain unsupported vendor extensions, macros, ActiveX controls, OLE objects, SmartArt, or complex animations. Such content may be degraded, ignored, or shown as a static preview.
- The viewer is read-only. It does not provide editing, saving, format conversion, print layout, or PDF export.

## Project structure

```text
src/
├── index.ts
└── office-file-viewer/
    ├── OfficeFileViewer.tsx  # Public component and file-loading orchestration
    ├── shell/                # Toolbar, preview dispatch, and shared states
    ├── services/
    │   ├── parsing/          # Parsing sessions, Worker protocol, runtime, and result assembly
    │   ├── doc/ docx/        # DOC/WPS and DOCX parsing
    │   ├── xls/ xlsx/        # XLS and XLSX parsing
    │   └── ppt/ pptx/        # PPT and PPTX parsing
    ├── formats/              # React renderers for each document format
    └── shared/
        ├── binary/           # Binary primitives such as CFB
        ├── officeart/        # OfficeArt drawing records
        ├── ooxml/            # ZIP, XML, relationships, themes, media, and chart adapters
        └── chart/            # ECharts rendering and failure fallback
```

Core data flow:

```text
File / URL / async loader
  → Detect filename and MIME type
  → Parse on the main thread or in a Web Worker
  → Normalize into TypeScript document models
  → Render with React format renderers
  → Display the browser preview
```

## Local development

```bash
yarn
yarn start
```

Build the component library and Dumi documentation:

```bash
yarn build
yarn docs:build
```

The project uses TypeScript, ESLint, Stylelint, and Prettier for static checks. Before publishing, test local files, remote URIs, Workers, fullscreen mode, and sample documents for each format in the target React and antd combinations.

## License

[MIT](./LICENSE)
