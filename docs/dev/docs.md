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
      uri={async (signal) => fetch('/files/demo.xlsx', { signal })}
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
      onWarning={(warning, file) => {
        console.warn('Preview warning', file.name, warning.code);
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
type OfficeFileViewerUriLoader = (
  signal?: AbortSignal,
) => Promise<File | Blob | string | Response>;

type OfficeFileViewerUri = File | string | OfficeFileViewerUriLoader;
```

Remote source rules:

- Cross-origin URLs must allow browser access through CORS.
- URLs should preferably end with a supported Office extension.
- An extensionless URL must expose a filename through `Content-Disposition` or a supported MIME type through `Content-Type`.
- An explicitly unsupported URL extension is rejected before download, even if its response body contains an Office file.
- Changing `uri` cancels the previous URL download and parse.
- An async loader receives an optional `AbortSignal`; pass it to `fetch` or another cancellable API. Existing zero-argument loaders remain compatible, and stale results cannot replace the latest source.
- Manually selecting a file also invalidates an in-flight remote or parsing result.
- After loading or parsing fails, the built-in error state can reload the most recent source.

<a id="component-api"></a>

## `OfficeFileViewer` API

| Prop                             | Type                                                 | Default           | Description                                                      |
| -------------------------------- | ---------------------------------------------------- | ----------------- | ---------------------------------------------------------------- |
| `locale`                         | `'zh-CN' \| 'en-US'`                                 | `'zh-CN'`         | Built-in viewer UI language                                      |
| `uri`                            | `OfficeFileViewerUri`                                | -                 | File source to preload; the file picker is shown when omitted    |
| `defaultFileName`                | `string`                                             | Localized message | Fallback name when the source does not provide a usable filename |
| `defaultZoom`                    | `number`                                             | `100`             | Initial zoom percentage, clamped from `25` to `300`              |
| `defaultViewState`               | `Partial<OfficeFileViewerViewState>`                 | -                 | Initial values for uncontrolled view fields                      |
| `viewState`                      | `Partial<OfficeFileViewerViewState>`                 | -                 | Per-field control of zoom, page, sidebars, and display mode      |
| `onViewStateChange`              | `(state, change) => void`                            | -                 | Called when the user requests a view-state change                |
| `defaultShowSpeakerNotes`        | `boolean`                                            | `false`           | Initial speaker-notes state in uncontrolled mode                 |
| `showSpeakerNotes`               | `boolean`                                            | -                 | Controlled speaker-notes visibility                              |
| `onSpeakerNotesVisibilityChange` | `(visible: boolean) => void`                         | -                 | Called when the presentation notes visibility changes            |
| `className`                      | `string`                                             | -                 | Additional class name for the viewer root                        |
| `height`                         | `CSSProperties['height']`                            | Parent height     | Viewer height; takes precedence over `style.height`              |
| `style`                          | `CSSProperties`                                      | -                 | Inline styles for the viewer root                                |
| `onFileParsed`                   | `(parsed: ParsedOfficeFile, file: File) => void`     | -                 | Called once when the complete materialized result is available   |
| `onPreviewReady`                 | `(info: OfficePreviewReadyInfo, file: File) => void` | -                 | Called once when the first usable preview is ready               |
| `onError`                        | `(error: Error, file?: File) => void`                | -                 | Called when loading, parsing, or a viewer operation fails        |
| `onWarning`                      | `(warning, file) => void`                            | -                 | Called for non-fatal parse, partial-preview, or font warnings    |
| `parseOptions`                   | `OfficeParseOptions`                                 | `{}`              | Worker strategy and optional resource limits                     |
| `imagePreview`                   | `boolean \| OfficeFileViewerImagePreviewOptions`     | `true`            | Content-image preview, download, and context-menu configuration  |
| `hyperlink`                      | `boolean`                                            | `true`            | Enables hyperlinks explicitly declared by the source document    |
| `search`                         | `false \| OfficeFileViewerSearchOptions`             | `{}`              | Full-document search runtime and initial matching options        |
| `review`                         | `false \| OfficeFileViewerReviewOptions`             | `{}`              | Read-only comments, revisions, footnotes, and endnotes           |
| `presentationMedia`              | `false \| OfficeFileViewerPresentationMediaOptions`  | `{}`              | Presentation media, external-source, and download policy         |
| `transitions`                    | `false \| 'source'`                                  | `false`           | Play supported source slide transitions                          |
| `fontOptions`                    | `OfficeFileViewerFontOptions`                        | `{}`              | Font aliases, fallback families, and missing-font diagnostics    |
| `onHyperlinkActivate`            | `(event: OfficeHyperlinkActivateEvent) => void`      | -                 | Called on valid activation and can prevent default navigation    |
| `onParseProgress`                | `(progress: ParseProgress) => void`                  | -                 | Called when the current parse stage or progress changes          |

### Content-image preview

Visible content images in DOC, DOCX, WPS, XLS, and XLSX support double-click, `Enter`, or Space to open the preview layer by default. The layer provides fit-to-window display, `10%` to `500%` zoom, panning, clockwise rotation, reset, and download. The custom image context menu contains only Preview and Download.

```ts | pure
type OfficeFileViewerImagePreviewOptions = {
  download?: boolean;
  contextMenu?: boolean;
};

type OfficeFileViewerImagePreviewConfig =
  | boolean
  | OfficeFileViewerImagePreviewOptions;
```

Omit the prop or pass `true` to enable every capability. Pass `false` to disable the additional image interactions. Object configuration can disable download or the custom context menu independently:

```tsx | pure
<OfficeFileViewer
  uri={file}
  imagePreview={{ download: false, contextMenu: true }}
/>
```

When `contextMenu` is disabled, double-click and keyboard preview remain available and the browser-native context menu is restored. Decorative header images, backgrounds, watermarks, page drawing layers, charts, and PPT/PPTX images are outside this interaction scope.

### Office hyperlinks

The viewer handles only links explicitly declared by the source Office file on text, cells, images, shapes, or action buttons. It does not turn ordinary URL text into links. To prevent accidental navigation while reading, activate a link with `Ctrl + click` on Windows and Linux or `Command + click` on macOS. A focused link can be activated with `Enter`; touch devices require a second tap within the confirmation window. Text, cell, shape, and button links also provide a context menu for direct opening or internal navigation, with copy support for safe external targets. Local paths, restricted targets, and images keep their existing context-menu behavior.

```tsx | pure
import {
  OfficeFileViewer,
  type OfficeHyperlinkActivateEvent,
} from 'office-file-viewer';

export default function OfficePreview({ file }: { file: File }) {
  const handleHyperlinkActivate = (event: OfficeHyperlinkActivateEvent) => {
    console.info(event.sourceType, event.sourceId, event.hyperlink);

    // Hand navigation to application routing, authorization, or audit logic.
    if (event.hyperlink.kind === 'external') {
      event.preventDefault();
    }
  };

  return (
    <OfficeFileViewer
      uri={file}
      hyperlink
      onHyperlinkActivate={handleHyperlinkActivate}
    />
  );
}
```

`OfficeHyperlinkActivateEvent` exposes `hyperlink`, `file`, `previewKind`, `sourceType`, `sourceId`, `defaultPrevented`, and `preventDefault()`. Default behavior is:

- `http` and `https` open in a new tab; `mailto` and `tel` are delegated to the browser or operating system.
- Word bookmarks, Excel sheets/cells/defined names, and PowerPoint slide actions navigate inside the current viewer.
- Local paths, UNC paths, and relative targets without a reliable remote source URL are reported to the callback only and are not opened automatically.
- `javascript`, `data`, `vbscript`, and unknown executable protocols are always blocked from default execution; a host callback cannot bypass this boundary.
- Modifier-click activates an image link; ordinary double-click and the context menu continue to follow the content-image preview configuration.

Set `hyperlink={false}` to remove link focus, hints, and activation behavior without changing source colors, underlines, or layout.

### Full-document search

DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX expose the toolbar search action and `Ctrl + F` (`Command + F` on macOS). Queries use cancellable incremental scanning, so generated results are navigable before a large document has been scanned completely. `Esc` closes the sidebar, and switching files clears the previous query and results.

```ts | pure
type OfficeFileViewerSearchOptions = {
  defaultVisible?: boolean;
  matchCase?: boolean;
  wholeWord?: boolean;
};
```

Search is enabled when the prop is omitted. Pass `false` to remove its action, keyboard shortcut, and runtime. Object values are uncontrolled defaults; use `viewState.searchVisible` for controlled visibility. Search covers Word body content, Excel cells, and visible text on PowerPoint slides. Word comments and PowerPoint speaker notes are currently excluded.

### Read-only review and Word notes

Review is enabled by default. Word shows the revision-view selector only when comments or revisions exist; comments and revisions render directly in the page-side markup rail without a Review button or fixed panel. Excel and PowerPoint expose the Review entry only when comments exist.

```ts | pure
type OfficeFileViewerReviewOptions = {
  defaultPanelVisible?: boolean;
  defaultRevisionMode?: 'final' | 'markup' | 'original';
  showComments?: boolean;
  showNotes?: boolean;
};
```

`defaultPanelVisible` and `viewState.reviewPanelVisible` control only the Excel and PowerPoint review panel. Word page-side markup is controlled by `wordRevisionMode`.

- DOCX restores comment threads, insertions, deletions, moves, and common formatting revisions with final, markup, and original read-only projections.
- Word revision modes are switched from the toolbar. Markup places comments and revisions in one page-side rail: comments use blue author markers, revisions use red author markers, and dashed leaders connect both kinds directly to their text while the rail resolves vertical collisions. The physical page and one base rail are centered as a stable group; additional rail columns grow only to the right, so scrolling does not shift the page horizontally. It uses neither rounded cards nor a separate review list. Narrow containers retain the rail through horizontal scrolling instead of covering body content.
- DOC/WPS restores recoverable insert/delete metadata and binary comments. Unsupported property-level original formatting keeps the final format and emits a stable warning.
- DOC, DOCX, and WPS footnotes appear on the reference page, long footnotes continue in order, and endnotes appear at the end. Search excludes comments and notes by default.
- Excel cell comments and PowerPoint slide comments use the same review panel. Selecting an item switches the worksheet or slide before focusing its target.
- `review={false}` removes the review runtime without changing the final body projection.

### Spreadsheet business semantics

XLS/XLSX restore frozen panes, Table/AutoFilter metadata, cell comments, and a safe conditional-formatting subset: `cellIs`, two/three-color scales, data bars, common icon sets, duplicate/unique values, Top10, and above/below average. Rules that require range-wide statistics are scanned in Source tiles and cached per rule, so scrolling does not recalculate against only the visible window. Unsupported formula rules are retained as summaries and emit warnings. The viewer never re-filters, re-sorts, or fully recalculates the workbook.

### Presentation media and slide transitions

```ts | pure
type OfficeFileViewerPresentationMediaOptions = {
  allowExternal?: boolean;
  download?: boolean;
};

type OfficeFileViewerPresentationTransitions = false | 'source';
```

Embedded audio and video use native browser controls with `autoplay={false}` and `preload="metadata"`. A media URL is created only for the active slide, and leaving the slide pauses playback and releases its reference. `presentationMedia={false}` disables playback; downloads default to enabled. External media makes no request by default. `allowExternal: true` permits only HTTP(S); `javascript`, `data`, `vbscript`, local paths, and unknown executable schemes remain blocked.

Transitions are disabled by default. With `transitions="source"`, toolbar previous/next actions restore `fade`, `push`, `wipe`, `split`, `cover`, and `uncover`. Thumbnail, search, and comment navigation does not trigger transitions. Rapid navigation cancels the old animation, and reduced-motion preference switches immediately.

### Font aliases and fallback

The package does not bundle or download Office fonts. Rendering keeps the source font first, then adds built-in aliases, host aliases, global fallback families, and a CSS generic family. Hosts can extend the mapping for their deployment environment:

```tsx | pure
<OfficeFileViewer
  uri={file}
  fontOptions={{
    aliases: {
      'Source Office Font': ['Available Corporate Font', 'Arial'],
    },
    fallbackFamilies: ['Noto Sans CJK SC', 'sans-serif'],
    warnOnMissing: true,
  }}
  onWarning={(warning) => {
    if (warning.code === 'FONT_FALLBACK_APPLIED') {
      console.warn(warning.requestedFamily, warning.candidates);
    }
  }}
/>
```

`aliases` override case-insensitively matching built-in aliases, and `fallbackFamilies` are appended to every font chain. `warnOnMissing` defaults to `true`, but diagnostics run only when `onWarning` is supplied and the browser supports the Font Loading API. Checks are batched after the first usable preview, and each missing family is reported once per document session with the stable code `FONT_FALLBACK_APPLIED`. Missing fonts fall back without blocking preview.

### Unified view state

Use `defaultViewState` for uncontrolled initial values. `viewState` controls only the fields that are present, while omitted fields remain internally managed:

```ts | pure
type OfficeFileViewerViewState = {
  zoom: number;
  activeSlideIndex: number;
  activeSheetId?: string;
  wordOutlineVisible: boolean;
  searchVisible: boolean;
  reviewPanelVisible: boolean;
  speakerNotesVisible: boolean;
  wordRevisionMode: 'final' | 'markup' | 'original';
  spreadsheetViewMode: 'source' | 'reading';
};
```

The first `onViewStateChange` argument is the complete view state after applying the request. The second is a single-field `{ key, value }` change. Invalid zoom values fall back and are clamped from `25` to `300`; slide indices and worksheet IDs are validated against the current file.

The legacy `defaultZoom`, `defaultShowSpeakerNotes`, `showSpeakerNotes`, and `onSpeakerNotesVisibilityChange` props remain compatible. When both APIs are provided, matching fields in `defaultViewState` override legacy defaults, and `viewState.speakerNotesVisible` overrides `showSpeakerNotes`.

### Callback timing

- `onPreviewReady` fires when either a complete model (`mode: 'materialized'`) or an on-demand source (`mode: 'source'`) can render its first usable preview.
- `onFileParsed` fires after the complete materialized result is ready; progressive intermediate results do not trigger it.
- `onParseProgress` can fire many times and may omit exact counts or `percent`.
- `onWarning` uses stable `code`, `previewKind`, and `source` fields to distinguish parser, retained-partial-preview, hyperlink, and font warnings; it does not replace `onError`.
- `onError` can receive no `file` when an error occurs before a usable file has been resolved. Resource-policy failures can be inspected through `OfficeResourceLimitError.code`.

## Parsing configuration and progress

```ts | pure
type WorkerMode = 'auto' | 'always' | 'never';

type OfficeParseResourcePolicy = {
  maxFileBytes?: number;
  maxArchiveEntries?: number;
  maxArchiveEntryBytes?: number;
  maxArchiveInflatedBytes?: number;
  maxCompressionRatio?: number;
  timeoutMs?: number;
};

type OfficeParseOptions = {
  worker?: WorkerMode;
  workerFactory?: () => Worker;
  resourcePolicy?: OfficeParseResourcePolicy;
};
```

| Mode       | Behavior for all seven formats                                                                                                                                                                  |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `'auto'`   | Default. Legacy formats prefer a Worker; small OOXML files use complete models, large OOXML files use Worker-owned on-demand sources; only Worker startup failures fall back to the main thread |
| `'always'` | Forces a Worker for every format. An unavailable Worker, load failure, or protocol mismatch fails instead of falling back                                                                       |
| `'never'`  | Disables Workers. Parsing runs on the main thread, while large files can still use on-demand sources and virtual rendering                                                                      |

Each active parse session owns its Worker and resource reader. Source changes, cancellation, and unmounting cancel requests and release the Worker, archive reader, and Blob URLs. Searches plus on-demand page, worksheet, slide, and lazy-resource reads stay inside the Worker for large files; the main thread receives only the structured data currently needed for display. `workerFactory` is intended for hosts with special asset paths or Content Security Policy requirements; most applications should use the built-in factory.

Vite pre-bundles dependencies in its development server. Only when development explicitly forces `worker: 'always'`, add the package to `optimizeDeps.exclude` so Vite preserves and processes the built-in Worker asset URL. Production builds and the default `auto` mode do not require this setting.

```ts | pure
export default {
  optimizeDeps: { exclude: ['office-file-viewer'] },
};
```

The file-profile thresholds used by `auto` select complete parsing, on-demand reads, virtual rendering, and task-yielding strategies. They are not file-size limits and never reject a file by themselves. Small files intentionally keep the simpler complete-model path to avoid unnecessary Worker messaging and lazy-read overhead.

### Resource limits

No file-size, timeout, or OOXML-archive limits are enabled by default, so the library does not reject valid business files without host input. Configure limits for untrusted files according to the target device and application boundary:

```tsx | pure
<OfficeFileViewer
  parseOptions={{
    resourcePolicy: {
      maxFileBytes: 100 * 1024 * 1024,
      maxArchiveEntries: 10_000,
      maxArchiveEntryBytes: 64 * 1024 * 1024,
      maxArchiveInflatedBytes: 512 * 1024 * 1024,
      maxCompressionRatio: 500,
      timeoutMs: 120_000,
    },
  }}
/>
```

`maxFileBytes` and `timeoutMs` apply to every format. Archive entry, inflated-size, and compression-ratio limits apply to DOCX, XLSX, and PPTX. Every configured value must be finite and greater than zero. When a limit is hit, `onError` receives an `OfficeResourceLimitError` with `code`, `limit`, `actual`, and an optional `path`. `timeoutMs` requests cancellation; synchronous main-thread parsing can observe it only after yielding control.

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

### `.ppt` compatibility parser

The root entry keeps `parsePpt(file)` for fully parsing a single unencrypted `.ppt` file into a `PresentationDocument`. This compatibility API materializes the whole presentation on the main thread and does not provide Worker selection, parse progress, cancellation, or resource limits. Prefer `createOfficeParseSession` for new integrations.

```ts | pure
import { disposePresentationDocument, parsePpt } from 'office-file-viewer';

const presentation = await parsePpt(file);

try {
  console.info(presentation.slides.length);
} finally {
  disposePresentationDocument(presentation);
}
```

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

| Document type      | Extension | Main parser coverage                                                                                                   |
| ------------------ | --------- | ---------------------------------------------------------------------------------------------------------------------- |
| Word OOXML         | `.docx`   | Body layout, graphics, links, outlines, comments, revisions, footnotes, and endnotes                                   |
| Word 97-2003       | `.doc`    | Binary body, tables, graphics, links, outlines, and recoverable comments, revisions, footnotes, and endnotes           |
| WPS Writer         | `.wps`    | Reuses the DOC binary pipeline, prioritizing body content, resources, and recoverable review semantics                 |
| Excel OOXML        | `.xlsx`   | Worksheets, graphics, links, frozen panes, tables/filters, comments, and common conditional formatting                 |
| Excel 97-2003      | `.xls`    | BIFF8 cells, OfficeArt, charts, links, panes, comments, and recoverable filter, table, and conditional-format metadata |
| PowerPoint OOXML   | `.pptx`   | Masters/layouts, graphics, links, comments, notes, embedded/external media, and six common slide transitions           |
| PowerPoint 97-2003 | `.ppt`    | Binary masters, graphics, links, comments, notes, recoverable media, common transitions, and static fallbacks          |

Common chart coverage includes line, column, pie, doughnut, area, scatter, bubble, radar, and map charts. Embedded document snapshots are used when possible for unsupported or damaged chart content.

Viewer interactions include:

- Zoom from `25%` to `300%` using manual numeric input, common presets, and `10%` step controls.
- Visible content images in DOC/DOCX/WPS and XLS/XLSX support preview, zoom, rotation, download, and a custom context menu.
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
- Page-level DOC/WPS OfficeArt canvases are currently displayed as a single SVG image, so independent link hit areas for child shapes cannot be retained precisely. These links emit a non-fatal degradation warning; field links and body bookmarks are unaffected.
- OOXML files can contain unsupported macros, ActiveX controls, OLE objects, SmartArt, vendor extensions, or complex animations. Such content may be ignored, degraded, or represented by an embedded static snapshot.
- Review is read-only: comments cannot be added, replied to, deleted, resolved, or written back. DOC/WPS property-level revisions that cannot restore original values keep final formatting.
- Excel does not re-run filters, sorting, or a complete formula engine. Unsupported conditional-format formulas are retained only as rule summaries.
- PowerPoint object animations, triggers, timelines, media synchronization, and codec transcoding are outside the current scope.
- The viewer is read-only. It does not edit, save, convert, print-layout, or export Office files to PDF or images.
- Externally linked images and dynamic map data can still require network access. If map data fails, an embedded snapshot is used when available; otherwise the viewer shows an explicit failure state.

### Browser and performance boundaries

The component targets modern browsers and depends on APIs such as `File`, `fetch`, `DOMParser`, `AbortController`, `IntersectionObserver`, `ResizeObserver`, Blob URLs, Canvas, Web Workers, and the Fullscreen API.

All seven formats can move parsing work into a Worker. Large DOCX, XLSX, and PPTX files additionally keep their archive reader in the Worker and load pages, ranges, slides, and resources on demand. Workers and virtual rendering reduce main-thread long tasks and one-time model cost, but they do not eliminate memory used by the source file, loaded content, or browser graphics resources.

The library does not enable a maximum source size, ZIP entry count, individual entry size, or total decompressed size by default, and internal large-file thresholds never reject a file. For untrusted files, hosts should explicitly configure `parseOptions.resourcePolicy` according to their devices and threat model.

### Untrusted files and remote access

- Validate file size, extension, MIME type, and source before parsing untrusted input.
- Add server-side malware scanning or content policy checks when required by the application.
- Remote files remain subject to browser CORS, authentication, and Content Security Policy rules; the viewer does not proxy downloads or bypass those controls.
- Parsing local files happens in the browser and does not actively upload them, but the host application remains responsible for its own logging, telemetry, and data-handling policy.
