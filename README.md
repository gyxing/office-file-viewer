# Office File Viewer

English | [简体中文](./README.zh-CN.md) | <a href="https://gyxing.github.io/office-file-viewer/" target="_blank" rel="noopener noreferrer">Live Demo</a> | [Documentation](https://gyxing.github.io/office-file-viewer/docs)

> A browser-based React component for offline preview of DOC/DOCX/DOCM/DOTX/WPS, XLS/XLSX/XLSM/XLTX, and PPT/PPTX/PPTM/POTX files.

`office-file-viewer` downloads, parses, and renders Office files entirely in the browser. It requires no companion document-conversion service and does not actively upload local files.

The package provides a unified viewer for Word documents, Excel spreadsheets, and PowerPoint presentations, including full-document search, read-only review, footnotes/endnotes, spreadsheet business semantics, secure media playback, slide transitions, zoom, fullscreen mode, and speaker notes.

> This is an independent parsing and rendering implementation, not the native Microsoft Office or WPS Office layout engine. Complex documents can render differently from desktop applications. Review the [complete limitations](https://gyxing.github.io/office-file-viewer/docs#limitations) before use.

## Features

- **Browser-only parsing**: Suitable for intranets, offline environments, and privacy-sensitive workflows.
- **Three Office families**: 13 common Word, Excel, and PowerPoint extensions.
- **Unified React component**: Shared loading, error, empty, zoom, and fullscreen interactions.
- **Themes and watermarks**: Built-in light, dark, and system modes support semantic token overrides and a performant text watermark over document content.
- **Reusable viewer shell**: Host-rendered content can reuse the same toolbar, zoom, fullscreen, theme, and watermark capabilities.
- **Flexible sources**: Accepts a local `File`, remote URL, or async loader with a cancellation signal.
- **Progressive preview**: Supported formats can render completed content while parsing continues.
- **Full-document search**: Every supported extension provides cancellable incremental search, result navigation, case matching, and whole-word matching.
- **Controlled view state**: Hosts can control percentage/fit zoom, active pages, worksheets, sidebars, and display modes per field.
- **Content-image actions**: Standalone content images in Word, Excel, and PowerPoint support double-click preview plus preview and download actions from a custom context menu; PowerPoint combines resource type, object metadata, and display size to identify small icons, which expose neither image preview nor the custom preview menu, while master, layout, and background images remain static.
- **Source hyperlinks**: Text, cell, image, shape, and action-button links use safe `Ctrl`/`Command` modifier activation and can be intercepted by the host.
- **Read-only review**: Word comments and revisions use a native-style page-side markup rail with dashed leaders connected directly to the body, without a fixed review list. DOCX supports final, markup, and original views, while DOC/WPS restores recoverable review semantics. Excel and PowerPoint comments continue to use the shared review panel.
- **Word notes**: All Word formats support footnotes, endnotes, reference navigation, and continuation pages for long footnotes.
- **Spreadsheet semantics**: Frozen panes, Table/AutoFilter, cell comments, and a common conditional-formatting subset are restored without modifying workbook data.
- **Presentation media and transitions**: Embedded audio/video is loaded only for the active slide and never autoplayed. External media is blocked by default, and common source transitions are opt-in.
- **Adaptive Workers**: Every supported extension reuses its parser type in a Worker. The default mode selects a complete model or an on-demand source from the file profile and safely falls back if Worker startup fails.
- **Font fallback and host fonts**: Source fonts, aliases, fallback chains, and opt-in host font resources are resolved consistently, with structured warnings when a family or resource is unavailable.
- **Structured errors**: Input, download, format, encryption, Worker, resource, and parse failures expose stable codes and stages.
- **Resource management**: The component handles cancellation, subscriptions, Workers, and Blob URLs, with optional host-configured parse limits.
- **Built-in viewer interface**: Includes scoped controls and styles for file selection, navigation, zoom, and fullscreen.

## Installation

```bash
npm install office-file-viewer
```

With Yarn:

```bash
yarn add office-file-viewer
```

`react` and `react-dom` are peer dependencies supplied by the host. The package is ESM-only and component styles are built as scoped CSS. Import Office-preview APIs from `office-file-viewer`; when only the reusable shell is needed, the stable `office-file-viewer/layout` subpath is also available.

## Version compatibility

| Item          | Requirement | Notes                                                  |
| ------------- | ----------- | ------------------------------------------------------ |
| React         | `>=16.9.0`  | Hooks-capable React versions                           |
| ReactDOM      | `>=16.9.0`  | Keep the same major version as React                   |
| Module format | ESM-only    | Use an ESM-capable browser build tool                  |
| Browsers      | Modern      | Recommended: Chromium 100+, Firefox 115+, Safari 16.4+ |

Current peer dependency ranges:

```text
react: >=16.9.0
react-dom: >=16.9.0
```

## Quick start

When `uri` is omitted, the component shows its built-in file picker:

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer height="80vh" />;
}
```

See the [complete quick start](https://gyxing.github.io/office-file-viewer/docs#quick-start), [`OfficeFileViewer` API](https://gyxing.github.io/office-file-viewer/docs#component-api), and [advanced parsing API](https://gyxing.github.io/office-file-viewer/docs#advanced-api) for URI sources, full-document search, font fallback, controlled view state, Worker modes, resource limits, low-level sessions, and cleanup.

## Supported formats

| Category   | Extensions                                | Main capabilities                                                                                              |
| ---------- | ----------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| Word       | `.doc`, `.docx`, `.docm`, `.dotx`, `.wps` | Body content, tables, graphics, links, outlines, comments, revisions, footnotes, and endnotes                  |
| Excel      | `.xls`, `.xlsx`, `.xlsm`, `.xltx`         | Worksheets, styles, graphics, links, frozen panes, tables/filters, comments, and common conditional formatting |
| PowerPoint | `.ppt`, `.pptx`, `.pptm`, `.potx`         | Slides, masters, graphics, links, comments, notes, audio/video, and common slide transitions                   |

Macro-enabled files expose only visible document content. Macros are never loaded or executed, and hosts receive a `MACRO_CONTENT_IGNORED` warning. Coverage does not guarantee complete restoration of every Office version, vendor extension, embedded object, animation, or complex layout.

## Limitations

- The viewer is read-only and does not edit, save, convert, print-layout, or export Office files.
- Remote files remain subject to browser CORS, authentication, and Content Security Policy rules.
- Internal optimization thresholds never reject large files. Very large or complex files automatically use on-demand reads and virtual rendering, but can still consume significant memory or briefly reduce responsiveness.
- The package does not bundle Office fonts; final layout depends on fonts available to the browser or fallback/host font resources configured by the host. URL resources still follow browser CORS/CSP rules.
- Review, filtering, media, and transition support is read-only. The viewer does not write comments, execute filters, macros, ActiveX, OLE, or object-level animations.

Read the [full performance, security, and rendering boundaries](https://gyxing.github.io/office-file-viewer/docs#limitations).

## Local development

```bash
yarn
yarn start
```

Build the library and documentation:

```bash
yarn build
yarn docs:build
```

Run the complete project checks:

```bash
yarn run check
```

## License

[MIT](./LICENSE)
