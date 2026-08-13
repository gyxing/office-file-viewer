# Office File Viewer

English | [简体中文](./README.zh-CN.md) | <a href="https://gyxing.github.io/office-file-viewer/" target="_blank" rel="noopener noreferrer">Live Demo</a> | [Documentation](https://gyxing.github.io/office-file-viewer/docs)

> A browser-based React component for offline preview of DOC/DOCX/WPS, XLS/XLSX, and PPT/PPTX files.

`office-file-viewer` downloads, parses, and renders Office files entirely in the browser. It requires no companion document-conversion service and does not actively upload local files.

The package provides a unified viewer for Word documents, Excel spreadsheets, and PowerPoint presentations, including file selection, remote loading, parsing progress, full-document search, zoom, fullscreen mode, content-image preview and download, document outlines, spreadsheet display modes, worksheet tabs, slide navigation, and speaker notes.

> This is an independent parsing and rendering implementation, not the native Microsoft Office or WPS Office layout engine. Complex documents can render differently from desktop applications. Review the [complete limitations](https://gyxing.github.io/office-file-viewer/docs#limitations) before use.

## Features

- **Browser-only parsing**: Suitable for intranets, offline environments, and privacy-sensitive workflows.
- **Seven formats**: DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX.
- **Unified React component**: Shared loading, error, empty, zoom, and fullscreen interactions.
- **Flexible sources**: Accepts a local `File`, remote URL, or async loader with a cancellation signal.
- **Progressive preview**: Supported formats can render completed content while parsing continues.
- **Full-document search**: All seven formats support cancellable incremental search, result navigation, case matching, and whole-word matching.
- **Controlled view state**: Hosts can control zoom, active pages, worksheets, sidebars, and display modes per field.
- **Content-image actions**: DOC/DOCX/WPS and XLS/XLSX support double-click preview plus preview and download actions from a custom context menu.
- **Source hyperlinks**: Text, cell, image, shape, and action-button links use safe `Ctrl`/`Command` modifier activation and can be intercepted by the host.
- **Adaptive Workers**: All seven formats support Workers. The default mode selects a complete model or an on-demand source from the file profile and safely falls back if Worker startup fails.
- **Font fallback**: Source fonts, aliases, and fallback chains are resolved consistently, with optional structured warnings for fonts missing in the current browser.
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

`react` and `react-dom` are peer dependencies supplied by the host. The package is ESM-only, component styles are built as scoped CSS, and public APIs should be imported only from `office-file-viewer`.

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

| Category   | Extensions              | Main capabilities                                                                                                       |
| ---------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Word       | `.doc`, `.docx`, `.wps` | Text, formatting, lists, tables, images, charts, shapes, links, and document outlines                                   |
| Excel      | `.xls`, `.xlsx`         | Worksheets, values, styles, merged cells, dimensions, images, charts, links, worksheet tabs, and original/reading modes |
| PowerPoint | `.ppt`, `.pptx`         | Slides, masters, text, shapes, images, tables, charts, links, navigation, and speaker notes                             |

Coverage does not guarantee complete restoration of every Office version, vendor extension, macro, embedded object, animation, or complex layout.

## Limitations

- The viewer is read-only and does not edit, save, convert, print-layout, or export Office files.
- Remote files remain subject to browser CORS, authentication, and Content Security Policy rules.
- Internal optimization thresholds never reject large files. Very large or complex files automatically use on-demand reads and virtual rendering, but can still consume significant memory or briefly reduce responsiveness.
- The package does not bundle Office fonts; final layout depends on fonts available to the browser or fallback fonts configured by the host.

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
