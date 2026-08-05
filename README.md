# Office File Viewer

English | [简体中文](./README.zh-CN.md) | <a href="https://gyxing.github.io/office-file-viewer/" target="_blank" rel="noopener noreferrer">Live Demo</a> | [Documentation](https://gyxing.github.io/office-file-viewer/docs)

> A browser-based React component for offline preview of DOC/DOCX/WPS, XLS/XLSX, and PPT/PPTX files.

`office-file-viewer` downloads, parses, and renders Office files entirely in the browser. It requires no companion document-conversion service and does not actively upload local files.

The package provides a unified viewer for Word documents, Excel spreadsheets, and PowerPoint presentations, including file selection, remote loading, parsing progress, zoom, fullscreen mode, document outlines, worksheet tabs, slide navigation, and speaker notes.

> This is an independent parsing and rendering implementation, not the native Microsoft Office or WPS Office layout engine. Complex documents can render differently from desktop applications. Review the [complete limitations](https://gyxing.github.io/office-file-viewer/docs#limitations) before use.

## Features

- **Browser-only parsing**: Suitable for intranets, offline environments, and privacy-sensitive workflows.
- **Seven formats**: DOC, DOCX, WPS, XLS, XLSX, PPT, and PPTX.
- **Unified React component**: Shared loading, error, empty, zoom, and fullscreen interactions.
- **Flexible sources**: Accepts a local `File`, remote URL, or async loader.
- **Progressive preview**: Supported formats can render completed content while parsing continues.
- **Worker support**: Legacy DOC/WPS, XLS, and PPT parsing can run in a Web Worker.
- **Resource management**: The component handles cancellation, subscriptions, Workers, and Blob URLs.
- **Host integration**: Supports antd v4, v5, and v6 and inherits the host `ConfigProvider`.

## Installation

```bash
npm install office-file-viewer antd react react-dom
```

With Yarn:

```bash
yarn add office-file-viewer antd react react-dom
```

`react`, `react-dom`, and `antd` are peer dependencies supplied by the host. The package is ESM-only, component styles are built as CSS, and public APIs should be imported only from `office-file-viewer`.

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

When using antd v4, load its stylesheet in the host entry:

```tsx
import 'antd/dist/antd.css';
```

antd v5 and v6 do not require that stylesheet.

## Quick start

When `uri` is omitted, the component shows its built-in file picker:

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer height="80vh" />;
}
```

See the [complete quick start](https://gyxing.github.io/office-file-viewer/docs#quick-start), [`OfficeFileViewer` API](https://gyxing.github.io/office-file-viewer/docs#component-api), and [advanced parsing API](https://gyxing.github.io/office-file-viewer/docs#advanced-api) for URI sources, callbacks, Worker modes, low-level sessions, and cleanup.

## Supported formats

| Category   | Extensions              | Main capabilities                                                                        |
| ---------- | ----------------------- | ---------------------------------------------------------------------------------------- |
| Word       | `.doc`, `.docx`, `.wps` | Text, formatting, lists, tables, images, charts, shapes, links, and document outlines    |
| Excel      | `.xls`, `.xlsx`         | Worksheets, values, styles, merged cells, dimensions, images, charts, and worksheet tabs |
| PowerPoint | `.ppt`, `.pptx`         | Slides, masters, text, shapes, images, tables, charts, navigation, and speaker notes     |

Coverage does not guarantee complete restoration of every Office version, vendor extension, macro, embedded object, animation, or complex layout.

## Limitations

- The viewer is read-only and does not edit, save, convert, print-layout, or export Office files.
- Remote files remain subject to browser CORS, authentication, and Content Security Policy rules.
- Large or complex files can use significant memory or briefly reduce responsiveness; hosts should validate untrusted file sizes, types, and sources.

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
