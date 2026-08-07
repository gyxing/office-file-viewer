# Office File Viewer

[English](./README.md) | 简体中文 | <a href="https://gyxing.github.io/office-file-viewer/zh-CN/" target="_blank" rel="noopener noreferrer">在线示例</a> | [完整文档](https://gyxing.github.io/office-file-viewer/zh-CN/docs)

> 一个面向 React 的纯浏览器 Office 文件预览组件，支持 DOC/DOCX/WPS、XLS/XLSX 和 PPT/PPTX。

`office-file-viewer` 在浏览器内完成 Office 文件下载、解析和渲染，不需要配套的文档转换服务，也不会主动上传本地文件。

组件使用统一界面预览 Word 文档、Excel 表格和 PowerPoint 演示文稿，包含文件选择、远程加载、解析进度、缩放、全屏、文档大纲、电子表格显示模式、工作表标签、幻灯片导航和演讲者备注。

> 这是独立实现的解析与渲染引擎，并非 Microsoft Office 或 WPS Office 的原生排版引擎。复杂文档可能与桌面应用存在差异，使用前请阅读[完整限制说明](https://gyxing.github.io/office-file-viewer/zh-CN/docs#limitations)。

## 主要特性

- **纯浏览器解析**：适用于内网、离线环境和隐私敏感场景。
- **七种文件格式**：支持 DOC、DOCX、WPS、XLS、XLSX、PPT 和 PPTX。
- **统一 React 组件**：各格式共用加载、错误、空状态、缩放和全屏交互。
- **多种文件来源**：接受本地 `File`、远程 URL 或异步加载函数。
- **渐进预览**：受支持格式可以在解析继续进行时提前展示已完成内容。
- **Worker 支持**：旧格式 DOC/WPS、XLS 和 PPT 可以在 Web Worker 中解析。
- **资源管理**：组件负责取消任务、订阅、Worker 和 Blob URL 的生命周期。
- **内置预览界面**：提供作用域隔离的文件选择、导航、缩放和全屏控件及样式。

## 安装

```bash
npm install office-file-viewer
```

使用 Yarn：

```bash
yarn add office-file-viewer
```

`react` 和 `react-dom` 是由宿主提供的 peer dependency。当前包仅发布 ESM，组件样式已构建为作用域隔离的 CSS，公共 API 只能从 `office-file-viewer` 根入口导入。

## 版本兼容

| 项目     | 要求       | 说明                          |
| -------- | ---------- | ----------------------------- |
| React    | `>=16.9.0` | 支持 Hooks 的 React 版本      |
| ReactDOM | `>=16.9.0` | 建议与 React 保持相同主版本   |
| 模块格式 | 仅 ESM     | 使用支持 ESM 的浏览器构建工具 |

当前 peer dependency 范围：

```text
react: >=16.9.0
react-dom: >=16.9.0
```

## 快速接入

不传 `uri` 时，组件显示内置文件选择入口：

```tsx
import { OfficeFileViewer } from 'office-file-viewer';

export default function OfficePreview() {
  return <OfficeFileViewer height="80vh" />;
}
```

查看[完整快速接入](https://gyxing.github.io/office-file-viewer/zh-CN/docs#quick-start)、[`OfficeFileViewer` API](https://gyxing.github.io/office-file-viewer/zh-CN/docs#component-api) 和[高级解析 API](https://gyxing.github.io/office-file-viewer/zh-CN/docs#advanced-api)，了解 URI 来源、回调、Worker 模式、底层会话和资源释放。

## 支持格式

| 分类       | 扩展名                  | 主要能力                                                                  |
| ---------- | ----------------------- | ------------------------------------------------------------------------- |
| Word       | `.doc`、`.docx`、`.wps` | 文本、格式、列表、表格、图片、图表、形状、链接和文档大纲                  |
| Excel      | `.xls`、`.xlsx`         | 工作表、值、样式、合并单元格、尺寸、图片、图表、工作表标签和原始/阅读模式 |
| PowerPoint | `.ppt`、`.pptx`         | 幻灯片、母版、文本、形状、图片、表格、图表、导航和演讲者备注              |

支持范围不代表可以完整还原所有 Office 版本、厂商扩展、宏、嵌入对象、动画或复杂布局。

## 限制说明

- 预览器为只读组件，不提供编辑、保存、格式转换、打印排版或文件导出。
- 远程文件仍受浏览器 CORS、身份认证和内容安全策略约束。
- 超大或复杂文件可能占用较多内存或短暂降低响应速度；宿主应校验不可信文件的大小、类型和来源。

请阅读完整的[性能、安全与渲染边界](https://gyxing.github.io/office-file-viewer/zh-CN/docs#limitations)。

## 本地开发

```bash
yarn
yarn start
```

构建组件库和文档：

```bash
yarn build
yarn docs:build
```

执行项目完整检查：

```bash
yarn run check
```

## License

[MIT](./LICENSE)
