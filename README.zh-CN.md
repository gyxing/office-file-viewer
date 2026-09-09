# Office File Viewer

[English](./README.md) | 简体中文 | <a href="https://gyxing.github.io/office-file-viewer/zh-CN/" target="_blank" rel="noopener noreferrer">在线示例</a> | [完整文档](https://gyxing.github.io/office-file-viewer/zh-CN/docs)

[![npm 版本](https://img.shields.io/npm/v/office-file-viewer.svg)](https://www.npmjs.com/package/office-file-viewer) [![npm 下载量](https://img.shields.io/npm/dm/office-file-viewer.svg)](https://www.npmjs.com/package/office-file-viewer) [![CI](https://github.com/gyxing/office-file-viewer/actions/workflows/validate.yml/badge.svg)](https://github.com/gyxing/office-file-viewer/actions/workflows/validate.yml) [![许可证](https://img.shields.io/npm/l/office-file-viewer.svg)](./LICENSE)

> 一个面向 React 的纯浏览器 Office 文件预览组件，支持 DOC/DOCX/DOCM/DOTX/WPS、XLS/XLSX/XLSM/XLTX 和 PPT/PPTX/PPTM/POTX。

`office-file-viewer` 在浏览器内完成 Office 文件下载、解析和渲染，不需要配套的文档转换服务，也不会主动上传本地文件。

组件使用统一界面预览 Word 文档、Excel 表格和 PowerPoint 演示文稿，包含全文查找、只读审阅、脚注/尾注、电子表格业务语义、安全媒体播放、页级切换、缩放、全屏和演讲者备注。

> 这是独立实现的解析与渲染引擎，并非 Microsoft Office 或 WPS Office 的原生排版引擎。复杂文档可能与桌面应用存在差异，使用前请阅读[完整限制说明](https://gyxing.github.io/office-file-viewer/zh-CN/docs#limitations)。

## 项目背景与价值

企业协作、知识库和文件管理系统经常需要预览 Office 文档。传统方案通常依赖服务端转码或第三方服务，会增加部署成本，也可能带来额外的文件隐私和数据合规问题。

`office-file-viewer` 采用浏览器优先的方案：本地文件可以在浏览器内完成解析和渲染，不需要部署额外的文档转换服务，也不会主动上传文件。项目希望成为可复用的开源基础组件，并明确说明宏、外部媒体、远程资源、大文件和复杂布局等安全与渲染边界。

## 稳定公共入口

包继续保留默认 `OfficeFileViewer` 入口，并增加稳定的 ESM 子路径：

```ts
import { OfficeFileViewer } from 'office-file-viewer';
import { OfficeViewerShell } from 'office-file-viewer/layout';
import { createOfficeParseSession } from 'office-file-viewer/core';
import { createOfficePluginRegistry } from 'office-file-viewer/plugins';
import { OfficeViewerProvider } from 'office-file-viewer/plugins';
import { exportOriginalOfficeFile } from 'office-file-viewer/export';
import 'office-file-viewer/styles.css';

const registry = createOfficePluginRegistry({ includeBuiltIns: true });
```

`core` 提供解析、快照、资源、生命周期和能力描述；`plugins` 提供实例级注册表及可选的 `OfficeViewerProvider`；`export` 提供原始文件导出和导出器契约；`layout` 提供可复用组合式 Shell。首版 `plugins`、`export` 和 Editor 预留契约标记为实验性。不要导入未文档化的 `services`、`formats`、`shared` 或 `dist` 深层路径。

```tsx
<OfficeViewerProvider registry={registry}>
  <OfficeFileViewer uri={file} />
</OfficeViewerProvider>
```

`registry` 应在应用作用域内复用，并在应用销毁时调用 `registry.dispose()`。

如果宿主构建器已经处理根入口的 CSS 副作用，不要再次导入 `styles.css`，避免重复规则。通过 `style` 传入的 CSS 变量优先于主题选项。

## 主要特性

- **纯浏览器解析**：适用于内网、离线环境和隐私敏感场景。
- **三类 Office 文件**：支持 Word、Excel、PowerPoint 的 13 种常用扩展名。
- **统一 React 组件**：各格式共用加载、错误、空状态、缩放和全屏交互。
- **主题与水印**：内置浅色、深色和跟随系统模式，可覆盖语义令牌，并为文档内容区添加高性能文字水印。
- **可复用预览外壳**：宿主自定义内容也可复用相同工具栏、缩放、全屏、主题和水印能力。
- **多种文件来源**：接受本地 `File`、远程 URL 或支持取消信号的异步加载函数。
- **渐进预览**：受支持格式可以在解析继续进行时提前展示已完成内容。
- **全文查找**：全部支持格式均提供可取消的增量查找、结果导航、大小写匹配和全词匹配。
- **可控视图状态**：可以按字段控制缩放、适应宽度/页面、当前页、工作表、侧栏和显示模式。
- **内容图片交互**：Word、Excel 和 PowerPoint 的独立内容图片支持双击预览，以及右键预览和下载；PowerPoint 根据资源类型、对象元数据和显示尺寸综合识别小图标，小图标不提供图片预览入口或自定义预览菜单，母版、版式和背景图片保持静态展示。
- **源文档超链接**：支持文字、单元格、图片、形状和按钮链接，以 `Ctrl`/`Command` 修饰单击安全激活，并支持宿主接管。
- **只读审阅**：Word 批注和修订采用原生风格页侧标记区，通过虚线直接连接正文，不创建固定审阅列表；DOCX 支持最终态、标记态和原始态，DOC/WPS 恢复可识别的审阅语义。Excel 和 PowerPoint 批注继续使用统一审阅面板。
- **Word 注释内容**：全部 Word 格式支持脚注、尾注、引用跳转和超长脚注续页。
- **Excel 业务语义**：支持冻结窗格、Table/AutoFilter、单元格批注，以及常用条件格式的静态还原。
- **PowerPoint 媒体与切换**：内嵌音视频按当前页加载且不自动播放；外部媒体默认阻止，常用页级切换可按源文件显式开启。
- **自适应 Worker**：全部支持格式均可复用对应解析类型的 Worker；默认模式会按文件画像自动选择完整模型或按需数据源，并在 Worker 无法启动时安全回退。
- **字体回退与宿主字体**：统一解析源字体、字体别名和回退链，可按需注册宿主字体资源，并在字体不可用时提供结构化警告。
- **结构化错误**：输入、下载、格式、加密文件、Worker、资源和解析失败使用稳定错误码与阶段。
- **资源管理**：组件负责取消任务、订阅、Worker 和 Blob URL 的生命周期，并支持宿主配置解析资源上限。
- **内置预览界面**：提供作用域隔离的文件选择、导航、缩放和全屏控件及样式。

## 安装

```bash
npm install office-file-viewer
```

使用 Yarn：

```bash
yarn add office-file-viewer
```

`react` 和 `react-dom` 是由宿主提供的 peer dependency。当前包仅发布 ESM，组件样式已构建为作用域隔离的 CSS。Office 预览 API 从 `office-file-viewer` 导入；只使用可复用外壳时，也可从稳定子路径 `office-file-viewer/layout` 导入。

## 版本兼容

| 项目     | 要求       | 说明                                           |
| -------- | ---------- | ---------------------------------------------- |
| React    | `>=16.9.0` | 支持 Hooks 的 React 版本                       |
| ReactDOM | `>=16.9.0` | 建议与 React 保持相同主版本                    |
| 模块格式 | 仅 ESM     | 使用支持 ESM 的浏览器构建工具                  |
| 浏览器   | 现代浏览器 | 建议 Chromium 100+、Firefox 115+、Safari 16.4+ |

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

查看[完整快速接入](https://gyxing.github.io/office-file-viewer/zh-CN/docs#quick-start)、[`OfficeFileViewer` API](https://gyxing.github.io/office-file-viewer/zh-CN/docs#component-api) 和[高级解析 API](https://gyxing.github.io/office-file-viewer/zh-CN/docs#advanced-api)，了解 URI 来源、全文查找、字体回退、受控视图状态、Worker 模式、资源限制、底层会话和资源释放。

## 支持格式

| 分类       | 扩展名                                    | 主要能力                                                           |
| ---------- | ----------------------------------------- | ------------------------------------------------------------------ |
| Word       | `.doc`、`.docx`、`.docm`、`.dotx`、`.wps` | 正文、表格、图形、链接、大纲、批注、修订、脚注与尾注               |
| Excel      | `.xls`、`.xlsx`、`.xlsm`、`.xltx`         | 工作表、样式、图形、链接、冻结窗格、Table/筛选、批注和常用条件格式 |
| PowerPoint | `.ppt`、`.pptx`、`.pptm`、`.potx`         | 幻灯片、母版、图形、链接、批注、备注、音视频与常用页级切换         |

宏启用文件只读取可见文档内容，不加载或执行宏，并通过 `MACRO_CONTENT_IGNORED` 警告通知宿主。支持范围不代表可以完整还原所有 Office 版本、厂商扩展、嵌入对象、动画或复杂布局。

## 限制说明

- Viewer 界面为只读组件，不提供编辑、保存、格式转换、打印排版或导出为 PDF/图片；原始文件复制通过实验性的 `office-file-viewer/export` 入口提供。
- 远程文件仍受浏览器 CORS、身份认证和内容安全策略约束。
- 组件不会因内部优化阈值拒绝大文件；超大或复杂文件会自动采用按需读取与虚拟渲染，但仍可能占用较多内存或短暂降低响应速度。
- 组件不捆绑 Office 字体；最终排版效果取决于当前浏览器可用字体或宿主配置的回退/宿主字体资源，URL 字体仍受浏览器 CORS/CSP 规则约束。
- 审阅、筛选、媒体和切换均为只读还原；不写回批注，不执行筛选、宏、ActiveX、OLE 或对象级动画。

请阅读完整的[性能、安全与渲染边界](https://gyxing.github.io/office-file-viewer/zh-CN/docs#limitations)。

## 社区与安全

- 提交 Pull Request 前，请阅读 [CONTRIBUTING.md](./CONTRIBUTING.md)。
- 社区参与请遵守 [CODE_OF_CONDUCT.md](./CODE_OF_CONDUCT.md)。
- 安全问题请按照 [SECURITY.md](./SECURITY.md) 中的说明反馈。

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
