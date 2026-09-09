import { Button } from 'antd';
import {
  OfficeViewerLayout,
  useOfficeViewerLayout,
  type OfficeFileViewerToolbarOptions,
  type OfficeViewerLayoutContentScaling,
  type OfficeViewerThemeOptions,
  type OfficeViewerWatermark,
} from 'office-file-viewer';
import type { CSSProperties, ReactElement } from 'react';
import React, { useCallback, useState } from 'react';
import './smoke-test.layout.less';

/** 自定义内容外壳烟测视图属性。 */
export type OfficeLayoutSmokePreviewProps = {
  /** 工具栏显示的宿主文件名。 */
  fileName: string;
  /** 外壳预览高度。 */
  height?: CSSProperties['height'];
  /** 初始缩放百分比。 */
  defaultZoom: number;
  /** 是否由烟测页控制缩放状态。 */
  controlledZoom: boolean;
  /** 缩放由外壳还是宿主内容负责。 */
  contentScaling: OfficeViewerLayoutContentScaling;
  /** 工具栏显示配置。 */
  toolbar?: false | OfficeFileViewerToolbarOptions;
  /** 是否追加使用外壳上下文的调试操作。 */
  toolbarExtra: boolean;
  /** 外壳主题配置。 */
  theme: OfficeViewerThemeOptions;
  /** 内容区水印配置。 */
  watermark: OfficeViewerWatermark;
  /** 选择本地文件后的回调。 */
  onFileSelect(file: File): void;
  /** 缩放变化回调。 */
  onZoomChange(zoom: number): void;
  /** 全屏状态变化回调。 */
  onFullscreenChange(fullscreen: boolean): void;
  /** 全屏请求失败回调。 */
  onFullscreenError(error: Error): void;
  /** 点击追加工具栏操作后的回调。 */
  onToolbarExtra(): void;
};

type LayoutToolbarExtraProps = {
  /** 点击复位操作后的回调。 */
  onAction(): void;
};

/** 验证 toolbarExtra 内部可以直接消费外壳动作。 */
function LayoutToolbarExtra({ onAction }: LayoutToolbarExtraProps) {
  const { actions } = useOfficeViewerLayout();

  return (
    <Button
      size="small"
      onClick={() => {
        actions.changeZoom(100);
        onAction();
      }}
    >
      复位缩放
    </Button>
  );
}

type LayoutSmokeDocumentProps = {
  /** 缩放由外壳还是调试内容负责。 */
  contentScaling: OfficeViewerLayoutContentScaling;
};

/** 渲染固定宿主内容，并展示 useOfficeViewerLayout 的实时状态。 */
function LayoutSmokeDocument({
  contentScaling,
}: LayoutSmokeDocumentProps): ReactElement {
  const { state, actions, meta } = useOfficeViewerLayout();
  const manualScale = state.zoom / 100;
  const manualStyle: CSSProperties | undefined =
    contentScaling === 'manual'
      ? {
          transform: `scale(${manualScale})`,
          transformOrigin: 'top left',
          width: `${100 / manualScale}%`,
        }
      : undefined;

  return (
    <div className="office-layout-smoke-stage" style={manualStyle}>
      <article className="office-layout-smoke-document">
        <header className="office-layout-smoke-document-header">
          <div>
            <p className="office-layout-smoke-eyebrow">HOST CONTENT</p>
            <h2>宿主业务文档预览</h2>
            <p>
              这部分内容完全由宿主提供，外壳只统一工具栏、缩放、全屏、主题与水印。
            </p>
          </div>
          <span className="office-layout-smoke-version">Layout API</span>
        </header>

        <section className="office-layout-smoke-runtime" aria-live="polite">
          <div>
            <span>当前缩放</span>
            <strong>{state.zoom}%</strong>
          </div>
          <div>
            <span>全屏状态</span>
            <strong>{state.isFullscreen ? '已全屏' : '普通视图'}</strong>
          </div>
          <div>
            <span>缩放职责</span>
            <strong>{meta.contentScaling}</strong>
          </div>
          <div className="office-layout-smoke-runtime-actions">
            <Button size="small" onClick={actions.zoomOut}>
              内容内缩小
            </Button>
            <Button size="small" onClick={actions.zoomIn}>
              内容内放大
            </Button>
          </div>
        </section>

        <section className="office-layout-smoke-summary">
          <div className="office-layout-smoke-summary-card">
            <span>统一交互</span>
            <strong>缩放 · 全屏 · 打开文件</strong>
            <p>宿主内容不必重复实现查看器的公共操作。</p>
          </div>
          <div className="office-layout-smoke-summary-card">
            <span>统一外观</span>
            <strong>主题 · 水印 · 工具栏</strong>
            <p>参数修改后可在当前页面即时对比效果。</p>
          </div>
          <div className="office-layout-smoke-summary-card">
            <span>开放能力</span>
            <strong>状态 · 动作 · 环境信息</strong>
            <p>子内容可通过 Hook 接入外壳控制能力。</p>
          </div>
        </section>

        <section className="office-layout-smoke-section">
          <h3>接入能力检查</h3>
          <table>
            <thead>
              <tr>
                <th>验证项</th>
                <th>当前配置</th>
                <th>预期表现</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>内容缩放</td>
                <td>{contentScaling}</td>
                <td>工具栏与内容内操作同步更新</td>
              </tr>
              <tr>
                <td>全屏能力</td>
                <td>
                  {meta.fullscreenSupported ? '浏览器支持' : '浏览器不支持'}
                </td>
                <td>全屏前后主题与水印保持一致</td>
              </tr>
              <tr>
                <td>宿主内容</td>
                <td>普通 React DOM</td>
                <td>不依赖任何 Office 解析模型</td>
              </tr>
            </tbody>
          </table>
        </section>

        <section className="office-layout-smoke-section office-layout-smoke-checklist">
          <h3>人工验证建议</h3>
          <ol>
            <li>切换浅色、深色和跟随系统，检查工具栏与工作区颜色。</li>
            <li>启用水印并调整颜色、字号和间距，检查内容区覆盖效果。</li>
            <li>切换托管/手动缩放与受控/非受控状态，检查缩放连续性。</li>
            <li>进入全屏后操作缩放，再退出全屏并检查状态与事件日志。</li>
          </ol>
        </section>
      </article>
    </div>
  );
}

/** 使用正式 OfficeViewerLayout 组合宿主调试内容。 */
export function OfficeLayoutSmokePreview({
  fileName,
  height,
  defaultZoom,
  controlledZoom,
  contentScaling,
  toolbar,
  toolbarExtra,
  theme,
  watermark,
  onFileSelect,
  onZoomChange,
  onFullscreenChange,
  onFullscreenError,
  onToolbarExtra,
}: OfficeLayoutSmokePreviewProps): ReactElement {
  const [zoom, setZoom] = useState(defaultZoom);
  const handleZoomChange = useCallback(
    (nextZoom: number) => {
      if (controlledZoom) setZoom(nextZoom);
      onZoomChange(nextZoom);
    },
    [controlledZoom, onZoomChange],
  );

  return (
    <OfficeViewerLayout
      className="office-layout-smoke-viewer"
      fileName={fileName}
      height={height}
      defaultZoom={defaultZoom}
      zoom={controlledZoom ? zoom : undefined}
      contentScaling={contentScaling}
      toolbar={toolbar}
      toolbarExtra={
        toolbarExtra ? (
          <LayoutToolbarExtra onAction={onToolbarExtra} />
        ) : undefined
      }
      theme={theme}
      watermark={watermark}
      onFileSelect={onFileSelect}
      onZoomChange={handleZoomChange}
      onFullscreenChange={onFullscreenChange}
      onFullscreenError={onFullscreenError}
    >
      <LayoutSmokeDocument contentScaling={contentScaling} />
    </OfficeViewerLayout>
  );
}
