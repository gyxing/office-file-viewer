import React from 'react';
import type { PlaygroundConfig } from './playground-config';
import type { PlaygroundContent } from './playground-content';

type PlaygroundControlPanelProps = {
  /** 当前所有可调参数。 */
  config: PlaygroundConfig;
  /** 当前语言的控件文案。 */
  content: PlaygroundContent;
  /** 合并单次参数修改。 */
  onChange(patch: Partial<PlaygroundConfig>): void;
  /** 恢复在线体验页默认参数。 */
  onReset(): void;
};

type PlaygroundToggleProps = {
  /** 开关旁显示的文字。 */
  label: string;
  /** 当前开关状态。 */
  checked: boolean;
  /** 用户切换后的回调。 */
  onChange(checked: boolean): void;
};

/** 使用原生复选框提供无依赖且可键盘操作的布尔参数控件。 */
function PlaygroundToggle({ label, checked, onChange }: PlaygroundToggleProps) {
  return (
    <label className="office-viewer-playground-toggle">
      <input
        type="checkbox"
        checked={checked}
        onChange={(event) => onChange(event.target.checked)}
      />
      <span aria-hidden="true" />
      <strong>{label}</strong>
    </label>
  );
}

type PlaygroundColorFieldProps = {
  /** 颜色参数名称。 */
  label: string;
  /** 当前十六进制颜色值。 */
  value: string;
  /** 用户选色后的回调。 */
  onChange(value: string): void;
};

/** 同时展示可视选色器和精确颜色值。 */
function PlaygroundColorField({
  label,
  value,
  onChange,
}: PlaygroundColorFieldProps) {
  return (
    <label className="office-viewer-playground-field">
      <span>{label}</span>
      <span className="office-viewer-playground-color-control">
        <input
          type="color"
          value={value}
          onChange={(event) => onChange(event.target.value)}
        />
        <code>{value}</code>
      </span>
    </label>
  );
}

/** 渲染在线体验页的精简参数集合，修改后由父级立即应用。 */
export function PlaygroundControlPanel({
  config,
  content,
  onChange,
  onReset,
}: PlaygroundControlPanelProps) {
  return (
    <aside
      className="office-viewer-playground-controls"
      aria-labelledby="playground-controls-title"
    >
      <div className="office-viewer-playground-controls-header">
        <h2 id="playground-controls-title">{content.controlsTitle}</h2>
        <button type="button" onClick={onReset}>
          {content.reset}
        </button>
      </div>

      <fieldset>
        <legend>{content.component}</legend>
        <div
          className="office-viewer-playground-segmented"
          role="group"
          aria-label={content.component}
        >
          <button
            type="button"
            aria-pressed={config.target === 'viewer'}
            onClick={() => onChange({ target: 'viewer' })}
          >
            {content.targetViewer}
          </button>
          <button
            type="button"
            aria-pressed={config.target === 'layout'}
            onClick={() => onChange({ target: 'layout' })}
          >
            {content.targetLayout}
          </button>
        </div>
      </fieldset>

      <fieldset>
        <legend>{content.appearance}</legend>
        <label className="office-viewer-playground-field">
          <span>{content.themeMode}</span>
          <select
            value={config.themeMode}
            onChange={(event) => {
              const themeMode = event.target
                .value as PlaygroundConfig['themeMode'];
              // 同步一组可辨识的工作区底色，避免深色主题仍被浅色自定义值覆盖。
              onChange({
                themeMode,
                workspaceColor: themeMode === 'dark' ? '#111827' : '#eef1f6',
              });
            }}
          >
            <option value="light">{content.themeLight}</option>
            <option value="dark">{content.themeDark}</option>
            <option value="system">{content.themeSystem}</option>
          </select>
        </label>
        <div className="office-viewer-playground-field-grid">
          <PlaygroundColorField
            label={content.primaryColor}
            value={config.primaryColor}
            onChange={(primaryColor) => onChange({ primaryColor })}
          />
          <PlaygroundColorField
            label={content.workspaceColor}
            value={config.workspaceColor}
            onChange={(workspaceColor) => onChange({ workspaceColor })}
          />
        </div>
      </fieldset>

      <fieldset>
        <legend>{content.watermark}</legend>
        <PlaygroundToggle
          label={content.watermarkEnabled}
          checked={config.watermarkEnabled}
          onChange={(watermarkEnabled) => onChange({ watermarkEnabled })}
        />
        {config.watermarkEnabled ? (
          <div className="office-viewer-playground-nested-controls">
            <label className="office-viewer-playground-field">
              <span>{content.watermarkContent}</span>
              <input
                type="text"
                value={config.watermarkContent}
                onChange={(event) =>
                  onChange({ watermarkContent: event.target.value })
                }
              />
            </label>
            <PlaygroundColorField
              label={content.watermarkColor}
              value={config.watermarkColor}
              onChange={(watermarkColor) => onChange({ watermarkColor })}
            />
            <div className="office-viewer-playground-field-grid">
              <label className="office-viewer-playground-field">
                <span>{content.watermarkOpacity}</span>
                <input
                  type="number"
                  min="0.02"
                  max="1"
                  step="0.02"
                  value={config.watermarkOpacity}
                  onChange={(event) =>
                    onChange({
                      watermarkOpacity: Math.min(
                        1,
                        Math.max(0.02, Number(event.target.value) || 0.02),
                      ),
                    })
                  }
                />
              </label>
              <label className="office-viewer-playground-field">
                <span>{content.watermarkRotate}</span>
                <span className="office-viewer-playground-input-suffix">
                  <input
                    type="number"
                    min="-90"
                    max="90"
                    step="1"
                    value={config.watermarkRotate}
                    onChange={(event) =>
                      onChange({
                        watermarkRotate: Math.min(
                          90,
                          Math.max(-90, Number(event.target.value) || 0),
                        ),
                      })
                    }
                  />
                  <span>°</span>
                </span>
              </label>
            </div>
          </div>
        ) : null}
      </fieldset>

      <fieldset>
        <legend>{content.toolbar}</legend>
        <label className="office-viewer-playground-field">
          <span>{content.toolbarMode}</span>
          <select
            value={config.toolbarMode}
            onChange={(event) =>
              onChange({
                toolbarMode: event.target
                  .value as PlaygroundConfig['toolbarMode'],
              })
            }
          >
            <option value="default">{content.toolbarDefault}</option>
            <option value="custom">{content.toolbarCustom}</option>
            <option value="hidden">{content.toolbarHidden}</option>
          </select>
        </label>
        {config.toolbarMode === 'custom' ? (
          <div className="office-viewer-playground-toggle-grid">
            <PlaygroundToggle
              label={content.toolbarFileName}
              checked={config.toolbarFileName}
              onChange={(toolbarFileName) => onChange({ toolbarFileName })}
            />
            <PlaygroundToggle
              label={content.toolbarOpenFile}
              checked={config.toolbarOpenFile}
              onChange={(toolbarOpenFile) => onChange({ toolbarOpenFile })}
            />
            {config.target === 'viewer' ? (
              <>
                <PlaygroundToggle
                  label={content.toolbarWordOutline}
                  checked={config.toolbarWordOutline}
                  onChange={(toolbarWordOutline) =>
                    onChange({ toolbarWordOutline })
                  }
                />
                <PlaygroundToggle
                  label={content.toolbarWordRevisionMode}
                  checked={config.toolbarWordRevisionMode}
                  onChange={(toolbarWordRevisionMode) =>
                    onChange({ toolbarWordRevisionMode })
                  }
                />
                <PlaygroundToggle
                  label={content.toolbarSpreadsheetViewMode}
                  checked={config.toolbarSpreadsheetViewMode}
                  onChange={(toolbarSpreadsheetViewMode) =>
                    onChange({ toolbarSpreadsheetViewMode })
                  }
                />
                <PlaygroundToggle
                  label={content.toolbarPresentationNavigation}
                  checked={config.toolbarPresentationNavigation}
                  onChange={(toolbarPresentationNavigation) =>
                    onChange({ toolbarPresentationNavigation })
                  }
                />
                <PlaygroundToggle
                  label={content.toolbarSpeakerNotes}
                  checked={config.toolbarSpeakerNotes}
                  onChange={(toolbarSpeakerNotes) =>
                    onChange({ toolbarSpeakerNotes })
                  }
                />
                <PlaygroundToggle
                  label={content.toolbarSearch}
                  checked={config.toolbarSearch}
                  onChange={(toolbarSearch) => onChange({ toolbarSearch })}
                />
                <PlaygroundToggle
                  label={content.toolbarReview}
                  checked={config.toolbarReview}
                  onChange={(toolbarReview) => onChange({ toolbarReview })}
                />
              </>
            ) : null}
            <PlaygroundToggle
              label={content.toolbarZoom}
              checked={config.toolbarZoom}
              onChange={(toolbarZoom) => onChange({ toolbarZoom })}
            />
            <PlaygroundToggle
              label={content.toolbarFullscreen}
              checked={config.toolbarFullscreen}
              onChange={(toolbarFullscreen) => onChange({ toolbarFullscreen })}
            />
          </div>
        ) : null}
      </fieldset>

      <fieldset>
        <legend>{content.preview}</legend>
        <div className="office-viewer-playground-field-grid">
          <label className="office-viewer-playground-field">
            <span>{content.zoom}</span>
            <span className="office-viewer-playground-input-suffix">
              <input
                type="number"
                min="20"
                max="400"
                step="10"
                value={config.zoom}
                onChange={(event) =>
                  onChange({
                    zoom: Math.min(
                      400,
                      Math.max(20, Number(event.target.value) || 100),
                    ),
                  })
                }
              />
              <span>%</span>
            </span>
          </label>
          <label className="office-viewer-playground-field">
            <span>{content.previewHeight}</span>
            <span className="office-viewer-playground-input-suffix">
              <input
                type="number"
                min="420"
                max="1000"
                step="20"
                value={config.previewHeight}
                onChange={(event) =>
                  onChange({
                    previewHeight: Math.min(
                      1000,
                      Math.max(420, Number(event.target.value) || 640),
                    ),
                  })
                }
              />
              <span>px</span>
            </span>
          </label>
        </div>

        {config.target === 'viewer' ? (
          <div className="office-viewer-playground-toggle-grid">
            <PlaygroundToggle
              label={content.search}
              checked={config.searchEnabled}
              onChange={(searchEnabled) => onChange({ searchEnabled })}
            />
            <PlaygroundToggle
              label={content.review}
              checked={config.reviewEnabled}
              onChange={(reviewEnabled) => onChange({ reviewEnabled })}
            />
            <PlaygroundToggle
              label={content.imagePreview}
              checked={config.imagePreviewEnabled}
              onChange={(imagePreviewEnabled) =>
                onChange({ imagePreviewEnabled })
              }
            />
          </div>
        ) : (
          <div className="office-viewer-playground-nested-controls">
            <PlaygroundToggle
              label={content.controlledZoom}
              checked={config.layoutControlledZoom}
              onChange={(layoutControlledZoom) =>
                onChange({ layoutControlledZoom })
              }
            />
            <label className="office-viewer-playground-field">
              <span>{content.contentScaling}</span>
              <select
                value={config.layoutContentScaling}
                onChange={(event) =>
                  onChange({
                    layoutContentScaling: event.target
                      .value as PlaygroundConfig['layoutContentScaling'],
                  })
                }
              >
                <option value="managed">{content.managedScaling}</option>
                <option value="manual">{content.manualScaling}</option>
              </select>
            </label>
          </div>
        )}
      </fieldset>
    </aside>
  );
}
