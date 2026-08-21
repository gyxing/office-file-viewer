import type {
  ChangeEvent,
  KeyboardEvent,
  MouseEvent,
  ReactElement,
} from 'react';
import React, { useEffect, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import { OfficeButton } from '../shared/ui/OfficeButton';
import { ZoomInIcon, ZoomOutIcon } from '../shared/ui/OfficeIcons';
import { OfficeTooltip } from '../shared/ui/OfficeTooltip';
import {
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './constants';
import { normalizeOfficeZoom } from './normalizeOfficeZoom';
import type { OfficeFileViewerZoomMode } from './viewState';

/** 工具栏通用缩放能力。 */
export type ZoomControls = {
  /** 当前预览缩放比例。 */
  value: number;
  /** 当前固定比例或自适应缩放模式。 */
  mode: OfficeFileViewerZoomMode;
  /** 当前格式允许选择的自适应缩放模式。 */
  fitModes: readonly Exclude<OfficeFileViewerZoomMode, 'percentage'>[];
  /** 当前是否已有可执行增减操作的文档。 */
  hasDocument: boolean;
  /** 减小预览比例。 */
  zoomOut(): void;
  /** 放大预览比例。 */
  zoomIn(): void;
  /** 设置指定预览比例。 */
  change(value: number): void;
  /** 切换固定比例或自适应缩放模式。 */
  changeMode(mode: OfficeFileViewerZoomMode): void;
};

/** 缩放操作组件属性。 */
type ZoomControlProps = {
  /** 当前可用的缩放能力。 */
  controls: ZoomControls;
};

/** 缩放输入仅接受空值或十进制整数。 */
const ZOOM_INPUT_PATTERN = /^\d*$/;

let zoomControlSequence = 0;

/** 为 React 16 环境生成稳定的组合框列表标识。 */
function createZoomListId() {
  zoomControlSequence += 1;
  return `office-file-zoom-list-${zoomControlSequence}`;
}

/** 将有效整数限制到预览器支持的缩放范围。 */
function normalizeZoomInput(value: string): number | undefined {
  if (!value || !/^\d+$/.test(value)) return undefined;
  const parsedValue = Number(value);
  if (!Number.isFinite(parsedValue)) return undefined;
  return normalizeOfficeZoom(parsedValue);
}

/** 提供可输入、可选择并带固定百分号后缀的缩放操作。 */
export function ZoomControl({ controls }: ZoomControlProps): ReactElement {
  const messages = useOfficeFileViewerMessages();
  const inputRef = useRef<HTMLInputElement>(null);
  const [listId] = useState(createZoomListId);
  const [inputValue, setInputValue] = useState(String(controls.value));
  const [open, setOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const fitOptionCount = controls.fitModes.length;
  const optionCount = fitOptionCount + OFFICE_ZOOM_LEVELS.length;

  useEffect(() => {
    setInputValue(String(controls.value));
  }, [controls.value]);

  const closeOptions = () => {
    setOpen(false);
    setActiveIndex(-1);
  };

  const commitZoom = (nextInputValue = inputValue) => {
    const normalizedValue = normalizeZoomInput(nextInputValue);
    if (normalizedValue === undefined) {
      setInputValue(String(controls.value));
      return;
    }
    setInputValue(String(normalizedValue));
    controls.change(normalizedValue);
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const nextInputValue = event.currentTarget.value;
    if (!ZOOM_INPUT_PATTERN.test(nextInputValue)) return;
    setInputValue(nextInputValue);
    setActiveIndex(-1);
    setOpen(true);
  };

  const moveActiveOption = (direction: 1 | -1) => {
    setOpen(true);
    setActiveIndex((currentIndex) => {
      if (currentIndex < 0) {
        const matchingIndex = OFFICE_ZOOM_LEVELS.indexOf(Number(inputValue));
        if (matchingIndex >= 0) return fitOptionCount + matchingIndex;
        return direction > 0 ? 0 : optionCount - 1;
      }
      return (currentIndex + direction + optionCount) % optionCount;
    });
  };

  const selectFitMode = (
    mode: Exclude<OfficeFileViewerZoomMode, 'percentage'>,
  ) => {
    controls.changeMode(mode);
    closeOptions();
    inputRef.current?.focus();
  };
  const handleInputKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      moveActiveOption(event.key === 'ArrowDown' ? 1 : -1);
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      setInputValue(String(controls.value));
      closeOptions();
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      if (activeIndex >= 0 && activeIndex < fitOptionCount) {
        selectFitMode(controls.fitModes[activeIndex]);
        return;
      }
      const selectedValue =
        activeIndex >= fitOptionCount
          ? OFFICE_ZOOM_LEVELS[activeIndex - fitOptionCount]
          : undefined;
      commitZoom(
        selectedValue === undefined ? inputValue : String(selectedValue),
      );
      closeOptions();
      return;
    }
    if (event.key === 'Tab') {
      commitZoom();
      closeOptions();
    }
  };

  const handleInputBlur = () => {
    commitZoom();
    closeOptions();
  };

  const handleOptionMouseDown = (event: MouseEvent<HTMLButtonElement>) => {
    // 阻止输入框先触发 blur，避免选项点击被关闭动作吞掉。
    event.preventDefault();
  };

  const selectPreset = (value: number) => {
    commitZoom(String(value));
    closeOptions();
    inputRef.current?.focus();
  };

  return (
    <div className="office-file-zoom-control">
      <div className="office-file-zoom-combobox">
        <div className="office-file-zoom-input-wrap">
          <input
            ref={inputRef}
            className="office-file-zoom-input"
            value={inputValue}
            aria-activedescendant={
              open && activeIndex >= 0
                ? `${listId}-option-${activeIndex}`
                : undefined
            }
            aria-autocomplete="list"
            aria-controls={listId}
            aria-expanded={open}
            aria-label={messages.toolbar.zoomLevel}
            autoComplete="off"
            inputMode="numeric"
            role="combobox"
            onBlur={handleInputBlur}
            onChange={handleInputChange}
            onClick={() => setOpen(true)}
            onFocus={() => setOpen(true)}
            onKeyDown={handleInputKeyDown}
          />
          <span className="office-file-zoom-input__suffix" aria-hidden="true">
            %
          </span>
        </div>
        {open ? (
          <div id={listId} className="office-file-zoom-options" role="listbox">
            {controls.fitModes.map((mode, index) => (
              <button
                id={`${listId}-option-${index}`}
                key={mode}
                className="office-file-zoom-option office-file-zoom-option--fit"
                type="button"
                aria-selected={controls.mode === mode}
                role="option"
                tabIndex={-1}
                onClick={() => selectFitMode(mode)}
                onMouseDown={handleOptionMouseDown}
                onMouseEnter={() => setActiveIndex(index)}
              >
                {mode === 'fit-width'
                  ? messages.toolbar.fitWidth
                  : messages.toolbar.fitPage}
              </button>
            ))}
            {OFFICE_ZOOM_LEVELS.map((value, index) => (
              <button
                id={`${listId}-option-${fitOptionCount + index}`}
                key={value}
                className="office-file-zoom-option"
                type="button"
                aria-selected={
                  activeIndex === fitOptionCount + index ||
                  (controls.mode === 'percentage' &&
                    Number(inputValue) === value)
                }
                role="option"
                tabIndex={-1}
                onClick={() => selectPreset(value)}
                onMouseDown={handleOptionMouseDown}
                onMouseEnter={() => setActiveIndex(fitOptionCount + index)}
              >
                {value}%
              </button>
            ))}
          </div>
        ) : null}
      </div>
      <OfficeTooltip content={messages.toolbar.zoomOut}>
        <OfficeButton
          aria-label={messages.toolbar.zoomOut}
          icon={<ZoomOutIcon className="office-file-zoom-icon" />}
          disabled={!controls.hasDocument || controls.value <= OFFICE_MIN_ZOOM}
          onClick={controls.zoomOut}
        />
      </OfficeTooltip>
      <OfficeTooltip content={messages.toolbar.zoomIn}>
        <OfficeButton
          aria-label={messages.toolbar.zoomIn}
          icon={<ZoomInIcon className="office-file-zoom-icon" />}
          disabled={!controls.hasDocument || controls.value >= OFFICE_MAX_ZOOM}
          onClick={controls.zoomIn}
        />
      </OfficeTooltip>
    </div>
  );
}
