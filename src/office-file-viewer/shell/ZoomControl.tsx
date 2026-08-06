import { AutoComplete, Button, Input, Space, Tooltip } from 'antd';
import type { KeyboardEvent } from 'react';
import React, { useEffect, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import {
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
  OFFICE_ZOOM_LEVELS,
} from './constants';
import { ZoomInIcon, ZoomOutIcon } from './icons';

/** 工具栏通用缩放能力。 */
export type ZoomControls = {
  /** 当前预览缩放比例。 */
  value: number;
  /** 当前是否已有可执行增减操作的文档。 */
  hasDocument: boolean;
  /** 减小预览比例。 */
  zoomOut(): void;
  /** 放大预览比例。 */
  zoomIn(): void;
  /** 设置指定预览比例。 */
  change(value: number): void;
};

type ZoomControlProps = {
  /** 当前可用的缩放能力。 */
  controls: ZoomControls;
};

/** 缩放下拉列表保持固定档位，自定义输入值不会追加到列表。 */
const OFFICE_ZOOM_OPTIONS = OFFICE_ZOOM_LEVELS.map((value) => ({
  value: String(value),
  label: `${value}%`,
}));

/** 缩放输入仅接受空值或十进制整数，空值在提交时恢复当前比例。 */
const ZOOM_INPUT_PATTERN = /^\d*$/;

/** 将有效整数限制到预览器支持的缩放范围。 */
function normalizeZoomInput(value: string): number | undefined {
  if (!value || !/^\d+$/.test(value)) return undefined;
  const parsedValue = Number(value);
  if (!Number.isFinite(parsedValue)) return undefined;
  return Math.min(OFFICE_MAX_ZOOM, Math.max(OFFICE_MIN_ZOOM, parsedValue));
}

/** 提供可输入、可选择并带固定百分号后缀的缩放操作。 */
export function ZoomControl({ controls }: ZoomControlProps) {
  const messages = useOfficeFileViewerMessages();
  // 只有方向键主动进入选项导航后，Enter 才交给 AutoComplete 选择建议项。
  const keyboardOptionNavigationRef = useRef(false);
  const [inputValue, setInputValue] = useState(String(controls.value));

  useEffect(() => {
    setInputValue(String(controls.value));
  }, [controls.value]);

  const commitZoom = (nextInputValue = inputValue) => {
    const normalizedValue = normalizeZoomInput(nextInputValue);
    if (normalizedValue === undefined) {
      setInputValue(String(controls.value));
      return;
    }
    setInputValue(String(normalizedValue));
    controls.change(normalizedValue);
  };

  const handleInputChange = (nextInputValue: string) => {
    keyboardOptionNavigationRef.current = false;
    if (ZOOM_INPUT_PATTERN.test(nextInputValue)) {
      setInputValue(nextInputValue);
    }
  };

  const handleInputKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      keyboardOptionNavigationRef.current = true;
      return;
    }
    if (event.key !== 'Enter' || keyboardOptionNavigationRef.current) return;
    event.preventDefault();
    event.stopPropagation();
    commitZoom();
    event.currentTarget.blur();
  };

  const handleSelect = (value: string) => {
    keyboardOptionNavigationRef.current = false;
    commitZoom(value);
  };

  const handleBlur = () => {
    keyboardOptionNavigationRef.current = false;
    commitZoom();
  };

  return (
    <Space size={8}>
      <AutoComplete
        value={inputValue}
        className="office-file-toolbar__zoom"
        options={OFFICE_ZOOM_OPTIONS}
        defaultActiveFirstOption={false}
        filterOption={false}
        onChange={handleInputChange}
        onSelect={handleSelect}
      >
        <Input
          aria-label={messages.toolbar.zoomLevel}
          inputMode="numeric"
          suffix={<span aria-hidden="true">%</span>}
          onBlur={handleBlur}
          onKeyDownCapture={handleInputKeyDown}
        />
      </AutoComplete>
      <Tooltip title={messages.toolbar.zoomOut}>
        <Button
          aria-label={messages.toolbar.zoomOut}
          icon={<ZoomOutIcon />}
          disabled={!controls.hasDocument || controls.value <= OFFICE_MIN_ZOOM}
          onClick={controls.zoomOut}
        />
      </Tooltip>
      <Tooltip title={messages.toolbar.zoomIn}>
        <Button
          aria-label={messages.toolbar.zoomIn}
          icon={<ZoomInIcon />}
          disabled={!controls.hasDocument || controls.value >= OFFICE_MAX_ZOOM}
          onClick={controls.zoomIn}
        />
      </Tooltip>
    </Space>
  );
}
