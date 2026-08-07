import type { ChangeEvent } from 'react';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { SpreadsheetViewMode } from '../services/spreadsheet/viewMode';

/** 电子表格显示模式控件属性。 */
type SpreadsheetViewModeControlProps = {
  /** 当前采用的显示模式。 */
  value: SpreadsheetViewMode;
  /** 当前是否禁止切换显示模式。 */
  disabled: boolean;
  /** 显示模式变化时触发。 */
  onChange: (value: SpreadsheetViewMode) => void;
};

/** 在原始版式和完整内容阅读模式之间切换。 */
function SpreadsheetViewModeControlComponent({
  value,
  disabled,
  onChange,
}: SpreadsheetViewModeControlProps) {
  const messages = useOfficeFileViewerMessages();

  const handleChange = (event: ChangeEvent<HTMLSelectElement>) => {
    onChange(event.currentTarget.value as SpreadsheetViewMode);
  };

  return (
    <select
      className="office-file-spreadsheet-view-mode"
      aria-label={messages.spreadsheet.viewMode}
      value={value}
      disabled={disabled}
      onChange={handleChange}
    >
      <option value="source">{messages.spreadsheet.sourceViewMode}</option>
      <option value="reading">{messages.spreadsheet.readingViewMode}</option>
    </select>
  );
}

export const SpreadsheetViewModeControl = memo(
  SpreadsheetViewModeControlComponent,
);
