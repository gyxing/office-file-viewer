import type { ChangeEvent } from 'react';
import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { WordRevisionMode } from '../services/annotations/types';

/** Word 修订投影模式控件属性。 */
type WordRevisionModeControlProps = {
  /** 当前采用的最终态、标记态或原始态。 */
  value: WordRevisionMode;
  /** 当前是否禁止切换修订投影。 */
  disabled: boolean;
  /** 修订投影变化时触发。 */
  onChange(value: WordRevisionMode): void;
};

/** 在工具栏中切换 Word 修订内容的只读投影。 */
function WordRevisionModeControlComponent({
  value,
  disabled,
  onChange,
}: WordRevisionModeControlProps) {
  const messages = useOfficeFileViewerMessages();
  const handleChange = (event: ChangeEvent<HTMLSelectElement>) => {
    onChange(event.currentTarget.value as WordRevisionMode);
  };

  return (
    <select
      className="office-file-word-revision-mode"
      data-testid="office-word-revision-mode"
      aria-label={messages.review.revisionMode}
      value={value}
      disabled={disabled}
      onChange={handleChange}
    >
      <option value="final">{messages.review.revisionFinal}</option>
      <option value="markup">{messages.review.revisionMarkup}</option>
      <option value="original">{messages.review.revisionOriginal}</option>
    </select>
  );
}

export const WordRevisionModeControl = memo(WordRevisionModeControlComponent);
