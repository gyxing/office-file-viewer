import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../locale';
import type { ParseProgress } from '../services/parsing';
import { OfficeNotice } from './Notice';

/** Office解析状态组件属性。 */
type OfficeParseStatusProps = {
  /** 当前解析阶段及其完成度信息。 */
  progress?: ParseProgress;
  /** 是否正在展示解析失败后保留的部分内容。 */
  warning?: boolean;
};

/** 将输入标准化为 `normalizePercent` 返回的结构。 */
function normalizePercent(progress: ParseProgress | undefined) {
  if (progress?.percent !== undefined) {
    return Math.max(0, Math.min(100, progress.percent * 100));
  }
  if (progress?.total && progress.completed !== undefined) {
    return Math.max(
      0,
      Math.min(100, (progress.completed / progress.total) * 100),
    );
  }
  return undefined;
}

/** OfficeParseStatus 在预览内容上方展示非阻塞解析进度或不完整警告。 */
function OfficeParseStatusComponent({
  progress,
  warning,
}: OfficeParseStatusProps) {
  const messages = useOfficeFileViewerMessages();
  if (warning) {
    return (
      <div className="office-file-parse-status" role="alert">
        <OfficeNotice
          type="warning"
          title={messages.progress.partialTitle}
          description={messages.progress.partialDescription}
        />
      </div>
    );
  }
  if (!progress) return null;

  const percent = normalizePercent(progress);
  const barClassName = [
    'office-file-parse-status__bar',
    percent === undefined ? 'office-file-parse-status__bar--indeterminate' : '',
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <div className="office-file-parse-status" role="status" aria-live="polite">
      <div className="office-file-parse-status__label">
        {messages.progress.stages[progress.stage]}
      </div>
      <div
        className="office-file-parse-status__track"
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={percent === undefined ? undefined : Math.round(percent)}
      >
        <div
          className={barClassName}
          style={percent === undefined ? undefined : { width: `${percent}%` }}
        />
      </div>
    </div>
  );
}

export const OfficeParseStatus = memo(OfficeParseStatusComponent);
