import React, { memo } from 'react';
import { useOfficeFileViewerMessages } from '../../locale';
import type { SpeakerNotesModel } from '../../services/presentation/types';
import { useOfficeFontResolver } from '../../shared/fonts/OfficeFontProvider';
import { useSpeakerNotesResize } from './useSpeakerNotesResize';

/** 定义演讲者备注面板可接收的属性。 */
type PptxSpeakerNotesProps = {
  /** 当前幻灯片序号。 */
  slideIndex: number;
  /** 当前幻灯片解析出的演讲者备注。 */
  notes?: SpeakerNotesModel;
};

/** 渲染当前幻灯片的演讲者备注，并保留段落与项目符号结构。 */
function PptxSpeakerNotesComponent({
  slideIndex,
  notes,
}: PptxSpeakerNotesProps) {
  const messages = useOfficeFileViewerMessages();
  const resolveFontFamily = useOfficeFontResolver();
  const { height, maxHeight, panelRef, handleKeyDown, handleMouseDown } =
    useSpeakerNotesResize();

  return (
    <section
      ref={panelRef}
      className="office-file-pptx-speaker-notes"
      aria-label={messages.presentation.notesRegion}
      style={{ height }}
    >
      <div
        className="office-file-pptx-speaker-notes__resize-handle"
        role="separator"
        aria-label={messages.presentation.resizeNotes}
        aria-orientation="horizontal"
        aria-valuemin={120}
        aria-valuemax={Math.round(maxHeight)}
        aria-valuenow={Math.round(height)}
        tabIndex={0}
        onKeyDown={handleKeyDown}
        onMouseDown={handleMouseDown}
      >
        <span aria-hidden="true" />
      </div>
      <div className="office-file-pptx-speaker-notes__content">
        {notes?.paragraphs.length ? (
          notes.paragraphs.map((paragraph, index) => {
            const level = Math.min(5, Math.max(0, paragraph.level ?? 0));
            const bullet = paragraph.bullet?.none
              ? undefined
              : paragraph.bullet?.char;
            return (
              <p
                className={`office-file-pptx-speaker-notes__paragraph office-file-pptx-speaker-notes__paragraph--level-${level}`}
                key={`${slideIndex}-${index}`}
              >
                {bullet ? (
                  <span
                    className="office-file-pptx-speaker-notes__bullet"
                    aria-hidden="true"
                    style={{
                      fontFamily: resolveFontFamily(
                        paragraph.bullet?.fontFamily,
                      ),
                    }}
                  >
                    {bullet}
                  </span>
                ) : null}
                {paragraph.runs.map((run, runIndex) => (
                  <span key={`${index}-${runIndex}`}>{run.text}</span>
                ))}
              </p>
            );
          })
        ) : (
          <div className="office-file-pptx-speaker-notes__empty">
            {messages.presentation.emptyNotes}
          </div>
        )}
      </div>
    </section>
  );
}

export const PptxSpeakerNotes = memo(PptxSpeakerNotesComponent);
