import type {
  SpeakerNotesModel,
  TextElement,
  ThemeModel,
} from '../../presentation/types';

import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { parsePptDrawing } from '../drawing';
import type { PptEditChain, PptParseContext } from '../types';
import type { PptNotesDescriptor } from './readSlideLists';

/** 一条已通过 NotesAtom 精确关联到幻灯片的备注记录。 */
export type PptNotesModel = {
  /** NotesAtom 指向的正式幻灯片 ID。 */
  slideIdRef: number;
  /** 过滤页眉页脚后得到的演讲者正文。 */
  speakerNotes?: SpeakerNotesModel;
};

/** 读取一个 NotesContainer，并只保留 Tx_TYPE_NOTES 文本。 */
export async function readPptNotes(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  descriptor: PptNotesDescriptor,
  theme: ThemeModel,
  fonts: Map<number, string>,
  context: PptParseContext,
): Promise<PptNotesModel | undefined> {
  const offset = editChain.persistOffsets.get(descriptor.persistId);
  if (offset === undefined) return undefined;

  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (!record || record.type !== PPT_RECORD.NOTES) {
    context.warnings.push({
      code: 'PPT_NOTES_CORRUPT',
      message: `备注页 ${descriptor.notesId} 不是有效的 NotesContainer`,
      offset,
    });
    return undefined;
  }

  let slideIdRef = 0;
  let drawing: Uint8Array | undefined;
  const children = new PptRecordReader(
    documentStream,
    record.dataOffset,
    record.endOffset,
  );
  for (const child of children.records()) {
    if (child.type === PPT_RECORD.NOTES_ATOM && child.length >= 4) {
      slideIdRef = new DataView(
        child.data.buffer,
        child.data.byteOffset,
        child.data.byteLength,
      ).getUint32(0, true);
    } else if (child.type === PPT_RECORD.PP_DRAWING) {
      drawing = child.data;
    }
  }
  if (!slideIdRef) return undefined;

  const elements = drawing
    ? (await parsePptDrawing(drawing, theme, fonts, context)).elements
    : [];
  const paragraphs = elements
    .filter(
      (element): element is TextElement =>
        element.type === 'text' && element.placeholderType === 'notes',
    )
    .flatMap((element) => element.paragraphs);
  const plainText = paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');

  return {
    slideIdRef,
    speakerNotes: plainText.trim() ? { paragraphs, plainText } : undefined,
  };
}
