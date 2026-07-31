import type { SpeakerNotesModel, ThemeModel } from '../../presentation/types';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';
import type {
  PptBinaryDocument,
  PptEditChain,
  PptParseContext,
  PptSlideModel,
} from '../types';
import { readPptFonts } from './readFonts';
import { readPptMaster } from './readMaster';
import { readPptNotes } from './readNotes';
import { readPptSlide } from './readSlide';
import { readPptSlideLists } from './readSlideLists';

/** PPT 缺少页面设置时使用的默认幻灯片宽度。 */
const DEFAULT_SLIDE_WIDTH = 960;
/** PPT 缺少页面设置时使用的默认幻灯片高度。 */
const DEFAULT_SLIDE_HEIGHT = 540;
/** PowerPoint 主坐标固定为 576 dpi，统一换算到浏览器的 96 dpi。 */
const MASTER_UNIT_TO_PX = 96 / 576;

/** PPT 文档的页面尺寸、主题和母版。 */
export type PptDocumentStructure = Pick<
  PptBinaryDocument,
  'width' | 'height' | 'theme' | 'masters'
>;

/** 完整解析和按页 Source 共用的文档结构读取结果。 */
export type PptDocumentReadStructure = PptDocumentStructure & {
  /** 按字体编号索引的字体族名称。 */
  fonts: Map<number, string>;
  /** 按源顺序排列的轻量描述信息。 */
  descriptors: ReturnType<typeof readPptSlideLists>;
};

/** 根 DocumentContainer 内无需读取独立母版记录的基础结构。 */
export type PptDocumentBaseStructure = Omit<
  PptDocumentReadStructure,
  'masters'
>;

/** PPT 文档和幻灯片解析进度的异步观察器。 */
export type PptDocumentObserver = {
  /** 接收 PPT 文档的结构与资源目录。 */
  structure(value: PptDocumentStructure): Promise<void>;
  /** 接收解析产生的单张幻灯片。 */
  slide(index: number, slide: PptSlideModel): Promise<void>;
};

/** 读取文档大小。 */
export function readDocumentSize(
  documentStream: Uint8Array,
  documentRecordOffset: number,
) {
  const documentRecord = new PptRecordReader(
    documentStream,
    documentRecordOffset,
    documentStream.length,
  ).readRecord()!;
  const reader = new PptRecordReader(
    documentStream,
    documentRecord.dataOffset,
    documentRecord.endOffset,
  );
  for (const child of reader.records()) {
    if (child.type !== PPT_RECORD.DOCUMENT_ATOM || child.length < 8) continue;
    const view = new DataView(
      child.data.buffer,
      child.data.byteOffset,
      child.data.byteLength,
    );
    const width = view.getInt32(0, true) * MASTER_UNIT_TO_PX;
    const height = view.getInt32(4, true) * MASTER_UNIT_TO_PX;
    if (width > 0 && height > 0) return { width, height };
  }
  return { width: DEFAULT_SLIDE_WIDTH, height: DEFAULT_SLIDE_HEIGHT };
}

/** 从一个已读取的 DocumentContainer 恢复尺寸、字体、主题和页面引用。 */
export function readPptDocumentBaseStructure(
  documentStream: Uint8Array,
  documentOffset: number,
  context: PptParseContext,
): PptDocumentBaseStructure {
  const documentRecord = new PptRecordReader(
    documentStream,
    documentOffset,
    documentStream.length,
  ).readRecord();
  if (!documentRecord || documentRecord.type !== PPT_RECORD.DOCUMENT) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '无法读取 PowerPoint 根文档对象',
      { offset: documentOffset, recordType: documentRecord?.type },
    );
  }
  const { width, height } = readDocumentSize(documentStream, documentOffset);
  const fonts = readPptFonts(documentStream, documentRecord);
  const defaultFont = fonts.get(0) ?? fonts.values().next().value;
  const theme: ThemeModel = {
    colorScheme: {
      lt1: '#ffffff',
      dk1: '#000000',
      accent1: '#4472c4',
      accent2: '#ed7d31',
    },
    fontScheme: { minorLatin: defaultFont },
    colorMap: {
      bg1: 'lt1',
      tx1: 'dk1',
      accent1: 'accent1',
      accent2: 'accent2',
    },
  };
  const descriptors = readPptSlideLists(
    documentStream,
    documentRecord,
    context,
  );
  return { width, height, theme, fonts, descriptors };
}

/** 只恢复根文档、字体、页面顺序和母版，不进入 Slide 或备注正文。 */
export function readPptDocumentStructure(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  context: PptParseContext,
): PptDocumentReadStructure {
  const documentOffset = editChain.persistOffsets.get(
    editChain.documentPersistId,
  )!;
  const { width, height, theme, fonts, descriptors } =
    readPptDocumentBaseStructure(documentStream, documentOffset, context);
  const masters = new Map(
    descriptors.masters
      .map((descriptor) => {
        const master = readPptMaster(
          documentStream,
          editChain,
          descriptor,
          theme,
          fonts,
          context,
        );
        return master ? ([master.id, master] as const) : undefined;
      })
      .filter(
        (entry): entry is readonly [number, NonNullable<typeof entry>[1]] =>
          Boolean(entry),
      ),
  );
  return { width, height, theme, masters, fonts, descriptors };
}

/** 从最终 persist 目录恢复文档、母版和正式幻灯片顺序。 */
export async function readPptBinaryDocument(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  context: PptParseContext,
  observer?: PptDocumentObserver,
): Promise<PptBinaryDocument> {
  const { width, height, theme, masters, fonts, descriptors } =
    readPptDocumentStructure(documentStream, editChain, context);
  await observer?.structure({ width, height, theme, masters });
  const speakerNotesBySlideId = new Map<number, SpeakerNotesModel>();
  for (const descriptor of descriptors.notes) {
    const notes = readPptNotes(
      documentStream,
      editChain,
      descriptor,
      theme,
      fonts,
      context,
    );
    if (notes?.speakerNotes) {
      speakerNotesBySlideId.set(notes.slideIdRef, notes.speakerNotes);
    }
  }
  const slides = [];
  for (const descriptor of descriptors.slides) {
    const slide = readPptSlide(
      documentStream,
      editChain,
      descriptor,
      width,
      height,
      theme,
      fonts,
      context,
    );
    if (slide) {
      slide.speakerNotes = speakerNotesBySlideId.get(descriptor.slideId);
      // 幻灯片局部背景优先于母版；两者都未声明时才使用主题背景。
      slide.background ??= masters.get(slide.masterId ?? Number.NaN)
        ?.background ?? {
        fill: theme.colorScheme.lt1 ?? '#ffffff',
      };
      slides.push(slide);
      await observer?.slide(slides.length - 1, slide);
    }
    await context.yieldIfNeeded();
  }
  if (!slides.length) {
    throw new PptParseError(
      'PPT_NO_VALID_SLIDES',
      'PPT 文件中没有可预览的有效幻灯片',
    );
  }

  return {
    width,
    height,
    theme,
    masters,
    slides,
    externalObjects: new Map(),
    warnings: context.warnings,
  };
}
