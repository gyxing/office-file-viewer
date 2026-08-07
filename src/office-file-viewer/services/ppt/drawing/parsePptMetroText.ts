import {
  OFFICE_ART_RECORD,
  type OfficeArtRecord,
} from '../../../shared/officeart';
import { openOfficeArchive } from '../../../shared/ooxml/archive';
import { descendantsByLocalName } from '../../../shared/ooxml/xml';
import { parsePptxTextElement } from '../../pptx/parsePptxSlide';
import type { TextElement, ThemeModel } from '../../presentation/types';
import type { PptParseContext } from '../types';
import { parsePptMetroShapeXml } from './parsePptMetroShapeXml';
import { readPptOfficeArtProperties } from './readOfficeArtProperties';

/** OfficeArt 中保存兼容 DrawingML 压缩包的 metroBlob 属性编号。 */
const PPT_METRO_BLOB_PROPERTY_ID = 0x03a9;
/** PowerPoint 写入 metroBlob 时使用的标准形状 XML 路径。 */
const PPT_METRO_SHAPE_PATH = 'drs/shapexml.xml';

function findMetroBlob(record: OfficeArtRecord) {
  const propertyRecord = record.children?.find(
    (child) => child.type === OFFICE_ART_RECORD.TERTIARY_FOPT,
  );
  return readPptOfficeArtProperties(propertyRecord).get(
    PPT_METRO_BLOB_PROPERTY_ID,
  )?.complexData;
}

/**
 * 按需读取旧版 PPT 形状内嵌的 DrawingML 文本，保留新版 Office 才能表达的文字样式。
 */
export async function parsePptMetroText(
  record: OfficeArtRecord,
  index: number,
  theme: ThemeModel,
  context: PptParseContext,
): Promise<TextElement | undefined> {
  const metroBlob = findMetroBlob(record);
  if (!metroBlob?.length) return undefined;

  let archive: Awaited<ReturnType<typeof openOfficeArchive>> | undefined;
  try {
    archive = await openOfficeArchive(metroBlob);
    const shapePath = archive.has(PPT_METRO_SHAPE_PATH)
      ? PPT_METRO_SHAPE_PATH
      : archive
          .list()
          .find(({ path }) => path.toLowerCase().endsWith('/shapexml.xml'))
          ?.path;
    if (!shapePath) return undefined;

    const shape = parsePptMetroShapeXml(await archive.readText(shapePath));
    const hasText = descendantsByLocalName(shape, 't').some((node) =>
      Boolean(node.textContent?.trim()),
    );
    return hasText ? parsePptxTextElement(shape, index, theme) : undefined;
  } catch (error) {
    const message =
      error instanceof Error
        ? `PPT 兼容形状无法读取：${error.message}`
        : 'PPT 兼容形状无法读取';
    if (
      !context.warnings.some(
        (warning) =>
          warning.code === 'PPT_METRO_BLOB_CORRUPT' &&
          warning.message === message,
      )
    ) {
      context.warnings.push({
        code: 'PPT_METRO_BLOB_CORRUPT',
        message,
      });
    }
    return undefined;
  } finally {
    await archive?.close();
  }
}
