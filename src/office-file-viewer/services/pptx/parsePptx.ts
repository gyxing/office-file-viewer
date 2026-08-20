import { loadOfficeEntries, readXml } from '../../shared/ooxml/archive';
import {
  attr,
  childrenByLocalName,
  descendantByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeArchiveResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import { collectRenderableOfficeMedia } from '../media/officeMetafile';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import { throwIfParseAborted } from '../parsing/runtime/types';
import { parseSlideXml, readPptxDefaultTextStyle } from './parsePptxSlide';
import { throwIfPptxParseAborted } from './PptxPackageContext';
import {
  buildPptxPackageState,
  readPresentationLayouts,
  readPresentationSize,
  readTableStyles,
  readTheme,
} from './readPptxPresentationParts';
import type { PptxDocument, SlideModel } from './types';

/** 在幻灯片边界让出主线程，使大演示文稿切换或卸载可以及时取消。 */
async function pptxParseCheckpoint(signal?: AbortSignal) {
  throwIfPptxParseAborted(signal);
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  throwIfPptxParseAborted(signal);
}

/** 解析 PPTX 包并返回标准演示文稿模型。 */
export async function parsePptx(
  file: File,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
): Promise<PptxDocument> {
  // 解析顺序：包资源 -> 主题/表格样式 -> 母版/版式 -> 每页 slide，最终产出前端可直接渲染的模型。
  throwIfPptxParseAborted(signal);
  const entries = await loadOfficeEntries(file, { signal, resourcePolicy });
  await pptxParseCheckpoint(signal);
  const media = await collectRenderableOfficeMedia(entries, 'ppt/media/');
  await pptxParseCheckpoint(signal);
  const packageState = buildPptxPackageState(entries, {
    mediaByName: media.byName,
    mediaByPath: media.byPath,
  });
  const presentationXml = readXml(entries, 'ppt/presentation.xml');
  const presentationDoc = parseXml(presentationXml);
  const themeXml = readXml(entries, 'ppt/theme/theme1.xml');
  const theme = themeXml
    ? readTheme(themeXml)
    : { colorScheme: {}, fontScheme: {}, colorMap: {} };
  const defaultTextStyle = readPptxDefaultTextStyle(presentationXml, theme);
  const tableStyles = readTableStyles(
    readXml(entries, 'ppt/tableStyles.xml'),
    theme,
  );
  const size = presentationXml
    ? readPresentationSize(presentationXml)
    : { width: 960, height: 540 };
  const presentationRels =
    packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const slideIdList = descendantByLocalName(
    presentationDoc.documentElement,
    'sldIdLst',
  );
  const slideIds = childrenByLocalName(slideIdList, 'sldId');
  const { masterDefinitions, masterLayoutDefinitions } =
    readPresentationLayouts(entries, packageState, theme, tableStyles);
  const layoutDefinitions = Object.values(masterLayoutDefinitions).flat();
  const slides: SlideModel[] = [];
  const slideTargets = Object.fromEntries(
    slideIds.map((node, index) => {
      const relationshipId = attr(node, 'r:id');
      const target = relationshipId
        ? presentationRels[relationshipId]?.target
        : undefined;
      return [target ?? `ppt/slides/slide${index + 1}.xml`, index];
    }),
  );

  for (let index = 0; index < slideIds.length; index += 1) {
    await pptxParseCheckpoint(signal);
    const node = slideIds[index];
    const relId = attr(node, 'r:id');
    const relTarget = relId ? presentationRels[relId]?.target : undefined;
    const relPath = relTarget
      ? relTarget.replace(/^ppt\//, '')
      : `slides/slide${index + 1}.xml`;
    const slidePath = relTarget ?? `ppt/${relPath}`;
    const relsPath = `ppt/${relPath.replace(
      /^slides\//,
      'slides/_rels/',
    )}.rels`;
    slides.push(
      parseSlideXml(
        readXml(entries, slidePath),
        index + 1,
        size.width,
        size.height,
        packageState,
        theme,
        relsPath,
        layoutDefinitions,
        masterDefinitions,
        tableStyles,
        slideTargets,
        defaultTextStyle,
      ),
    );
  }

  throwIfPptxParseAborted(signal);
  const warnings = slides.flatMap((slide) => slide.warnings ?? []);
  return {
    width: size.width,
    height: size.height,
    theme,
    slides,
    warnings: warnings.length ? warnings : undefined,
  };
}

/** 通过统一运行时合同解析 PPTX，并输出完整演示文稿模型。 */
export const runPptxParser: OfficeFormatParser = async (
  file,
  { signal, resourcePolicy },
  sink,
) => {
  sink.progress({
    stage: 'content',
    percent: 0.05,
    message: '正在解析文件',
  });
  const document = await parsePptx(file, signal, resourcePolicy);
  throwIfParseAborted(signal);
  await sink.parsed({ kind: 'pptx', document });
  await sink.complete();
};
