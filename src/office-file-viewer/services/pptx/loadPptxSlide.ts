import { getOfficePartRelationshipsPath } from '../../shared/ooxml/relationships';
import type { SlideModel } from '../presentation/types';
import { parseSlideXml } from './parsePptxSlide';
import type {
  PptxPackageContext,
  PptxSlideDescriptor,
} from './PptxPackageContext';
import { throwIfPptxParseAborted } from './PptxPackageContext';

/** 补齐当前页直接引用的图表或 WPS 扩展 XML。 */
async function loadReferencedXml(
  context: PptxPackageContext,
  descriptor: PptxSlideDescriptor,
  signal?: AbortSignal,
) {
  const targets = Object.values(
    context.packageState.relationships[descriptor.relsPath] ?? {},
  )
    .map((relationship) => relationship.target)
    .filter(
      (target) =>
        target.endsWith('.xml') &&
        !target.includes('slideLayouts/') &&
        !target.includes('notesSlides/'),
    );
  const readTargets = async (paths: readonly string[]) => {
    await Promise.all(
      paths.map(async (target) => {
        if (
          !context.reader.has(target) ||
          context.packageState.entries.has(target)
        ) {
          return;
        }
        context.packageState.entries.set(
          target,
          await context.reader.readText(target, signal),
        );
      }),
    );
  };

  await readTargets(targets);

  // 图表可通过自身关系文件覆盖主题；按页加载时必须连同覆盖 XML 一并读取。
  const chartThemeTargets = targets.flatMap((target) =>
    Object.values(
      context.packageState.relationships[
        getOfficePartRelationshipsPath(target)
      ] ?? {},
    )
      .filter((relationship) =>
        relationship.type?.toLowerCase().endsWith('/themeoverride'),
      )
      .map((relationship) => relationship.target),
  );
  await readTargets(chartThemeTargets);
}

/** 只读取并解析目标 PPTX Slide，复用完整解析器的格式计算规则。 */
export async function loadPptxSlide(
  context: PptxPackageContext,
  descriptor: PptxSlideDescriptor,
  signal?: AbortSignal,
): Promise<SlideModel> {
  throwIfPptxParseAborted(signal);
  const xml = await context.reader.readText(descriptor.slidePath, signal);
  await loadReferencedXml(context, descriptor, signal);
  throwIfPptxParseAborted(signal);
  const slide = parseSlideXml(
    xml,
    descriptor.index,
    context.width,
    context.height,
    context.packageState,
    context.theme,
    descriptor.relsPath,
    [...context.layoutDefinitions],
    [...context.masterDefinitions],
    context.tableStyles,
    Object.fromEntries(
      context.descriptors.map((item, index) => [item.slidePath, index]),
    ),
    context.defaultTextStyle,
  );
  slide.hidden = descriptor.hidden;
  slide.speakerNotes = undefined;
  return slide;
}
