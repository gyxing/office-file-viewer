import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import {
  collectMedia,
  type OfficeRelationship,
} from '../../shared/ooxml/media';
import { readRelationships } from '../../shared/ooxml/relationships';
import { emuToPx } from '../../shared/ooxml/units';
import {
  attr,
  childByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import { markPresentationElementsPreviewable } from '../presentation/imagePreviewPolicy';
import {
  parsePptxVisualTree,
  readPptxPlaceholder,
  readPptxSlideBackground,
  readPptxTableCellStyle,
  readPptxTextPresetMap,
} from './parsePptxSlide';
import type {
  LayoutDefinition,
  MasterDefinition,
  PptxPackageState as PackageState,
  PlaceholderStyle,
  RelationshipMap,
  TableCellStyle,
  TableStyleMap,
  TableStyleVariantName,
} from './PptxPackageContext';
import type { ThemeModel } from './types';

/** 构建 PPTX 各部件解析共享的包状态。 */
export function buildPptxPackageState(
  entries: OfficeEntryMap,
  mediaSources?: Pick<PackageState, 'mediaByName' | 'mediaByPath'>,
): PackageState {
  const relationships: RelationshipMap = {};
  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = mediaSources
    ? { byName: mediaSources.mediaByName, byPath: mediaSources.mediaByPath }
    : collectMedia(entries, 'ppt/media/');
  const mediaUseCounts: Record<string, number> = {};
  Object.entries(relationships).forEach(([relsPath, rels]) => {
    if (!/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/iu.test(relsPath)) {
      return;
    }
    Object.values(rels).forEach((relationship) => {
      if (
        relationship.targetMode?.toLowerCase() === 'external' ||
        !/^ppt\/media\//iu.test(relationship.target)
      ) {
        return;
      }
      mediaUseCounts[relationship.target] =
        (mediaUseCounts[relationship.target] ?? 0) + 1;
    });
  });

  return {
    entries,
    relationships,
    mediaByName: media.byName,
    mediaByPath: media.byPath,
    mediaUseCounts,
  };
}

/** 输出 PPTX 包结构的诊断信息。 */
export function debugPptxPackage(entries: OfficeEntryMap) {
  const packageState = buildPptxPackageState(entries);
  return {
    relsCount: Object.keys(packageState.relationships).length,
    mediaCount: Object.keys(packageState.mediaByPath).length,
    mediaSample: Object.entries(packageState.mediaByPath)
      .slice(0, 5)
      .map(([path, src]) => ({
        path,
        hasSrc: Boolean(src),
        prefix: typeof src === 'string' ? src.slice(0, 30) : src.kind,
      })),
  };
}

/** 读取 PPTX 页面宽高并转换为渲染像素。 */
export function readPresentationSize(xml: string) {
  const doc = parseXml(xml);
  const sldSz = descendantByLocalName(doc.documentElement, 'sldSz');
  const cx = Number(attr(sldSz, 'cx') ?? 12800000);
  const cy = Number(attr(sldSz, 'cy') ?? 7200000);
  return { width: emuToPx(cx), height: emuToPx(cy) };
}

/** 读取 PPTX 主题颜色和字体配置。 */
export function readTheme(xml: string): ThemeModel {
  const doc = parseXml(xml);
  const colorScheme: Record<string, string> = {};
  const fontScheme: Record<string, string> = {};
  const colorMap = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
  };

  const colorNode = descendantByLocalName(doc.documentElement, 'clrScheme');
  descendantsByLocalName(colorNode, 'dk1')
    .concat(descendantsByLocalName(colorNode, 'lt1'))
    .concat(descendantsByLocalName(colorNode, 'dk2'))
    .concat(descendantsByLocalName(colorNode, 'lt2'))
    .concat(descendantsByLocalName(colorNode, 'accent1'))
    .concat(descendantsByLocalName(colorNode, 'accent2'))
    .concat(descendantsByLocalName(colorNode, 'accent3'))
    .concat(descendantsByLocalName(colorNode, 'accent4'))
    .concat(descendantsByLocalName(colorNode, 'accent5'))
    .concat(descendantsByLocalName(colorNode, 'accent6'))
    .concat(descendantsByLocalName(colorNode, 'hlink'))
    .concat(descendantsByLocalName(colorNode, 'folHlink'))
    .forEach((node) => {
      const child = node.firstElementChild;
      if (!child) return;
      const childName = child.localName.split(':').pop()?.toLowerCase();
      const value =
        childName === 'sysclr'
          ? attr(child, 'lastClr') ?? attr(child, 'val')
          : attr(child, 'val') ?? attr(child, 'lastClr');
      if (value) colorScheme[node.localName] = value;
    });

  const fontNode = descendantByLocalName(doc.documentElement, 'fontScheme');
  ['majorFont', 'minorFont'].forEach((bucket) => {
    const node = childByLocalName(fontNode, bucket);
    if (!node) return;
    const latin = childByLocalName(node, 'latin');
    const ea = childByLocalName(node, 'ea');
    const cs = childByLocalName(node, 'cs');
    const latinFamily = attr(latin, 'typeface');
    const eastAsiaFamily = attr(ea, 'typeface');
    const complexScriptFamily = attr(cs, 'typeface');
    // 同时保留分类槽和兼容字体栈，主题占位符才能按文字类别还原源字体。
    if (latinFamily) fontScheme[`${bucket}Latin`] = latinFamily;
    if (eastAsiaFamily) fontScheme[`${bucket}EastAsia`] = eastAsiaFamily;
    if (complexScriptFamily) {
      fontScheme[`${bucket}ComplexScript`] = complexScriptFamily;
    }
    fontScheme[bucket] = [latinFamily, eastAsiaFamily, complexScriptFamily]
      .filter(Boolean)
      .join(', ');
  });

  return { colorScheme, fontScheme, colorMap };
}

/** 读取 PPTX 表格样式及其区域变体。 */
export function readTableStyles(xml: string, theme: ThemeModel): TableStyleMap {
  if (!xml) return {};
  const doc = parseXml(xml);
  const result: TableStyleMap = {};
  descendantsByLocalName(doc.documentElement, 'tblStyle').forEach(
    (styleNode) => {
      const styleId = attr(styleNode, 'styleId');
      if (!styleId) return;
      const variants: Partial<Record<TableStyleVariantName, TableCellStyle>> =
        {};
      (
        [
          'wholeTbl',
          'band1H',
          'band2H',
          'band1V',
          'band2V',
          'firstRow',
          'lastRow',
          'firstCol',
          'lastCol',
        ] as TableStyleVariantName[]
      ).forEach((variantName) => {
        const variantNode = childByLocalName(styleNode, variantName);
        if (!variantNode) return;
        variants[variantName] = readPptxTableCellStyle(variantNode, theme);
      });
      result[styleId] = {
        styleId,
        styleName: attr(styleNode, 'styleName') ?? undefined,
        variants,
      };
    },
  );
  return result;
}

function readMaster(
  xml: string,
  theme: ThemeModel,
  relPath: string,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
  tableStyles?: TableStyleMap,
): MasterDefinition {
  const doc = parseXml(xml);
  const cSld = childByLocalName(doc.documentElement, 'cSld');
  const bg = childByLocalName(cSld, 'bg');
  const background = readPptxSlideBackground(bg, theme, packageState, rels);
  const placeholders: Record<string, PlaceholderStyle> = {};
  descendantsByLocalName(cSld, 'sp').forEach((node) => {
    const ph = descendantByLocalName(node, 'ph');
    if (!ph) return;
    const style = readPptxPlaceholder(node, theme);
    const key = `${style.type ?? 'body'}:${style.idx ?? '0'}`;
    placeholders[key] = style;
  });
  const textPresets = readPptxTextPresetMap(
    childByLocalName(doc.documentElement, 'txStyles'),
    theme,
  );
  const elements = markPresentationElementsPreviewable(
    parsePptxVisualTree(
      childByLocalName(cSld, 'spTree'),
      theme,
      packageState,
      rels,
      `master-${relPath}`,
      undefined,
      tableStyles,
      false,
    ),
    false,
  );
  return { path: relPath, placeholders, textPresets, background, elements };
}

function readLayout(
  xml: string,
  theme: ThemeModel,
  relPath: string,
  masterPath: string,
  packageState: PackageState,
  rels: Record<string, OfficeRelationship>,
  tableStyles?: TableStyleMap,
): LayoutDefinition {
  const doc = parseXml(xml);
  const cSld = childByLocalName(doc.documentElement, 'cSld');
  const bg = childByLocalName(cSld, 'bg');
  const background = readPptxSlideBackground(bg, theme, packageState, rels);
  const placeholders: Record<string, PlaceholderStyle> = {};
  descendantsByLocalName(cSld, 'sp').forEach((node) => {
    const ph = descendantByLocalName(node, 'ph');
    if (!ph) return;
    const style = readPptxPlaceholder(node, theme);
    const key = `${style.type ?? 'body'}:${style.idx ?? '0'}`;
    placeholders[key] = style;
  });
  const textPresets = readPptxTextPresetMap(
    childByLocalName(doc.documentElement, 'txStyles'),
    theme,
  );
  const elements = markPresentationElementsPreviewable(
    parsePptxVisualTree(
      childByLocalName(cSld, 'spTree'),
      theme,
      packageState,
      rels,
      `layout-${relPath}`,
      undefined,
      tableStyles,
      false,
    ),
    false,
  );
  return {
    path: relPath,
    masterPath,
    showMasterShapes: attr(doc.documentElement, 'showMasterSp') !== '0',
    placeholders,
    textPresets,
    background,
    elements,
  };
}

/** 读取 PPTX 母版、版式和占位符继承关系。 */
export function readPresentationLayouts(
  entries: OfficeEntryMap,
  packageState: PackageState,
  theme: ThemeModel,
  tableStyles?: TableStyleMap,
) {
  // 先读取 master，再读取它关联的 layout；后续 slide 会按 layout/master 继承占位符和默认样式。
  const presentationRels =
    packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const masterDefinitions: MasterDefinition[] = [];
  const masterLayoutDefinitions: Record<string, LayoutDefinition[]> = {};

  Object.entries(presentationRels).forEach(([, relationship]) => {
    const target = relationship.target;
    if (!target.includes('slideMasters/')) return;
    const xmlPath = target.startsWith('ppt/') ? target : `ppt/${target}`;
    const relPath = xmlPath
      .replace(/^ppt\/slideMasters\//, 'ppt/slideMasters/_rels/')
      .replace(/\.xml$/, '.xml.rels');
    const masterRels = packageState.relationships[relPath] ?? {};
    const master = readMaster(
      readXml(entries, xmlPath),
      theme,
      xmlPath,
      packageState,
      masterRels,
      tableStyles,
    );
    masterDefinitions.push(master);
    masterLayoutDefinitions[xmlPath] = Object.values(masterRels)
      .map((relationship) => relationship.target)
      .filter((item) => item.includes('slideLayouts/'))
      .map((layoutTarget) => {
        const layoutPath = layoutTarget.startsWith('ppt/')
          ? layoutTarget
          : `ppt/${layoutTarget}`;
        const layoutRelPath = layoutPath
          .replace(/^ppt\/slideLayouts\//, 'ppt/slideLayouts/_rels/')
          .concat('.rels');
        return readLayout(
          readXml(entries, layoutPath),
          theme,
          layoutPath,
          xmlPath,
          packageState,
          packageState.relationships[layoutRelPath] ?? {},
          tableStyles,
        );
      });
  });

  return { masterDefinitions, masterLayoutDefinitions };
}
