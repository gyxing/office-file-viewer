import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readBinary, readXml } from '../../shared/ooxml/archive';
import { collectMedia } from '../../shared/ooxml/media';
import { readRelationships } from '../../shared/ooxml/relationships';
import { readOfficeTheme } from '../../shared/ooxml/theme';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantsByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import type { OfficeArchiveResourcePolicy } from '../../shared/resource/OfficeResourcePolicy';
import type { OfficeFormatParser } from '../parsing/formatParserRegistry';
import { throwIfParseAborted } from '../parsing/runtime/types';
import { loadXlsxEntries } from './archive';
import { xlsxImageBytesToDataUrl } from './createXlsxImageResource';
import {
  parseMaterializedXlsxSheet,
  readMaterializedXlsxSharedStrings,
  type MaterializedXlsxPackageState,
} from './parseXlsxSheet';
import type { XlsxSheet, XlsxWorkbook } from './types';
import { decodeMojibake, parseStyles } from './xlsxCellFormatting';

// XLSX 中工作表、drawing、chart、media 分散在不同 XML，通过关系表统一解析引用路径。
function buildPackageState(
  entries: OfficeEntryMap,
): MaterializedXlsxPackageState {
  const relationships: MaterializedXlsxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = collectMedia(entries, 'xl/media/');
  // 浏览器不原生支持 metafile；物化路径在建包时一次转换，按需路径则延迟到图片可见时转换。
  for (const [path] of entries) {
    if (!/^xl\/media\/.*\.(?:emf|wmf)$/i.test(path)) continue;
    const binary = readBinary(entries, path);
    if (!binary) continue;
    try {
      const dataUrl = xlsxImageBytesToDataUrl(path, binary);
      media.byPath[path] = dataUrl;
      media.byName[path.split('/').pop() ?? path] = dataUrl;
    } catch {
      // 个别损坏的 metafile 保留原资源，不能阻断其余工作表解析。
    }
  }

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
    theme: readOfficeTheme(readXml(entries, 'xl/theme/theme1.xml')),
  };
}

function throwIfXlsxParseAborted(signal?: AbortSignal) {
  if (!signal?.aborted) return;
  const error = new Error('XLSX 解析已取消');
  error.name = 'AbortError';
  throw error;
}

/** 在工作表边界让出主线程，使大工作簿切换或卸载可以及时取消。 */
async function xlsxParseCheckpoint(signal?: AbortSignal) {
  throwIfXlsxParseAborted(signal);
  await new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
  throwIfXlsxParseAborted(signal);
}

/** 解析 XLSX 包并返回标准工作簿模型。 */
export async function parseXlsx(
  file: File,
  signal?: AbortSignal,
  resourcePolicy?: OfficeArchiveResourcePolicy,
): Promise<XlsxWorkbook> {
  // sharedStrings 和 styles 是全工作簿共享数据，先解析后再逐个 sheet 套用。
  throwIfXlsxParseAborted(signal);
  const entries = await loadXlsxEntries(file, { signal, resourcePolicy });
  await xlsxParseCheckpoint(signal);
  const packageState = buildPackageState(entries);
  const workbookXml = readXml(entries, 'xl/workbook.xml');
  const workbookRels =
    packageState.relationships['xl/_rels/workbook.xml.rels'] ?? {};
  const sharedStrings = readMaterializedXlsxSharedStrings(
    readXml(entries, 'xl/sharedStrings.xml'),
  );
  const styleBook = parseStyles(
    readXml(entries, 'xl/styles.xml'),
    packageState.theme,
  );
  const workbookDoc = parseXml(workbookXml);
  const definedNames = Object.fromEntries(
    descendantsByLocalName(workbookDoc.documentElement, 'definedName')
      .map((node) => [attr(node, 'name'), node.textContent?.trim()] as const)
      .filter(
        (entry): entry is readonly [string, string] =>
          Boolean(entry[0] && entry[1]),
      ),
  );
  const sheetEntries = childrenByLocalName(
    childByLocalName(workbookDoc.documentElement, 'sheets'),
    'sheet',
  )
    .map((node, sourceIndex) => ({ node, sourceIndex }))
    .filter(({ node }) => {
      // Excel/WPS 的隐藏工作表不属于用户可见预览内容。
      const state = attr(node, 'state');
      return state !== 'hidden' && state !== 'veryHidden';
    });
  const sheets: XlsxSheet[] = [];
  for (let index = 0; index < sheetEntries.length; index += 1) {
    await xlsxParseCheckpoint(signal);
    const { node: sheetNode, sourceIndex } = sheetEntries[index];
    const relId = attr(sheetNode, 'r:id') ?? attr(sheetNode, 'id') ?? '';
    const rel = workbookRels[relId];
    const path = rel?.target ?? `xl/worksheets/sheet${sourceIndex + 1}.xml`;
    sheets.push(
      parseMaterializedXlsxSheet(
        readXml(entries, path),
        {
          id: attr(sheetNode, 'sheetId') ?? String(sourceIndex + 1),
          name: decodeMojibake(
            attr(sheetNode, 'name') ?? `Sheet ${sourceIndex + 1}`,
          ),
          path,
        },
        sharedStrings,
        styleBook,
        packageState,
      ),
    );
  }

  throwIfXlsxParseAborted(signal);
  return { sheets, definedNames };
}

/** 通过统一运行时合同解析 XLSX，并输出完整工作簿模型。 */
export const runXlsxParser: OfficeFormatParser = async (
  file,
  { signal, resourcePolicy },
  sink,
) => {
  sink.progress({
    stage: 'content',
    percent: 0.05,
    message: '正在解析文件',
  });
  const workbook = await parseXlsx(file, signal, resourcePolicy);
  throwIfParseAborted(signal);
  await sink.parsed({ kind: 'xlsx', workbook });
  await sink.complete();
};
