/** DOC/WPS 样式表中与大纲语义有关的最小样式定义。 */
type DocOutlineStyleDefinition = {
  /** 样式在 STSH.rglpstd 中的索引。 */
  id: number;
  /** StdfBase.sti 提供的跨语言内置样式标识。 */
  invariantStyleId: number;
  /** StdfBase.stk，1 表示段落样式。 */
  kind: number;
  /** 样式继承的父样式索引。 */
  baseStyleId?: number;
  /** 源样式名，仅用于识别 Word 的 TOC 样式。 */
  name?: string;
  /** 样式直接声明的大纲级别。 */
  outlineLevel?: number;
};

/** DOC/WPS 大纲解析使用的样式目录。 */
export type DocStyleOutlineCatalog = {
  /** 按 istd 索引的样式定义。 */
  styles: Map<number, DocOutlineStyleDefinition>;
  /** 样式继承后的大纲级别缓存。 */
  outlineLevelCache: Map<number, number | null>;
  /** 样式继承后的 TOC 判定缓存。 */
  tocStyleCache: Map<number, boolean>;
  /** 不阻断正文解析的样式表警告。 */
  warnings: string[];
};

/** 创建空样式目录，让异常 DOC 仍可继续按正文降级预览。 */
function createEmptyCatalog(): DocStyleOutlineCatalog {
  return {
    styles: new Map(),
    outlineLevelCache: new Map(),
    tocStyleCache: new Map(),
    warnings: [],
  };
}

/** 安全读取小端 16 位无符号整数。 */
function readUint16(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 2 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getUint16(offset, true);
}

/** 安全读取小端 16 位有符号整数。 */
function readInt16(bytes: Uint8Array, offset: number) {
  if (offset < 0 || offset + 2 > bytes.length) return undefined;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset,
    bytes.byteLength,
  ).getInt16(offset, true);
}

/** 按 SPRM 的 spra 位读取操作数长度。 */
function readSprmOperandSize(sprm: number, bytes: Uint8Array, offset: number) {
  const sizeCode = (sprm >> 13) & 0x7;
  if (sizeCode === 0 || sizeCode === 1) return 1;
  if (sizeCode === 2 || sizeCode === 4 || sizeCode === 5) return 2;
  if (sizeCode === 3) return 4;
  if (sizeCode === 6 || sizeCode === 7) return 1 + (bytes[offset] ?? 0);
  return 0;
}

/** 遍历段落 SPRM，并按出现顺序更新大纲级别。 */
function applyParagraphOutlineSprms(
  bytes: Uint8Array,
  initialStyleId: number,
  initialLevel: number,
  resolveStyleLevel: (styleId: number) => number,
) {
  let styleId = initialStyleId;
  let level = initialLevel;
  let didSetOutline = false;
  let offset = 0;

  while (offset + 2 <= bytes.length) {
    const sprm = readUint16(bytes, offset);
    if (sprm === undefined) break;
    offset += 2;
    const operandSize = readSprmOperandSize(sprm, bytes, offset);
    if (!operandSize || offset + operandSize > bytes.length) break;

    if (sprm === 0x4600 && operandSize >= 2) {
      styleId = readUint16(bytes, offset) ?? styleId;
      level = resolveStyleLevel(styleId);
    } else if (sprm === 0x2640 && !(styleId >= 1 && styleId <= 9)) {
      const candidate = bytes[offset];
      if (candidate !== undefined && candidate <= 9) {
        level = candidate;
        didSetOutline = true;
      }
    } else if (sprm === 0x2602) {
      const increment = new DataView(
        bytes.buffer,
        bytes.byteOffset + offset,
        operandSize,
      ).getInt8(0);
      if (styleId >= 1 && styleId <= 9) {
        styleId = Math.min(9, Math.max(1, styleId + increment));
        level = styleId - 1;
      } else if (level !== 9) {
        level = Math.min(9, Math.max(0, level + increment));
      }
      didSetOutline = true;
    }
    offset += operandSize;
  }

  return { styleId, level, didSetOutline };
}

/** Word 的目录样式名来自 STSH，而不是正文文本，因此可安全排除。 */
function isDirectTocStyle(style: DocOutlineStyleDefinition) {
  const primaryName = style.name?.split(',')[0].trim() ?? '';
  return /^(?:toc|目录)\s*[1-9]$/i.test(primaryName);
}

/** 沿样式继承链解析最终大纲级别。 */
function resolveStyleOutlineLevel(
  styleId: number,
  catalog: DocStyleOutlineCatalog,
  seen: Set<number> = new Set(),
): number {
  if (styleId >= 1 && styleId <= 9) return styleId - 1;
  const cached = catalog.outlineLevelCache.get(styleId);
  if (cached !== undefined) return cached ?? 9;
  if (seen.has(styleId)) {
    catalog.outlineLevelCache.set(styleId, null);
    return 9;
  }
  const style = catalog.styles.get(styleId);
  if (!style || style.kind !== 1) {
    catalog.outlineLevelCache.set(styleId, null);
    return 9;
  }
  seen.add(styleId);
  const level =
    style.outlineLevel ??
    (style.baseStyleId !== undefined
      ? resolveStyleOutlineLevel(style.baseStyleId, catalog, seen)
      : 9);
  catalog.outlineLevelCache.set(styleId, level === 9 ? null : level);
  return level;
}

/** 沿样式继承链判断是否属于 Word 自动目录样式。 */
function resolveTocStyle(
  styleId: number,
  catalog: DocStyleOutlineCatalog,
  seen: Set<number> = new Set(),
): boolean {
  const cached = catalog.tocStyleCache.get(styleId);
  if (cached !== undefined) return cached;
  if (seen.has(styleId)) {
    catalog.tocStyleCache.set(styleId, false);
    return false;
  }
  const style = catalog.styles.get(styleId);
  if (!style) {
    catalog.tocStyleCache.set(styleId, false);
    return false;
  }
  seen.add(styleId);
  const isToc =
    isDirectTocStyle(style) ||
    (style.baseStyleId !== undefined &&
      resolveTocStyle(style.baseStyleId, catalog, seen));
  catalog.tocStyleCache.set(styleId, isToc);
  return isToc;
}

/** 从一个 LPStd 负载读取段落样式的大纲元数据。 */
function parseStyleDefinition(
  data: Uint8Array,
  styleId: number,
  stdStart: number,
  stdLength: number,
  baseLength: number,
): DocOutlineStyleDefinition | undefined {
  const stdEnd = stdStart + stdLength;
  if (stdLength < baseLength || baseLength < 10 || stdEnd > data.length)
    return undefined;

  const invariantAndFlags = readUint16(data, stdStart);
  const kindAndBase = readUint16(data, stdStart + 2);
  const countAndNext = readUint16(data, stdStart + 4);
  if (
    invariantAndFlags === undefined ||
    kindAndBase === undefined ||
    countAndNext === undefined
  )
    return undefined;

  const invariantStyleId = invariantAndFlags & 0x0fff;
  const kind = kindAndBase & 0x000f;
  const rawBaseStyleId = kindAndBase >>> 4;
  const baseStyleId = rawBaseStyleId === 0x0fff ? undefined : rawBaseStyleId;
  const formattingSetCount = countAndNext & 0x000f;
  const nameOffset = stdStart + baseLength;
  const nameLength = readUint16(data, nameOffset);
  if (nameLength === undefined) return undefined;
  const nameByteLength = nameLength * 2;
  const nameStart = nameOffset + 2;
  const nameEnd = nameStart + nameByteLength;
  if (nameEnd + 2 > stdEnd) return undefined;
  const name = new TextDecoder('utf-16le')
    .decode(data.slice(nameStart, nameEnd))
    .replace(/\u0000+$/g, '');

  const definition: DocOutlineStyleDefinition = {
    id: styleId,
    invariantStyleId,
    kind,
    baseStyleId,
    name,
  };
  if (kind !== 1 || formattingSetCount < 1) return definition;

  const papxLengthOffset = nameEnd + 2;
  const papxLength = readUint16(data, papxLengthOffset);
  if (papxLength === undefined) return definition;
  const papxStart = papxLengthOffset + 2;
  const papxEnd = papxStart + papxLength;
  if (papxEnd > stdEnd || papxLength < 2) return definition;

  const initialLevel = styleId >= 1 && styleId <= 9 ? styleId - 1 : 9;
  const parsed = applyParagraphOutlineSprms(
    data.slice(papxStart + 2, papxEnd),
    styleId,
    initialLevel,
    () => 9,
  );
  if (parsed.didSetOutline && parsed.level >= 0 && parsed.level <= 9)
    definition.outlineLevel = parsed.level;
  return definition;
}

/**
 * 从 FIB 指向的 STSH 读取样式大纲目录。
 * 解析失败只记录警告，不能阻断旧版文档的正文降级预览。
 */
export function parseDocStyleOutlineCatalog(
  tableStream: Uint8Array,
  fcStshf: number,
  lcbStshf: number,
): DocStyleOutlineCatalog {
  const catalog = createEmptyCatalog();
  if (!fcStshf || !lcbStshf) return catalog;
  if (fcStshf < 0 || fcStshf + lcbStshf > tableStream.length) {
    catalog.warnings.push('DOC 样式表范围超出 Table 流，已忽略大纲样式。');
    return catalog;
  }

  const data = tableStream.slice(fcStshf, fcStshf + lcbStshf);
  const headerLength = readUint16(data, 0);
  if (headerLength === undefined || headerLength < 4) {
    catalog.warnings.push('DOC 样式表头无效，已忽略大纲样式。');
    return catalog;
  }
  const styleCount = readUint16(data, 2) ?? 0;
  const baseLength = readUint16(data, 4) ?? 0;
  if (baseLength < 10 || 2 + headerLength > data.length) {
    catalog.warnings.push('DOC 样式表基础结构无效，已忽略大纲样式。');
    return catalog;
  }

  let offset = 2 + headerLength;
  for (let styleId = 0; styleId < styleCount; styleId += 1) {
    const stdLength = readInt16(data, offset);
    if (stdLength === undefined || stdLength < 0) {
      catalog.warnings.push('DOC 样式定义长度无效，已停止读取后续样式。');
      break;
    }
    const stdStart = offset + 2;
    if (stdStart + stdLength > data.length) {
      catalog.warnings.push('DOC 样式定义超出样式表范围，已停止读取后续样式。');
      break;
    }
    if (stdLength > 0) {
      const definition = parseStyleDefinition(
        data,
        styleId,
        stdStart,
        stdLength,
        baseLength,
      );
      if (definition) catalog.styles.set(styleId, definition);
    }
    offset = stdStart + stdLength + (stdLength % 2);
  }
  return catalog;
}

/**
 * 读取 PAPX 的最终大纲级别。
 * Heading 1～9 的固定 istd 优先级高于 sprmPOutLvl，符合 MS-DOC 规则。
 */
export function readDocParagraphOutlineLevel(
  grpprl: Uint8Array,
  catalog: DocStyleOutlineCatalog,
): number | undefined {
  const initialStyleId = readUint16(grpprl, 0);
  if (initialStyleId === undefined) return undefined;
  const initialLevel = resolveStyleOutlineLevel(initialStyleId, catalog);
  const parsed = applyParagraphOutlineSprms(
    grpprl.slice(Math.min(2, grpprl.length)),
    initialStyleId,
    initialLevel,
    (styleId) => resolveStyleOutlineLevel(styleId, catalog),
  );
  if (resolveTocStyle(parsed.styleId, catalog)) return undefined;
  return parsed.level >= 0 && parsed.level <= 8 ? parsed.level : undefined;
}
