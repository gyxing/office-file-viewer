import { openOfficeArchive } from '../../shared/ooxml/archive';
import { readRelationships } from '../../shared/ooxml/relationships';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  parseXml,
} from '../../shared/ooxml/xml';
import { createBrowserSafeSfnt } from '../fonts/createBrowserSafeSfnt';

/** 同时解压和编译过多字体会形成明显的主线程与内存峰值。 */
const DOCX_EMBEDDED_FONT_LOAD_CONCURRENCY = 3;

/** 为并行存在的文档生成互不覆盖的内嵌字体族。 */
let embeddedFontSessionSequence = 0;

/** OOXML 内嵌字体声明与 CSS 字体变体的对应关系。 */
const DOCX_EMBEDDED_FONT_VARIANTS = [
  { element: 'embedRegular', style: 'normal', weight: '400' },
  { element: 'embedBold', style: 'normal', weight: '700' },
  { element: 'embedItalic', style: 'italic', weight: '400' },
  { element: 'embedBoldItalic', style: 'italic', weight: '700' },
] as const;

/** 从 DOCX 字体表读取的单个内嵌字体变体。 */
type DocxEmbeddedFontDescriptor = {
  /** 文档样式引用的字体族名称。 */
  family: string;
  /** 字体文件在 OOXML 包内的规范路径。 */
  path: string;
  /** 用于还原 ODTTF 前 32 字节的 OOXML 字体密钥。 */
  key: Uint8Array;
  /** CSS 字体样式。 */
  style: 'normal' | 'italic';
  /** CSS 字重。 */
  weight: '400' | '700';
};

/** 当前文档注册到 Font Loading API 的字体会话。 */
export type DocxEmbeddedFontSession = {
  /** 源字体名到当前文档内嵌字体族的映射。 */
  aliases: Readonly<Record<string, string>>;
  /** 从所属文档移除本次注册的全部字体。 */
  dispose(): void;
};

function createAbortError() {
  const error = new Error('DOCX 内嵌字体加载已取消');
  error.name = 'AbortError';
  return error;
}

function throwIfAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createAbortError();
}

/** 将 OOXML GUID 字符串转换为字体解密使用的 16 字节密钥。 */
function parseFontKey(value?: string) {
  const hex = value?.replace(/[{}-]/g, '');
  if (!hex || !/^[0-9a-f]{32}$/i.test(hex)) return undefined;
  return Uint8Array.from({ length: 16 }, (_, index) =>
    Number.parseInt(hex.slice(index * 2, index * 2 + 2), 16),
  );
}

/** 按 ECMA-376 规则还原 ODTTF 被异或混淆的文件头。 */
function deobfuscateFont(bytes: Uint8Array, key: Uint8Array) {
  const decoded = bytes.slice();
  const limit = Math.min(32, decoded.length);
  for (let index = 0; index < limit; index += 1) {
    decoded[index] ^= key[15 - (index % 16)];
  }
  return decoded;
}

/** 解析字体表及关系文件，只保留可以从当前包读取的内嵌字体。 */
function readEmbeddedFontDescriptors(
  fontTableXml: string,
  relationshipXml: string,
  hasEntry: (path: string) => boolean,
) {
  const relationships = readRelationships(
    relationshipXml,
    'word/_rels/fontTable.xml.rels',
  );
  const root = parseXml(fontTableXml).documentElement;
  const descriptors: DocxEmbeddedFontDescriptor[] = [];
  const identities = new Set<string>();

  childrenByLocalName(root, 'font').forEach((fontNode) => {
    const family = attr(fontNode, 'w:name') ?? attr(fontNode, 'name');
    if (!family) return;
    DOCX_EMBEDDED_FONT_VARIANTS.forEach((variant) => {
      const embed = childByLocalName(fontNode, variant.element);
      const relationId = attr(embed, 'r:id') ?? attr(embed, 'id');
      const key = parseFontKey(
        attr(embed, 'w:fontKey') ?? attr(embed, 'fontKey'),
      );
      const path = relationId ? relationships[relationId]?.target : undefined;
      if (!path || !key || !hasEntry(path)) return;
      const identity = `${family}\u0000${variant.style}\u0000${variant.weight}`;
      if (identities.has(identity)) return;
      identities.add(identity);
      descriptors.push({
        family,
        path,
        key,
        style: variant.style,
        weight: variant.weight,
      });
    });
  });
  return descriptors;
}

/** 加载并注册 DOCX 包内嵌字体；不支持 Font Loading API 时静默跳过。 */
export async function loadDocxEmbeddedFonts(
  file: File,
  ownerDocument: Document,
  signal?: AbortSignal,
): Promise<DocxEmbeddedFontSession> {
  const FontFaceConstructor =
    typeof FontFace === 'undefined' ? undefined : FontFace;
  const fontSet = ownerDocument.fonts;
  if (!FontFaceConstructor || !fontSet) {
    return { aliases: {}, dispose: () => undefined };
  }

  throwIfAborted(signal);
  const reader = await openOfficeArchive(file, { signal });
  const loadedFaces: FontFace[] = [];
  const loadedAliases: Record<string, string> = {};
  try {
    if (
      !reader.has('word/fontTable.xml') ||
      !reader.has('word/_rels/fontTable.xml.rels')
    ) {
      return { aliases: {}, dispose: () => undefined };
    }
    const [fontTableXml, relationshipXml] = await Promise.all([
      reader.readText('word/fontTable.xml', signal),
      reader.readText('word/_rels/fontTable.xml.rels', signal),
    ]);
    const descriptors = readEmbeddedFontDescriptors(
      fontTableXml,
      relationshipXml,
      (path) => reader.has(path),
    );
    embeddedFontSessionSequence += 1;
    const sessionPrefix = `office-file-viewer-embedded-${embeddedFontSessionSequence}`;
    const familyAliases = new Map<string, string>();
    descriptors.forEach((descriptor) => {
      const identity = descriptor.family.toLocaleLowerCase();
      if (!familyAliases.has(identity)) {
        familyAliases.set(
          identity,
          `${sessionPrefix}-${familyAliases.size + 1}`,
        );
      }
    });
    let nextIndex = 0;
    const loadNext = async () => {
      while (nextIndex < descriptors.length) {
        throwIfAborted(signal);
        const descriptor = descriptors[nextIndex];
        const familyIdentity = descriptor.family.toLocaleLowerCase();
        const embeddedFamily = familyAliases.get(familyIdentity)!;
        nextIndex += 1;
        let decoded: Uint8Array | undefined;
        try {
          const encoded = await reader.readBinary(descriptor.path, signal);
          decoded = deobfuscateFont(encoded, descriptor.key);
          const face = new FontFaceConstructor(
            embeddedFamily,
            decoded.buffer as ArrayBuffer,
            {
              style: descriptor.style,
              weight: descriptor.weight,
            },
          );
          const loaded = await face.load();
          throwIfAborted(signal);
          fontSet.add(loaded);
          loadedFaces.push(loaded);
          loadedAliases[familyIdentity] = embeddedFamily;
        } catch {
          if (signal?.aborted) throw createAbortError();
          const browserSafeFont = decoded
            ? createBrowserSafeSfnt(decoded)
            : undefined;
          if (browserSafeFont) {
            try {
              const fallbackFace = new FontFaceConstructor(
                embeddedFamily,
                browserSafeFont.buffer as ArrayBuffer,
                { style: descriptor.style, weight: descriptor.weight },
              );
              const loaded = await fallbackFace.load();
              fontSet.add(loaded);
              loadedFaces.push(loaded);
              loadedAliases[familyIdentity] = embeddedFamily;
              continue;
            } catch {
              // 精简字体仍被拒绝时继续使用宿主字体映射或系统回退字体。
            }
          }
          // 单个字体损坏或浏览器拒绝该字体时保留其余字体和系统回退链。
        }
      }
    };
    await Promise.all(
      Array.from(
        {
          length: Math.min(
            DOCX_EMBEDDED_FONT_LOAD_CONCURRENCY,
            descriptors.length,
          ),
        },
        () => loadNext(),
      ),
    );
  } catch (error) {
    loadedFaces.forEach((face) => fontSet.delete(face));
    throw error;
  } finally {
    await reader.close();
  }

  let disposed = false;
  return {
    aliases: loadedAliases,
    dispose() {
      if (disposed) return;
      disposed = true;
      loadedFaces.forEach((face) => fontSet.delete(face));
    },
  };
}
