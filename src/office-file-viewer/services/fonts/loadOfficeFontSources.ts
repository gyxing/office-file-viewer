import type { OfficeFileViewerFontSource } from './types';

/** 单个宿主字体资源加载失败时保留的诊断信息。 */
export type OfficeFontSourceLoadFailure = {
  /** 加载失败的字体资源声明。 */
  source: OfficeFileViewerFontSource;
  /** 浏览器或资源请求返回的原始错误。 */
  error: unknown;
};

/** 一组宿主字体资源的异步加载结果。 */
export type OfficeFontSourceLoadResult = {
  /** 已成功注册到文档字体集合的资源数量。 */
  loaded: number;
  /** 未能注册的资源及其原因。 */
  failures: readonly OfficeFontSourceLoadFailure[];
};

/** 与一个预览会话绑定的宿主字体资源生命周期。 */
export type OfficeFontSourceSession = {
  /** 所有资源完成尝试后兑现，单个失败不会令 Promise 拒绝。 */
  ready: Promise<OfficeFontSourceLoadResult>;
  /** 移除本次注册的字体并释放临时 Blob URL。 */
  dispose(): void;
};

type FontFaceConstructor = new (
  family: string,
  source: string,
  descriptors?: FontFaceDescriptors,
) => FontFace;

type ObjectUrlApi = {
  createObjectURL: (object: Blob) => string;
  revokeObjectURL?: (url: string) => void;
};

/** 为 Blob 字体对象生成进程内稳定身份，避免内联配置重复加载。 */
let fontSourceObjectSequence = 0;
/** 保存 Blob 对象身份但不延长其生命周期。 */
const fontSourceObjectIds = new WeakMap<object, number>();

function getSourceIdentity(src: string | Blob) {
  if (typeof src === 'string') return `string:${src}`;
  let identity = fontSourceObjectIds.get(src);
  if (identity === undefined) {
    fontSourceObjectSequence += 1;
    identity = fontSourceObjectSequence;
    fontSourceObjectIds.set(src, identity);
  }
  return `object:${identity}`;
}

function normalizeSourceFamily(family: string) {
  const value = family.trim();
  if (!value) throw new Error('字体族名称不能为空');
  return value;
}

/** 生成稳定的字体资源签名，避免宿主每次渲染创建数组时重复加载字体。 */
export function getOfficeFontSourcesKey(
  sources: readonly OfficeFileViewerFontSource[] | undefined,
) {
  return (sources ?? [])
    .map(
      (source) =>
        `${source.family.trim().toLocaleLowerCase()}\u0000${getSourceIdentity(
          source.src,
        )}\u0000${source.weight ?? ''}\u0000${source.style ?? ''}\u0000${
          source.stretch ?? ''
        }\u0000${source.unicodeRange ?? ''}`,
    )
    .join('\u0001');
}

function getFontFaceConstructor(ownerDocument: Document) {
  const ownerWindow = ownerDocument.defaultView as
    | (Window & { FontFace?: FontFaceConstructor })
    | null;
  if (ownerWindow?.FontFace) return ownerWindow.FontFace;
  return typeof FontFace === 'undefined'
    ? undefined
    : (FontFace as unknown as FontFaceConstructor);
}

function getObjectUrlApi(ownerDocument: Document) {
  const ownerWindow = ownerDocument.defaultView;
  if (ownerWindow?.URL?.createObjectURL) {
    return ownerWindow.URL as unknown as ObjectUrlApi;
  }
  const globalUrl =
    typeof URL === 'undefined'
      ? undefined
      : (URL as unknown as Partial<ObjectUrlApi>);
  if (globalUrl?.createObjectURL) return globalUrl as ObjectUrlApi;
  return undefined;
}

/** 将宿主传入的原始 URL 转换为 FontFace 可接受的 CSS source。 */
function normalizeStringSource(source: string) {
  const value = source.trim();
  if (!value) throw new Error('字体资源地址不能为空');
  if (/^(?:url|local)\s*\(/i.test(value)) return value;
  return `url(${JSON.stringify(value)})`;
}

/** 生成 FontFace 描述符，同时保留宿主未提供的默认值。 */
function createDescriptors(source: OfficeFileViewerFontSource) {
  const descriptors: FontFaceDescriptors = {};
  if (source.weight !== undefined) descriptors.weight = String(source.weight);
  if (source.style !== undefined) descriptors.style = source.style;
  if (source.stretch !== undefined) descriptors.stretch = source.stretch;
  if (source.unicodeRange !== undefined) {
    descriptors.unicodeRange = source.unicodeRange;
  }
  return descriptors;
}

/** 为当前预览会话加载宿主字体资源，并在销毁时撤销注册。 */
export function createOfficeFontSourceSession(
  sources: readonly OfficeFileViewerFontSource[] | undefined,
  ownerDocument: Document,
): OfficeFontSourceSession {
  const loadedFaces: FontFace[] = [];
  const objectUrls: string[] = [];
  let disposed = false;
  const fontSet = ownerDocument.fonts;
  const FontFaceCtor = getFontFaceConstructor(ownerDocument);
  const objectUrlApi = getObjectUrlApi(ownerDocument);
  const uniqueSources: OfficeFileViewerFontSource[] = [];
  (sources ?? []).forEach((source) => {
    const duplicate = uniqueSources.some(
      (current) =>
        current.family.trim().toLocaleLowerCase() ===
          source.family.trim().toLocaleLowerCase() &&
        current.src === source.src &&
        current.weight === source.weight &&
        current.style === source.style &&
        current.stretch === source.stretch &&
        current.unicodeRange === source.unicodeRange,
    );
    if (!duplicate) uniqueSources.push(source);
  });

  const ready = (async (): Promise<OfficeFontSourceLoadResult> => {
    if (!uniqueSources.length || !FontFaceCtor || !fontSet) {
      // 旧浏览器没有 Font Loading API 时保留原有字体回退行为，不制造伪错误。
      return { loaded: 0, failures: [] };
    }
    const failures: OfficeFontSourceLoadFailure[] = [];
    let loaded = 0;
    for (const source of uniqueSources) {
      if (disposed) break;
      let cssSource: string;
      try {
        if (typeof source.src === 'string') {
          cssSource = normalizeStringSource(source.src);
        } else {
          if (!objectUrlApi) throw new Error('浏览器不支持 Blob 字体资源');
          const objectUrl = objectUrlApi.createObjectURL(source.src);
          objectUrls.push(objectUrl);
          cssSource = `url(${JSON.stringify(objectUrl)})`;
        }
        const face = new FontFaceCtor(
          normalizeSourceFamily(source.family),
          cssSource,
          createDescriptors(source),
        );
        const loadedFace = await face.load();
        if (disposed) break;
        fontSet.add(loadedFace);
        loadedFaces.push(loadedFace);
        loaded += 1;
      } catch (error) {
        if (!disposed) failures.push({ source, error });
      }
    }
    return { loaded, failures };
  })();

  return {
    ready,
    dispose() {
      if (disposed) return;
      disposed = true;
      loadedFaces.forEach((face) => {
        if (fontSet && typeof fontSet.delete === 'function') {
          fontSet.delete(face);
        }
      });
      objectUrls.forEach((objectUrl) =>
        objectUrlApi?.revokeObjectURL?.(objectUrl),
      );
    },
  };
}
