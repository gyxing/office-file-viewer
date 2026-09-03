import type {
  ImageElement,
  PresentationImagePreviewMetadata,
  PresentationImageSourceKind,
  SlideElement,
} from './types';

/** 方形小图标允许的最大边长，单位为标准化渲染像素。 */
const COMPACT_ICON_MAX_SIZE = 48;
/** 狭长小图标允许的最大短边，单位为标准化渲染像素。 */
const NARROW_ICON_MAX_SHORT_SIDE = 24;
/** 狭长小图标允许的最大长边，单位为标准化渲染像素。 */
const NARROW_ICON_MAX_LONG_SIDE = 128;
/** 需要结合语义信号继续判断的中等图标最大短边。 */
const RELAXED_ICON_MAX_SHORT_SIDE = 96;
/** 需要结合语义信号继续判断的中等图标最大长边。 */
const RELAXED_ICON_MAX_LONG_SIDE = 192;
/** 轻量重复资源的最大字节数，避免把重复使用的大图误判为图标。 */
const SMALL_REUSED_RESOURCE_MAX_BYTES = 96 * 1024;
/** 中等尺寸 PNG/GIF/WebP 图片作为图标候选时的最大字节数。 */
const SMALL_RASTER_ICON_MAX_BYTES = 32 * 1024;
/** 图片名称中可作为图标提示的中英文语义词。 */
const ICON_SEMANTIC_PATTERN =
  /(?:icon|glyph|pictogram|symbol|arrow|button|logo|图标|箭头|按钮|符号|徽标|标志)/iu;

/** 根据源路径或 SVG 扩展节点归类演示文稿图片资源。 */
export function detectPresentationImageSourceKind(
  target?: string,
  hasSvgBlip = false,
  mimeType?: string,
): PresentationImageSourceKind {
  if (
    hasSvgBlip ||
    /\.(?:svg|emf|wmf|pict)(?:[?#]|$)/iu.test(target ?? '') ||
    /image\/(?:svg\+xml|x-(?:emf|wmf))/iu.test(
      `${target ?? ''} ${mimeType ?? ''}`,
    )
  ) {
    return 'vector';
  }
  if (
    /\.(?:png|jpe?g|gif|webp|bmp|dib|tiff?)(?:[?#]|$)/iu.test(target ?? '') ||
    /image\/(?:png|jpe?g|gif|webp|bmp|tiff?)/iu.test(
      `${target ?? ''} ${mimeType ?? ''}`,
    )
  ) {
    return 'raster';
  }
  return 'unknown';
}

function hasIconSemantic(
  metadata?: PresentationImagePreviewMetadata,
  alt?: string,
) {
  return ICON_SEMANTIC_PATTERN.test(
    [metadata?.objectName, metadata?.objectDescription, alt]
      .filter(Boolean)
      .join(' '),
  );
}

function isCompactIcon(width: number, height: number) {
  const shortSide = Math.min(width, height);
  const longSide = Math.max(width, height);
  return (
    (width <= COMPACT_ICON_MAX_SIZE && height <= COMPACT_ICON_MAX_SIZE) ||
    (shortSide <= NARROW_ICON_MAX_SHORT_SIDE &&
      longSide <= NARROW_ICON_MAX_LONG_SIDE)
  );
}

function isRelaxedIconBounds(width: number, height: number) {
  const shortSide = Math.min(width, height);
  const longSide = Math.max(width, height);
  return (
    shortSide <= RELAXED_ICON_MAX_SHORT_SIDE &&
    longSide <= RELAXED_ICON_MAX_LONG_SIDE
  );
}

function isLikelyIcon(
  width: number,
  height: number,
  metadata?: PresentationImagePreviewMetadata,
  sourceKind?: PresentationImageSourceKind,
  alt?: string,
  mimeType?: string,
  resourceSize?: number,
) {
  if (isCompactIcon(width, height)) return true;
  if (!isRelaxedIconBounds(width, height)) return false;
  if (hasIconSemantic(metadata, alt)) return true;
  if (sourceKind === 'vector') return true;
  if (
    sourceKind === 'raster' &&
    /image\/(?:png|gif|webp)/iu.test(mimeType ?? '') &&
    resourceSize !== undefined &&
    resourceSize <= SMALL_RASTER_ICON_MAX_BYTES
  ) {
    return true;
  }
  return Boolean(
    metadata?.resourceReuseCount &&
      metadata.resourceReuseCount >= 2 &&
      resourceSize !== undefined &&
      resourceSize <= SMALL_REUSED_RESOURCE_MAX_BYTES,
  );
}

function inferSourceKind(
  element: Pick<ImageElement, 'src' | 'previewMetadata'>,
) {
  if (
    element.previewMetadata?.sourceKind &&
    element.previewMetadata.sourceKind !== 'unknown'
  ) {
    return element.previewMetadata.sourceKind;
  }
  if (typeof element.src === 'string') {
    return detectPresentationImageSourceKind(element.src);
  }
  if (element.src.kind === 'lazy') {
    return detectPresentationImageSourceKind(
      undefined,
      false,
      element.src.mimeType,
    );
  }
  return detectPresentationImageSourceKind(element.src.url);
}

function inferResourceSize(
  element: Pick<ImageElement, 'src' | 'previewMetadata'>,
) {
  if (element.previewMetadata?.resourceSize !== undefined) {
    return element.previewMetadata.resourceSize;
  }
  return element.src &&
    typeof element.src !== 'string' &&
    element.src.kind === 'lazy'
    ? element.src.size
    : undefined;
}

function inferMimeType(element: Pick<ImageElement, 'src' | 'previewMetadata'>) {
  if (element.previewMetadata?.mimeType) {
    return element.previewMetadata.mimeType;
  }
  if (typeof element.src !== 'string' && element.src.kind === 'lazy') {
    return element.src.mimeType;
  }
  if (typeof element.src === 'string') {
    return /^data:([^;,]+)/iu.exec(element.src)?.[1];
  }
  return undefined;
}

function getImageResourceKey(element: ImageElement) {
  const metadataKey = element.previewMetadata?.resourceKey;
  if (metadataKey) return metadataKey;
  if (typeof element.src === 'string') return element.src || undefined;
  return element.src.kind === 'url' ? element.src.url : element.src.id;
}

function collectImageResourceUse(
  elements: readonly SlideElement[],
  counts: Map<string, number>,
) {
  elements.forEach((element) => {
    if (element.type === 'group') {
      collectImageResourceUse(element.children, counts);
      return;
    }
    if (element.type !== 'image' || element.previewable === false) return;
    const key = getImageResourceKey(element);
    if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
  });
}

/** 为同一页重复使用的图片补充复用次数，作为图标识别的辅助信号。 */
export function annotatePresentationImageReuse(
  elements: readonly SlideElement[],
): SlideElement[] {
  const counts = new Map<string, number>();
  collectImageResourceUse(elements, counts);
  if (!Array.from(counts.values()).some((count) => count >= 2)) {
    return elements.slice();
  }
  const applyReuse = (items: readonly SlideElement[]): SlideElement[] =>
    items.map((element) => {
      if (element.type === 'group') {
        const children = applyReuse(element.children);
        return children.some(
          (child, index) => child !== element.children[index],
        )
          ? { ...element, children }
          : element;
      }
      if (element.type !== 'image' || element.previewable === false) {
        return element;
      }
      const key = getImageResourceKey(element);
      const resourceReuseCount = Math.max(
        element.previewMetadata?.resourceReuseCount ?? 0,
        key ? counts.get(key) ?? 0 : 0,
      );
      if (!key || !resourceReuseCount || resourceReuseCount < 2) return element;
      return {
        ...element,
        previewMetadata: {
          ...element.previewMetadata,
          resourceKey: key,
          resourceReuseCount,
        },
      };
    });
  return applyReuse(elements);
}

/** 判断演示文稿图片是否应提供预览交互，使用多重来源信号减少图标误触发。 */
export function canPreviewPresentationImage(
  element: Pick<
    ImageElement,
    'width' | 'height' | 'previewable' | 'previewMetadata' | 'src' | 'alt'
  >,
): boolean {
  if (element.previewable === false) return false;

  const { width, height } = element;
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0
  ) {
    // 尺寸异常时保留原有能力，避免解析降级误伤可预览的正文图片。
    return true;
  }

  return !isLikelyIcon(
    width,
    height,
    element.previewMetadata,
    inferSourceKind(element),
    element.alt,
    inferMimeType(element),
    inferResourceSize(element),
  );
}

/** 为演示文稿中的图片统一设置是否允许独立预览，避免渲染层依赖来源路径或元素编号。 */
export function markPresentationElementsPreviewable(
  elements: readonly SlideElement[],
  previewable: boolean,
): SlideElement[] {
  return elements.map((element) => {
    if (element.type === 'image') {
      return { ...element, previewable };
    }
    if (element.type === 'group') {
      return {
        ...element,
        children: markPresentationElementsPreviewable(
          element.children,
          previewable,
        ),
      };
    }
    return element;
  });
}
