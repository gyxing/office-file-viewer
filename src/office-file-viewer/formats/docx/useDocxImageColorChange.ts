import { useEffect, useRef, useState } from 'react';
import type {
  DocxImageColorChange,
  DocxImageCrop,
} from '../../services/docx/types';

/** 颜色匹配允许吸收有损图片在边缘产生的轻微压缩误差。 */
const DRAWING_COLOR_CHANGE_TOLERANCE = 4;

/** 颜色替换后的浏览器资源状态。 */
type DocxImageEffectUrlState = {
  /** 应用于当前图片节点的浏览器地址。 */
  url?: string;
  /** 图片效果是否仍在计算。 */
  loading: boolean;
  /** 图片效果计算失败时保留的错误。 */
  error?: Error;
};

function isColorChannelMatch(actual: number, expected: number) {
  return Math.abs(actual - expected) <= DRAWING_COLOR_CHANGE_TOLERANCE;
}

function parseHexColor(color: string) {
  const match = color.match(/^#([0-9a-f]{6})$/i);
  if (!match) return undefined;
  const value = Number.parseInt(match[1], 16);
  return {
    red: (value >> 16) & 255,
    green: (value >> 8) & 255,
    blue: value & 255,
  };
}

function loadImage(url: string, ownerDocument: Document) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = ownerDocument.createElement('img');
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('DOCX 图片效果资源加载失败'));
    image.src = url;
  });
}

function canvasToPngBlob(canvas: HTMLCanvasElement) {
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('DOCX 图片颜色替换导出失败'));
    }, 'image/png');
  });
}

/** 按 DrawingML clrChange 语义替换像素颜色，并保留抗锯齿边缘的半透明过渡。 */
async function createColorChangedImageUrl(
  sourceUrl: string,
  effect: DocxImageColorChange,
  ownerDocument: Document,
) {
  const from = parseHexColor(effect.from);
  const to = parseHexColor(effect.to);
  if (!from || !to) return undefined;

  const image = await loadImage(sourceUrl, ownerDocument);
  const canvas = ownerDocument.createElement('canvas');
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const context = canvas.getContext('2d', { willReadFrequently: true });
  if (!context) throw new Error('当前浏览器无法处理 DOCX 图片颜色替换');

  context.drawImage(image, 0, 0);
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
  const sourceAlpha = effect.useAlpha ? 255 : undefined;
  for (let index = 0; index < pixels.data.length; index += 4) {
    if (
      !isColorChannelMatch(pixels.data[index], from.red) ||
      !isColorChannelMatch(pixels.data[index + 1], from.green) ||
      !isColorChannelMatch(pixels.data[index + 2], from.blue) ||
      (sourceAlpha !== undefined &&
        !isColorChannelMatch(pixels.data[index + 3], sourceAlpha))
    ) {
      continue;
    }
    const distance = Math.max(
      Math.abs(pixels.data[index] - from.red),
      Math.abs(pixels.data[index + 1] - from.green),
      Math.abs(pixels.data[index + 2] - from.blue),
    );
    const matchRatio = 1 - distance / (DRAWING_COLOR_CHANGE_TOLERANCE + 1);
    pixels.data[index] = Math.round(
      pixels.data[index] + (to.red - pixels.data[index]) * matchRatio,
    );
    pixels.data[index + 1] = Math.round(
      pixels.data[index + 1] + (to.green - pixels.data[index + 1]) * matchRatio,
    );
    pixels.data[index + 2] = Math.round(
      pixels.data[index + 2] + (to.blue - pixels.data[index + 2]) * matchRatio,
    );
    pixels.data[index + 3] = Math.round(
      pixels.data[index + 3] * (1 - matchRatio * (1 - effect.alpha)),
    );
  }
  context.putImageData(pixels, 0, 0);
  const urlApi = ownerDocument.defaultView?.URL ?? URL;
  return urlApi.createObjectURL(await canvasToPngBlob(canvas));
}

/** 在图片确实声明 clrChange 时才按需生成处理后的地址，并在卸载时释放资源。 */
export function useDocxImageColorChange(
  sourceUrl: string | undefined,
  effect: DocxImageColorChange | undefined,
): DocxImageEffectUrlState {
  const generationRef = useRef(0);
  const [state, setState] = useState<DocxImageEffectUrlState>({
    loading: Boolean(sourceUrl && effect),
  });

  useEffect(() => {
    const generation = ++generationRef.current;
    if (!sourceUrl || !effect || typeof document === 'undefined') {
      setState({ url: sourceUrl, loading: false });
      return undefined;
    }

    const ownerDocument = document;
    const urlApi = ownerDocument.defaultView?.URL ?? URL;
    let generatedUrl: string | undefined;
    setState({ loading: true });
    void createColorChangedImageUrl(sourceUrl, effect, ownerDocument).then(
      (url) => {
        generatedUrl = url;
        if (generation === generationRef.current) {
          setState({ url: url ?? sourceUrl, loading: false });
        } else if (url) {
          urlApi.revokeObjectURL(url);
        }
      },
      (error) => {
        if (generation !== generationRef.current) return;
        setState({
          url: sourceUrl,
          loading: false,
          error: error instanceof Error ? error : new Error(String(error)),
        });
      },
    );
    return () => {
      if (generatedUrl) urlApi.revokeObjectURL(generatedUrl);
    };
  }, [effect, sourceUrl]);

  return state;
}

/** 将内联图片裁剪模型转换为容器内的绝对定位样式。 */
export function resolveDocxCroppedImageStyle(crop: DocxImageCrop) {
  const visibleWidth = Math.max(0.01, 1 - crop.left - crop.right);
  const visibleHeight = Math.max(0.01, 1 - crop.top - crop.bottom);
  return {
    position: 'absolute' as const,
    left: `${-(crop.left / visibleWidth) * 100}%`,
    top: `${-(crop.top / visibleHeight) * 100}%`,
    width: `${100 / visibleWidth}%`,
    height: `${100 / visibleHeight}%`,
    maxWidth: 'none',
    objectFit: 'fill' as const,
  };
}