import type { OfficeResourceSource } from '../resource-store';

/** 演示文稿媒体的浏览器播放类别。 */
export type PresentationMediaKind = 'audio' | 'video';

/** OfficeFileViewer 对演示文稿媒体的只读运行配置。 */
export type OfficeFileViewerPresentationMediaOptions = Readonly<{
  /** 是否允许加载源文件声明的 HTTP(S) 外部媒体，默认禁止。 */
  allowExternal?: boolean;
  /** 浏览器能够取得媒体地址时是否显示下载操作，默认显示。 */
  download?: boolean;
}>;

/** 演示文稿中的内嵌或外部媒体来源。 */
export type PresentationMediaSource = Readonly<{
  /** 浏览器应使用音频还是视频控件。 */
  kind: PresentationMediaKind;
  /** 内嵌资源由资源仓库管理，外部资源受安全配置约束。 */
  sourceKind: 'embedded' | 'external';
  /** 内嵌资源引用或外部地址。 */
  source: string | OfficeResourceSource;
  /** 源文件声明或根据扩展名推断出的 MIME。 */
  mimeType?: string;
  /** 下载时使用的文件名。 */
  fileName?: string;
  /** 是否循环播放；自动播放始终由查看器禁止。 */
  loop?: boolean;
}>;

/** 根据 Office 媒体文件名推断浏览器可识别的 MIME。 */
export function getPresentationMediaMimeType(path: string) {
  const lower = path.toLowerCase().split(/[?#]/, 1)[0];
  if (lower.endsWith('.mp4') || lower.endsWith('.m4v')) return 'video/mp4';
  if (lower.endsWith('.webm')) return 'video/webm';
  if (lower.endsWith('.ogv') || lower.endsWith('.ogg')) return 'video/ogg';
  if (lower.endsWith('.avi')) return 'video/x-msvideo';
  if (lower.endsWith('.mov')) return 'video/quicktime';
  if (lower.endsWith('.mp3')) return 'audio/mpeg';
  if (lower.endsWith('.wav')) return 'audio/wav';
  if (lower.endsWith('.m4a') || lower.endsWith('.aac')) return 'audio/mp4';
  if (lower.endsWith('.oga')) return 'audio/ogg';
  if (lower.endsWith('.mid') || lower.endsWith('.midi')) return 'audio/midi';
  if (lower.endsWith('.svg')) return 'image/svg+xml';
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.bmp') || lower.endsWith('.dib')) return 'image/bmp';
  if (lower.endsWith('.tif') || lower.endsWith('.tiff')) return 'image/tiff';
  return 'application/octet-stream';
}
