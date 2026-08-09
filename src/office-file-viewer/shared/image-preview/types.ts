import type { OfficeResourceSource } from '../../services/resource-store';

/** 控制图片预览附加交互的可选能力。 */
export type OfficeFileViewerImagePreviewOptions = {
  /** 是否允许从预览层或右键菜单下载原始图片，默认开启。 */
  download?: boolean;
  /** 是否以自定义的“预览、下载”菜单替代图片原生右键菜单，默认开启。 */
  contextMenu?: boolean;
};

/** 图片预览功能支持整体开关或按能力配置。 */
export type OfficeFileViewerImagePreviewConfig =
  | boolean
  | OfficeFileViewerImagePreviewOptions;

/** 预览器内部统一使用的图片资源描述。 */
export type OfficeImagePreviewTarget = {
  /** 当前图片在所属文档内的稳定标识。 */
  id: string;
  /** 可直接访问或按需物化的图片资源。 */
  source: string | OfficeResourceSource;
  /** 图片无法显示时使用的替代文本。 */
  alt?: string;
  /** 优先用于预览标题和下载文件名的图片名称。 */
  name?: string;
  /** 用于推断下载扩展名的图片 MIME 类型。 */
  mimeType?: string;
};

/** 图片预览配置归一化后的内部状态。 */
export type ResolvedOfficeImagePreviewOptions = {
  /** 是否启用图片预览交互。 */
  enabled: boolean;
  /** 是否允许下载原始图片。 */
  download: boolean;
  /** 是否启用图片自定义右键菜单。 */
  contextMenu: boolean;
};
