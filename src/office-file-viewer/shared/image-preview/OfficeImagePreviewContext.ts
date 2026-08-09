import { createContext, useContext } from 'react';
import type {
  OfficeImagePreviewTarget,
  ResolvedOfficeImagePreviewOptions,
} from './types';

/** 图片右键菜单相对浏览器视口的触发坐标。 */
export type OfficeImageContextMenuPoint = {
  /** 触发位置的横坐标。 */
  clientX: number;
  /** 触发位置的纵坐标。 */
  clientY: number;
};

/** 图片节点调用单实例预览层时使用的内部接口。 */
export type OfficeImagePreviewContextValue = {
  /** 当前生效的图片交互配置。 */
  options: ResolvedOfficeImagePreviewOptions;
  /** 打开图片预览层。 */
  openPreview(target: OfficeImagePreviewTarget, trigger: HTMLElement): void;
  /** 在指定位置打开图片右键菜单。 */
  openContextMenu(
    target: OfficeImagePreviewTarget,
    point: OfficeImageContextMenuPoint,
    trigger: HTMLElement,
  ): void;
};

/** 当前预览器实例向图片节点提供的交互上下文。 */
export const OfficeImagePreviewContext =
  createContext<OfficeImagePreviewContextValue | null>(null);

/** 获取当前预览器实例提供的图片交互能力。 */
export function useOfficeImagePreviewContext() {
  return useContext(OfficeImagePreviewContext);
}
