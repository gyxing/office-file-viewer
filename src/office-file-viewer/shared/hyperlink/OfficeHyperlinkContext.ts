import { createContext, useContext } from 'react';
import type {
  OfficeHyperlinkActivationRequest,
  OfficeHyperlinkNavigator,
} from './types';

/** 超链接右键菜单触发点的视口坐标。 */
export type OfficeHyperlinkContextMenuPoint = {
  /** 触发点相对浏览器视口左侧的位置。 */
  clientX: number;
  /** 触发点相对浏览器视口顶部的位置。 */
  clientY: number;
};

/** 格式渲染节点使用的实例级超链接能力。 */
export type OfficeHyperlinkContextValue = {
  /** 当前查看器是否启用源文档链接交互。 */
  enabled: boolean;
  /** 当前平台要求按住的修饰键名称。 */
  modifierLabel: 'Ctrl' | 'Command';
  /** 当前平台对应的鼠标激活提示。 */
  activationHint: string;
  /** 提交一次经过输入规则校验的链接激活。 */
  activate(request: OfficeHyperlinkActivationRequest): void;
  /** 在链接支持受控菜单时打开菜单，并返回本次右键是否已被接管。 */
  openContextMenu(
    request: OfficeHyperlinkActivationRequest,
    point: OfficeHyperlinkContextMenuPoint,
    trigger: HTMLElement,
  ): boolean;
  /** 注册当前文档族的内部导航器，并返回清理函数。 */
  registerNavigator(
    family: 'word' | 'spreadsheet' | 'presentation',
    navigator: OfficeHyperlinkNavigator,
  ): () => void;
};

/** 当前 OfficeFileViewer 实例提供的链接交互上下文。 */
export const OfficeHyperlinkContext =
  createContext<OfficeHyperlinkContextValue | null>(null);

/** 获取当前查看器实例提供的超链接交互能力。 */
export function useOfficeHyperlinkContext() {
  return useContext(OfficeHyperlinkContext);
}
