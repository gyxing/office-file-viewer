import type { PreviewKind } from '../../services/preview';

/** Word 文档内部可定位的书签目标。 */
export type OfficeWordHyperlinkTarget = {
  /** 内部书签名称。 */
  bookmark: string;
};

/** 电子表格内部可定位的工作表、单元格或名称目标。 */
export type OfficeSpreadsheetHyperlinkTarget = {
  /** 目标工作表名称；省略时沿用当前工作表。 */
  sheetName?: string;
  /** 目标单元格或单元格区域引用。 */
  cellRef?: string;
  /** 目标工作簿定义名称。 */
  definedName?: string;
};

/** 演示文稿内部可定位的幻灯片目标。 */
export type OfficePresentationHyperlinkTarget = {
  /** 目标幻灯片稳定标识。 */
  slideId?: string;
  /** 目标幻灯片零基索引。 */
  slideIndex?: number;
  /** 相对当前页执行的标准跳转动作。 */
  action?: 'first' | 'last' | 'next' | 'previous';
};

/** Office 文档内部链接支持的三类导航目标。 */
export type OfficeInternalHyperlinkTarget =
  | ({ family: 'word' } & OfficeWordHyperlinkTarget)
  | ({ family: 'spreadsheet' } & OfficeSpreadsheetHyperlinkTarget)
  | ({ family: 'presentation' } & OfficePresentationHyperlinkTarget);

/** 解析器输出且可在线程间传递的统一 Office 超链接。 */
export type OfficeHyperlink =
  | {
      /** 外部链接的协议语义。 */
      kind: 'external' | 'email' | 'phone' | 'file';
      /** 源文档声明的目标地址。 */
      target: string;
      /** Office 中为链接声明的悬停说明。 */
      screenTip?: string;
    }
  | {
      /** 文档内部导航链接。 */
      kind: 'internal';
      /** 已按文档族标准化的内部目标。 */
      target: OfficeInternalHyperlinkTarget;
      /** Office 中为链接声明的悬停说明。 */
      screenTip?: string;
    };

/** 宿主可用于区分链接来源对象的稳定类型。 */
export type OfficeHyperlinkSourceType =
  | 'text'
  | 'cell'
  | 'image'
  | 'shape'
  | 'button';

/** 链接所在渲染对象的稳定身份。 */
export type OfficeHyperlinkSource = {
  /** 链接所在对象类型。 */
  type: OfficeHyperlinkSourceType;
  /** 链接所在对象的稳定标识。 */
  id: string;
};

/** OfficeFileViewer 向宿主报告的一次链接激活事务。 */
export type OfficeHyperlinkActivateEvent = {
  /** 解析器输出的标准链接。 */
  hyperlink: OfficeHyperlink;
  /** 当前正在预览的文件。 */
  file: File;
  /** 当前文件的具体预览格式。 */
  previewKind: PreviewKind;
  /** 链接所在对象类型。 */
  sourceType: OfficeHyperlinkSourceType;
  /** 链接所在对象的稳定标识。 */
  sourceId: string;
  /** 宿主是否已阻止组件执行默认导航。 */
  readonly defaultPrevented: boolean;
  /** 阻止组件执行默认导航，但不改变安全策略。 */
  preventDefault(): void;
};

/** 内部导航器完成定位时返回 true，无法定位时返回 false。 */
export type OfficeHyperlinkNavigator = (
  target: OfficeInternalHyperlinkTarget,
) => boolean | Promise<boolean>;

/** 链接被激活时采用的输入方式。 */
export type OfficeHyperlinkActivationMode =
  | 'mouse'
  | 'keyboard'
  | 'touch'
  | 'context-menu';

/** 格式渲染节点向共享 Provider 提交的链接激活请求。 */
export type OfficeHyperlinkActivationRequest = {
  /** 当前节点绑定的标准链接。 */
  hyperlink: OfficeHyperlink;
  /** 当前节点的稳定来源信息。 */
  source: OfficeHyperlinkSource;
  /** 用户本次使用的输入方式。 */
  mode: OfficeHyperlinkActivationMode;
};
