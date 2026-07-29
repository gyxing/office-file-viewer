/** Word 文档由源格式明确声明的大纲条目。 */
export type WordOutlineItem = {
  /** 大纲条目的稳定标识。 */
  id: string;
  /** 目录中展示的段落文本。 */
  text: string;
  /** 源文档大纲级别，使用从 0 开始的内部表示。 */
  level: number;
  /** 正文中对应段落块的稳定标识。 */
  targetBlockId: string;
};

/** Word 大纲侧栏使用的树节点。 */
export type WordOutlineTreeNode = WordOutlineItem & {
  /** Ant Design Tree 使用的稳定键。 */
  key: string;
  /** 当前条目的直接子级。 */
  children?: WordOutlineTreeNode[];
  /** 懒加载树用于声明当前节点是否确定没有后代。 */
  isLeaf?: boolean;
};

/** Word 文档提前画像和渐进解析共同维护的轻量统计。 */
export type WordPerformanceStats = {
  estimatedPageCount?: number;
  outlineCount: number;
  paragraphCount: number;
  tableRowCount: number;
  imageCount: number;
  drawingCount: number;
  textLength: number;
  largestXmlSize?: number;
  slowPagination: boolean;
};

/** 根据统计结果选择大纲与页面渲染模式。 */
export type WordPerformanceProfile = {
  renderWeight: number;
  outlineMode: 'normal' | 'virtual';
  pageMode: 'normal' | 'windowed';
};
