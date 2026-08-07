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
  /** 大纲树用于识别节点的稳定键。 */
  key: string;
  /** 当前条目的直接子级。 */
  children?: WordOutlineTreeNode[];
  /** 懒加载树用于声明当前节点是否确定没有后代。 */
  isLeaf?: boolean;
};

/** Word 文档提前画像和渐进解析共同维护的轻量统计。 */
export type WordPerformanceStats = {
  /** 根据来源分页信息估算的页面数量。 */
  estimatedPageCount?: number;
  /** 大纲数量。 */
  outlineCount: number;
  /** 段落数量。 */
  paragraphCount: number;
  /** 表格 行数量。 */
  tableRowCount: number;
  /** 图片数量。 */
  imageCount: number;
  /** 绘图数量。 */
  drawingCount: number;
  /** 正文包含的字符总数。 */
  textLength: number;
  /** 压缩包内最大 XML 部件解压后的大小，单位为字节。 */
  largestXmlSize?: number;
  /** 分页过程是否已经触发慢任务判定。 */
  slowPagination: boolean;
};

/** 根据统计结果选择大纲与页面渲染模式。 */
export type WordPerformanceProfile = {
  /** 用于衡量文档渲染开销的综合权重。 */
  renderWeight: number;
  /** 大纲 模式。 */
  outlineMode: 'normal' | 'virtual';
  /** 当前使用的页面加载模式。 */
  pageMode: 'normal' | 'windowed';
};
