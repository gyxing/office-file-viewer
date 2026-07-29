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
  children: WordOutlineTreeNode[];
};
