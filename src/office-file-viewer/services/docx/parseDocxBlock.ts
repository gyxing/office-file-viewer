import type {
  DocxBlock,
  DocxPage,
  DocxPageContent,
  DocxTableBlock,
} from './types';

/** DOCX 流式块解析器产生的事件。 */
export type DocxBlockParseEvent =
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'blocks';
      /** 按源文档顺序排列的内容块。 */
      blocks: DocxBlock[];
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'page-boundary';
      /** 当前关联的页面模型。 */
      page: DocxPage;
      /** 当前页面的页眉、页脚和首页变体。 */
      regions?: Pick<
        DocxPageContent,
        'headers' | 'footerPageNumbers' | 'differentFirstPage'
      >;
    };

/** DOCX 流式块解析器完成后的结果。 */
export type DocxBlockParseResult = {
  /** 流式解析当前块时产生的页面和内容事件。 */
  events: DocxBlockParseEvent[];
  /** 前一个分页边界是否由源文档显式指定。 */
  previousBoundaryWasExplicit: boolean;
};

/** DOCX 流式块解析器依赖的格式解析操作。 */
export type DocxBlockParseOperations<TContext> = {
  /** 判断段落是否声明了显式分页。 */
  hasExplicitPageBreak(node: Element): boolean;
  /** 找出表格中声明分页前置的行索引。 */
  readTablePageBreakRows(node: Element): number[];
  /** 判断段落是否仅包含分页标记。 */
  isPageBreakOnlyParagraph(node: Element): boolean;
  /** 判断 XML 节点是否为 DOCX 段落。 */
  isParagraph(node: Element): boolean;
  /** 判断 XML 节点是否为 DOCX 表格。 */
  isTable(node: Element): boolean;
  /** 将段落节点解析为一个或多个内容块。 */
  readParagraphBlocks(
    node: Element,
    id: string,
    context: TContext,
  ): DocxBlock[];
  /** 将 DOCX 表格节点解析为标准表格块。 */
  parseTable(node: Element, id: string, context: TContext): DocxTableBlock;
  /** 为表格及其内部对象追加页面偏移。 */
  offsetTable(
    table: DocxTableBlock,
    previousBlock: DocxBlock | undefined,
  ): DocxTableBlock;
  /** 读取段落末尾携带的分节属性。 */
  readParagraphSection(node: Element): Element | null;
  /** 根据分节属性生成页面尺寸与页边距。 */
  readSectionPage(section: Element): DocxPage;
  /** 读取分节对应的页眉和页脚区域。 */
  readSectionRegions(
    section: Element,
    context: TContext,
  ): Pick<
    DocxPageContent,
    'headers' | 'footerPageNumbers' | 'differentFirstPage'
  >;
};

/** 解析单个 DOCX 块级节点时使用的选项。 */
type ParseDocxBlockOptions<TContext> = {
  /** 当前需要解析的 OOXML 元素节点。 */
  node: Element;
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 当前操作共享的上下文。 */
  context: TContext;
  /** 源节点没有页面设置时使用的默认页面模型。 */
  defaultPage: DocxPage;
  /** 当前内容块之前的相邻内容块。 */
  previousBlock?: DocxBlock;
  /** 前一个分页边界是否由源文档显式指定。 */
  previousBoundaryWasExplicit: boolean;
  /** 解析块级节点所需的格式读取操作。 */
  operations: DocxBlockParseOperations<TContext>;
};

/** 把单个 body 直接子元素转换为块和分页边界，供物化与流式路径共用。 */
export function parseDocxBlock<TContext>({
  node,
  index,
  context,
  defaultPage,
  previousBlock,
  previousBoundaryWasExplicit: _previousBoundaryWasExplicit,
  operations,
}: ParseDocxBlockOptions<TContext>): DocxBlockParseResult {
  const events: DocxBlockParseEvent[] = [];
  const explicitPageBreak = operations.hasExplicitPageBreak(node);
  const tableBreakRows = operations.isTable(node)
    ? operations.readTablePageBreakRows(node)
    : [];
  if (operations.isPageBreakOnlyParagraph(node)) {
    events.push({ type: 'page-boundary', page: defaultPage });
    return { events, previousBoundaryWasExplicit: true };
  }

  if (operations.isParagraph(node)) {
    const blocks = operations.readParagraphBlocks(
      node,
      `p-${index + 1}`,
      context,
    );
    if (blocks.length) events.push({ type: 'blocks', blocks });
  }
  if (operations.isTable(node)) {
    const table = operations.offsetTable(
      operations.parseTable(node, `table-${index + 1}`, context),
      previousBlock,
    );
    let rowStart = 0;
    for (const rowEnd of tableBreakRows) {
      if (rowEnd > rowStart) {
        events.push({
          type: 'blocks',
          blocks: [
            {
              ...table,
              id: `${table.id}-page-${rowStart + 1}-${rowEnd}`,
              sourceBlockId: table.sourceBlockId ?? table.id,
              rows: table.rows.slice(rowStart, rowEnd),
            },
          ],
        });
      }
      events.push({ type: 'page-boundary', page: defaultPage });
      rowStart = rowEnd;
    }
    if (rowStart < table.rows.length) {
      events.push({
        type: 'blocks',
        blocks: [
          rowStart
            ? {
                ...table,
                id: `${table.id}-page-${rowStart + 1}-${table.rows.length}`,
                sourceBlockId: table.sourceBlockId ?? table.id,
                rows: table.rows.slice(rowStart),
              }
            : table,
        ],
      });
    }
  }

  const section = operations.readParagraphSection(node);
  if (section) {
    events.push({
      type: 'page-boundary',
      page: operations.readSectionPage(section),
      regions: operations.readSectionRegions(section, context),
    });
  } else if (explicitPageBreak) {
    events.push({ type: 'page-boundary', page: defaultPage });
  }
  return {
    events,
    previousBoundaryWasExplicit: !section && explicitPageBreak,
  };
}
