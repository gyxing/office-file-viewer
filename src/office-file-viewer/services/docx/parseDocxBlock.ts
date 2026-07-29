import type {
  DocxBlock,
  DocxPage,
  DocxPageContent,
  DocxTableBlock,
} from './types';

export type DocxBlockParseEvent =
  | { type: 'blocks'; blocks: DocxBlock[] }
  | {
      type: 'page-boundary';
      page: DocxPage;
      regions?: Pick<
        DocxPageContent,
        'headers' | 'footerPageNumbers' | 'differentFirstPage'
      >;
    };

export type DocxBlockParseResult = {
  events: DocxBlockParseEvent[];
  previousBoundaryWasExplicit: boolean;
};

export type DocxBlockParseOperations<TContext> = {
  hasRenderedPageBreak(node: Element): boolean;
  hasExplicitPageBreak(node: Element): boolean;
  readTablePageBreakRows(node: Element): number[];
  isPageBreakOnlyParagraph(node: Element): boolean;
  isParagraph(node: Element): boolean;
  isTable(node: Element): boolean;
  readParagraphBlocks(
    node: Element,
    id: string,
    context: TContext,
  ): DocxBlock[];
  parseTable(node: Element, id: string, context: TContext): DocxTableBlock;
  offsetTable(
    table: DocxTableBlock,
    previousBlock: DocxBlock | undefined,
  ): DocxTableBlock;
  readParagraphSection(node: Element): Element | null;
  readSectionPage(section: Element): DocxPage;
  readSectionRegions(
    section: Element,
    context: TContext,
  ): Pick<
    DocxPageContent,
    'headers' | 'footerPageNumbers' | 'differentFirstPage'
  >;
};

type ParseDocxBlockOptions<TContext> = {
  node: Element;
  index: number;
  context: TContext;
  defaultPage: DocxPage;
  previousBlock?: DocxBlock;
  previousBoundaryWasExplicit: boolean;
  operations: DocxBlockParseOperations<TContext>;
};

/** 把单个 body 直接子元素转换为块和分页边界，供物化与流式路径共用。 */
export function parseDocxBlock<TContext>({
  node,
  index,
  context,
  defaultPage,
  previousBlock,
  previousBoundaryWasExplicit,
  operations,
}: ParseDocxBlockOptions<TContext>): DocxBlockParseResult {
  const events: DocxBlockParseEvent[] = [];
  const renderedPageBreak = operations.hasRenderedPageBreak(node);
  const explicitPageBreak = operations.hasExplicitPageBreak(node);
  const tableBreakRows = operations.isTable(node)
    ? operations.readTablePageBreakRows(node)
    : [];
  if (
    renderedPageBreak &&
    !tableBreakRows.length &&
    !previousBoundaryWasExplicit
  ) {
    events.push({ type: 'page-boundary', page: defaultPage });
  }
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
