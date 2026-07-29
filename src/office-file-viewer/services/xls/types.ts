import type { SpreadsheetWarning } from '../spreadsheet/types';
import type { Biff8Record } from './biff8/Biff8Reader';

/** 描述 Biff8SheetType 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8SheetType =
  | 'worksheet'
  | 'chart'
  | 'macro'
  | 'dialog'
  | 'unknown';

/** 描述 Biff8SheetDescriptor 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8SheetDescriptor = {
  /** Biff8SheetDescriptor 在所属文档或任务中的唯一标识。 */
  id: string;
  /** Biff8SheetDescriptor 的可读名称。 */
  name: string;
  /** 工作表子流在 Workbook 流中的起始字节偏移。 */
  streamOffset: number;
  /** 工作表可见性，veryHidden 表示无法通过普通界面取消隐藏。 */
  visibility: 'visible' | 'hidden' | 'veryHidden';
  /** 用于区分 Biff8SheetDescriptor 不同结构分支的类型标识。 */
  type: Biff8SheetType;
};

/** 描述 Biff8Font 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8Font = {
  /** Biff8Font 的可读名称。 */
  name: string;
  /** 字体高度，单位为 twip（1/20 磅）。 */
  heightTwips: number;
  /** 字体颜色在工作簿调色板中的索引。 */
  colorIndex: number;
  /** 是否使用粗体渲染 Biff8Font。 */
  bold: boolean;
  /** 是否使用斜体渲染 Biff8Font。 */
  italic: boolean;
  /** 是否为 Biff8Font 绘制下划线。 */
  underline: boolean;
};

/** 描述 XLS/BIFF8 解析使用的样式参数。 */
export type Biff8BorderStyle = {
  /** BIFF8 边框线型代码。 */
  style: number;
  /** 边框颜色在工作簿调色板中的索引。 */
  colorIndex: number;
};

/** 描述 BIFF8 XF 中独立保存的对角边框位和线型。 */
export type Biff8DiagonalBorderStyle = Biff8BorderStyle & {
  /** 是否绘制左下到右上的对角线。 */
  up: boolean;
  /** 是否绘制左上到右下的对角线。 */
  down: boolean;
};

/** 描述 Biff8CellFormat 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8CellFormat = {
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  fontIndex: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  formatIndex: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  parentStyleIndex: number;
  /** 是否为命名样式记录，而不是普通单元格格式。 */
  isStyle: boolean;
  /** BIFF8 水平对齐枚举代码。 */
  horizontalAlign?: number;
  /** BIFF8 垂直对齐枚举代码。 */
  verticalAlign?: number;
  /** 是否允许单元格文本自动换行。 */
  wrapText?: boolean;
  /** BIFF8 填充图案枚举代码。 */
  fillPattern?: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  fillForegroundColorIndex?: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  fillBackgroundColorIndex?: number;
  /** Biff8CellFormat 关联的 leftBorder 结构；字段形状由 Biff8BorderStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  leftBorder?: Biff8BorderStyle;
  /** Biff8CellFormat 关联的 rightBorder 结构；字段形状由 Biff8BorderStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  rightBorder?: Biff8BorderStyle;
  /** Biff8CellFormat 关联的 topBorder 结构；字段形状由 Biff8BorderStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  topBorder?: Biff8BorderStyle;
  /** Biff8CellFormat 关联的 bottomBorder 结构；字段形状由 Biff8BorderStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  bottomBorder?: Biff8BorderStyle;
  /** XF 中的对角线方向、线型和调色板颜色。 */
  diagonalBorder?: Biff8DiagonalBorderStyle;
};

/** 描述 Biff8DefinedName 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8DefinedName = {
  /** Biff8DefinedName 在所属文档或任务中的唯一标识。 */
  id: number;
  /** Biff8DefinedName 的可读名称。 */
  name: string;
  /** 已定义名称对应的原始公式令牌字节。 */
  tokens: Uint8Array;
};

/** 描述 Biff8WorkbookGlobals 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8WorkbookGlobals = {
  /** Biff8WorkbookGlobals 包含的 sheets 有序集合。 */
  sheets: Biff8SheetDescriptor[];
  /** BIFF8 ExternSheet 中可解析为当前工作簿的工作表引用。 */
  externalSheets: Array<{
    /** 外部引用表指向的首个工作表索引。 */
    firstSheetIndex?: number;
    /** 外部引用表指向的末个工作表索引。 */
    lastSheetIndex?: number;
  }>;
  /** Biff8WorkbookGlobals 包含的 sharedStrings 有序集合。 */
  sharedStrings: string[];
  /** Biff8WorkbookGlobals 包含的 fonts 有序集合。 */
  fonts: Biff8Font[];
  /** Biff8WorkbookGlobals 按业务键索引的 formats 映射。 */
  formats: Map<number, string>;
  /** Biff8WorkbookGlobals 包含的 cellFormats 有序集合。 */
  cellFormats: Biff8CellFormat[];
  /** Biff8WorkbookGlobals 包含的 palette 集合。 */
  palette: string[];
  /** 是否使用 1904 日期系统；false 表示使用 1900 日期系统。 */
  date1904: boolean;
  /** Biff8WorkbookGlobals 包含的 definedNames 有序集合。 */
  definedNames: Biff8DefinedName[];
  /** Biff8WorkbookGlobals 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
  /** 工作簿是否包含 VBA 项目或相关宏记录。 */
  hasVba: boolean;
  /** 工作簿遗留字符串采用的 Windows 代码页编号。 */
  codePage?: number;
  /** Biff8WorkbookGlobals 包含的 drawingGroupRecords 有序集合。 */
  drawingGroupRecords: Biff8RecordSequence[];
};

/** 描述 Biff8RecordSequence 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8RecordSequence = {
  /** Biff8RecordSequence 在源文件记录中的数字标识。 */
  recordId: number;
  /** Biff8RecordSequence 在源二进制流中的字节偏移。 */
  offset: number;
  /** Biff8RecordSequence 包含的 chunks 有序集合。 */
  chunks: Uint8Array[];
};

/** 描述 Biff8ChartSubstream 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ChartSubstream = {
  /** Biff8ChartSubstream 在源二进制流中的字节偏移。 */
  offset: number;
  /** Biff8ChartSubstream 包含的 records 有序集合。 */
  records: Biff8Record[];
};

/** 描述 Biff8Cell 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8Cell = {
  /** 单元格在工作表中的零基行索引。 */
  row: number;
  /** 单元格在工作表中的零基列索引。 */
  column: number;
  /** 单元格引用的 XF 格式记录索引。 */
  xfIndex: number;
  /** Biff8Cell 保存的解析值或业务值。 */
  value: string | number | boolean | null;
  /** 公式缓存值或普通单元格值的解析类型。 */
  cachedType: 'string' | 'number' | 'boolean' | 'error' | 'blank';
  /** Biff8Cell 的 formula 文本值。 */
  formula?: string;
  /** Biff8Cell 的 formulaTokens 文本值。 */
  formulaTokens?: string;
};

/** 描述 Biff8RowInfo 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8RowInfo = {
  /** Biff8RowInfo 在所属集合中的位置索引。 */
  index: number;
  /** 行高，单位为 twip（1/20 磅）。 */
  heightTwips?: number;
  /** 是否隐藏 Biff8RowInfo；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
  /** 行分组的大纲层级。 */
  outlineLevel?: number;
};

/** 描述 Biff8ColumnInfo 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ColumnInfo = {
  /** 列格式范围的起始列索引，包含该列。 */
  firstColumn: number;
  /** 列格式范围的结束列索引，包含该列。 */
  lastColumn: number;
  /** 列宽，以默认字体字符宽度为单位。 */
  widthCharacters: number;
  /** 是否隐藏 Biff8ColumnInfo；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
  /** 列分组的大纲层级。 */
  outlineLevel?: number;
  /** 该列范围默认引用的 XF 格式记录索引。 */
  xfIndex?: number;
};

/** 描述 Biff8Merge 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8Merge = {
  /** 合并区域起始行索引，包含该行。 */
  startRow: number;
  /** 合并区域起始列索引，包含该列。 */
  startColumn: number;
  /** 合并区域结束行索引，包含该行。 */
  endRow: number;
  /** 合并区域结束列索引，包含该列。 */
  endColumn: number;
};

/** 描述 Biff8Worksheet 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8Worksheet = {
  /** Biff8Worksheet 关联的 descriptor 结构；字段形状由 Biff8SheetDescriptor 定义。 */
  descriptor: Biff8SheetDescriptor;
  /** Biff8Worksheet 包含的 cells 有序集合。 */
  cells: Biff8Cell[];
  /** Biff8Worksheet 包含的 rows 有序集合。 */
  rows: Biff8RowInfo[];
  /** Biff8Worksheet 包含的 columns 有序集合。 */
  columns: Biff8ColumnInfo[];
  /** Biff8Worksheet 包含的 merges 有序集合。 */
  merges: Biff8Merge[];
  /** 工作表默认列宽，以默认字体字符宽度为单位。 */
  defaultColumnWidth: number;
  /** 工作表默认行高，单位为 twip（1/20 磅）。 */
  defaultRowHeightTwips: number;
  /** 工作表已使用区域的行列边界；结束位置采用开区间。 */
  dimensions?: {
    /** Biff8Worksheet 的 firstRow 数值；具体语义遵循对应源文件格式。 */
    firstRow: number;
    /** Biff8Worksheet 的 lastRowExclusive 数值；具体语义遵循对应源文件格式。 */
    lastRowExclusive: number;
    /** Biff8Worksheet 的 firstColumn 数值；具体语义遵循对应源文件格式。 */
    firstColumn: number;
    /** Biff8Worksheet 的 lastColumnExclusive 数值；具体语义遵循对应源文件格式。 */
    lastColumnExclusive: number;
  };
  /** 是否包含 DrawingRecords 对应的数据。 */
  hasDrawingRecords: boolean;
  /** 是否包含 ChartRecords 对应的数据。 */
  hasChartRecords: boolean;
  /** Biff8Worksheet 包含的 chartSubstreams 有序集合。 */
  chartSubstreams: Biff8ChartSubstream[];
  /** Biff8Worksheet 包含的 drawingRecords 有序集合。 */
  drawingRecords: Biff8RecordSequence[];
  /** Biff8Worksheet 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};

/** 描述 Biff8ChartSheet 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8ChartSheet = {
  /** Biff8ChartSheet 关联的 descriptor 结构；字段形状由 Biff8SheetDescriptor 定义。 */
  descriptor: Biff8SheetDescriptor;
  /** Biff8ChartSheet 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
};

/** 描述 Biff8Workbook 在 XLS/BIFF8 解析中的数据结构。 */
export type Biff8Workbook = {
  /** Biff8Workbook 关联的 globals 结构；字段形状由 Biff8WorkbookGlobals 定义。 */
  globals: Biff8WorkbookGlobals;
  /** Biff8Workbook 包含的 worksheets 有序集合。 */
  worksheets: Biff8Worksheet[];
  /** Biff8Workbook 包含的 chartSheets 有序集合。 */
  chartSheets: Biff8ChartSheet[];
  /** Biff8Workbook 解析时产生但不阻止继续预览的警告集合。 */
  warnings: SpreadsheetWarning[];
};
