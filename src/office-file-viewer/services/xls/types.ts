import type { SpreadsheetWarning } from '../spreadsheet/types';
import type { Biff8Record } from './biff8/Biff8Reader';

/** BIFF8 工作表子流支持的内容类型。 */
export type Biff8SheetType =
  | 'worksheet'
  | 'chart'
  | 'macro'
  | 'dialog'
  | 'unknown';

/** BIFF8 工作表子流的位置、名称和可见状态。 */
export type Biff8SheetDescriptor = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** 面向用户展示的名称。 */
  name: string;
  /** 工作表子流在 Workbook 流中的起始字节偏移。 */
  streamOffset: number;
  /** 工作表可见性，veryHidden 表示无法通过普通界面取消隐藏。 */
  visibility: 'visible' | 'hidden' | 'veryHidden';
  /** 工作表子流的内容类型。 */
  type: Biff8SheetType;
};

/** BIFF8 字体表中的单个字体定义。 */
export type Biff8Font = {
  /** 面向用户展示的名称。 */
  name: string;
  /** 字体高度，单位为 twip（1/20 磅）。 */
  heightTwips: number;
  /** 字体颜色在工作簿调色板中的索引。 */
  colorIndex: number;
  /** 是否使用粗体。 */
  bold: boolean;
  /** 是否使用斜体。 */
  italic: boolean;
  /** 是否绘制下划线。 */
  underline: boolean;
};

/** BIFF8 单元格边框的线型和调色板颜色。 */
export type Biff8BorderStyle = {
  /** BIFF8 边框线型代码。 */
  style: number;
  /** 边框颜色在工作簿调色板中的索引。 */
  colorIndex: number;
};

/** BIFF8 对角边框的方向、线型和颜色。 */
export type Biff8DiagonalBorderStyle = Biff8BorderStyle & {
  /** 是否绘制左下到右上的对角线。 */
  up: boolean;
  /** 是否绘制左上到右下的对角线。 */
  down: boolean;
};

/** BIFF8 XF 记录解析出的单元格格式。 */
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
  /** 是否缩小文字以适应单元格宽度。 */
  shrinkToFit?: boolean;
  /** BIFF8 填充图案枚举代码。 */
  fillPattern?: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  fillForegroundColorIndex?: number;
  /** 当前字段引用的工作簿样式表索引，具体目标由属性名确定。 */
  fillBackgroundColorIndex?: number;
  /** 单元格左边框的线型和颜色。 */
  leftBorder?: Biff8BorderStyle;
  /** 单元格右边框的线型和颜色。 */
  rightBorder?: Biff8BorderStyle;
  /** 单元格上边框的线型和颜色。 */
  topBorder?: Biff8BorderStyle;
  /** 单元格下边框的线型和颜色。 */
  bottomBorder?: Biff8BorderStyle;
  /** XF 中的对角线方向、线型和调色板颜色。 */
  diagonalBorder?: Biff8DiagonalBorderStyle;
};

/** BIFF8 工作簿已定义名称及其公式令牌。 */
export type Biff8DefinedName = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** 面向用户展示的名称。 */
  name: string;
  /** 已定义名称对应的原始公式令牌字节。 */
  tokens: Uint8Array;
};

/** BIFF8 工作簿全局子流中的共享表和元数据。 */
export type Biff8WorkbookGlobals = {
  /** 按工作簿顺序排列的工作表描述信息。 */
  sheets: Biff8SheetDescriptor[];
  /** BIFF8 ExternSheet 中可解析为当前工作簿的工作表引用。 */
  externalSheets: Array<{
    /** 外部引用表指向的首个工作表索引。 */
    firstSheetIndex?: number;
    /** 外部引用表指向的末个工作表索引。 */
    lastSheetIndex?: number;
  }>;
  /** 工作簿共享字符串表中的文本。 */
  sharedStrings: string[];
  /** 工作簿字体表中的字体定义。 */
  fonts: Biff8Font[];
  /** 按格式编号索引的数字格式代码。 */
  formats: Map<number, string>;
  /** 工作簿 XF 表中的单元格格式。 */
  cellFormats: Biff8CellFormat[];
  /** 工作簿使用的自定义调色板颜色。 */
  palette: string[];
  /** 是否使用 1904 日期系统；false 表示使用 1900 日期系统。 */
  date1904: boolean;
  /** 工作簿内声明的名称和公式令牌。 */
  definedNames: Biff8DefinedName[];
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
  /** 工作簿是否包含 VBA 项目或相关宏记录。 */
  hasVba: boolean;
  /** 工作簿遗留字符串采用的 Windows 代码页编号。 */
  codePage?: number;
  /** 工作簿级 OfficeArt 绘图组记录。 */
  drawingGroupRecords: Biff8RecordSequence[];
};

/** 由主记录及其 CONTINUE 记录组成的连续字节块。 */
export type Biff8RecordSequence = {
  /** 该组合记录的首个 BIFF8 类型编号。 */
  recordId: number;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 由 CONTINUE 记录拼接前的原始字节块。 */
  chunks: Uint8Array[];
};

/** BIFF8 内嵌图表的记录子流。 */
export type Biff8ChartSubstream = {
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 当前图表子流中的 BIFF8 记录。 */
  records: Biff8Record[];
};

/** BIFF8 单元格的值、公式缓存和格式索引。 */
export type Biff8Cell = {
  /** 单元格在工作表中的零基行索引。 */
  row: number;
  /** 单元格在工作表中的零基列索引。 */
  column: number;
  /** 单元格引用的 XF 格式记录索引。 */
  xfIndex: number;
  /** 单元格公式缓存值或普通值。 */
  value: string | number | boolean | null;
  /** 公式缓存值或普通单元格值的解析类型。 */
  cachedType: 'string' | 'number' | 'boolean' | 'error' | 'blank';
  /** 可供展示的 A1 形式公式文本。 */
  formula?: string;
  /** 无法完整解码公式时保留的令牌描述。 */
  formulaTokens?: string;
};

/** BIFF8 ROW 记录中的行高和显示状态。 */
export type Biff8RowInfo = {
  /** 工作表中的零基行索引。 */
  index: number;
  /** ROW 记录保存的行高，单位为 twip（1/20 磅）。 */
  heightTwips?: number;
  /** 是否由用户手动设置行高；否则该高度为 Excel 自动计算结果。 */
  customHeight?: boolean;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
  /** 行分组的大纲层级。 */
  outlineLevel?: number;
};

/** BIFF8 COLINFO 记录覆盖的列区间和样式。 */
export type Biff8ColumnInfo = {
  /** 列格式范围的起始列索引，包含该列。 */
  firstColumn: number;
  /** 列格式范围的结束列索引，包含该列。 */
  lastColumn: number;
  /** 列宽，以默认字体字符宽度为单位。 */
  widthCharacters: number;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
  /** 列分组的大纲层级。 */
  outlineLevel?: number;
  /** 该列范围默认引用的 XF 格式记录索引。 */
  xfIndex?: number;
};

/** BIFF8 合并单元格的行列边界。 */
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

/** 解析后的 BIFF8 普通工作表。 */
export type Biff8Worksheet = {
  /** 当前工作表对应的子流描述信息。 */
  descriptor: Biff8SheetDescriptor;
  /** 按源记录读取的非空单元格。 */
  cells: Biff8Cell[];
  /** 显式声明高度或状态的行信息。 */
  rows: Biff8RowInfo[];
  /** 显式声明宽度或状态的列区间。 */
  columns: Biff8ColumnInfo[];
  /** 当前工作表声明的合并区域。 */
  merges: Biff8Merge[];
  /** 工作表默认列宽，以默认字体字符宽度为单位。 */
  defaultColumnWidth: number;
  /** 工作表默认行高，单位为 twip（1/20 磅）。 */
  defaultRowHeightTwips: number;
  /** 工作表已使用区域的行列边界；结束位置采用开区间。 */
  dimensions?: {
    /** 已使用区域的起始行索引。 */
    firstRow: number;
    /** 已使用区域结束行的开区间索引。 */
    lastRowExclusive: number;
    /** 已使用区域的起始列索引。 */
    firstColumn: number;
    /** 已使用区域结束列的开区间索引。 */
    lastColumnExclusive: number;
  };
  /** 是否包含 OfficeArt 绘图记录。 */
  hasDrawingRecords: boolean;
  /** 是否包含 BIFF8 图表记录。 */
  hasChartRecords: boolean;
  /** 工作表内嵌图表的 BIFF8 子流。 */
  chartSubstreams: Biff8ChartSubstream[];
  /** 工作表级 OfficeArt 绘图记录。 */
  drawingRecords: Biff8RecordSequence[];
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};

/** 解析后的 BIFF8 独立图表工作表。 */
export type Biff8ChartSheet = {
  /** 当前图表工作表对应的子流描述信息。 */
  descriptor: Biff8SheetDescriptor;
  /** 对应的 BIFF8 图表子流。 */
  substream: Biff8ChartSubstream;
};

/** 解析后的 BIFF8 工作簿及其工作表。 */
export type Biff8Workbook = {
  /** 工作簿所有工作表共享的字体、样式和字符串表。 */
  globals: Biff8WorkbookGlobals;
  /** 解析成功的普通工作表。 */
  worksheets: Biff8Worksheet[];
  /** 解析成功的独立图表工作表。 */
  chartSheets: Biff8ChartSheet[];
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: SpreadsheetWarning[];
};
