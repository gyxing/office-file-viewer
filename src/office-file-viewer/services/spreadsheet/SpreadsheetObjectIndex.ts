import type { SpreadsheetAxisIndex } from './SpreadsheetAxisIndex';
import type {
  SpreadsheetChart,
  SpreadsheetImage,
  SpreadsheetRange,
} from './types';

type SpreadsheetFloatingObject = SpreadsheetImage | SpreadsheetChart;

/** 按完整工作表像素坐标查询与虚拟窗口相交的浮动对象。 */
export class SpreadsheetObjectIndex<T extends SpreadsheetFloatingObject> {
  constructor(
    private readonly objects: readonly T[],
    private readonly rowIndex: SpreadsheetAxisIndex,
    private readonly columnIndex: SpreadsheetAxisIndex,
  ) {}

  query(range: SpreadsheetRange) {
    const left = this.columnIndex.offsetAt(range.startColumn);
    const top = this.rowIndex.offsetAt(range.startRow);
    const right = this.columnIndex.offsetAt(range.endColumn + 1);
    const bottom = this.rowIndex.offsetAt(range.endRow + 1);
    return this.objects.filter(
      (item) =>
        item.x + item.width >= left &&
        item.x <= right &&
        item.y + item.height >= top &&
        item.y <= bottom,
    );
  }
}
