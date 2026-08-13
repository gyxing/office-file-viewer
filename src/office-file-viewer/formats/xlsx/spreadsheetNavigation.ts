import type { MutableRefObject } from 'react';

/** 电子表格当前渲染模式对内部链接暴露的滚动能力。 */
export type SpreadsheetNavigationController = {
  /** 当前控制器所属工作表的稳定标识。 */
  sheetId: string;
  /** 将指定的一基行列坐标滚动到可视区域。 */
  scrollToCell(rowIndex: number, columnIndex: number): boolean;
};

/** 工作表切换后等待新网格控制器挂载的最大帧数。 */
const SPREADSHEET_NAVIGATION_MAX_FRAMES = 90;

/** 等待指定工作表的网格控制器完成挂载，供链接与查找导航共同复用。 */
export function waitForSpreadsheetNavigationController(
  controllerRef: MutableRefObject<SpreadsheetNavigationController | undefined>,
  sheetId: string,
) {
  return new Promise<SpreadsheetNavigationController | undefined>((resolve) => {
    let frame = 0;
    const inspect = () => {
      const controller = controllerRef.current;
      if (controller?.sheetId === sheetId) {
        resolve(controller);
        return;
      }
      frame += 1;
      if (frame >= SPREADSHEET_NAVIGATION_MAX_FRAMES) {
        resolve(undefined);
        return;
      }
      requestAnimationFrame(inspect);
    };
    inspect();
  });
}

/** 把 A1 或 A1:B2 引用转换为用于导航的左上角坐标。 */
export function parseSpreadsheetCellReference(reference: string) {
  const normalized = reference.replace(/\$/g, '').trim();
  const match = /^([A-Za-z]{1,3})(\d+)(?::[A-Za-z]{1,3}\d+)?$/.exec(normalized);
  if (!match) return undefined;
  const columnIndex = match[1]
    .toUpperCase()
    .split('')
    .reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0);
  const rowIndex = Number(match[2]);
  return rowIndex > 0 && columnIndex > 0
    ? { rowIndex, columnIndex }
    : undefined;
}
