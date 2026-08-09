import type { MutableRefObject } from 'react';
import { useCallback, useEffect } from 'react';
import type { SpreadsheetSourceSnapshot } from '../../services/spreadsheet/SpreadsheetSource';
import type {
  OfficeInternalHyperlinkTarget,
  OfficeSpreadsheetHyperlinkTarget,
} from '../../shared/hyperlink';
import { useOfficeHyperlinkContext } from '../../shared/hyperlink';
import {
  parseSpreadsheetCellReference,
  type SpreadsheetNavigationController,
} from './spreadsheetNavigation';

/** 工作表切换后等待新网格控制器挂载的最大帧数。 */
const SPREADSHEET_NAVIGATION_MAX_FRAMES = 90;

/** 去掉 Excel 对定义名称目标增加的公式前缀。 */
function normalizeDefinedNameTarget(value: string) {
  return value.trim().replace(/^=/, '');
}

function unquoteSheetName(value: string) {
  const trimmed = value.trim();
  return trimmed.startsWith("'") && trimmed.endsWith("'")
    ? trimmed.slice(1, -1).replace(/''/g, "'")
    : trimmed;
}

/** 将定义名称或显式目标归一化为工作表名称与单元格引用。 */
function resolveSpreadsheetLocation(
  target: OfficeSpreadsheetHyperlinkTarget,
  snapshot: SpreadsheetSourceSnapshot,
) {
  let sheetName = target.sheetName;
  let cellRef = target.cellRef;
  if (target.definedName) {
    const definedNameEntry = Object.entries(snapshot.definedNames ?? {}).find(
      ([name]) => name.toLowerCase() === target.definedName?.toLowerCase(),
    );
    if (!definedNameEntry) return undefined;
    const location = normalizeDefinedNameTarget(definedNameEntry[1]);
    const separator = location.lastIndexOf('!');
    if (separator < 0) return undefined;
    sheetName = unquoteSheetName(location.slice(0, separator));
    cellRef = location.slice(separator + 1);
  }
  const position = cellRef ? parseSpreadsheetCellReference(cellRef) : undefined;
  return position ? { sheetName, ...position } : undefined;
}

function waitForSheetController(
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

/** 注册 Excel 内部链接导航，兼容工作表切换、定义名称与虚拟网格。 */
export function useSpreadsheetHyperlinkNavigation({
  snapshot,
  activeSheetId,
  onSelectSheet,
  navigationControllerRef,
}: {
  /** 当前工作簿的轻量快照。 */
  snapshot: SpreadsheetSourceSnapshot;
  /** 当前激活工作表标识。 */
  activeSheetId?: string;
  /** 请求切换工作表。 */
  onSelectSheet: (sheetId: string) => void;
  /** 当前活动网格暴露的定位控制器。 */
  navigationControllerRef: MutableRefObject<
    SpreadsheetNavigationController | undefined
  >;
}) {
  const hyperlinkContext = useOfficeHyperlinkContext();
  const navigate = useCallback(
    async (target: OfficeInternalHyperlinkTarget) => {
      if (target.family !== 'spreadsheet') return false;
      const location = resolveSpreadsheetLocation(target, snapshot);
      if (!location) return false;
      const targetSheet = location.sheetName
        ? snapshot.sheets.find(
            (sheet) =>
              sheet.name.toLowerCase() === location.sheetName?.toLowerCase(),
          )
        : snapshot.sheets.find((sheet) => sheet.id === activeSheetId) ??
          snapshot.sheets[0];
      if (!targetSheet || targetSheet.kind !== 'worksheet') return false;
      if (targetSheet.id !== activeSheetId) onSelectSheet(targetSheet.id);
      const controller = await waitForSheetController(
        navigationControllerRef,
        targetSheet.id,
      );
      return Boolean(
        controller?.scrollToCell(location.rowIndex, location.columnIndex),
      );
    },
    [activeSheetId, navigationControllerRef, onSelectSheet, snapshot],
  );

  useEffect(() => {
    if (!hyperlinkContext) return undefined;
    return hyperlinkContext.registerNavigator('spreadsheet', navigate);
  }, [hyperlinkContext, navigate]);
}
