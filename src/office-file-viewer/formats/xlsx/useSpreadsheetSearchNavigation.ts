import type { MutableRefObject } from 'react';
import { useCallback } from 'react';
import type { OfficeSearchTarget } from '../../services/search/types';
import { useOfficeSearchNavigatorRegistration } from '../search/OfficeSearchContext';
import {
  type SpreadsheetNavigationController,
  waitForSpreadsheetNavigationController,
} from './spreadsheetNavigation';

/** 注册 Excel 查找结果导航，切换工作表后等待对应网格再定位单元格。 */
export function useSpreadsheetSearchNavigation({
  activeSheetId,
  onSelectSheet,
  navigationControllerRef,
}: {
  /** 当前激活工作表标识。 */
  activeSheetId?: string;
  /** 请求切换工作表。 */
  onSelectSheet: (sheetId: string) => void;
  /** 当前活动网格暴露的定位控制器。 */
  navigationControllerRef: MutableRefObject<
    SpreadsheetNavigationController | undefined
  >;
}) {
  const navigate = useCallback(
    async (target: OfficeSearchTarget) => {
      if (target.kind !== 'spreadsheet') return false;
      if (target.sheetId !== activeSheetId) onSelectSheet(target.sheetId);
      const controller = await waitForSpreadsheetNavigationController(
        navigationControllerRef,
        target.sheetId,
      );
      return Boolean(
        controller?.scrollToCell(target.rowIndex, target.columnIndex),
      );
    },
    [activeSheetId, navigationControllerRef, onSelectSheet],
  );

  useOfficeSearchNavigatorRegistration('spreadsheet', navigate);
}
