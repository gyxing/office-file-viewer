import type { MutableRefObject } from 'react';
import { useCallback } from 'react';
import type { OfficeAnnotation } from '../../services/annotations/types';
import { useOfficeAnnotationNavigation } from '../../shared/annotations';
import {
  type SpreadsheetNavigationController,
  waitForSpreadsheetNavigationController,
} from './spreadsheetNavigation';

/** 注册 Excel 批注导航，切换工作表后定位对应单元格。 */
export function useSpreadsheetAnnotationNavigation({
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
    async (annotation: OfficeAnnotation) => {
      if (annotation.target.kind !== 'spreadsheet-cell') return false;
      if (annotation.target.sheetId !== activeSheetId) {
        onSelectSheet(annotation.target.sheetId);
      }
      const controller = await waitForSpreadsheetNavigationController(
        navigationControllerRef,
        annotation.target.sheetId,
      );
      return Boolean(
        controller?.scrollToCell(
          annotation.target.row,
          annotation.target.column,
        ),
      );
    },
    [activeSheetId, navigationControllerRef, onSelectSheet],
  );
  useOfficeAnnotationNavigation('spreadsheet-cell', navigate);
}
