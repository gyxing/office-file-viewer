import type { OfficeAnnotationTarget } from '../../services/annotations/types';
import { useOfficeAnnotationNavigatorRegistration } from './OfficeAnnotationProvider';
import type { OfficeAnnotationNavigator } from './useOfficeAnnotationController';

/** 把格式专属定位函数注册到当前查看器的审阅运行时。 */
export function useOfficeAnnotationNavigation(
  kind: OfficeAnnotationTarget['kind'],
  navigator: OfficeAnnotationNavigator | undefined,
) {
  useOfficeAnnotationNavigatorRegistration(kind, navigator);
}
