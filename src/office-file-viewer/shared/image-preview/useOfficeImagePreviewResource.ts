import { useMemo } from 'react';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../services/resource-store';
import type { OfficeImagePreviewTarget } from './types';

/** 在菜单或弹层存活期间独立持有图片资源，避免虚拟列表卸载后地址失效。 */
export function useOfficeImagePreviewResource(
  target: OfficeImagePreviewTarget | undefined,
  generation = 0,
) {
  const source = useMemo<OfficeResourceSource | undefined>(() => {
    if (!target) return undefined;
    return typeof target.source === 'string'
      ? { kind: 'url', url: target.source }
      : { ...target.source };
  }, [generation, target]);
  return useOfficeResourceUrl(source);
}
