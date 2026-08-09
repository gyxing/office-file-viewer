import {
  OFFICE_DEFAULT_ZOOM,
  OFFICE_MAX_ZOOM,
  OFFICE_MIN_ZOOM,
} from './constants';

/** 将缩放比例约束到组件支持的范围，并为无效值提供稳定回退。 */
export function normalizeOfficeZoom(
  value: number,
  fallback = OFFICE_DEFAULT_ZOOM,
): number {
  const normalizedFallback = Number.isFinite(fallback)
    ? Math.min(OFFICE_MAX_ZOOM, Math.max(OFFICE_MIN_ZOOM, fallback))
    : OFFICE_DEFAULT_ZOOM;
  if (!Number.isFinite(value)) return normalizedFallback;
  return Math.min(OFFICE_MAX_ZOOM, Math.max(OFFICE_MIN_ZOOM, value));
}
