import {
  LegendComponent,
  TitleComponent,
  TooltipComponent,
} from 'echarts/components';
import { init, registerMap, use } from 'echarts/core';
import { CanvasRenderer } from 'echarts/renderers';
import type { OfficeChartModel } from '../officeChartTypes';
import {
  resolveOfficeEChartsCapabilities,
  type OfficeEChartsCapability,
} from './officeEChartsCapabilities';

/** 图表视图实际使用的窄 ECharts 运行时。 */
export type OfficeEChartsRuntime = {
  /** 创建 ECharts 实例。 */
  init: typeof init;
  /** 注册地图 GeoJSON 数据。 */
  registerMap: typeof registerMap;
};

/** 公共渲染器和通用组件在页面生命周期内只需注册一次。 */
let sharedRegistered = false;

/** 各能力缓存加载 Promise，以合并并发请求并复用成功结果。 */
const capabilityPromises = new Map<OfficeEChartsCapability, Promise<void>>();

/** 注册所有图表族共用的渲染器和基础组件。 */
function ensureSharedCapability(): void {
  if (sharedRegistered) return;
  use([CanvasRenderer, TooltipComponent, LegendComponent, TitleComponent]);
  sharedRegistered = true;
}

/** 静态路径确保构建器能为每个能力族生成独立异步块。 */
const capabilityLoaders: Record<OfficeEChartsCapability, () => Promise<void>> =
  {
    cartesian: () =>
      import('./registerCartesianCapability').then((module) =>
        module.registerCartesianCapability(),
      ),
    pie: () =>
      import('./registerPieCapability').then((module) =>
        module.registerPieCapability(),
      ),
    radar: () =>
      import('./registerRadarCapability').then((module) =>
        module.registerRadarCapability(),
      ),
    map: () =>
      import('./registerMapCapability').then((module) =>
        module.registerMapCapability(),
      ),
    labels: () =>
      import('./registerLabelCapability').then((module) =>
        module.registerLabelCapability(),
      ),
  };

/** 加载单个能力；失败结果不缓存，允许后续视图重新尝试。 */
function ensureCapability(capability: OfficeEChartsCapability): Promise<void> {
  const existing = capabilityPromises.get(capability);
  if (existing) return existing;

  const loading = capabilityLoaders[capability]().catch((error) => {
    capabilityPromises.delete(capability);
    throw error;
  });
  capabilityPromises.set(capability, loading);
  return loading;
}

/** 在真实绘图前加载并注册当前模型需要的 ECharts 能力。 */
export async function ensureOfficeEChartsRuntime(
  chart: OfficeChartModel,
): Promise<OfficeEChartsRuntime> {
  ensureSharedCapability();
  const capabilities = resolveOfficeEChartsCapabilities(chart);
  await Promise.all(capabilities.map(ensureCapability));
  return { init, registerMap };
}
