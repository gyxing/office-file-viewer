import { MapChart } from 'echarts/charts';
import { GraphicComponent, VisualMapComponent } from 'echarts/components';
import { use } from 'echarts/core';

/** 注册地图、分级着色和地图说明图形能力。 */
export function registerMapCapability(): void {
  use([MapChart, VisualMapComponent, GraphicComponent]);
}
