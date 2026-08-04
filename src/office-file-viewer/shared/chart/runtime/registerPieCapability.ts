import { PieChart } from 'echarts/charts';
import { GraphicComponent } from 'echarts/components';
import { use } from 'echarts/core';

/** 注册饼图、环形图和复合饼图所需能力。 */
export function registerPieCapability(): void {
  use([PieChart, GraphicComponent]);
}
