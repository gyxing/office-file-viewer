import { BarChart, LineChart, ScatterChart } from 'echarts/charts';
import { AxisPointerComponent, GridComponent } from 'echarts/components';
import { use } from 'echarts/core';

/** 注册 Office 笛卡尔图和组合图所需能力。 */
export function registerCartesianCapability(): void {
  use([LineChart, BarChart, ScatterChart, GridComponent, AxisPointerComponent]);
}
