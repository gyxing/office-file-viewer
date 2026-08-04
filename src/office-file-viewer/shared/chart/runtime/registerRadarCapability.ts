import { RadarChart } from 'echarts/charts';
import { use } from 'echarts/core';

/** 注册雷达图能力；RadarChart 会安装对应坐标组件。 */
export function registerRadarCapability(): void {
  use([RadarChart]);
}
