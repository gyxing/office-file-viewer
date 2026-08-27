import { use } from 'echarts/core';
import { LabelLayout } from 'echarts/features';

/** 仅在源图表存在可见数据标签时注册标签布局能力。 */
export function registerLabelCapability(): void {
  use([LabelLayout]);
}
