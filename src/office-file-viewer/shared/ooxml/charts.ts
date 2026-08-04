// 保留既有导入路径，解析器暂不感知图表模型和渲染模块已经解耦。
export type {
  OfficeChartModel,
  OfficeChartSeries,
  OfficeChartType,
  OfficeDataLabels,
} from '../chart/officeChartTypes';
export { decodeMojibake, parseOfficeChartXml } from './parseOfficeChartXml';
