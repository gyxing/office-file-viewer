import { decodeMojibake, type OfficeChartModel } from './charts';
import { attr, descendantsByLocalName } from './xml';

/** 描述 WpsJsonObject 在 OOXML 公共解析中的数据结构。 */
type WpsJsonObject = Record<string, unknown>;

/** 描述 WpsChartCommon 在 OOXML 公共解析中的数据结构。 */
type WpsChartCommon = {
  /** WpsChartCommon 对外展示的标题。 */
  title?: string;
  /** 是否显示图表图例。 */
  showLegend: boolean;
  /** WpsChartCommon 的图例停靠位置；未提供时使用来源格式或渲染器的默认行为。 */
  legendPosition?: OfficeChartModel['legendPosition'];
  /** WpsChartCommon 的图例尺寸和文字样式；未提供时使用来源格式或渲染器的默认行为。 */
  legendStyle?: OfficeChartModel['legendStyle'];
  /** 是否显示图表数据标签。 */
  showDataLabels: boolean;
  /** WpsChartCommon 的图表数据标签显示配置。 */
  dataLabels: NonNullable<OfficeChartModel['dataLabels']>;
  /** WpsChartCommon 的 snapshotSrc 文本值。 */
  snapshotSrc?: string;
};

const WPS_CHINA_MAP_URL =
  'https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json';

/** 读取 `readWebExtensionProperties` 所需的源数据，供 OOXML 公共解析使用。 */
function readWebExtensionProperties(webExtensionNode: Element) {
  const properties: Record<string, string> = {};
  descendantsByLocalName(webExtensionNode, 'property').forEach(
    (propertyNode) => {
      const key = attr(propertyNode, 'key');
      const value = attr(propertyNode, 'value');
      if (key && value !== undefined) {
        properties[key] = value;
      }
    },
  );
  return properties;
}

/** 解码 `decodeMojibakeDeep` 接收的源数据。 */
function decodeMojibakeDeep<TValue>(value: TValue): TValue {
  if (typeof value === 'string') {
    return decodeMojibake(value) as TValue;
  }
  if (Array.isArray(value)) {
    return value.map((item) => decodeMojibakeDeep(item)) as TValue;
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [
        key,
        decodeMojibakeDeep(item),
      ]),
    ) as TValue;
  }
  return value;
}

/** 解析 `parseJsonProperty` 接收的数据，并返回 OOXML 公共解析结果。 */
function parseJsonProperty<TValue>(
  properties: Record<string, string>,
  key: string,
): TValue | undefined {
  const raw = properties[key];
  if (!raw) return undefined;

  try {
    // WPS 扩展属性是无固定结构的第三方 JSON，只在解析边界做受控类型断言。
    return decodeMojibakeDeep(JSON.parse(raw)) as TValue;
  } catch {
    try {
      return decodeMojibakeDeep(
        JSON.parse(raw.replace(/[?\uFFFD]quot;/g, '"')),
      ) as TValue;
    } catch {
      return undefined;
    }
  }
}
/** 将输入标准化为 `normalizeLegendPosition` 返回的结构。 */
function normalizeLegendPosition(
  value: unknown,
): OfficeChartModel['legendPosition'] | undefined {
  if (typeof value !== 'string') return undefined;
  const lower = value.toLowerCase();
  if (lower.includes('bottom')) return 'bottom';
  if (lower.includes('top')) return 'top';
  if (lower.includes('left')) return 'left';
  if (lower.includes('right')) return 'right';
  return undefined;
}

/** 将输入标准化为 `normalizeChartColor` 返回的结构。 */
function normalizeChartColor(value: unknown) {
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object') {
    const color =
      (
        value as {
          /** WPS 图表样式配置 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
          color?: unknown;
          /** 局部转换过程中使用的 rgb 配置；未提供时使用来源格式或渲染器的默认行为。 */
          rgb?: unknown;
        }
      ).color ??
      (
        value as {
          /** 局部转换过程中使用的 rgb 配置；未提供时使用来源格式或渲染器的默认行为。 */
          rgb?: unknown;
        }
      ).rgb;
    if (typeof color === 'string') return color;
  }
  return undefined;
}

/** 提取并汇总 `collectChartColors` 返回的数据。 */
function collectChartColors(
  style: WpsJsonObject | undefined,
  fallbackKey: 'seriesThemeColor' | 'fill',
) {
  if (!style) return [];
  if (
    fallbackKey === 'seriesThemeColor' &&
    Array.isArray(style.seriesThemeColor)
  ) {
    return style.seriesThemeColor
      .map(normalizeChartColor)
      .filter((color): color is string => Boolean(color));
  }

  const fill = style.fill as
    | {
        /** 当前结构 包含的 props 有序集合。 */ props?: unknown[];
      }
    | undefined;
  if (fallbackKey === 'fill' && Array.isArray(fill?.props)) {
    return fill.props
      .map((item) =>
        normalizeChartColor(
          (
            item as
              | {
                  /** WPS 图表样式配置 的前景或文本颜色，使用标准化 CSS 颜色值；未提供时沿用来源格式或渲染器的默认规则。 */
                  color?: unknown;
                }
              | undefined
          )?.color,
        ),
      )
      .filter((color): color is string => Boolean(color));
  }
  return [];
}

/** 读取 `readWpsLegendStyle` 所需的源数据，供 OOXML 公共解析使用。 */
function readWpsLegendStyle(
  legend: unknown,
): OfficeChartModel['legendStyle'] | undefined {
  if (!legend || typeof legend !== 'object') return undefined;
  const legendObject = legend as WpsJsonObject;
  const textStyle =
    legendObject.textStyle && typeof legendObject.textStyle === 'object'
      ? (legendObject.textStyle as WpsJsonObject)
      : legendObject;
  const fontSize = Number(textStyle.fontSize);
  const itemWidth = Number(legendObject.itemWidth);
  const itemHeight = Number(legendObject.itemHeight);
  const fontStyle: 'normal' | 'italic' | 'oblique' | undefined =
    textStyle.fontStyle === 'normal' ||
    textStyle.fontStyle === 'italic' ||
    textStyle.fontStyle === 'oblique'
      ? textStyle.fontStyle
      : undefined;
  const fontWeight:
    | 'normal'
    | 'bold'
    | 'bolder'
    | 'lighter'
    | number
    | undefined =
    textStyle.fontWeight === 'normal' ||
    textStyle.fontWeight === 'bold' ||
    textStyle.fontWeight === 'bolder' ||
    textStyle.fontWeight === 'lighter' ||
    typeof textStyle.fontWeight === 'number'
      ? textStyle.fontWeight
      : undefined;
  const normalizedTextStyle = {
    color: normalizeChartColor(textStyle.color),
    fontFamily:
      typeof textStyle.fontFamily === 'string'
        ? textStyle.fontFamily
        : undefined,
    fontSize: Number.isFinite(fontSize) && fontSize > 0 ? fontSize : undefined,
    fontStyle,
    fontWeight,
  };

  return {
    ...(Number.isFinite(itemWidth) && itemWidth > 0 ? { itemWidth } : {}),
    ...(Number.isFinite(itemHeight) && itemHeight > 0 ? { itemHeight } : {}),
    ...(Object.values(normalizedTextStyle).some((value) => value !== undefined)
      ? { textStyle: normalizedTextStyle }
      : {}),
  };
}
/** 读取 `readPercent` 所需的源数据，供 OOXML 公共解析使用。 */
function readPercent(value: unknown) {
  if (typeof value !== 'string') return undefined;
  const parsed = Number(value.replace(/%$/, ''));
  return Number.isFinite(parsed) ? parsed : undefined;
}

/** 读取 `readRadiusPair` 所需的源数据，供 OOXML 公共解析使用。 */
function readRadiusPair(value: unknown): [string, string] | undefined {
  if (
    !Array.isArray(value) ||
    typeof value[0] !== 'string' ||
    typeof value[1] !== 'string'
  )
    return undefined;
  return [value[0], value[1]];
}

/** 读取 `readWpsSeriesStyle` 所需的源数据，供 OOXML 公共解析使用。 */
function readWpsSeriesStyle(style: WpsJsonObject | undefined) {
  return Array.isArray(style?.series) &&
    style.series[0] &&
    typeof style.series[0] === 'object'
    ? (style.series[0] as WpsJsonObject)
    : undefined;
}

/** 读取 `readWpsPiePointStyles` 所需的源数据，供 OOXML 公共解析使用。 */
function readWpsPiePointStyles(
  seriesStyle: WpsJsonObject | undefined,
  count: number,
) {
  const itemStyle = seriesStyle?.itemStyle;
  if (!itemStyle || typeof itemStyle !== 'object') return undefined;
  const borderColor = normalizeChartColor(
    (
      itemStyle as {
        /** 当前局部结构 的边框颜色配置；未提供时使用来源格式或渲染器的默认行为。 */
        borderColor?: unknown;
      }
    ).borderColor,
  );
  const borderWidth = Number(
    (
      itemStyle as {
        /** WPS 图表样式配置的 borderWidth 渲染数值，单位为标准化渲染像素。 */
        borderWidth?: unknown;
      }
    ).borderWidth,
  );
  if (!borderColor && !Number.isFinite(borderWidth)) return undefined;
  return Array.from({ length: count }, () => ({
    borderColor,
    borderWidth: Number.isFinite(borderWidth) ? borderWidth : undefined,
  }));
}

/** 读取 `readCommonChartSettings` 所需的源数据，供 OOXML 公共解析使用。 */
function readCommonChartSettings(
  chartStyle: WpsJsonObject | undefined,
  snapshotSrc?: string,
): WpsChartCommon {
  const title =
    chartStyle?.title && typeof chartStyle.title === 'object'
      ? (chartStyle.title as {
          /** 当前结构 携带或渲染的文本内容。 */
          text?: unknown;
          /** 局部转换过程中使用的 show 配置；未提供时使用来源格式或渲染器的默认行为。 */
          show?: unknown;
        })
      : undefined;
  const legend = chartStyle?.legend;
  const legendObject =
    legend && typeof legend === 'object'
      ? (legend as WpsJsonObject)
      : undefined;
  const label = chartStyle?.label;
  const labelStyle =
    label && typeof label === 'object' ? (label as WpsJsonObject) : undefined;
  const textLabel =
    labelStyle?.textLabel && typeof labelStyle.textLabel === 'object'
      ? (labelStyle.textLabel as WpsJsonObject)
      : undefined;
  const numberLabel =
    labelStyle?.numberLabel && typeof labelStyle.numberLabel === 'object'
      ? (labelStyle.numberLabel as WpsJsonObject)
      : undefined;
  const position =
    typeof labelStyle?.position === 'string'
      ? labelStyle.position
      : typeof textLabel?.position === 'string'
      ? textLabel.position
      : typeof numberLabel?.position === 'string'
      ? numberLabel.position
      : undefined;
  const separator =
    typeof labelStyle?.separator === 'string'
      ? labelStyle.separator
      : typeof textLabel?.separator === 'string'
      ? textLabel.separator
      : typeof numberLabel?.separator === 'string'
      ? numberLabel.separator
      : undefined;
  return {
    title:
      typeof title?.show === 'boolean' &&
      title.show !== false &&
      typeof title.text === 'string'
        ? decodeMojibake(title.text)
        : undefined,
    showLegend: legendObject?.show !== false,
    legendPosition: normalizeLegendPosition(legendObject?.position),
    legendStyle: readWpsLegendStyle(legend),
    showDataLabels: Boolean(
      labelStyle?.show || numberLabel?.show || textLabel?.show,
    ),
    dataLabels: {
      position,
      separator,
      showVal: Boolean(labelStyle?.show ?? numberLabel?.show),
      showCatName: Boolean(
        labelStyle?.showCategoryName ??
          textLabel?.show ??
          numberLabel?.showCatName,
      ),
      showSerName: Boolean(
        labelStyle?.showSeriesName ??
          textLabel?.showSerName ??
          numberLabel?.showSerName,
      ),
      showPercent: Boolean(labelStyle?.showPercent ?? numberLabel?.showPercent),
      showLeaderLines: Boolean(
        labelStyle?.showLeaderLines ?? numberLabel?.showLeaderLines,
      ),
    },
    snapshotSrc,
  };
}

/** 解析 `parseDemoChart` 接收的数据，并返回 OOXML 公共解析结果。 */
function parseDemoChart(
  properties: Record<string, string>,
  demoData: WpsJsonObject,
  style: WpsJsonObject | undefined,
  extStyle: WpsJsonObject | undefined,
  chartStyle: WpsJsonObject | undefined,
  common: WpsChartCommon,
): OfficeChartModel | undefined {
  if (!Array.isArray(demoData.data) || !Array.isArray(demoData.data[0]))
    return undefined;

  const rows = demoData.data
    .slice(1)
    .filter((row): row is unknown[] => Array.isArray(row));
  const headers = demoData.data[0] as unknown[];
  const isPie = String(properties.type ?? '')
    .toLowerCase()
    .includes('pie');
  const pieType =
    typeof style?.pieType === 'string' ? style.pieType.toLowerCase() : '';
  const radius = readRadiusPair(style?.radius);
  const seriesStyle = readWpsSeriesStyle(extStyle) ?? readWpsSeriesStyle(style);
  const roseType =
    seriesStyle?.roseType === 'radius' || seriesStyle?.roseType === 'area'
      ? seriesStyle.roseType
      : undefined;
  const categories = rows
    .map((row) => decodeMojibake(String(row[0] ?? '').trim()))
    .filter(Boolean);
  const seriesNames = headers
    .slice(1)
    .map((header, index) =>
      decodeMojibake(String(header ?? `Series ${index + 1}`).trim()),
    );
  const palette = collectChartColors(chartStyle, 'seriesThemeColor');
  const chartType: OfficeChartModel['type'] =
    isPie && (pieType.includes('doughnut') || radius || roseType)
      ? 'doughnut'
      : isPie
      ? 'pie'
      : 'line';
  const isPieChart = chartType === 'pie' || chartType === 'doughnut';
  const piePointStyles = isPieChart
    ? readWpsPiePointStyles(seriesStyle, categories.length)
    : undefined;
  const sourceSeries: OfficeChartModel['series'] = seriesNames.map(
    (name, index) => ({
      name,
      values: rows.map((row) => Number(row[index + 1] ?? 0) || 0),
      type: isPieChart
        ? 'pie'
        : style?.areaStyle &&
          typeof style.areaStyle === 'object' &&
          (
            style.areaStyle as {
              /** 局部转换过程中使用的 show 配置；未提供时使用来源格式或渲染器的默认行为。 */
              show?: unknown;
            }
          ).show
        ? 'area'
        : 'line',
      color: palette[index],
      smooth: Boolean(
        (style?.smooth as boolean | undefined) ??
          (Array.isArray(style?.series)
            ? (
                style.series as Array<{
                  /** 局部转换过程中使用的 smooth 配置；未提供时使用来源格式或渲染器的默认行为。 */
                  smooth?: unknown;
                }>
              )[0]?.smooth
            : undefined),
      ),
      marker:
        Array.isArray(style?.symbol) &&
        (
          style.symbol[0] as {
            /** 局部转换过程中使用的 show 配置；未提供时使用来源格式或渲染器的默认行为。 */
            show?: unknown;
          }
        )?.show === false
          ? {
              symbol: 'none',
              size:
                Number(
                  (
                    style.symbol[0] as {
                      /** 局部转换过程中使用的 size 配置；未提供时使用来源格式或渲染器的默认行为。 */
                      size?: unknown;
                    }
                  )?.size,
                ) || 6,
            }
          : {
              size:
                Number(
                  Array.isArray(style?.series)
                    ? (
                        style.series[0] as {
                          /** 局部转换过程中使用的 symbolSize 配置；未提供时使用来源格式或渲染器的默认行为。 */
                          symbolSize?: unknown;
                        }
                      )?.symbolSize
                    : undefined,
                ) || 6,
            },
    }),
  );

  const series = sourceSeries.length
    ? sourceSeries.map((item) =>
        isPieChart
          ? {
              ...item,
              pointColors: palette.length ? palette : undefined,
              pointStyles: piePointStyles,
            }
          : item,
      )
    : [
        {
          name: 'Series 1',
          values: rows.map((row) => Number(row[1] ?? 0) || 0),
          type: isPieChart ? ('pie' as const) : ('line' as const),
          pointColors: palette.length ? palette : undefined,
          pointStyles: piePointStyles,
        },
      ];

  return {
    type: chartType,
    categories,
    series,
    ...common,
    holeSize: chartType === 'doughnut' ? readPercent(radius?.[0]) : undefined,
    radius: roseType ? radius : undefined,
    roseType,
    startAngle:
      isPieChart && Number.isFinite(Number(style?.startAngle))
        ? Number(style?.startAngle)
        : isPieChart
        ? 0
        : undefined,
  };
}

/** 解析 `parseMapChart` 接收的数据，并返回 OOXML 公共解析结果。 */
function parseMapChart(
  properties: Record<string, string>,
  mapData:
    | {
        /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
        data?: unknown[];
        /** 局部转换过程中使用的 props 配置；未提供时使用来源格式或渲染器的默认行为。 */
        props?: WpsJsonObject;
      }
    | undefined,
  chartStyle: WpsJsonObject | undefined,
  common: WpsChartCommon,
): OfficeChartModel | undefined {
  if (properties.dschart_type?.toLowerCase() !== 'map') return undefined;
  const table =
    Array.isArray(mapData?.data) && Array.isArray(mapData.data[0])
      ? mapData.data[0]
      : undefined;
  if (!table || table.length < 2) return undefined;

  const rows = table
    .slice(1)
    .filter((row): row is unknown[] => Array.isArray(row));
  const categories = rows
    .map((row) => decodeMojibake(String(row[0] ?? '').trim()))
    .filter(Boolean);
  const header = table[0];
  const tierIndex = Array.isArray(header) && header.length > 2 ? 2 : undefined;
  const seriesName =
    Array.isArray(header) && typeof header[1] === 'string'
      ? decodeMojibake(header[1])
      : 'Series 1';
  const colors = collectChartColors(chartStyle, 'fill');
  const tiers =
    tierIndex !== undefined
      ? rows.map((row) => decodeMojibake(String(row[tierIndex] ?? '').trim()))
      : [];
  const tierNames = Array.from(new Set(tiers.filter(Boolean)));
  const pointColors = tierNames.length
    ? tiers
        .map((tier) => colors[tierNames.indexOf(tier)])
        .filter((color): color is string => Boolean(color))
    : colors;
  const mapRegion =
    chartStyle?.mapRegion && typeof chartStyle.mapRegion === 'object'
      ? (chartStyle.mapRegion as {
          /** 局部转换过程中使用的 country 配置；未提供时使用来源格式或渲染器的默认行为。 */
          country?: unknown;
          /** 局部转换过程中使用的 province 配置；未提供时使用来源格式或渲染器的默认行为。 */
          province?: unknown;
          /** 局部转换过程中使用的 city 配置；未提供时使用来源格式或渲染器的默认行为。 */
          city?: unknown;
        })
      : undefined;

  return {
    type: 'map',
    categories,
    series: [
      {
        name: seriesName,
        values: rows.map((row) => Number(row[1] ?? 0) || 0),
        type: 'map',
        pointColors: pointColors.length ? pointColors : undefined,
        pointLabels: tiers.length ? tiers : undefined,
      },
    ],
    ...common,
    mapSeriesName: seriesName,
    mapName: 'china',
    mapGeoJsonUrl: WPS_CHINA_MAP_URL,
    mapRegion: mapRegion
      ? decodeMojibake(
          String(
            mapRegion.city || mapRegion.province || mapRegion.country || '',
          ),
        )
      : undefined,
  };
}

/**
 * 将 WPS WebExtension 属性转换为统一 Office 图表模型。
 */
export function parseWpsWebExtensionChartModel(
  webExtensionNode: Element,
  snapshotSrc?: string,
): OfficeChartModel | undefined {
  const properties = readWebExtensionProperties(webExtensionNode);
  const demoData = parseJsonProperty<WpsJsonObject>(properties, 'demoData');
  const style = parseJsonProperty<WpsJsonObject>(properties, 'style');
  const extStyle = parseJsonProperty<WpsJsonObject>(properties, 'extStyle');
  const dschart = parseJsonProperty<WpsJsonObject>(properties, 'dschart');
  const mapData =
    dschart && typeof dschart === 'object'
      ? (
          dschart as {
            /** 局部转换过程中使用的 json 配置；未提供时使用来源格式或渲染器的默认行为。 */
            json?: {
              /** 当前结构 当前步骤需要处理的原始或标准化数据。 */
              data?: unknown[];
              /** 局部转换过程中使用的 props 配置；未提供时使用来源格式或渲染器的默认行为。 */
              props?: WpsJsonObject;
            };
          }
        ).json
      : undefined;
  const chartStyle = style ?? mapData?.props;
  const common = readCommonChartSettings(chartStyle, snapshotSrc);

  if (demoData)
    return parseDemoChart(
      properties,
      demoData,
      style,
      extStyle,
      chartStyle,
      common,
    );
  if (dschart) return parseMapChart(properties, mapData, chartStyle, common);
  return undefined;
}
