import type { ReactElement, ReactNode } from 'react';
import React, { createContext, useContext, useMemo, useState } from 'react';
import './index.less';
import { resolveOfficeViewerWatermark } from './resolveOfficeViewerWatermark';
import type {
  OfficeViewerWatermark,
  ResolvedOfficeViewerWatermark,
} from './types';

const OfficeWatermarkContext = createContext<
  ResolvedOfficeViewerWatermark | undefined
>(undefined);

let watermarkPatternSequence = 0;

/** 为 React 16 环境生成实例内稳定且互不冲突的 SVG 图案标识。 */
function createWatermarkPatternId(): string {
  watermarkPatternSequence += 1;
  return `office-file-watermark-pattern-${watermarkPatternSequence}`;
}

/** 水印上下文属性。 */
type OfficeWatermarkProviderProps = {
  /** 当前实例的水印配置。 */
  watermark?: OfficeViewerWatermark;
  /** 需要共享水印配置的预览结构。 */
  children: ReactNode;
};

/** 为各格式内容视口提供同一份已归一化水印配置。 */
export function OfficeWatermarkProvider({
  watermark,
  children,
}: OfficeWatermarkProviderProps): ReactElement {
  const resolvedWatermark = useMemo(
    () => resolveOfficeViewerWatermark(watermark),
    [watermark],
  );
  return (
    <OfficeWatermarkContext.Provider value={resolvedWatermark}>
      {children}
    </OfficeWatermarkContext.Provider>
  );
}

/** 根据文本尺寸估算重复图案宽度，避免长文本互相覆盖。 */
function resolvePatternWidth(watermark: ResolvedOfficeViewerWatermark): number {
  const longestLineLength = watermark.content.reduce(
    (length, line) => Math.max(length, Array.from(line).length),
    0,
  );
  return Math.max(
    120,
    longestLineLength * watermark.fontSize * 0.68 + watermark.gap[0],
  );
}

/** 根据多行文字和垂直间距计算重复图案高度。 */
function resolvePatternHeight(
  watermark: ResolvedOfficeViewerWatermark,
): number {
  const textHeight = watermark.content.length * watermark.fontSize * 1.5;
  return Math.max(80, textHeight + watermark.gap[1]);
}

/** 在内容视口上方绘制单个 SVG 重复图案，避免创建大量水印节点。 */
function OfficeWatermarkLayer(): ReactElement | null {
  const watermark = useContext(OfficeWatermarkContext);
  const [patternId] = useState(createWatermarkPatternId);
  if (!watermark) return null;

  const patternWidth = resolvePatternWidth(watermark);
  const patternHeight = resolvePatternHeight(watermark);
  const centerX = patternWidth / 2 + watermark.offset[0];
  const centerY = patternHeight / 2 + watermark.offset[1];
  const lineHeight = watermark.fontSize * 1.5;
  const firstLineY =
    centerY - ((watermark.content.length - 1) * lineHeight) / 2;

  return (
    <svg
      className="office-file-watermark-layer"
      aria-hidden="true"
      focusable="false"
      preserveAspectRatio="none"
    >
      <defs>
        <pattern
          id={patternId}
          width={patternWidth}
          height={patternHeight}
          patternUnits="userSpaceOnUse"
        >
          <g transform={`rotate(${watermark.rotate} ${centerX} ${centerY})`}>
            {watermark.content.map((line, index) => (
              <text
                key={`${line}-${index}`}
                x={centerX}
                y={firstLineY + index * lineHeight}
                fill={watermark.color}
                fillOpacity={watermark.opacity}
                fontFamily={watermark.fontFamily}
                fontSize={watermark.fontSize}
                fontWeight={watermark.fontWeight}
                textAnchor="middle"
                dominantBaseline="middle"
              >
                {line}
              </text>
            ))}
          </g>
        </pattern>
      </defs>
      <rect width="100%" height="100%" fill={`url(#${patternId})`} />
    </svg>
  );
}

/** 水印内容视口属性。 */
type OfficeWatermarkSurfaceProps = {
  /** 参与原格式布局的额外类名。 */
  className?: string;
  /** 需要被水印覆盖的内容。 */
  children: ReactNode;
};

/** 创建稳定内容视口；无水印时也保持相同 DOM 与滚动布局。 */
export function OfficeWatermarkSurface({
  className,
  children,
}: OfficeWatermarkSurfaceProps): ReactElement {
  return (
    <div
      className={['office-file-watermark-surface', className]
        .filter(Boolean)
        .join(' ')}
    >
      {children}
      <OfficeWatermarkLayer />
    </div>
  );
}
