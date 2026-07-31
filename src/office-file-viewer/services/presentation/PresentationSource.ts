import type { PresentationPerformanceProfile } from './presentationPerformance';
import type {
  PresentationWarning,
  SlideModel,
  SpeakerNotesModel,
  ThemeModel,
} from './types';

/** 描述单页幻灯片在按需数据源中的轻量状态。 */
export type PresentationSlideDescriptor = {
  /** 幻灯片在当前文稿中的稳定标识。 */
  id: string;
  /** 幻灯片从 1 开始的展示序号。 */
  index: number;
  /** 当前页是否被源文档标记为隐藏。 */
  hidden: boolean;
  /** 当前页是否存在演讲者备注。 */
  hasSpeakerNotes: boolean;
  /** 未完整解析时用于性能画像的元素数量估值。 */
  estimatedElementCount: number;
  /** 当前页内容版本，重试或重新解析时递增。 */
  revision: number;
  /** 当前页尚未读取、已就绪或读取失败。 */
  status: 'estimated' | 'ready' | 'error';
  /** 当前页读取失败时保留的诊断信息。 */
  errorMessage?: string;
};

/** 演示文稿按需数据源对 React 暴露的不可变快照。 */
export type PresentationSourceSnapshot = {
  /** 快照版本，描述符变化时递增。 */
  revision: number;
  /** 演示文稿画布宽度。 */
  width: number;
  /** 演示文稿画布高度。 */
  height: number;
  /** 演示文稿使用的主题。 */
  theme: ThemeModel;
  /** 幻灯片总数。 */
  slideCount: number;
  /** 按源文档顺序排列的轻量幻灯片描述符。 */
  slides: readonly PresentationSlideDescriptor[];
  /** 不阻止预览继续进行的解析警告。 */
  warnings?: readonly PresentationWarning[];
  /** 当前数据源选定的缩略图和幻灯片性能模式。 */
  performance: PresentationPerformanceProfile;
};

/** 为普通文稿和大型按需文稿提供统一的幻灯片读取协议。 */
export interface PresentationSource {
  /** 返回当前可观察状态的只读快照。 */
  getSnapshot(): PresentationSourceSnapshot;
  /** 订阅状态快照变化，并返回取消订阅函数。 */
  subscribe(listener: () => void): () => void;
  /** 读取指定索引的幻灯片模型。 */
  getSlide(index: number, signal?: AbortSignal): Promise<SlideModel>;
  /** 读取指定幻灯片的演讲者备注。 */
  getSpeakerNotes(
    index: number,
    signal?: AbortSignal,
  ): Promise<SpeakerNotesModel | undefined>;
  /** 确保指定内容范围已经开始加载或可用。 */
  ensureRange(start: number, end: number, signal?: AbortSignal): Promise<void>;
  /** 保留指定可视范围并回收远离窗口的缓存内容。 */
  retainRange(start: number, end: number): () => void;
  /** 重新加载此前失败的指定内容。 */
  retry(index: number): void;
  /** 幂等释放当前对象持有的资源和订阅。 */
  dispose(): Promise<void>;
}

/** 创建与浏览器 AbortSignal 语义一致的取消错误。 */
export function createPresentationAbortError(message = '演示文稿读取已取消') {
  const error = new Error(message);
  error.name = 'AbortError';
  return error;
}

/** 在进入可取消的幻灯片读取步骤前统一检查取消状态。 */
export function throwIfPresentationAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createPresentationAbortError();
}
