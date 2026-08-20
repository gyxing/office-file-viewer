import type { OfficeChartModel } from '../../shared/ooxml/charts';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import type { PresentationMediaSource } from '../presentation/mediaTypes';
import type { PresentationTransition } from '../presentation/transitionTypes';
import type {
  PresentationAnnotation,
  PresentationWarning,
  SlideBackground,
  SlideElement,
  SpeakerNotesModel,
  TextStyle,
  ThemeModel,
} from '../presentation/types';

/** PPT 文档级超链接表中的原始目标信息。 */
export type PptHyperlinkEntry = {
  /** ExHyperlinkAtom 声明的链接编号。 */
  id: number;
  /** 链接指向的文件、网址或内部幻灯片标识。 */
  target?: string;
  /** 链接在目标文件内部的子地址。 */
  location?: string;
  /** PowerPoint 为链接保存的可读名称。 */
  friendlyName?: string;
};

/** PPT 二进制流中的单条记录。 */
export type PptRecord = {
  /** 消息或数据结构采用的协议版本号。 */
  version: number;
  /** 当前二进制记录的实例字段。 */
  instance: number;
  /** PPT 二进制记录的类型编号。 */
  type: number;
  /** 当前数据的长度。 */
  length: number;
  /** 在所属数据范围中的偏移位置。 */
  offset: number;
  /** 记录正文相对源流起点的字节偏移。 */
  dataOffset: number;
  /** 记录结束位置相对源流起点的字节偏移。 */
  endOffset: number;
  /** 当前记录正文的原始字节。 */
  data: Uint8Array;
};

/** 汇总 PPT 二进制解析各步骤共享的上下文。 */
export type PptParseContext = {
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: PresentationWarning[];
  /** 持有的图片、字体或对象 URL 等资源；文档释放时需同步清理。 */
  resources: PortableResource[];
  /** 为解析期资源分配标识时使用的递增序号。 */
  resourceSequence: number;
  /** 按图片对象编号索引的资源地址。 */
  blipUrls: Map<number, string>;
  /** 按图表对象编号索引的标准图表模型。 */
  charts: Map<
    number,
    {
      /** 图表渲染相关文案。 */
      chart: OfficeChartModel;
      /** 面向用户展示的标题。 */
      title?: string;
    }
  >;
  /** 按 ExHyperlink 编号索引的文档级链接表。 */
  hyperlinks: Map<number, PptHyperlinkEntry>;
  /** 按外部对象编号索引的可播放音视频来源。 */
  presentationMedia: Map<number, PresentationMediaSource>;
  /** PPT 幻灯片标识到演示顺序索引的映射。 */
  slideIndexById: Map<number, number>;
  /** 长任务主动让出主线程时间片的函数。 */
  yieldIfNeeded: () => Promise<void>;
};

/** PPT 持久化对象标识到流偏移的映射。 */
export type PptPersistObjectMap = Map<number, number>;

/** PPT 最近一次保存对应的持久化对象编辑链。 */
export type PptEditChain = {
  /** PPT 根文档对象的持久化标识。 */
  documentPersistId: number;
  /** 分配后续持久化对象标识时使用的起始值。 */
  persistIdSeed: number;
  /** 持久化对象标识到源流偏移的映射。 */
  persistOffsets: PptPersistObjectMap;
  /** 按新到旧顺序记录的 UserEditAtom 流偏移。 */
  editOffsets: number[];
};

/** PPT 幻灯片的持久化标识、页面标识和顺序。 */
export type PptSlideDescriptor = {
  /** PPT 对象在持久化目录中的标识。 */
  persistId: number;
  /** 幻灯片在 PPT 文档列表中的标识。 */
  slideId: number;
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
};

/** PPT 母版的背景、文字默认值和绘制元素。 */
export type PptMasterModel = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** PPT 对象在持久化目录中的标识。 */
  persistId: number;
  /** 当前页面、幻灯片或元素的背景配置。 */
  background?: SlideBackground;
  /** 母版提供的默认文字样式。 */
  textDefaults?: TextStyle;
  /** 按绘制顺序排列的演示文稿元素。 */
  elements: SlideElement[];
};

/** PPT 外部对象及其静态预览引用。 */
export type PptExternalObject = {
  /** 在所属集合中的唯一标识。 */
  id: number;
  /** PPT 对象在持久化目录中的标识。 */
  persistId?: number;
  /** 面向用户展示的名称。 */
  name?: string;
  /** 用于区分联合类型分支的类型标识。 */
  type?: string;
  /** 外部对象静态预览图片的 Blip 标识。 */
  previewBlipId?: number;
};

/** 解析后的 PPT 幻灯片模型。 */
export type PptSlideModel = {
  /** 在所属集合中的唯一标识。 */
  id: string;
  /** PPT 对象在持久化目录中的标识。 */
  persistId: number;
  /** 幻灯片在 PPT 文档列表中的标识。 */
  slideId: number;
  /** 在所属集合中的零基索引。 */
  index: number;
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 幻灯片引用的母版标识。 */
  masterId?: number;
  /** 是否隐藏当前项目。 */
  hidden?: boolean;
  /** 当前页面、幻灯片或元素的背景配置。 */
  background?: SlideBackground;
  /** 当前幻灯片关联的演讲者备注正文。 */
  speakerNotes?: SpeakerNotesModel;
  /** 当前幻灯片包含的只读批注。 */
  annotations?: PresentationAnnotation[];
  /** 从源文件恢复的页级切换。 */
  transition?: PresentationTransition;
  /** 当前页局部降级说明。 */
  warnings?: PresentationWarning[];
  /** 按绘制顺序排列的演示文稿元素。 */
  elements: SlideElement[];
  /** 当前对象在 PowerPoint Document 流中的字节偏移。 */
  sourceOffset: number;
};

/** 解析后的 PPT 文档、母版和幻灯片集合。 */
export type PptBinaryDocument = {
  /** 宽度，单位为标准化渲染像素。 */
  width: number;
  /** 高度，单位为标准化渲染像素。 */
  height: number;
  /** 当前文档使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** 按母版标识索引的 PPT 母版模型。 */
  masters: Map<number, PptMasterModel>;
  /** 按演示文稿顺序排列的幻灯片。 */
  slides: PptSlideModel[];
  /** 按对象标识索引的 PPT 外部对象。 */
  externalObjects: Map<number, PptExternalObject>;
  /** 解析时产生但不阻止继续预览的警告。 */
  warnings: PresentationWarning[];
};

/** 创建一次解析独占的 warning、可传输资源与时间片上下文。 */
export function createPptParseContext(
  yieldIfNeeded: () => Promise<void>,
): PptParseContext {
  return {
    warnings: [],
    resources: [],
    resourceSequence: 0,
    blipUrls: new Map<number, string>(),
    charts: new Map<
      number,
      {
        /** 关联的图表模型。 */
        chart: OfficeChartModel;
        /** 图表标题。 */
        title?: string;
      }
    >(),
    hyperlinks: new Map<number, PptHyperlinkEntry>(),
    presentationMedia: new Map<number, PresentationMediaSource>(),
    slideIndexById: new Map<number, number>(),
    yieldIfNeeded,
  };
}

/** 为 PPT 资源分配会话内稳定且不冲突的标识。 */
export function createPptResourceId(
  context: PptParseContext,
  category: string,
) {
  context.resourceSequence += 1;
  return `ppt:${category}:${context.resourceSequence}`;
}

/** 注册可传输资源，并返回只在解析模型中使用的资源引用。 */
export function registerPptResource(
  context: PptParseContext,
  resource: PortableResource,
) {
  context.resources.push(resource);
  return createResourceReference(resource.id);
}
