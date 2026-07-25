import type { OfficeChartModel } from '../../shared/ooxml/charts';
import { createResourceReference } from '../parsing/assembly/resourceReferences';
import type { PortableResource } from '../parsing/protocol/messages';
import type {
  PresentationWarning,
  SlideBackground,
  SlideElement,
  TextStyle,
  ThemeModel,
} from '../presentation/types';

/** 表示PPT 二进制解析读取到的一条记录。 */
export type PptRecord = {
  /** 消息或数据结构采用的协议版本号。 */
  version: number;
  /** PptRecord 从源格式读取的 instance 枚举或标识值。 */
  instance: number;
  /** 用于区分 PptRecord 不同结构分支的类型标识。 */
  type: number;
  /** PptRecord 对应二进制记录或数据块的字节长度。 */
  length: number;
  /** PptRecord 在源二进制流中的字节偏移。 */
  offset: number;
  /** PptRecord 在对应二进制流中的字节偏移。 */
  dataOffset: number;
  /** PptRecord 在对应二进制流中的字节偏移。 */
  endOffset: number;
  /** PptRecord 当前步骤需要处理的原始或标准化数据。 */
  data: Uint8Array;
};

/** 汇总PPT 二进制解析当前步骤需要共享的上下文。 */
export type PptParseContext = {
  /** PptParseContext 解析时产生但不阻止继续预览的警告集合。 */
  warnings: PresentationWarning[];
  /** PptParseContext 持有的图片、字体或对象 URL 等资源；文档释放时需同步清理。 */
  resources: PortableResource[];
  /** PptParseContext 用于分配后续资源或持久化标识的递增值。 */
  resourceSequence: number;
  /** PptParseContext 按业务键索引的 blipUrls 映射。 */
  blipUrls: Map<number, string>;
  /** PptParseContext 按业务键索引的 charts 映射。 */
  charts: Map<
    number,
    {
      /** PptParseContext 当前关联的图表模型。 */
      chart: OfficeChartModel;
      /** PptParseContext 对外展示的标题。 */
      title?: string;
    }
  >;
  /** PptParseContext 执行 yieldIfNeeded 操作时调用的函数。 */
  yieldIfNeeded: () => Promise<void>;
};

/** 描述 PptPersistObjectMap 在 PPT 二进制解析中的数据结构。 */
export type PptPersistObjectMap = Map<number, number>;

/** 描述 PptEditChain 在 PPT 二进制解析中的数据结构。 */
export type PptEditChain = {
  /** PptEditChain 在源文件记录中的数字标识。 */
  documentPersistId: number;
  /** PptEditChain 用于分配后续资源或持久化标识的递增值。 */
  persistIdSeed: number;
  /** PptEditChain 关联的 persistOffsets 结构；字段形状由 PptPersistObjectMap 定义。 */
  persistOffsets: PptPersistObjectMap;
  /** PptEditChain 包含的 editOffsets 有序集合。 */
  editOffsets: number[];
};

/** 描述 PptSlideDescriptor 在 PPT 二进制解析中的数据结构。 */
export type PptSlideDescriptor = {
  /** PptSlideDescriptor 在源文件记录中的数字标识。 */
  persistId: number;
  /** PptSlideDescriptor 在源文件记录中的数字标识。 */
  slideId: number;
  /** PptSlideDescriptor 在所属集合中的位置索引。 */
  index: number;
  /** 是否隐藏 PptSlideDescriptor；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
};

/** 描述 PPT 二进制解析使用的标准化模型。 */
export type PptMasterModel = {
  /** PptMasterModel 在所属文档或任务中的唯一标识。 */
  id: number;
  /** PptMasterModel 在源文件记录中的数字标识。 */
  persistId: number;
  /** PptMasterModel 的背景填充模型；未提供时使用来源格式或渲染器的默认行为。 */
  background?: SlideBackground;
  /** PptMasterModel 关联的 textDefaults 结构；字段形状由 TextStyle 定义；未提供时使用来源格式或渲染器的默认行为。 */
  textDefaults?: TextStyle;
  /** PptMasterModel 包含的 elements 有序集合。 */
  elements: SlideElement[];
};

/** 描述 PptExternalObject 在 PPT 二进制解析中的数据结构。 */
export type PptExternalObject = {
  /** PptExternalObject 在所属文档或任务中的唯一标识。 */
  id: number;
  /** PptExternalObject 在源文件记录中的数字标识。 */
  persistId?: number;
  /** PptExternalObject 的可读名称。 */
  name?: string;
  /** 用于区分 PptExternalObject 不同结构分支的类型标识。 */
  type?: string;
  /** PptExternalObject 在源文件记录中的数字标识。 */
  previewBlipId?: number;
};

/** 描述 PPT 二进制解析使用的标准化模型。 */
export type PptSlideModel = {
  /** PptSlideModel 在所属文档或任务中的唯一标识。 */
  id: string;
  /** PptSlideModel 在源文件记录中的数字标识。 */
  persistId: number;
  /** PptSlideModel 在源文件记录中的数字标识。 */
  slideId: number;
  /** PptSlideModel 在所属集合中的位置索引。 */
  index: number;
  /** PptSlideModel 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** PptSlideModel 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
  /** PptSlideModel 在源文件记录中的数字标识。 */
  masterId?: number;
  /** 是否隐藏 PptSlideModel；未提供时沿用来源格式或渲染器的默认规则。 */
  hidden?: boolean;
  /** PptSlideModel 的背景填充模型；未提供时使用来源格式或渲染器的默认行为。 */
  background?: SlideBackground;
  /** PptSlideModel 包含的 elements 有序集合。 */
  elements: SlideElement[];
  /** PptSlideModel 在对应二进制流中的字节偏移。 */
  sourceOffset: number;
};

/** 描述 PPT 二进制解析生成的标准化文档模型。 */
export type PptBinaryDocument = {
  /** PptBinaryDocument 的 width 几何值，单位遵循对应 Office 二进制记录定义。 */
  width: number;
  /** PptBinaryDocument 的 height 几何值，单位遵循对应 Office 二进制记录定义。 */
  height: number;
  /** PptBinaryDocument 使用的主题颜色和字体配置。 */
  theme: ThemeModel;
  /** PptBinaryDocument 按业务键索引的 masters 映射。 */
  masters: Map<number, PptMasterModel>;
  /** PptBinaryDocument 包含的 slides 有序集合。 */
  slides: PptSlideModel[];
  /** PptBinaryDocument 按业务键索引的 externalObjects 映射。 */
  externalObjects: Map<number, PptExternalObject>;
  /** PptBinaryDocument 解析时产生但不阻止继续预览的警告集合。 */
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
        /** 当前结构 当前关联的图表模型。 */
        chart: OfficeChartModel;
        /** 当前结构 对外展示的标题。 */
        title?: string;
      }
    >(),
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
