import type {
  OfficeAnnotation,
  OfficeWordAnnotationTarget,
} from '../../annotations/types';

/** Word 可恢复的修订内容类别。 */
export type WordRevisionKind =
  | 'insert'
  | 'delete'
  | 'move-from'
  | 'move-to'
  | 'format';

/** 单处 Word 修订的作者、时间与语义类别。 */
export type WordRevision = Readonly<{
  /** 在当前文档中稳定且唯一的修订标识。 */
  id: string;
  /** 插入、删除、移动或格式变化。 */
  kind: WordRevisionKind;
  /** 修订作者；源文件缺失时保持为空。 */
  author?: string;
  /** 源文件提供的 ISO 日期或原始日期文本。 */
  createdAt?: string;
}>;

/** 可供审阅面板定位和展示的 Word 修订记录。 */
export type WordRevisionRecord = WordRevision &
  Readonly<{
    /** 修订正文的短摘要；格式修订没有可见文字时保持为空。 */
    excerpt: string;
    /** 首个修订片段在正文中的稳定位置。 */
    target: OfficeWordAnnotationTarget;
  }>;

/** 行内内容关联的批注范围和一个或多个修订。 */
export type WordInlineReview = Readonly<{
  /** 覆盖当前行内内容的批注标识。 */
  annotationIds?: readonly string[];
  /** 按外层到内层顺序覆盖当前内容的修订。 */
  revisions?: readonly WordRevision[];
}>;

/** Word 审阅语义降级时提供的稳定诊断。 */
export type WordReviewWarning = Readonly<{
  /** 供宿主稳定识别降级类型的代码。 */
  code: string;
  /** 面向日志展示的可读说明。 */
  message: string;
}>;

/** Word 文档级批注、修订和注释摘要。 */
export type WordReviewDocument = Readonly<{
  /** 按源文档顺序排列的批注及回复。 */
  annotations: readonly OfficeAnnotation[];
  /** 按正文首次出现顺序排列、可供审阅面板定位的修订。 */
  revisions: readonly WordRevisionRecord[];
  /** 当前文档包含的修订记录数量。 */
  revisionCount: number;
  /** 当前文档包含的脚注和尾注数量。 */
  noteCount: number;
  /** 是否支持最终态、标记态和原始态切换。 */
  supportsRevisionModes: boolean;
  /** 无法完整恢复但不阻断预览的诊断。 */
  warnings: readonly WordReviewWarning[];
}>;
