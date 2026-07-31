import type { DocBlock } from '../../doc/types';
import type { SlideModel } from '../../presentation/types';
import type { ParsedOfficeFile } from '../../preview';
import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import type {
  PortableDocMetadata,
  PortablePresentationMetadata,
  PortableResource,
} from '../protocol/messages';
import type { ParseProgress } from '../types';

/** 为同一次解析运行时提供稳定会话标识和统一取消信号。 */
export type RuntimeContext = {
  /** 当前文档解析会话的标识。 */
  documentSessionId: string;
  /** 用于取消当前异步操作的信号。 */
  signal: AbortSignal;
};

/** 定义解析运行时输出分块结果时调用的接收接口。 */
export type RuntimeSink = {
  /** 接收并转发解析进度。 */
  progress(progress: ParseProgress): void;
  /** 接收解析器产生的可移植资源分块。 */
  resource(resource: PortableResource): Promise<void>;
  /** 接收解析产生的单张工作表。 */
  sheet(
    index: number,
    revision: number,
    sheet: SpreadsheetSheet,
  ): Promise<void>;
  /** 接收演示文稿的主体元数据。 */
  presentationMetadata(metadata: PortablePresentationMetadata): Promise<void>;
  /** 接收解析产生的单张幻灯片。 */
  slide(index: number, slide: SlideModel): Promise<void>;
  /** 接收文字文档的主体元数据。 */
  documentMetadata(metadata: PortableDocMetadata): Promise<void>;
  /** 接收文字文档的连续内容块。 */
  documentBlocks(startIndex: number, blocks: DocBlock[]): Promise<void>;
  /** 接收解析完成的标准化文件结果。 */
  parsed(parsed: ParsedOfficeFile): Promise<void>;
  /** 通知接收方解析任务已经完成。 */
  complete(warnings?: SpreadsheetWarning[]): void | Promise<void>;
  /** 通知接收方解析任务发生错误。 */
  error(error: unknown): void;
};

/** 创建跨运行时一致的取消错误。 */
export function createParseAbortError() {
  const error = new Error('文件解析已取消');
  error.name = 'AbortError';
  return error;
}
