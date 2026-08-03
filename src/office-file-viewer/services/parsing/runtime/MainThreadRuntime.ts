import type { PreviewKind } from '../formatDefinitions';
import { loadOfficeFormatParser } from '../formatParserRegistry';
import type { RuntimeContext, RuntimeSink } from './types';
import { throwIfParseAborted } from './types';

/** 使用统一事件接口在主线程执行格式解析。 */
export class MainThreadRuntime {
  async run(
    file: File,
    kind: PreviewKind,
    context: RuntimeContext,
    sink: RuntimeSink,
  ) {
    try {
      throwIfParseAborted(context.signal);
      const parser = await loadOfficeFormatParser(kind);
      await parser(file, context, sink);
    } catch (error) {
      sink.error(error);
      throw error;
    }
  }

  dispose() {}
}
