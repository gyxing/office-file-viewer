import { SaxesParser } from 'saxes';

import type { OfficeXmlEvent } from './OfficeXmlEventReader';

/** 创建统一 AbortError，避免底层 Reader 取消原因泄漏为普通 Error。 */
function createXmlAbortError() {
  const error = new Error('XML 读取已取消');
  error.name = 'AbortError';
  return error;
}

function throwIfAborted(signal?: AbortSignal) {
  if (signal?.aborted) throw createXmlAbortError();
}

function yieldToMainThread() {
  return new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
}

/**
 * 使用 saxes 增量解码 UTF-8 XML，避免把完整主 XML 拼接为字符串。
 */
export async function* readOfficeXmlEventsWithSaxes(
  stream: ReadableStream<Uint8Array>,
  signal?: AbortSignal,
): AsyncGenerator<OfficeXmlEvent> {
  throwIfAborted(signal);
  const decoder = new TextDecoder('utf-8', { fatal: true });
  const parser = new SaxesParser({ xmlns: true });
  const pendingEvents: OfficeXmlEvent[] = [];
  let parserError: Error | undefined;

  parser.on('opentag', (tag) => {
    const attributes = new Map<string, string>();
    Object.values(tag.attributes).forEach((attribute) => {
      attributes.set(attribute.name, attribute.value);
    });
    pendingEvents.push({
      type: 'open',
      localName: tag.local,
      namespaceUri: tag.uri || undefined,
      attributes,
    });
  });
  parser.on('text', (text) => {
    if (text) pendingEvents.push({ type: 'text', text });
  });
  parser.on('cdata', (text) => {
    if (text) pendingEvents.push({ type: 'text', text });
  });
  parser.on('closetag', (tag) => {
    pendingEvents.push({
      type: 'close',
      localName: tag.local,
      namespaceUri: tag.uri || undefined,
    });
  });
  parser.on('error', (error) => {
    parserError = error;
  });

  const reader = stream.getReader();
  let completed = false;
  try {
    while (true) {
      throwIfAborted(signal);
      const result = await reader.read();
      throwIfAborted(signal);
      if (result.done) break;

      parser.write(decoder.decode(result.value, { stream: true }));
      if (parserError) throw parserError;
      while (pendingEvents.length) {
        yield pendingEvents.shift()!;
      }
      throwIfAborted(signal);
      await yieldToMainThread();
    }

    parser.write(decoder.decode());
    parser.close();
    if (parserError) throw parserError;
    while (pendingEvents.length) {
      yield pendingEvents.shift()!;
    }
    completed = true;
  } catch (error) {
    if (signal?.aborted) throw createXmlAbortError();
    throw error;
  } finally {
    if (!completed) {
      const reason = signal?.aborted ? createXmlAbortError() : undefined;
      await reader.cancel(reason).catch(() => undefined);
    }
    reader.releaseLock();
  }
}
