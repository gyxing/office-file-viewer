/** 描述流式 OOXML 读取器产出的命名空间安全事件。 */
export type OfficeXmlEvent =
  | {
      type: 'open';
      localName: string;
      namespaceUri?: string;
      attributes: ReadonlyMap<string, string>;
    }
  | { type: 'text'; text: string }
  | { type: 'close'; localName: string; namespaceUri?: string };

/**
 * 按 XML 元素边界读取事件；SAX 实现只在大 XML 路径实际调用时动态加载。
 */
export async function* readOfficeXmlEvents(
  stream: ReadableStream<Uint8Array>,
  signal?: AbortSignal,
): AsyncGenerator<OfficeXmlEvent> {
  const { readOfficeXmlEventsWithSaxes } = await import(
    './SaxesOfficeXmlEventReader'
  );
  yield* readOfficeXmlEventsWithSaxes(stream, signal);
}
