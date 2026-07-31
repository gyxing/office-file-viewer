/** 描述流式 OOXML 读取器产出的命名空间安全事件。 */
export type OfficeXmlEvent =
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'open';
      /** 不含命名空间前缀的 XML 节点名称。 */
      localName: string;
      /** XML 节点所属的命名空间地址。 */
      namespaceUri?: string;
      /** 当前开始标签上的属性键值。 */
      attributes: ReadonlyMap<string, string>;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'text';
      /** 文本内容。 */
      text: string;
    }
  | {
      /** 用于区分联合类型分支的类型标识。 */
      type: 'close';
      /** 不含命名空间前缀的 XML 节点名称。 */
      localName: string;
      /** XML 节点所属的命名空间地址。 */
      namespaceUri?: string;
    };

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
