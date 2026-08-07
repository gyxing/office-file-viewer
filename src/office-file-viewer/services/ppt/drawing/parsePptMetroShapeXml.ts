/** 轻量 XML 属性，覆盖 DrawingML 文本解析会访问的字段。 */
type PptMetroXmlAttribute = {
  /** 带命名空间前缀的原始属性名。 */
  name: string;
  /** 去除命名空间前缀后的属性名。 */
  localName: string;
  /** 已完成 XML 实体解码的属性值。 */
  value: string;
};

type PptMetroXmlPart = string | PptMetroXmlElement;

function readLocalName(name: string) {
  return name.split(':').pop() ?? name;
}

function decodeXmlEntities(value: string) {
  return value.replace(
    /&(?:amp|lt|gt|quot|apos|nbsp|#\d+|#x[\da-f]+);/gi,
    (entity) => {
      const normalized = entity.toLowerCase();
      if (normalized === '&amp;') return '&';
      if (normalized === '&lt;') return '<';
      if (normalized === '&gt;') return '>';
      if (normalized === '&quot;') return '"';
      if (normalized === '&apos;') return "'";
      if (normalized === '&nbsp;') return '\u00a0';
      const radix = normalized.startsWith('&#x') ? 16 : 10;
      const raw = normalized.slice(radix === 16 ? 3 : 2, -1);
      const codePoint = Number.parseInt(raw, radix);
      return Number.isFinite(codePoint) &&
        codePoint >= 0 &&
        codePoint <= 0x10ffff
        ? String.fromCodePoint(codePoint)
        : entity;
    },
  );
}

/**
 * Worker 中没有 DOMParser；该元素只实现既有 PPTX 文本解析所需的只读 DOM 子集。
 */
class PptMetroXmlElement {
  /** 与浏览器 Element 一致的元素节点编号。 */
  readonly nodeType = 1;
  /** 带命名空间前缀的原始标签名。 */
  readonly nodeName: string;
  /** 去除命名空间前缀后的标签名。 */
  readonly localName: string;
  /** 当前元素的只读属性集合。 */
  readonly attributes: PptMetroXmlAttribute[];
  /** 当前元素的直接子元素。 */
  readonly children: PptMetroXmlElement[] = [];
  private readonly parts: PptMetroXmlPart[] = [];

  constructor(name: string, attributes: PptMetroXmlAttribute[]) {
    this.nodeName = name;
    this.localName = readLocalName(name);
    this.attributes = attributes;
  }

  get firstElementChild() {
    return this.children[0] ?? null;
  }

  get textContent(): string {
    return this.parts
      .map((part) => (typeof part === 'string' ? part : part.textContent))
      .join('');
  }

  appendElement(element: PptMetroXmlElement) {
    this.children.push(element);
    this.parts.push(element);
  }

  appendText(text: string) {
    if (text) this.parts.push(text);
  }

  getAttribute(name: string) {
    return (
      this.attributes.find((attribute) => attribute.name === name)?.value ??
      null
    );
  }

  getElementsByTagName(name: string) {
    const normalized = readLocalName(name).toLowerCase();
    const result: PptMetroXmlElement[] = [];
    const visit = (element: PptMetroXmlElement) => {
      for (const child of element.children) {
        if (
          name === '*' ||
          child.nodeName.toLowerCase() === name.toLowerCase() ||
          child.localName.toLowerCase() === normalized
        ) {
          result.push(child);
        }
        visit(child);
      }
    };
    visit(this);
    return result;
  }

  getElementsByTagNameNS(_namespace: string, localName: string) {
    return this.getElementsByTagName(localName);
  }

  querySelector(selector: string) {
    return this.querySelectorAll(selector)[0] ?? null;
  }

  querySelectorAll(selector: string) {
    return this.getElementsByTagName(selector);
  }
}

function findTagEnd(xml: string, start: number) {
  let quote = '';
  for (let index = start; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) quote = '';
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      return index;
    }
  }
  return -1;
}

function parseAttributes(source: string) {
  const attributes: PptMetroXmlAttribute[] = [];
  const expression = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let match: RegExpExecArray | null;
  while ((match = expression.exec(source))) {
    const name = match[1];
    attributes.push({
      name,
      localName: readLocalName(name),
      value: decodeXmlEntities(match[2] ?? match[3] ?? ''),
    });
  }
  return attributes;
}

/** 将 Office 生成的 metroBlob 形状 XML 转为可复用 PPTX 文本解析器的元素树。 */
export function parsePptMetroShapeXml(xml: string): Element {
  const stack: PptMetroXmlElement[] = [];
  let root: PptMetroXmlElement | undefined;
  let offset = xml.charCodeAt(0) === 0xfeff ? 1 : 0;

  while (offset < xml.length) {
    if (xml.startsWith('<!--', offset)) {
      const end = xml.indexOf('-->', offset + 4);
      if (end < 0) throw new Error('PPT 兼容形状包含未闭合的 XML 注释');
      offset = end + 3;
      continue;
    }
    if (xml.startsWith('<?', offset)) {
      const end = xml.indexOf('?>', offset + 2);
      if (end < 0) throw new Error('PPT 兼容形状包含未闭合的 XML 声明');
      offset = end + 2;
      continue;
    }
    if (xml.startsWith('<![CDATA[', offset)) {
      const end = xml.indexOf(']]>', offset + 9);
      if (end < 0) throw new Error('PPT 兼容形状包含未闭合的 CDATA');
      stack[stack.length - 1]?.appendText(xml.slice(offset + 9, end));
      offset = end + 3;
      continue;
    }
    if (xml.startsWith('<!', offset)) {
      // Office 形状不需要 DTD；忽略声明也避免解析外部实体。
      const end = findTagEnd(xml, offset + 2);
      if (end < 0) throw new Error('PPT 兼容形状包含未闭合的 XML 声明');
      offset = end + 1;
      continue;
    }
    if (xml[offset] !== '<') {
      const end = xml.indexOf('<', offset);
      const textEnd = end < 0 ? xml.length : end;
      stack[stack.length - 1]?.appendText(
        decodeXmlEntities(xml.slice(offset, textEnd)),
      );
      offset = textEnd;
      continue;
    }

    const end = findTagEnd(xml, offset + 1);
    if (end < 0) throw new Error('PPT 兼容形状包含未闭合的 XML 标签');
    const source = xml.slice(offset + 1, end).trim();
    if (source.startsWith('/')) {
      stack.pop();
      offset = end + 1;
      continue;
    }

    const selfClosing = source.endsWith('/');
    const body = selfClosing ? source.slice(0, -1).trimEnd() : source;
    const name = body.match(/^([^\s/>]+)/)?.[1];
    if (!name) throw new Error('PPT 兼容形状包含无效的 XML 标签');
    const element = new PptMetroXmlElement(
      name,
      parseAttributes(body.slice(name.length)),
    );
    stack[stack.length - 1]?.appendElement(element);
    root ??= element;
    if (!selfClosing) stack.push(element);
    offset = end + 1;
  }

  if (!root) throw new Error('PPT 兼容形状缺少 XML 根元素');
  return root as unknown as Element;
}
