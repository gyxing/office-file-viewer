import { SaxesParser } from 'saxes';

/** Worker 内部使用的轻量 XML 属性。 */
type OfficeXmlAttribute = {
  /** 源 XML 中包含前缀的属性名。 */
  name: string;
  /** 不包含命名空间前缀的属性名。 */
  localName: string;
  /** 属性所属命名空间。 */
  namespaceURI: string | null;
  /** 属性原始值。 */
  value: string;
};

type OfficeXmlNode = OfficeXmlElement | OfficeXmlText;

/** 在 Worker 中保存 XML 文本节点。 */
class OfficeXmlText {
  readonly nodeType = 3;
  parentNode: OfficeXmlElement | undefined;

  constructor(readonly data: string) {}

  get textContent() {
    return this.data;
  }

  cloneNode() {
    return new OfficeXmlText(this.data);
  }
}

/** 提取限定名称中的本地名称。 */
function localNameOf(qualifiedName: string) {
  return qualifiedName.split(':').pop() ?? qualifiedName;
}

/** 在 Worker 中提供 OOXML 解析实际依赖的 Element 子集。 */
class OfficeXmlElement {
  readonly nodeType = 1;
  readonly childNodes: OfficeXmlNode[] = [];
  parentNode: OfficeXmlElement | OfficeXmlDocument | undefined;
  readonly attributes: OfficeXmlAttribute[];

  constructor(
    readonly nodeName: string,
    readonly localName: string,
    readonly namespaceURI: string | null,
    attributes: OfficeXmlAttribute[] = [],
  ) {
    this.attributes = attributes;
  }

  get tagName() {
    return this.nodeName;
  }

  get parentElement() {
    return this.parentNode instanceof OfficeXmlElement ? this.parentNode : null;
  }

  get children() {
    return this.childNodes.filter(
      (node): node is OfficeXmlElement => node instanceof OfficeXmlElement,
    );
  }

  get textContent(): string {
    return this.childNodes.map((node) => node.textContent).join('');
  }

  set textContent(value: string) {
    this.childNodes.length = 0;
    if (value) this.appendChild(new OfficeXmlText(value));
  }

  appendChild<TNode extends OfficeXmlNode>(node: TNode): TNode {
    node.parentNode = this;
    this.childNodes.push(node);
    return node;
  }

  getAttribute(name: string) {
    return (
      this.attributes.find((attribute) => attribute.name === name)?.value ??
      null
    );
  }

  getAttributeNS(namespaceURI: string | null, localName: string) {
    return (
      this.attributes.find(
        (attribute) =>
          attribute.namespaceURI === namespaceURI &&
          attribute.localName === localName,
      )?.value ?? null
    );
  }

  hasAttribute(name: string) {
    return this.attributes.some((attribute) => attribute.name === name);
  }

  setAttribute(name: string, value: string) {
    this.setAttributeValue(name, null, value);
  }

  setAttributeNS(namespaceURI: string | null, name: string, value: string) {
    this.setAttributeValue(name, namespaceURI, value);
  }

  getElementsByTagName(name: string) {
    return this.collectDescendants(
      (element) =>
        name === '*' || element.nodeName === name || element.localName === name,
    );
  }

  getElementsByTagNameNS(namespaceURI: string | null, localName: string) {
    return this.collectDescendants(
      (element) =>
        (namespaceURI === '*' || element.namespaceURI === namespaceURI) &&
        (localName === '*' || element.localName === localName),
    );
  }

  querySelector(selector: string) {
    return this.querySelectorAll(selector)[0] ?? null;
  }

  querySelectorAll(selector: string) {
    const normalized = localNameOf(selector.trim());
    return this.collectDescendants(
      (element) => selector === '*' || element.localName === normalized,
    );
  }

  cloneNode(deep = false): OfficeXmlElement {
    const clone = new OfficeXmlElement(
      this.nodeName,
      this.localName,
      this.namespaceURI,
      this.attributes.map((attribute) => ({ ...attribute })),
    );
    if (deep) {
      this.childNodes.forEach((child) =>
        clone.appendChild(child.cloneNode(true)),
      );
    }
    return clone;
  }

  private setAttributeValue(
    name: string,
    namespaceURI: string | null,
    value: string,
  ) {
    const existing = this.attributes.find(
      (attribute) => attribute.name === name,
    );
    if (existing) {
      existing.value = value;
      return;
    }
    this.attributes.push({
      name,
      localName: localNameOf(name),
      namespaceURI,
      value,
    });
  }

  private collectDescendants(
    predicate: (element: OfficeXmlElement) => boolean,
  ) {
    const matches: OfficeXmlElement[] = [];
    const visit = (element: OfficeXmlElement) => {
      element.children.forEach((child) => {
        if (predicate(child)) matches.push(child);
        visit(child);
      });
    };
    visit(this);
    return matches;
  }
}

/** 在 Worker 中提供 OOXML 解析实际依赖的 XMLDocument 子集。 */
class OfficeXmlDocument {
  readonly nodeType = 9;
  documentElement!: OfficeXmlElement;

  get children() {
    return this.documentElement ? [this.documentElement] : [];
  }

  createElement(name: string) {
    return new OfficeXmlElement(name, localNameOf(name), null);
  }

  createElementNS(namespaceURI: string | null, name: string) {
    return new OfficeXmlElement(name, localNameOf(name), namespaceURI);
  }

  createTextNode(text: string) {
    return new OfficeXmlText(text);
  }

  appendChild(element: OfficeXmlElement) {
    element.parentNode = this;
    this.documentElement = element;
    return element;
  }

  importNode(node: OfficeXmlNode, deep = false) {
    return node.cloneNode(deep);
  }

  querySelector(selector: string) {
    if (
      selector === '*' ||
      this.documentElement.localName === localNameOf(selector)
    ) {
      return this.documentElement;
    }
    return this.documentElement.querySelector(selector);
  }

  querySelectorAll(selector: string) {
    const root = this.querySelector(selector);
    const descendants = this.documentElement.querySelectorAll(selector);
    return root === this.documentElement
      ? [this.documentElement, ...descendants]
      : descendants;
  }

  getElementsByTagName(name: string) {
    return this.querySelectorAll(name);
  }
}

/** 使用 saxes 同步构建 Worker 可用的轻量 XML 文档。 */
export function parseOfficeXmlDocument(xml: string): XMLDocument {
  const document = new OfficeXmlDocument();
  const stack: OfficeXmlElement[] = [];
  let parseError: Error | undefined;
  const parser = new SaxesParser({ xmlns: true });

  parser.on('opentag', (tag) => {
    const attributes = Object.values(tag.attributes).map((attribute) => ({
      name: attribute.name,
      localName: attribute.local,
      namespaceURI: attribute.uri || null,
      value: attribute.value,
    }));
    const element = new OfficeXmlElement(
      tag.name,
      tag.local,
      tag.uri || null,
      attributes,
    );
    const parent = stack[stack.length - 1];
    if (parent) parent.appendChild(element);
    else document.appendChild(element);
    stack.push(element);
  });
  parser.on('text', (text) => {
    if (text) stack[stack.length - 1]?.appendChild(new OfficeXmlText(text));
  });
  parser.on('cdata', (text) => {
    if (text) stack[stack.length - 1]?.appendChild(new OfficeXmlText(text));
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.on('error', (error) => {
    parseError = error;
  });

  parser.write(xml).close();
  if (parseError) throw parseError;
  if (!document.documentElement) throw new Error('XML 文档缺少根元素');

  // 适配层只实现项目解析器使用的 DOM 子集，边界由本模块集中收口。
  return document as unknown as XMLDocument;
}
