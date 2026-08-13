import { parseOfficeXmlDocument } from './OfficeXmlDom';

/** 解析 XML 字符串；文档包含语法错误时抛出异常。 */
export function parseXml(xml: string) {
  if (typeof DOMParser === 'undefined') {
    return parseOfficeXmlDocument(xml);
  }
  try {
    return new DOMParser().parseFromString(xml, 'application/xml');
  } catch {
    const document = new DOMParser().parseFromString(xml, 'text/html');
    const root =
      Array.from(document.body?.children ?? []).find(
        (node) => node.nodeType === 1,
      ) ??
      Array.from(document.children ?? []).find((node) => node.nodeType === 1) ??
      document.documentElement;
    return new Proxy(document, {
      get(target, prop, receiver) {
        if (prop === 'documentElement') return root;
        if (prop === 'querySelector')
          return (
            root?.querySelector?.bind(root) ?? target.querySelector.bind(target)
          );
        if (prop === 'querySelectorAll')
          return (
            root?.querySelectorAll?.bind(root) ??
            target.querySelectorAll.bind(target)
          );
        if (prop === 'getElementsByTagName') {
          return (
            root?.getElementsByTagName?.bind(root) ??
            target.getElementsByTagName.bind(target)
          );
        }
        return Reflect.get(target, prop, receiver);
      },
    });
  }
}

/** 读取节点文本并去除首尾空白；空节点返回空字符串。 */
export function textContent(node: Element | null | undefined) {
  return node?.textContent ?? '';
}

/** 读取元素属性；属性不存在时返回 undefined。 */
export function attr(node: Element | null | undefined, name: string) {
  if (!node) {
    return undefined;
  }

  const direct = node.getAttribute(name);
  if (direct !== null) {
    return direct;
  }

  const localName = name.includes(':') ? name.split(':').pop() : name;
  const attributes = node.attributes ? Array.from(node.attributes) : [];
  const matched = attributes.find(
    (item) => item.localName === localName || item.name === name,
  );
  return matched?.value;
}

/** 移除 XML 命名空间前缀并统一为小写本地名称。 */
function normalizedLocalName(node: Element) {
  return (
    node.localName.split(':').pop()?.toLowerCase() ??
    node.localName.toLowerCase()
  );
}

/** 判断元素是否匹配指定的 XML 本地名称。 */
export function matchesLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) return false;
  return normalizedLocalName(node) === localName.toLowerCase();
}

/** 按 XML 本地名称查找第一个后代元素。 */
export function descendantByXmlLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) return null;
  const normalized = localName.toLowerCase();
  return (
    Array.from(node.getElementsByTagName('*')).find(
      (child) =>
        child.nodeName.includes(':') &&
        normalizedLocalName(child) === normalized,
    ) ?? null
  );
}

/** 返回节点列表中的第一个元素节点。 */
export function firstElement(
  node: Element | null | undefined,
  selector: string,
) {
  return node?.querySelector(selector) ?? null;
}

/** 从节点列表中过滤出全部元素节点。 */
export function allElements(
  node: Element | null | undefined,
  selector: string,
) {
  return node ? Array.from(node.querySelectorAll(selector)) : [];
}

/** 按本地名称查找第一个直接子元素。 */
export function childByLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) {
    return null;
  }

  const normalized = localName.toLowerCase();
  return (
    Array.from(node.children).find(
      (child) =>
        normalizedLocalName(child) === normalized ||
        child.localName.toLowerCase() === normalized,
    ) ?? null
  );
}

/** 按本地名称返回全部直接子元素。 */
export function childrenByLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) {
    return [];
  }

  const normalized = localName.toLowerCase();
  return Array.from(node.children).filter(
    (child) =>
      normalizedLocalName(child) === normalized ||
      child.localName.toLowerCase() === normalized,
  );
}

/** 按本地名称查找第一个后代元素，并兼容命名空间前缀。 */
export function descendantByLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) {
    return null;
  }

  const normalized = localName.toLowerCase();

  // 首先尝试标准方法
  const standardMatch = Array.from(node.getElementsByTagName('*')).find(
    (child) =>
      normalizedLocalName(child) === normalized ||
      child.localName.toLowerCase() === normalized,
  );

  if (standardMatch) {
    return standardMatch;
  }

  // 对于 VML 元素（如 v:textbox），尝试使用命名空间 URI 查找
  // VML 命名空间: urn:schemas-microsoft-com:vml
  // WordprocessingML 命名空间: http://schemas.microsoft.com/office/word/2003/wordml
  const vmlNamespaces = [
    'urn:schemas-microsoft-com:vml',
    'http://schemas.microsoft.com/office/word/2003/wordml',
  ];

  for (const ns of vmlNamespaces) {
    try {
      const nsMatch = node.getElementsByTagNameNS(ns, normalized);
      if (nsMatch && nsMatch.length > 0) {
        return nsMatch[0];
      }
    } catch {
      // 某些环境可能不支持 getElementsByTagNameNS，忽略错误
    }
  }

  // 最后尝试通过完整的标签名查找（包括常见的 VML 前缀）
  const prefixes = ['v:', 'w:', 'o:'];
  for (const prefix of prefixes) {
    try {
      const prefixedMatch = node.getElementsByTagName(prefix + normalized);
      if (prefixedMatch && prefixedMatch.length > 0) {
        return prefixedMatch[0];
      }
    } catch {
      // 忽略错误
    }
  }

  return null;
}

/** 按本地名称返回全部后代元素，并兼容命名空间前缀。 */
export function descendantsByLocalName(
  node: Element | null | undefined,
  localName: string,
) {
  if (!node) {
    return [];
  }

  const normalized = localName.toLowerCase();

  // 首先尝试标准方法
  const standardMatches = Array.from(node.getElementsByTagName('*')).filter(
    (child) =>
      normalizedLocalName(child) === normalized ||
      child.localName.toLowerCase() === normalized,
  );

  if (standardMatches.length > 0) {
    return standardMatches;
  }

  // 对于 VML 元素，尝试使用命名空间 URI 查找
  const vmlNamespaces = [
    'urn:schemas-microsoft-com:vml',
    'http://schemas.microsoft.com/office/word/2003/wordml',
  ];

  for (const ns of vmlNamespaces) {
    try {
      const nsMatches = Array.from(node.getElementsByTagNameNS(ns, normalized));
      if (nsMatches.length > 0) {
        return nsMatches;
      }
    } catch {
      // 某些环境可能不支持 getElementsByTagNameNS，忽略错误
    }
  }

  // 最后尝试通过完整的标签名查找（包括常见的 VML 前缀）
  const prefixes = ['v:', 'w:', 'o:'];
  for (const prefix of prefixes) {
    try {
      const prefixedMatches = Array.from(
        node.getElementsByTagName(prefix + normalized),
      );
      if (prefixedMatches.length > 0) {
        return prefixedMatches;
      }
    } catch {
      // 忽略错误
    }
  }

  return [];
}
