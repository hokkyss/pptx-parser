import type { XmlChild, XmlElement, XmlParser, XmlRaw } from '@hokkyss/pptx-core';

export type { XmlChild, XmlElement, XmlParser, XmlRaw };

const ALWAYS_ARRAY_TAGS = new Set([
  'a:br', 'a:custClr',
  'a:extraClrScheme', 'a:gs',
  'a:p', 'a:r',
  'a:tc', 'a:tr',
  'br', 'c:pt',
  'c:ser', 'custClr',
  'cxnSp', 'Default',
  'extraClrScheme', 'graphicFrame',
  'grpSp', 'gs',
  'guide', 'handoutMasterId',
  'notesMasterId', 'Override',
  'p', 'p:cxnSp',
  'p:graphicFrame', 'p:grpSp',
  'p:handoutMasterId', 'p:notesMasterId',
  'p:pic', 'p:sldId',
  'p:sldMasterId', 'p:sp',
  'p15:guide',
  'pic',
  'pt',
  'r', 'Relationship',
  'ser', 'sldId',
  'sldMasterId', 'sp',
  'tc', 'tr',
]);

const UNESCAPE_MAP: Record<string, string> = {
  '&amp;': '&',
  '&apos;': '\'',
  '&gt;': '>',
  '&lt;': '<',
  '&quot;': '"',
};

/**
 * Decodes standard XML entities and numeric character references.
 * @param str Input string containing encoded XML entities.
 * @returns Decoded string.
 */
export function unescapeXml(str: string): string {
  if (!str.includes('&')) return str;
  return str.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (entity) => {
    const mapped = UNESCAPE_MAP[entity];
    if (mapped) return mapped;
    if (entity.startsWith('&#x') || entity.startsWith('&#X')) {
      const code = parseInt(entity.slice(3, -1), 16);
      return !isNaN(code) ? String.fromCodePoint(code) : entity;
    }
    if (entity.startsWith('&#')) {
      const code = parseInt(entity.slice(2, -1), 10);
      return !isNaN(code) ? String.fromCodePoint(code) : entity;
    }
    return entity;
  });
}

const ATTR_REGEX = /([a-zA-Z0-9_:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/**
 * Extracts and unescapes key-value pairs from an XML attribute string.
 * @param attrStr Raw attribute string inside an XML opening tag.
 * @returns Key-value map of attributes.
 */
function parseAttributes(attrStr: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!attrStr || attrStr.trim().length === 0) return attrs;
  let match: null | RegExpExecArray;
  ATTR_REGEX.lastIndex = 0;
  while ((match = ATTR_REGEX.exec(attrStr)) !== null) {
    const name = match[1];
    const val = unescapeXml(match[2] ?? match[3] ?? '');
    attrs[name] = val;
  }
  return attrs;
}

interface XmlElementInternal extends XmlElement {
  attrs: Record<string, string>;
  children: (string | XmlElementInternal)[];
  tag: string;
  [key: string]: unknown;
}

/**
 * Attaches child element to parent node properties for backwards compatibility.
 */
function attachChildProperty(parent: XmlElementInternal, tagName: string, node: XmlElementInternal): void {
  const colonIdx = tagName.indexOf(':');
  const localName = colonIdx !== -1 ? tagName.slice(colonIdx + 1) : tagName;
  const isArrayExpected = ALWAYS_ARRAY_TAGS.has(tagName) || ALWAYS_ARRAY_TAGS.has(localName);

  // Full tagName property (e.g. 'p:sp')
  const existing = parent[tagName];
  if (existing === undefined) {
    parent[tagName] = isArrayExpected ? [node] : node;
  } else if (Array.isArray(existing)) {
    existing.push(node);
  } else {
    parent[tagName] = [existing, node];
  }

  // Local tagName property (e.g. 'sp')
  if (localName !== tagName) {
    const existingLocal = parent[localName];
    if (existingLocal === undefined) {
      parent[localName] = isArrayExpected ? [node] : node;
    } else if (Array.isArray(existingLocal)) {
      existingLocal.push(node);
    } else {
      parent[localName] = [existingLocal, node];
    }
  }
}

/**
 * Fast, self-contained, isomorphic XML parser for OpenXML documents.
 * Preserves exact document order in `node.children` while maintaining backwards-compatible object properties.
 * @param xmlString Raw XML string to parse.
 * @returns Parsed XML element tree.
 */
export function parseXml(xmlString: string): XmlElement {
  if (!xmlString || xmlString.trim().length === 0) {
    return {} as unknown as XmlElement;
  }

  const len = xmlString.length;
  let pos = 0;
  const stack: XmlElementInternal[] = [];
  let root: undefined | XmlElementInternal;

  while (pos < len) {
    const nextOpen = xmlString.indexOf('<', pos);
    if (nextOpen === -1) {
      // Remaining trailing text
      if (stack.length > 0 && pos < len) {
        const text = xmlString.slice(pos);
        if (text.trim().length > 0) {
          const unescaped = unescapeXml(text);
          const parent = stack[stack.length - 1];
          parent.children.push(unescaped);
          parent['#text'] = parent['#text'] ? String(parent['#text']) + unescaped : unescaped;
        }
      }
      break;
    }

    // Text between previous position and next '<'
    if (nextOpen > pos && stack.length > 0) {
      const rawText = xmlString.slice(pos, nextOpen);
      // Keep whitespace if inside an element, but skip purely blank lines between tags
      const trimmed = rawText.trim();
      if (trimmed.length > 0 || rawText.includes('\n') === false) {
        const unescaped = unescapeXml(rawText);
        const parent = stack[stack.length - 1];
        parent.children.push(unescaped);
        parent['#text'] = parent['#text'] ? String(parent['#text']) + unescaped : unescaped;
      }
    }

    // Inspect tag type
    if (xmlString.startsWith('<!--', nextOpen)) {
      // Comment <!-- ... -->
      const endComment = xmlString.indexOf('-->', nextOpen + 4);
      pos = endComment === -1 ? len : endComment + 3;
      continue;
    }

    if (xmlString.startsWith('<?', nextOpen)) {
      // XML declaration or processing instruction <? ... ?>
      const endPI = xmlString.indexOf('?>', nextOpen + 2);
      pos = endPI === -1 ? len : endPI + 2;
      continue;
    }

    if (xmlString.startsWith('<![CDATA[', nextOpen)) {
      // CDATA block <![CDATA[ ... ]]>
      const endCDATA = xmlString.indexOf(']]>', nextOpen + 9);
      const cdataContent = endCDATA === -1 ? xmlString.slice(nextOpen + 9) : xmlString.slice(nextOpen + 9, endCDATA);
      if (stack.length > 0) {
        const parent = stack[stack.length - 1];
        parent.children.push(cdataContent);
        parent['#text'] = parent['#text'] ? String(parent['#text']) + cdataContent : cdataContent;
      }
      pos = endCDATA === -1 ? len : endCDATA + 3;
      continue;
    }

    if (xmlString.startsWith('<!', nextOpen)) {
      // DOCTYPE or other declaration
      const endDecl = xmlString.indexOf('>', nextOpen + 2);
      pos = endDecl === -1 ? len : endDecl + 1;
      continue;
    }

    if (xmlString[nextOpen + 1] === '/') {
      // Closing tag </tagName>
      const endClose = xmlString.indexOf('>', nextOpen + 2);
      if (endClose !== -1) {
        stack.pop();
        pos = endClose + 1;
      } else {
        pos = len;
      }
      continue;
    }

    // Opening or self-closing tag
    // Scan for matching '>' skipping quoted attribute values
    let inQuote: null | string = null;
    let endTag = nextOpen + 1;
    while (endTag < len) {
      const ch = xmlString[endTag];
      if (inQuote) {
        if (ch === inQuote) inQuote = null;
      } else if (ch === '"' || ch === '\'') {
        inQuote = ch;
      } else if (ch === '>') {
        break;
      }
      endTag++;
    }

    if (endTag >= len) break;

    const isSelfClosing = xmlString[endTag - 1] === '/';
    const tagContent = xmlString.slice(nextOpen + 1, isSelfClosing ? endTag - 1 : endTag).trim();

    const firstSpace = tagContent.search(/\s/);
    const tagName = firstSpace === -1 ? tagContent : tagContent.slice(0, firstSpace);
    const attrStr = firstSpace === -1 ? '' : tagContent.slice(firstSpace);
    const attrs = parseAttributes(attrStr);

    const node: XmlElementInternal = {
      attrs,
      children: [],
      tag: tagName,
    };

    // Attach attributes for both @_ and direct access
    for (const [k, v] of Object.entries(attrs)) {
      node['@_' + k] = v;
      node[k] = v;
    }

    if (!root) {
      root = node;
    }

    if (stack.length > 0) {
      const parent = stack[stack.length - 1];
      parent.children.push(node);
      attachChildProperty(parent, tagName, node);
    }

    if (!isSelfClosing) {
      stack.push(node);
    }

    pos = endTag + 1;
  }

  if (!root) {
    return {} as unknown as XmlElement;
  }

  const result: XmlElement = {
    ...root,
    [root.tag]: root,
  };
  const colonIdx = root.tag.indexOf(':');
  if (colonIdx !== -1) {
    result[root.tag.slice(colonIdx + 1)] = root;
  }

  return result;
}

/**
 * Creates a fast, isomorphic XML parser conforming to the OpenXML `XmlParser` interface.
 * @returns Frozen `XmlParser` instance.
 */
export function createXmlParser(): XmlParser {
  return Object.freeze({
    parse<T = XmlElement>(xmlString: string): T {
      return parseXml(xmlString) as unknown as T;
    },
  });
}

/** Default shared singleton `XmlParser` instance used across all parsers */
export const defaultXmlParser = createXmlParser();

/**
 * Namespace-agnostic helper to get a single child object node by local tag name.
 * Seamlessly handles prefix variants like `'p:sp'`, `'a:sp'`, `'sp'`.
 * @param node Parent XML object node.
 * @param targetName Target local tag name (e.g. `'sp'`, `'bodyPr'`, `'p'`).
 * @returns Child object node or `undefined` if not found.
 */
export function getXmlChild(node: Record<string, unknown> | undefined, targetName: string): Record<string, unknown> | undefined {
  if (!node || typeof node !== 'object') return undefined;

  // If node has children array, find first matching child in exact document order
  if ('children' in node && Array.isArray(node.children)) {
    for (const child of node.children) {
      if (typeof child === 'object' && child !== null && 'tag' in child) {
        const tag = (child as XmlElement).tag;
        const local = tag.includes(':') ? tag.split(':')[1] : tag;
        if (local === targetName || tag === targetName) {
          return child as Record<string, unknown>;
        }
      }
    }
  }

  // Fallback to property lookup
  for (const key of Object.keys(node)) {
    const localName = key.includes(':') ? key.split(':')[1] : key;
    if (localName === targetName || key === targetName) {
      const val = node[key];
      if (Array.isArray(val)) {
        return val[0] as Record<string, unknown>;
      }
      if (typeof val === 'object' && val !== null) {
        return val as Record<string, unknown>;
      }
    }
  }
  return undefined;
}

/**
 * Namespace-agnostic helper to get an array of child object nodes by local tag name.
 * Seamlessly handles prefix variants like `'p:sp'`, `'a:sp'`, `'sp'`.
 * @param node Parent XML object node.
 * @param targetName Target local tag name (e.g. `'sp'`, `'r'`, `'p'`).
 * @returns Array of matching child object nodes.
 */
export function getXmlChildren(node: Record<string, unknown> | undefined, targetName: string): Record<string, unknown>[] {
  if (!node || typeof node !== 'object') return [];

  // If node has children array, collect matching children in exact document order
  if ('children' in node && Array.isArray(node.children)) {
    const results: Record<string, unknown>[] = [];
    for (const child of node.children) {
      if (typeof child === 'object' && child !== null && 'tag' in child) {
        const tag = (child as XmlElement).tag;
        const local = tag.includes(':') ? tag.split(':')[1] : tag;
        if (local === targetName || tag === targetName) {
          results.push(child as Record<string, unknown>);
        }
      }
    }
    if (results.length > 0) return results;
  }

  // Fallback to property lookup
  const results: Record<string, unknown>[] = [];
  for (const key of Object.keys(node)) {
    const localName = key.includes(':') ? key.split(':')[1] : key;
    if (localName === targetName || key === targetName) {
      const val = node[key];
      if (Array.isArray(val)) {
        for (const item of val) {
          if (typeof item === 'object' && item !== null) {
            results.push(item as Record<string, unknown>);
          }
        }
      } else if (typeof val === 'object' && val !== null) {
        results.push(val as Record<string, unknown>);
      }
    }
  }
  return results;
}

/**
 * Helper to get attribute value from an XML node (checks both plain and @_ prefixed names).
 * @param node XML node object.
 * @param attrName Attribute name to retrieve.
 * @returns String attribute value or undefined.
 */
export function getXmlAttr(node: Record<string, unknown> | undefined, attrName: string): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  if ('attrs' in node && typeof node.attrs === 'object' && node.attrs !== null) {
    const attrs = node.attrs as Record<string, unknown>;
    if (attrs[attrName] !== undefined) return String(attrs[attrName]);
  }
  if (node['@_' + attrName] !== undefined) return String(node['@_' + attrName]);
  if (node[attrName] !== undefined && typeof node[attrName] !== 'object') return String(node[attrName]);
  return undefined;
}

/**
 * Helper to extract text content from an XML node or string value.
 * @param node XML node object or primitive value.
 * @returns String text content or undefined.
 */
export function getXmlText(node: unknown): string | undefined {
  if (node === undefined || node === null) return undefined;
  if (typeof node === 'string') return node;
  if (typeof node === 'number' || typeof node === 'boolean') return String(node);
  if (typeof node === 'object') {
    const el = node as Record<string, unknown>;
    if (el['#text'] !== undefined) return String(el['#text']);
    if ('children' in el && Array.isArray(el.children)) {
      const texts = el.children.filter((c): c is string => typeof c === 'string');
      if (texts.length > 0) return texts.join('');
    }
  }
  return undefined;
}
