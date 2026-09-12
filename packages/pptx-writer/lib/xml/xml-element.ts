export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Invalid XML 1.0 control characters forbidden in ECMA-376 / ISO-29500 OpenXML schemas.
 * Matches ASCII control characters \x00-\x08, \x0B-\x0C, \x0E-\x1F, \uFFFE, \uFFFF.
 */
// eslint-disable-next-line no-control-regex
export const INVALID_XML_CHARS_REGEX = /[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g;

const XML_ESCAPE_ATTR_MAP: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  '\'': '&apos;',
};

const XML_ESCAPE_TEXT_MAP: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
};

/**
 * Escapes XML attribute values according to XML 1.0 specification.
 */
export function escapeXmlAttr(val: string): string {
  return val
    .replace(INVALID_XML_CHARS_REGEX, '')
    .replace(/[&<>"']/g, (ch) => XML_ESCAPE_ATTR_MAP[ch] ?? ch);
}

/**
 * Escapes XML text content according to XML 1.0 specification.
 */
export function escapeXmlText(val: string): string {
  return val
    .replace(INVALID_XML_CHARS_REGEX, '')
    .replace(/[&<>]/g, (ch) => XML_ESCAPE_TEXT_MAP[ch] ?? ch);
}

/**
 * Sanitizes a string by stripping invalid XML 1.0 control characters that cause PowerPoint corruption.
 */
export function sanitizeXmlText<T>(val: T): T {
  if (typeof val === 'string') {
    return val.replace(INVALID_XML_CHARS_REGEX, '') as T;
  }
  return val;
}

/**
 * Raw XML wrapper node used when pre-formatted XML strings must be injected verbatim.
 */
export interface XmlRaw {
  raw: string;
}

/**
 * Creates a raw XML node that will be output verbatim without escaping.
 */
export function rawXml(xml: string): XmlRaw {
  return { raw: xml };
}

/**
 *
 */
export function isXmlRaw(val: unknown): val is XmlRaw {
  return typeof val === 'object' && val !== null && 'raw' in val && typeof (val as XmlRaw).raw === 'string';
}

/**
 * Strongly-typed representation of an XML element in an ordered AST.
 */
export interface XmlElement {
  tag: string;
  attrs?: Record<string, boolean | number | string | undefined>;
  children?: (number | string | XmlElement | XmlRaw)[];
}

export type XmlChild = false | null | number | string | undefined | XmlChild[] | XmlElement | XmlRaw;

/**
 *
 */
export function isXmlElement(val: unknown): val is XmlElement {
  return typeof val === 'object' && val !== null && 'tag' in val && typeof (val as Record<string, unknown>).tag === 'string';
}

/**
 *
 */
function flattenChildren(items: XmlChild[]): (number | string | XmlElement | XmlRaw)[] {
  const result: (number | string | XmlElement | XmlRaw)[] = [];
  for (const item of items) {
    if (item === null || item === undefined || item === false) {
      continue;
    }
    if (Array.isArray(item)) {
      result.push(...flattenChildren(item));
    } else {
      result.push(item);
    }
  }
  return result;
}

/**
 * Factory helper function to construct XmlElement nodes with ordered children.
 *
 * Supports flexible signatures:
 * - `el('tag')`
 * - `el('tag', { attr: 'val' })`
 * - `el('tag', [child1, child2])`
 * - `el('tag', child1)`
 * - `el('tag', 'text')`
 * - `el('tag', { attr: 'val' }, [child1, child2])`
 * - `el('tag', { attr: 'val' }, child1)`
 * - `el('tag', { attr: 'val' }, 'text')`
 */
export function el(
  tag: string,
  attrsOrChildren?: Record<string, boolean | number | string | undefined> | XmlChild,
  children?: XmlChild,
): XmlElement {
  if (attrsOrChildren === undefined && children === undefined) {
    return { tag };
  }

  let attrs: Record<string, boolean | number | string | undefined> | undefined;
  let rawChildren: XmlChild[] = [];

  if (
    attrsOrChildren === null
    || attrsOrChildren === false
    || Array.isArray(attrsOrChildren)
    || typeof attrsOrChildren === 'string'
    || typeof attrsOrChildren === 'number'
    || isXmlElement(attrsOrChildren)
    || isXmlRaw(attrsOrChildren)
  ) {
    if (attrsOrChildren !== null && attrsOrChildren !== false) {
      rawChildren = Array.isArray(attrsOrChildren) ? attrsOrChildren : [attrsOrChildren];
    }
    if (children !== undefined && children !== null && children !== false) {
      rawChildren.push(...(Array.isArray(children) ? children : [children]));
    }
  } else if (typeof attrsOrChildren === 'object') {
    attrs = attrsOrChildren;
    if (children !== undefined && children !== null && children !== false) {
      rawChildren = Array.isArray(children) ? children : [children];
    }
  }

  const flattened = flattenChildren(rawChildren);
  const element: XmlElement = { tag };
  if (attrs && Object.keys(attrs).length > 0) {
    element.attrs = attrs;
  }
  if (flattened.length > 0) {
    element.children = flattened;
  }
  return element;
}

/**
 * Renders an XmlElement or child node directly into a valid XML string.
 * Children and attributes are serialized in exact sequence.
 */
export function renderXml(node: XmlChild): string {
  if (node === null || node === undefined || node === false) {
    return '';
  }
  if (typeof node === 'string') {
    return escapeXmlText(node);
  }
  if (typeof node === 'number') {
    return String(node);
  }
  if (isXmlRaw(node)) {
    return node.raw;
  }
  if (Array.isArray(node)) {
    return node.map(renderXml).join('');
  }

  const { tag, attrs, children } = node;
  let attrStr = '';
  if (attrs) {
    for (const [rawKey, rawVal] of Object.entries(attrs)) {
      if (rawVal === undefined) continue;
      const key = rawKey.startsWith('@_') ? rawKey.slice(2) : rawKey;
      let attrVal: string;
      if (rawVal === true) {
        attrVal = '1';
      } else if (rawVal === false) {
        attrVal = '0';
      } else {
        attrVal = String(rawVal);
      }
      attrStr += ` ${key}="${escapeXmlAttr(attrVal)}"`;
    }
  }

  if (!children || children.length === 0) {
    return `<${tag}${attrStr}/>`;
  }

  let inner = '';
  for (let i = 0; i < children.length; i++) {
    inner += renderXml(children[i]);
  }
  return `<${tag}${attrStr}>${inner}</${tag}>`;
}

/**
 * Serializes an XmlElement AST into a complete XML string with XML declaration.
 */
export function serializeXml(root: XmlElement): string {
  return XML_DECLARATION + renderXml(root);
}
