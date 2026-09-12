/**
 * Raw XML wrapper node used when pre-formatted XML strings must be injected verbatim.
 */
export interface XmlRaw {
  raw: string;
}

/**
 * Strongly-typed representation of an XML element in an ordered AST.
 */
export interface XmlElement {
  tag: string;
  attrs?: Record<string, boolean | number | string | undefined>;
  children?: (number | string | XmlElement | XmlRaw)[];
  [key: string]: unknown;
}

export type XmlChild = false | null | number | string | undefined | XmlChild[] | XmlElement | XmlRaw;

/**
 * Fast, isomorphic XML parser interface for OpenXML schema parsing.
 */
export interface XmlParser {
  /**
   * Parses an OpenXML text string into an XmlElement or typed JavaScript object tree.
   * @template T Target object shape (defaults to `XmlElement`).
   * @param xmlString Raw XML content string.
   * @returns Parsed XmlElement or object tree representation of the XML document.
   */
  parse<T = XmlElement>(xmlString: string): T;
}
