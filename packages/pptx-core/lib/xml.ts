/**
 * Fast, isomorphic XML parser interface for OpenXML schema parsing.
 */
export interface XmlParser {
  /**
   * Parses an OpenXML text string into a typed JavaScript object tree.
   * @template T Target object shape (defaults to `Record<string, unknown>`).
   * @param xmlString Raw XML content string.
   * @returns Parsed JavaScript object tree representation of the XML document.
   */
  parse<T = Record<string, unknown>>(xmlString: string): T;
}
