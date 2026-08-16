import type { XmlParser } from '@hokkyss/pptx-core';
import { XMLParser } from 'fast-xml-parser';

export type { XmlParser };

/**
 * Creates a fast, isomorphic XML parser wrapping `fast-xml-parser` pre-configured with OpenXML schema rules.
 *
 * Configured to preserve XML attributes under `@_` prefix, text content under `#text`, and automatically force array structures for repeating OpenXML elements (`p:sp`, `a:p`, `a:r`, `a:tr`, `c:ser`, etc.).
 * @returns Frozen `XmlParser` instance.
 * @example
 * ```ts
 * const xmlParser = createXmlParser();
 * const obj = xmlParser.parse<Record<string, unknown>>(xmlString);
 * ```
 */
export function createXmlParser(): XmlParser {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    // Keep text nodes accessible under #text
    textNodeName: '#text',
    // Preserve element arrays for tags that occur multiple times in OpenXML
    isArray: (name) => {
      const arrayElements = [
        'p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp',
        'p:sldId', 'p:sldMasterId', 'p:notesMasterId', 'p:handoutMasterId',
        'a:p', 'a:r', 'a:tr', 'a:tc', 'c:ser', 'c:pt', 'Relationship',
        'Override', 'Default', 'a:gs', 'p15:guide', 'a:custClr', 'a:extraClrScheme',
      ];
      return arrayElements.includes(name);
    },
    // Preserve exact whitespace in text runs
    trimValues: false,
  });

  return Object.freeze({
    parse<T = Record<string, unknown>>(xmlString: string): T {
      if (!xmlString || xmlString.trim().length === 0) {
        return {} as T;
      }
      return parser.parse(xmlString) as T;
    },
  });
}

/** Default shared singleton `XmlParser` instance used across all parsers */
export const defaultXmlParser = createXmlParser();
