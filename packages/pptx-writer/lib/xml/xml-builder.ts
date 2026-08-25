import { XMLBuilder, type XmlBuilderOptions } from 'fast-xml-parser';

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Invalid XML 1.0 control characters forbidden in ECMA-376 / ISO-29500 OpenXML schemas.
 * Matches ASCII control characters \x00-\x08, \x0B-\x0C, \x0E-\x1F, \uFFFE, \uFFFF.
 */
// eslint-disable-next-line no-control-regex
export const INVALID_XML_CHARS_REGEX = /[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g;

/**
 * Sanitizes a string by stripping invalid XML 1.0 control characters that cause PowerPoint corruption.
 * @param val Input string or value.
 * @returns Sanitized string or original value.
 */
export function sanitizeXmlText<T>(val: T): T {
  if (typeof val === 'string') {
    return val.replace(INVALID_XML_CHARS_REGEX, '') as T;
  }
  return val;
}

/**
 * Creates a configured XMLBuilder instance tailored for OpenXML serialization.
 * Automatically sanitizes invalid XML 1.0 control characters across attributes and text content.
 * @param options
 */
export function createXmlBuilder(options?: Partial<XmlBuilderOptions>): XMLBuilder {
  return new XMLBuilder({
    attributeNamePrefix: '@_',
    format: false,
    ignoreAttributes: false,
    suppressBooleanAttributes: false,
    suppressEmptyNode: true,
    textNodeName: '#text',
    tagValueProcessor: (_tagName, tagValue) => (typeof tagValue === 'string' ? tagValue.replace(INVALID_XML_CHARS_REGEX, '') : tagValue),
    attributeValueProcessor: (_attrName, attrValue) => (typeof attrValue === 'string' ? attrValue.replace(INVALID_XML_CHARS_REGEX, '') : attrValue),
    ...options,
  });
}

/**
 * Serializes a JavaScript object representation of OpenXML nodes into a complete XML string.
 * @param rootObject The object to serialize.
 * @param options Optional builder settings.
 * @returns Serialized XML string with OpenXML declaration.
 */
export function serializeXml(rootObject: Record<string, unknown>, options?: Partial<XmlBuilderOptions>): string {
  const builder = createXmlBuilder(options);
  const xmlBody = builder.build(rootObject);
  return XML_DECLARATION + xmlBody;
}
