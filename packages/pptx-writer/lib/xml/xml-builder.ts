import { XMLBuilder, type XmlBuilderOptions } from 'fast-xml-parser';

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * Creates a configured XMLBuilder instance tailored for OpenXML serialization.
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
