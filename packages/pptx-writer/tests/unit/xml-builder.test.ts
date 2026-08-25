import { describe, expect, it } from 'vitest';
import { createXmlBuilder, serializeXml } from '../../lib/xml/xml-builder';

describe('XML Builder', () => {
  it('creates an xml builder instance with standard OpenXML options', () => {
    const builder = createXmlBuilder();
    expect(builder).toBeDefined();
    expect(typeof builder.build).toBe('function');
  });

  it('serializes a javascript object to XML with XML declaration', () => {
    const obj = {
      'p:sld': {
        '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p:cSld': {
          'p:spTree': {
            'p:nvGrpSpPr': {
              'p:cNvGrpSpPr': {},
              'p:cNvPr': { '@_id': '1', '@_name': '' },
              'p:nvPr': {},
            },
          },
        },
      },
    };

    const xml = serializeXml(obj);
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(xml).toContain('<p:sld');
    expect(xml).toContain('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"');
    expect(xml).toContain('<p:cNvPr id="1" name=""/>');
  });
});

import { sanitizeXmlText } from '../../lib/xml/xml-builder';

describe('sanitizeXmlText helper', () => {
  it('strips invalid xml chars from string and returns non-string value untouched', () => {
    expect(sanitizeXmlText('Hello\x00World\x08')).toBe('HelloWorld');
    expect(sanitizeXmlText(123)).toBe(123);
    expect(sanitizeXmlText(null)).toBe(null);

    const builder = createXmlBuilder();
    const xml = builder.build({
      item: {
        '@_numericAttr': 42,
        '#text': 100,
      },
    });
    expect(xml).toContain('numericAttr="42"');
    expect(xml).toContain('100');
  });
});
