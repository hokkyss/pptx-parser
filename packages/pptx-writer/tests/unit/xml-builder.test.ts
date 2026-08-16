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
