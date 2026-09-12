import { describe, expect, it } from 'vitest';
import { el, sanitizeXmlText, serializeXml } from '../../lib/xml/xml-builder';

describe('XML Builder Re-exports & Serialization', () => {
  it('serializes an XmlElement AST to XML with XML declaration', () => {
    const root = el('p:sld', {
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }, [
      el('p:cSld', [
        el('p:spTree', [
          el('p:nvGrpSpPr', [
            el('p:cNvPr', { id: 1, name: '' }),
            el('p:cNvGrpSpPr'),
            el('p:nvPr'),
          ]),
        ]),
      ]),
    ]);

    const xml = serializeXml(root);
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(xml).toContain('<p:sld');
    expect(xml).toContain('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"');
    expect(xml).toContain('<p:cNvPr id="1" name=""/>');
  });
});

describe('sanitizeXmlText helper', () => {
  it('strips invalid xml chars from string and returns non-string value untouched', () => {
    expect(sanitizeXmlText('Hello\x00World\x08')).toBe('HelloWorld');
    expect(sanitizeXmlText(123)).toBe(123);
    expect(sanitizeXmlText(null)).toBe(null);

    const node = el('item', { numericAttr: 42 }, 100);
    const xml = serializeXml(node);
    expect(xml).toContain('numericAttr="42"');
    expect(xml).toContain('100');
  });
});
