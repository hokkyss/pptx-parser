import { describe, expect, it } from 'vitest';
import {
  el,
  escapeXmlAttr,
  escapeXmlText,
  isXmlElement,
  rawXml,
  renderXml,
  sanitizeXmlText,
  serializeXml,
} from '../../lib/xml/xml-element';

describe('XmlElement AST and Compiler', () => {
  it('escapes XML text content correctly without escaping quotes', () => {
    expect(escapeXmlText('Fish & Chips <Salt> "Pepper"')).toBe('Fish &amp; Chips &lt;Salt&gt; "Pepper"');
    expect(escapeXmlText('Normal text')).toBe('Normal text');
  });

  it('escapes XML attribute values correctly including quotes', () => {
    expect(escapeXmlAttr('Fish & "Chips" <Salt> \'Pepper\'')).toBe('Fish &amp; &quot;Chips&quot; &lt;Salt&gt; &apos;Pepper&apos;');
  });

  it('sanitizes XML 1.0 invalid control characters in text and attributes', () => {
    const dirty = 'Hello\x00\x08World\x0B\x0C\x0E\x1F!';
    expect(sanitizeXmlText(dirty)).toBe('HelloWorld!');
    expect(escapeXmlText(dirty)).toBe('HelloWorld!');
    expect(escapeXmlAttr(dirty)).toBe('HelloWorld!');
  });

  it('identifies XmlElements accurately via isXmlElement', () => {
    expect(isXmlElement(el('a:p'))).toBe(true);
    expect(isXmlElement({ tag: 'custom' })).toBe(true);
    expect(isXmlElement(null)).toBe(false);
    expect(isXmlElement('string')).toBe(false);
    expect(isXmlElement({ notTag: 123 })).toBe(false);
  });

  it('constructs empty elements and renders self-closing tags', () => {
    const node = el('a:br');
    expect(node).toEqual({ tag: 'a:br' });
    expect(renderXml(node)).toBe('<a:br/>');
  });

  it('constructs elements with attributes (stripping leading @_ if present)', () => {
    const node = el('a:rPr', { '@_b': '1', i: '1', lang: 'en-US' });
    expect(renderXml(node)).toBe('<a:rPr b="1" i="1" lang="en-US"/>');
  });

  it('formats boolean attributes as "1" and "0"', () => {
    const node = el('p:cNvSpPr', { txBox: true, noGrp: false, empty: undefined });
    expect(renderXml(node)).toBe('<p:cNvSpPr txBox="1" noGrp="0"/>');
  });

  it('constructs elements with single child text', () => {
    const node = el('a:t', 'Hello World & Beyond');
    expect(renderXml(node)).toBe('<a:t>Hello World &amp; Beyond</a:t>');
  });

  it('constructs elements with attrs and children', () => {
    const node = el('a:pPr', { lvl: 1 }, [
      el('a:buChar', { char: '•' }),
    ]);
    expect(renderXml(node)).toBe('<a:pPr lvl="1"><a:buChar char="•"/></a:pPr>');
  });

  it('filters out null, undefined, and false conditional children and flattens arrays', () => {
    const isBold = false;
    const isItalic = true;
    const extraRuns = [el('a:r', el('a:t', 'Extra 1')), el('a:r', el('a:t', 'Extra 2'))];

    const node = el('a:p', [
      isBold && el('a:pPr'),
      null,
      undefined,
      false,
      el('a:r', [
        isItalic && el('a:rPr', { i: '1' }),
        el('a:t', 'Main'),
      ]),
      extraRuns,
    ]);

    expect(renderXml(node)).toBe(
      '<a:p><a:r><a:rPr i="1"/><a:t>Main</a:t></a:r><a:r><a:t>Extra 1</a:t></a:r><a:r><a:t>Extra 2</a:t></a:r></a:p>',
    );
  });

  it('inlines verbatim raw XML when rawXml is used', () => {
    const prebuilt = '<custom:raw foo="bar"><inner/></custom:raw>';
    const node = el('root', [
      rawXml(prebuilt),
    ]);
    expect(renderXml(node)).toBe('<root><custom:raw foo="bar"><inner/></custom:raw></root>');
  });

  it('prepends XML declaration when serializeXml is called', () => {
    const root = el('p:sld', { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main' }, [
      el('p:cSld'),
    ]);
    const xml = serializeXml(root);
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')).toBe(true);
    expect(xml).toContain('<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld/></p:sld>');
  });
});
