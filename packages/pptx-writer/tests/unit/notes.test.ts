import { describe, expect, it } from 'vitest';
import type { PptxTextBody } from '@hokkyss/pptx-core';
import { serializeNotesSlide, serializeNotesSlideRels } from '../../lib/serializers/notes-serializer';

describe('serializeNotesSlide', () => {
  it('returns a valid XML string containing p:notes root element', () => {
    const result = serializeNotesSlide('Hello world');
    expect(result).toContain('<?xml version="1.0"');
    expect(result).toContain('<p:notes');
    expect(result).toContain('</p:notes>');
  });

  it('includes all required XML namespaces', () => {
    const result = serializeNotesSlide('test');
    expect(result).toContain('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
    expect(result).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    expect(result).toContain('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"');
  });

  it('includes p:cSld > p:spTree wrapper structure', () => {
    const result = serializeNotesSlide('test');
    expect(result).toContain('<p:cSld>');
    expect(result).toContain('<p:spTree>');
    expect(result).toContain('</p:spTree>');
    expect(result).toContain('</p:cSld>');
  });

  it('includes slide image placeholder shape', () => {
    const result = serializeNotesSlide('test');
    expect(result).toContain('Slide Image Placeholder 1');
    expect(result).toContain('<p:ph type="sldImg"');
  });

  it('includes notes body placeholder shape', () => {
    const result = serializeNotesSlide('test');
    expect(result).toContain('Notes Placeholder 2');
    expect(result).toContain('<p:ph type="body"');
  });

  it('includes clrMapOvr > masterClrMapping element', () => {
    const result = serializeNotesSlide('test');
    expect(result).toContain('<p:clrMapOvr>');
    expect(result).toContain('<a:masterClrMapping');
  });

  it('embeds the text content from a plain string', () => {
    const result = serializeNotesSlide('Speaker notes here');
    expect(result).toContain('p:txBody');
    expect(result).toContain('Speaker notes here');
  });

  it('splits multi-line string into separate paragraphs', () => {
    const result = serializeNotesSlide('Line one\nLine two\nLine three');
    expect(result).toContain('Line one');
    expect(result).toContain('Line two');
    expect(result).toContain('Line three');
  });

  it('accepts a PptxTextBody object and serializes it via the structured path', () => {
    const textBody: PptxTextBody = {
      bodyProperties: {},
      paragraphs: [
        {
          properties: {},
          runs: [{ text: 'Structured note', properties: {} }],
        },
      ],
    };
    const result = serializeNotesSlide(textBody);
    expect(result).toContain('p:notes');
    expect(result).toContain('Structured note');
  });

  it('handles empty string input without throwing', () => {
    expect(() => serializeNotesSlide('')).not.toThrow();
    expect(serializeNotesSlide('')).toContain('p:notes');
  });

  it('handles PptxTextBody with multiple paragraphs', () => {
    const textBody: PptxTextBody = {
      bodyProperties: {},
      paragraphs: [
        { properties: {}, runs: [{ text: 'First paragraph', properties: {} }] },
        { properties: {}, runs: [{ text: 'Second paragraph', properties: {} }] },
      ],
    };
    const result = serializeNotesSlide(textBody);
    expect(result).toContain('First paragraph');
    expect(result).toContain('Second paragraph');
  });
});

describe('serializeNotesSlideRels', () => {
  it('returns a valid XML string with Relationships root', () => {
    const result = serializeNotesSlideRels(1);
    expect(result).toContain('<?xml version="1.0"');
    expect(result).toContain('<Relationships');
    expect(result).toContain('</Relationships>');
  });

  it('includes notesMaster relationship as rId1', () => {
    const result = serializeNotesSlideRels(1);
    expect(result).toContain('Id="rId1"');
    expect(result).toContain('notesMaster');
    expect(result).toContain('Target="../notesMasters/notesMaster1.xml"');
  });

  it('includes slide relationship as rId2 pointing to correct slide number', () => {
    const result = serializeNotesSlideRels(1);
    expect(result).toContain('Id="rId2"');
    expect(result).toContain('/relationships/slide"');
    expect(result).toContain('Target="../slides/slide1.xml"');
  });

  it('uses the correct slide number in target path', () => {
    expect(serializeNotesSlideRels(5)).toContain('Target="../slides/slide5.xml"');
    expect(serializeNotesSlideRels(10)).toContain('Target="../slides/slide10.xml"');
  });
});
