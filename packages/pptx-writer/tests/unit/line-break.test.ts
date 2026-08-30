import { describe, expect, it } from 'vitest';
import type { PptxParagraph, PptxShapeElement, PptxTextBody } from '@hokkyss/pptx-core';
import {
  serializeParagraph,
  serializeTextBody,
} from '../../lib/serializers/text-serializer';
import { serializeShape } from '../../lib/serializers/shape-serializer';
import { serializeSlide } from '../../lib/serializers/slide-serializer';

describe('Line Break and Granular Bullet Level Serializer', () => {
  it('serializes Shift+Enter line break sentinel { break: true } as <a:br>', () => {
    const paragraph: PptxParagraph = {
      properties: {
        level: 0,
      },
      runs: [
        { text: 'A1' },
        { break: true, properties: {}, text: '' },
        { text: 'a1 level, no bullet' },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    expect(typeof serialized).toBe('string');
    expect(serialized).toContain('<a:r><a:t>A1</a:t></a:r>');
    expect(serialized).toContain('<a:br');
    expect(serialized).toContain('<a:r><a:t>a1 level, no bullet</a:t></a:r>');

    // Check correct ordering: A1 before <a:br>, <a:br> before second text
    const a1Index = (serialized as string).indexOf('A1');
    const brIndex = (serialized as string).indexOf('<a:br');
    const secondTextIndex = (serialized as string).indexOf('a1 level, no bullet');

    expect(a1Index).toBeLessThan(brIndex);
    expect(brIndex).toBeLessThan(secondTextIndex);
  });

  it('serializes <a:br> carrying optional run properties in <a:rPr>', () => {
    const paragraph: PptxParagraph = {
      properties: {},
      runs: [
        { text: 'Before' },
        { break: true, properties: { bold: true, italic: true }, text: '' },
        { text: 'After' },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    expect(typeof serialized).toBe('string');
    expect(serialized).toMatch(/<a:br><a:rPr b="1" i="1"/);
  });

  it('preserves entity escaping (&, <, >) when serializing paragraphs with line breaks', () => {
    const paragraph: PptxParagraph = {
      properties: { level: 1 },
      runs: [
        { text: 'Fish & Chips <Salt>' },
        { break: true, properties: {}, text: '' },
        { text: 'Line 2 & More' },
      ],
    };

    const serialized = serializeParagraph(paragraph) as string;
    expect(serialized).toContain('Fish &amp; Chips &lt;Salt&gt;');
    expect(serialized).toContain('Line 2 &amp; More');
  });

  it('does NOT emit marL or indent when only level is specified (master inheritance)', () => {
    const paragraph: PptxParagraph = {
      properties: {
        level: 2,
      },
      runs: [
        { text: 'Deport Topic Lvl 2' },
      ],
    };

    const serialized = serializeParagraph(paragraph) as Record<string, unknown>;
    const pPr = serialized['a:pPr'] as Record<string, unknown>;
    expect(pPr).toBeDefined();
    expect(pPr['@_lvl']).toBe(2);
    expect(pPr['@_marL']).toBeUndefined();
    expect(pPr['@_indent']).toBeUndefined();
    expect(pPr['a:buChar']).toBeUndefined();
    expect(pPr['a:buNone']).toBeUndefined();
  });

  it('emits marL and indent only when explicit bullet is provided', () => {
    const paragraph: PptxParagraph = {
      properties: {
        bullet: { char: '•', type: 'char' },
        level: 1,
      },
      runs: [
        { text: 'Explicit Bullet Item' },
      ],
    };

    const serialized = serializeParagraph(paragraph) as Record<string, unknown>;
    const pPr = serialized['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_lvl']).toBe(1);
    expect(pPr['@_marL']).toBeDefined();
    expect(pPr['@_indent']).toBeDefined();
    expect(pPr['a:buChar']).toBeDefined();
  });

  it('serializes text body containing line breaks and plain paragraphs together', () => {
    const textBody: PptxTextBody = {
      bodyProperties: { verticalAlignment: 'top' },
      paragraphs: [
        {
          properties: { level: 0 },
          runs: [
            { text: 'A1' },
            { break: true, properties: {}, text: '' },
            { text: 'a1 continuation' },
          ],
        },
        {
          properties: { level: 1 },
          runs: [{ text: 'B1' }],
        },
      ],
    };

    const serialized = serializeTextBody(textBody);
    expect(typeof serialized).toBe('string');
    expect(serialized).toContain('<a:p><a:r><a:t>A1</a:t></a:r><a:br/><a:r><a:t>a1 continuation</a:t></a:r></a:p>');
    expect(serialized).toContain('<a:p><a:pPr lvl="1"/><a:r><a:t>B1</a:t></a:r></a:p>');
  });

  it('serializes complete shape containing line breaks into valid shape XML string', () => {
    const shape: PptxShapeElement = {
      elementType: 'shape',
      id: '3',
      name: 'Content Placeholder 2',
      placeholder: { idx: 1, type: 'body' },
      textBody: {
        bodyProperties: {},
        paragraphs: [
          {
            properties: { level: 0 },
            runs: [
              { text: 'Root' },
              { break: true, properties: {}, text: '' },
              { text: 'Continuation without bullet' },
            ],
          },
          {
            properties: { level: 1 },
            runs: [{ text: 'Sub level 1' }],
          },
        ],
      },
    };

    const serialized = serializeShape(shape);
    expect(typeof serialized).toBe('string');
    expect(serialized).toContain('<p:sp>');
    expect(serialized).toContain('<p:txBody>');
    expect(serialized).toContain('<a:br/>');
  });

  it('serializes full slide with mixed line-break shape and normal shapes', () => {
    const slide = {
      elements: [
        {
          elementType: 'shape' as const,
          id: '2',
          name: 'Title',
          textBody: {
            bodyProperties: {},
            paragraphs: [{ properties: {}, runs: [{ text: 'Slide Title' }] }],
          },
        },
        {
          elementType: 'shape' as const,
          id: '3',
          name: 'Body with Breaks',
          textBody: {
            bodyProperties: {},
            paragraphs: [
              {
                properties: { level: 0 },
                runs: [
                  { text: 'A1' },
                  { break: true, properties: {}, text: '' },
                  { text: 'a1 detail' },
                ],
              },
              {
                properties: { level: 1 },
                runs: [{ text: 'B1' }],
              },
            ],
          },
        },
      ],
      slideNumber: 1,
    };

    const slideXml = serializeSlide(slide);
    expect(slideXml).toContain('<?xml version="1.0"');
    expect(slideXml).toContain('<p:sld');
    expect(slideXml).toContain('Slide Title');
    expect(slideXml).toContain('A1');
    expect(slideXml).toContain('<a:br/>');
    expect(slideXml).toContain('a1 detail');
    expect(slideXml).toContain('B1');
  });
});
