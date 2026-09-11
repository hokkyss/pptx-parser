import { describe, expect, it } from 'vitest';
import { emu, emuDegree, type PptxParagraph, type PptxShapeElement, type PptxSlide, type PptxTextBody } from '@hokkyss/pptx-core';
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
        { text: 'A1', properties: {} },
        { break: true, properties: {}, text: '' },
        { text: 'a1 level, no bullet', properties: {} },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    expect(typeof serialized).toBe('object');
    expect(serialized['a:r']).toBeDefined();
    expect(serialized['a:br']).toBeDefined();
  });

  it('serializes <a:br> carrying optional run properties in <a:rPr>', () => {
    const paragraph: PptxParagraph = {
      properties: {},
      runs: [
        { text: 'Before', properties: {} },
        { break: true, properties: { bold: true, italic: true }, text: '' },
        { text: 'After', properties: {} },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    expect(typeof serialized).toBe('object');
    const brList = serialized['a:br'] as Record<string, unknown>[];
    expect(brList[0]?.['a:rPr']).toMatchObject({ '@_b': '1', '@_i': '1' });
  });

  it('preserves text content when serializing paragraphs with line breaks', () => {
    const paragraph: PptxParagraph = {
      properties: { level: 1 },
      runs: [
        { text: 'Fish & Chips <Salt>', properties: {} },
        { break: true, properties: {}, text: '' },
        { text: 'Line 2 & More', properties: {} },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    expect(typeof serialized).toBe('object');
    const textRuns = serialized['a:r'] as Record<string, unknown>[];
    expect(textRuns[0]?.['a:t']).toBe('Fish & Chips <Salt>');
    expect(textRuns[1]?.['a:t']).toBe('Line 2 & More');
  });

  it('does NOT emit marL or indent when only level is specified (master inheritance)', () => {
    const paragraph: PptxParagraph = {
      properties: {
        level: 2,
      },
      runs: [
        { text: 'Deport Topic Lvl 2', properties: {} },
      ],
    };

    const serialized = serializeParagraph(paragraph);
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
        { text: 'Explicit Bullet Item', properties: {} },
      ],
    };

    const serialized = serializeParagraph(paragraph);
    const pPr = serialized['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_lvl']).toBe(1);
    expect(pPr['@_marL']).toBeUndefined(); // Inherited from slide master
    expect(pPr['@_indent']).toBeUndefined();
    expect(pPr['a:buChar']).toBeDefined();
  });

  it('serializes text body containing line breaks and plain paragraphs together', () => {
    const textBody: PptxTextBody = {
      bodyProperties: { verticalAlignment: 'top' },
      paragraphs: [
        {
          properties: { level: 0 },
          runs: [
            { text: 'A1', properties: {} },
            { break: true, properties: {}, text: '' },
            { text: 'a1 continuation', properties: {} },
          ],
        },
        {
          properties: { level: 1 },
          runs: [{ text: 'B1', properties: {} }],
        },
      ],
    };

    const serialized = serializeTextBody(textBody);
    expect(typeof serialized).toBe('object');
    expect(serialized['a:p']).toBeDefined();
    const paragraphs = serialized['a:p'] as Record<string, unknown>[];
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.['a:br']).toBeDefined();
  });

  it('serializes complete shape containing line breaks into valid shape object', () => {
    const shape: PptxShapeElement = {
      type: 'shape',
      elementType: 'shape',
      isVisible: true,
      zIndex: 0,
      position: {
        x: emu(0),
        y: emu(0),
        cx: emu(10),
        cy: emu(10),
      },
      rotation: emuDegree(0),
      id: '3',
      name: 'Content Placeholder 2',
      placeholder: { idx: 1, type: 'body' },
      textBody: {
        bodyProperties: {},
        paragraphs: [
          {
            properties: { level: 0 },
            runs: [
              { text: 'Root', properties: {} },
              { break: true, properties: {}, text: '' },
              { text: 'Continuation without bullet', properties: {} },
            ],
          },
          {
            properties: { level: 1 },
            runs: [{ text: 'Sub level 1', properties: {} }],
          },
        ],
      },
    };

    const serialized = serializeShape(shape);
    expect(typeof serialized).toBe('object');
    expect(serialized['p:nvSpPr']).toBeDefined();
    expect(serialized['p:txBody']).toBeDefined();
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

    const slideXml = serializeSlide(slide as PptxSlide);
    expect(slideXml).toContain('<?xml version="1.0"');
    expect(slideXml).toContain('<p:sld');
    expect(slideXml).toContain('Slide Title');
    expect(slideXml).toContain('A1');
    expect(slideXml).toContain('<a:br/>');
    expect(slideXml).toContain('a1 detail');
    expect(slideXml).toContain('B1');
  });
});
