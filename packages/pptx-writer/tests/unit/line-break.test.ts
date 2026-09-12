import { describe, expect, it } from 'vitest';
import { emu, emuDegree, type PptxParagraph, type PptxShapeElement, type PptxSlide, type PptxTextBody } from '@hokkyss/pptx-core';
import {
  serializeParagraph,
  serializeTextBody,
} from '../../lib/serializers/text-serializer';
import { serializeShape } from '../../lib/serializers/shape-serializer';
import { serializeSlide } from '../../lib/serializers/slide-serializer';
import { renderXml } from '../../lib/xml/xml-element';

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
    expect(serialized.tag).toBe('a:p');
    const xml = renderXml(serialized);
    expect(xml).toContain('<a:r>');
    expect(xml).toContain('<a:br');
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
    const xml = renderXml(serialized);
    expect(xml).toContain('<a:br');
    expect(xml).toContain('b="1"');
    expect(xml).toContain('i="1"');
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
    const xml = renderXml(serialized);
    expect(xml).toContain('Fish &amp; Chips &lt;Salt&gt;');
    expect(xml).toContain('Line 2 &amp; More');
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
    const xml = renderXml(serialized);
    expect(xml).toContain('lvl="2"');
    expect(xml).not.toContain('marL=');
    expect(xml).not.toContain('indent=');
    expect(xml).not.toContain('<a:buChar');
    expect(xml).not.toContain('<a:buNone');
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
    const xml = renderXml(serialized);
    expect(xml).toContain('lvl="1"');
    expect(xml).toContain('marL=');
    expect(xml).toContain('indent=');
    expect(xml).toContain('<a:buChar');
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
    expect(serialized.tag).toBe('p:txBody');
    const xml = renderXml(serialized);
    expect(xml).toContain('A1');
    expect(xml).toContain('a1 continuation');
    expect(xml).toContain('B1');
    expect(xml).toContain('<a:br');
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
    expect(serialized.tag).toBe('p:sp');
    const xml = renderXml(serialized);
    expect(xml).toContain('<p:nvSpPr>');
    expect(xml).toContain('<p:txBody>');
    expect(xml).toContain('<a:br');
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
    expect(slideXml).toContain('<a:br');
    expect(slideXml).toContain('a1 detail');
    expect(slideXml).toContain('B1');
  });
});
