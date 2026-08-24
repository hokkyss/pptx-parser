import { describe, expect, it } from 'vitest';
import type { PptxTextBody } from '@hokkyss/pptx-core';
import { emu, hundredthsPoint } from '@hokkyss/pptx-core';
import {
  serializeBodyProperties,
  serializeBulletProperties,
  serializeColorNode,
  serializeFill,
  serializeParagraph,
  serializeRunProperties,
  serializeTextBody,
} from '../../lib/serializers/text-serializer';

describe('Text Body Serializer', () => {
  it('serializes text body with paragraph alignment, line spacing, runs, formatting, and colors', () => {
    const textBody: PptxTextBody = {
      bodyProperties: {
        verticalAlignment: 'middle',
        wrap: 'square',
        leftInset: emu(91440),
        rightInset: emu(91440),
        topInset: emu(45720),
        bottomInset: emu(45720),
      },
      paragraphs: [
        {
          properties: {
            alignment: 'center',
            bullet: { type: 'none' },
          },
          runs: [
            {
              text: 'Hello ',
              properties: {
                bold: true,
                fontFamily: 'Arial',
                fontSize: hundredthsPoint(2400),
                color: 'FF0000',
              },
            },
            {
              text: 'World!',
              properties: {
                italic: true,
                underline: true,
                fontFamily: 'Arial',
                fontSize: hundredthsPoint(2400),
                color: '0000FF',
              },
            },
          ],
        },
      ],
    };

    const xmlObject = serializeTextBody(textBody);
    expect(xmlObject).toBeDefined();

    const bodyPr = xmlObject['a:bodyPr'] as Record<string, unknown>;
    expect(bodyPr['@_anchor']).toBe('ctr');
    expect(bodyPr['@_wrap']).toBe('square');

    const paragraphs = xmlObject['a:p'] as Record<string, unknown>[];
    const p = paragraphs[0];
    const pPr = p['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_algn']).toBe('ctr');
    expect(pPr['a:buNone']).toBeDefined();

    const runs = p['a:r'] as Record<string, Record<string, unknown>>[];
    expect(runs[0]['a:t']).toBe('Hello ');
    expect(runs[0]['a:rPr']['@_b']).toBe('1');
    expect(runs[0]['a:rPr']['@_sz']).toBe(2400);
    expect((runs[0]['a:rPr']['a:solidFill'] as Record<string, Record<string, unknown>>)['a:srgbClr']['@_val']).toBe('FF0000');

    expect(runs[1]['a:t']).toBe('World!');
    expect(runs[1]['a:rPr']['@_i']).toBe('1');
    expect(runs[1]['a:rPr']['@_u']).toBe('sng');
    expect((runs[1]['a:rPr']['a:solidFill'] as Record<string, Record<string, unknown>>)['a:srgbClr']['@_val']).toBe('0000FF');
  });

  it('sanitizes invalid XML 1.0 control characters in text runs preventing PowerPoint corruption', () => {
    const textBody: PptxTextBody = {
      paragraphs: [
        {
          runs: [
            {
              text: 'Clean\x00Text\x08With\x0BControl\x0CChars\x1F!',
            },
          ],
        },
      ],
    };

    const xmlObject = serializeTextBody(textBody);
    const p = (xmlObject['a:p'] as Record<string, unknown>[])[0];
    const runs = p['a:r'] as Record<string, Record<string, unknown>>[];
    expect(runs[0]['a:t']).toBe('CleanTextWithControlChars!');
  });
});

describe('Text & Fill Serializer extended coverage', () => {
  it('serializes radial and path gradient fills', () => {
    const radialFill = serializeFill({
      type: 'gradient',
      gradient: {
        type: 'radial',
        flip: 'xy',
        rotateWithShape: false,
        stops: [{ color: { type: 'srgb', value: '000000' } }, { color: { type: 'srgb', value: 'FFFFFF' } }],
      },
    });
    expect(radialFill?.['a:gradFill']).toBeDefined();
    const gradFillNode = radialFill?.['a:gradFill'] as Record<string, string> | undefined;
    expect(gradFillNode?.['@_flip']).toBe('xy');

    const pathFill = serializeFill({
      type: 'gradient',
      gradient: {
        type: 'path',
        pathBounds: { left: 0.2, top: 0.2, right: 0.8, bottom: 0.8 },
        stops: [{ color: { type: 'srgb', value: 'FF0000' }, position: 0 }],
      },
    });
    const pathGradNode = pathFill?.['a:gradFill'] as Record<string, Record<string, string>> | undefined;
    expect(pathGradNode?.['a:path']?.['@_path']).toBe('rect');
  });

  it('serializes subscript, superscript, strikethrough, and baseline', () => {
    const rPr = serializeRunProperties({
      subscript: true,
      strikethrough: true,
      baseline: -25000,
    });
    expect(rPr['@_strike']).toBe('sngStrike');
    expect(rPr['@_baseline']).toBe(-25000);
  });
});

describe('Paragraph serializer bullets and empty runs', () => {
  it('computes marL and indent for numbered and char bullets', () => {
    const autoNumPara = serializeParagraph({
      properties: {
        bullet: { type: 'autoNum', autoNumType: 'arabicPeriod' },
        level: 2,
      },
      runs: [{ text: 'Item' }],
    });
    expect(autoNumPara['a:pPr']).toBeDefined();
    expect(autoNumPara['a:pPr']?.['@_marL']).toBe((2 * 228600) + 203200);

    const emptyPara = serializeParagraph({ runs: [] });
    expect(emptyPara['a:endParaRPr']).toEqual({});
  });
});

describe('Paragraph serializer margin and indent legacy fallbacks', () => {
  it('serializes margin and indent from legacy paragraph properties', () => {
    // @ts-expect-error Testing legacy margin/indent property fallbacks
    const p = serializeParagraph({ margin: 150000, indent: -50000, runs: [{ text: 'Legacy' }] });
    const pPr = p['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_marL']).toBe(150000);
    expect(pPr['@_indent']).toBe(-50000);
  });
});

describe('Text Serializer color objects and paragraph insets', () => {
  it('serializes scheme color objects on text runs and explicit margins/indents on paragraphs', () => {
    const rPr = serializeRunProperties({
      color: { type: 'scheme', value: 'accent1' },
      superscript: true,
    });
    expect(rPr['@_baseline']).toBe('30000');
    expect(rPr['a:solidFill']).toBeDefined();

    const p = serializeParagraph({
      properties: {
        leftMargin: emu(300000),
        firstLineIndent: emu(-150000),
      },
      runs: [{ text: 'Indented' }],
    });
    const pPr = p['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_marL']).toBe(300000);
    expect(pPr['@_indent']).toBe(-150000);
  });
});

describe('Text Serializer subscript baseline', () => {
  it('sets baseline to -25000 when subscript is true without explicit baseline', () => {
    const rPr = serializeRunProperties({ subscript: true });
    expect(rPr['@_baseline']).toBe('-25000');
  });
});

describe('Text Serializer char bullet and fill fallbacks', () => {
  it('serializes char bullet with explicit character and handles empty bodyProperties/fill fallbacks', () => {
    const charBullet = serializeBulletProperties({ type: 'char', char: '•' });
    expect(charBullet?.['a:buChar']).toEqual({ '@_char': '•' });

    expect(serializeBodyProperties(undefined)).toEqual({});
    // @ts-expect-error Testing unsupported fill type
    expect(serializeFill({ type: 'unsupported' })).toBeUndefined();
    expect(serializeFill(undefined)).toBeUndefined();
  });
});

describe('Text Serializer noFill and large angle gradients', () => {
  it('serializes noFill and raw angle values in gradient fills', () => {
    expect(serializeFill({ type: 'none' })).toEqual({ 'a:noFill': {} });

    const largeAngleFill = serializeFill({
      type: 'gradient',
      gradient: { angle: 5400000, stops: [] },
    });
    const lin = (largeAngleFill?.['a:gradFill'] as Record<string, Record<string, unknown>>)?.['a:lin'];
    expect(lin?.['@_ang']).toBe(5400000);
  });
});

describe('Text Serializer color node alpha and string fallbacks', () => {
  it('serializes color object with alpha and unrecognized color string fallbacks', () => {
    const clrWithAlpha = serializeColorNode({ type: 'srgb', value: 'FF0000', alpha: 50000 });
    expect(clrWithAlpha['a:srgbClr']).toBeDefined();
    expect((clrWithAlpha['a:srgbClr'] as Record<string, Record<string, unknown>>)['a:alpha']?.['@_val']).toBe(50000);

    const fallbackClr = serializeColorNode('custom-named-color');
    expect(fallbackClr['a:srgbClr']).toBeDefined();
  });
});
