import { describe, expect, it } from 'vitest';
import type { PptxTextBody } from '@hokkyss/pptx-core';
import { emu, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeTextBody } from '../../lib/serializers/text-serializer';

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
});
