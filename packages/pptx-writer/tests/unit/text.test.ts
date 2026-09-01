import { describe, expect, it } from 'vitest';
import type { PptxTextBody } from '@hokkyss/pptx-core';
import { emu, hundredthsPoint, thousandthsPercent } from '@hokkyss/pptx-core';
import {
  serializeBodyProperties,
  serializeBulletProperties,
  serializeColorNode,
  serializeFill,
  serializeHyperlink,
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
      bodyProperties: {},
      paragraphs: [
        {
          properties: {},
          runs: [
            {
              properties: {},
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
      gradient: {
        flip: 'xy',
        rotateWithShape: false,
        stops: [
          { color: { type: 'srgb', value: '000000' }, position: thousandthsPercent(0) },
          { color: { type: 'srgb', value: 'FFFFFF' }, position: thousandthsPercent(100000) },
        ],
        type: 'radial',
      },
      type: 'gradient',
    });
    expect(radialFill?.['a:gradFill']).toBeDefined();
    const gradFillNode = radialFill?.['a:gradFill'] as Record<string, string> | undefined;
    expect(gradFillNode?.['@_flip']).toBe('xy');

    const pathFill = serializeFill({
      gradient: {
        pathBounds: { bottom: 0.8, left: 0.2, right: 0.8, top: 0.2 },
        stops: [{ color: { type: 'srgb', value: 'FF0000' }, position: thousandthsPercent(0) }],
        type: 'path',
      },
      type: 'gradient',
    });
    const pathGradNode = pathFill?.['a:gradFill'] as Record<string, Record<string, string>> | undefined;
    expect(pathGradNode?.['a:path']?.['@_path']).toBe('rect');
  });

  it('serializes subscript, superscript, strikethrough, and baseline', () => {
    const rPr = serializeRunProperties({
      baseline: -25000,
      strikethrough: true,
      subscript: true,
    });
    expect(rPr['@_strike']).toBe('sngStrike');
    expect(rPr['@_baseline']).toBe(-25000);
  });
});

describe('Paragraph serializer bullets and empty runs', () => {
  it('serializes bullet nodes without forcing marL and indent overrides', () => {
    const autoNumPara = serializeParagraph({
      properties: {
        bullet: { autoNumType: 'arabicPeriod', type: 'autoNum' },
        level: 2,
      },
      runs: [{ properties: {}, text: 'Item' }],
    });
    expect(autoNumPara['a:pPr']).toBeDefined();
    const pPr = autoNumPara['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_lvl']).toBe(2);
    expect(pPr['a:buAutoNum']).toBeDefined();
    expect(pPr['@_marL']).toBeUndefined();
    expect(pPr['@_indent']).toBeUndefined();

    const emptyPara = serializeParagraph({ properties: {}, runs: [] });
    expect(emptyPara['a:endParaRPr']).toEqual({});
  });
});

describe('Paragraph serializer margin and indent legacy fallbacks', () => {
  it('serializes margin and indent from legacy paragraph properties', () => {
    // @ts-expect-error Testing legacy margin/indent property fallbacks
    const p = serializeParagraph({ indent: -50000, margin: 150000, runs: [{ properties: {}, text: 'Legacy' }] });
    const pPr = p['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_marL']).toBe(150000);
    expect(pPr['@_indent']).toBe(-50000);
  });
});

describe('Text Serializer color objects and paragraph insets', () => {
  it('serializes scheme color objects on text runs and explicit margins/indents on paragraphs', () => {
    const rPr = serializeRunProperties({
      color: 'accent1',
      superscript: true,
    });
    expect(rPr['@_baseline']).toBe('30000');
    expect(rPr['a:solidFill']).toBeDefined();

    const p = serializeParagraph({
      properties: {
        firstLineIndent: emu(-150000),
        leftMargin: emu(300000),
      },
      runs: [{ properties: {}, text: 'Indented' }],
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
    const charBullet = serializeBulletProperties({ char: '•', type: 'char' });
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
      gradient: { angle: 5400000, stops: [] },
      type: 'gradient',
    });
    const lin = (largeAngleFill?.['a:gradFill'] as Record<string, Record<string, unknown>>)?.['a:lin'];
    expect(lin?.['@_ang']).toBe(5400000);
  });
});

describe('Text Serializer color node alpha and string fallbacks', () => {
  it('serializes color object with alpha and unrecognized color string fallbacks', () => {
    const clrWithAlpha = serializeColorNode({ alpha: thousandthsPercent(50000), type: 'srgb', value: 'FF0000' });
    expect(clrWithAlpha['a:srgbClr']).toBeDefined();
    expect((clrWithAlpha['a:srgbClr'] as Record<string, Record<string, unknown>>)['a:alpha']?.['@_val']).toBe(50000);

    const fallbackClr = serializeColorNode('custom-named-color');
    expect(fallbackClr['a:srgbClr']).toBeDefined();
  });

  it('covers rich text runs, line spacing in points and percentage', () => {
    const tBodyObj = serializeTextBody({
      bodyProperties: { verticalAlignment: 'middle' },
      paragraphs: [
        {
          properties: {
            alignment: 'right',
            lineSpacing: hundredthsPoint(2000),
          },
          runs: [
            {
              properties: {
                bold: true,
                fontSize: hundredthsPoint(1600),
                italic: true,
                strikethrough: 'sngStrike',
                subscript: true,
                superscript: true,
                underline: 'sng',
              },
              text: 'Styled Run',
            },
          ],
        },
        {
          properties: {
            lineSpacing: hundredthsPoint(1200),
          },
          runs: [{ properties: {}, text: 'P% line spacing' }],
        },
      ],
    });
    expect(tBodyObj).toBeDefined();
  });

  it('covers gradient stop auto-interpolation, radial fillToRect fallbacks, and opacity overrides', () => {
    // Stops without explicit position (auto-interpolation)
    const autoStops = serializeFill({
      gradient: {
        rotateWithShape: true,
        stops: [
          // @ts-expect-error Testing missing position auto-interpolation
          { color: '#FF0000' },
          // @ts-expect-error Testing missing position auto-interpolation
          { color: '#00FF00' },
          // @ts-expect-error Testing missing position auto-interpolation
          { color: '#0000FF' },
        ],
        type: 'linear',
      },
      type: 'gradient',
    });
    expect(autoStops?.['a:gradFill']).toBeDefined();
    const gs = (autoStops?.['a:gradFill'] as Record<string, Record<string, Record<string, number>[]>>)?.['a:gsLst']?.['a:gs'];
    expect(gs[0]['@_pos']).toBe(0);
    expect(gs[1]['@_pos']).toBe(50000);
    expect(gs[2]['@_pos']).toBe(100000);

    // Single stop without position
    const singleStop = serializeFill({
      gradient: {
        rotateWithShape: false,
        // @ts-expect-error Testing missing position on single stop
        stops: [{ color: '#FF0000' }],
        type: 'linear',
      },
      type: 'gradient',
    });
    const singleGs = (singleStop?.['a:gradFill'] as Record<string, Record<string, Record<string, number>[]>>)?.['a:gsLst']?.['a:gs'];
    expect(singleGs[0]['@_pos']).toBe(0);

    // Radial gradient with default fillToRect (no pathBounds)
    const radialDefault = serializeFill({
      gradient: {
        stops: [{ color: '#111111', position: thousandthsPercent(0) }],
        type: 'radial',
      },
      type: 'gradient',
    });
    const fillToRect = (radialDefault?.['a:gradFill'] as Record<string, Record<string, Record<string, number>>>)?.['a:path']?.['a:fillToRect'];
    expect(fillToRect?.['@_l']).toBe(50000);

    // Radial gradient with pathBounds <= 1 scaling
    const radialBounded = serializeFill({
      gradient: {
        pathBounds: { bottom: 0.9, left: 0.1, right: 0.9, top: 0.1 },
        stops: [{ color: '#222222', position: thousandthsPercent(0) }],
        type: 'radial',
      },
      type: 'gradient',
    });
    const scaledBounds = (radialBounded?.['a:gradFill'] as Record<string, Record<string, Record<string, number>>>)?.['a:path']?.['a:fillToRect'];
    expect(scaledBounds?.['@_l']).toBe(10000);

    // Solid fill with opacity override
    const colorWithOpacity = serializeColorNode('#ABCDEF', 0.75);
    expect((colorWithOpacity['a:srgbClr'] as Record<string, Record<string, number>>)['a:alpha']?.['@_val']).toBe(75000);

    const colorWithAlphaVal = serializeColorNode('#ABCDEF', 80000);
    expect((colorWithAlphaVal['a:srgbClr'] as Record<string, Record<string, number>>)['a:alpha']?.['@_val']).toBe(80000);
  });

  it('covers run color object, underline/strikethrough styles, and hyperlink actions', () => {
    // Run with color object
    // @ts-expect-error Testing object color on text run
    const rPrWithColorObj = serializeRunProperties({ color: { type: 'srgb', value: '10B981' } });
    expect(rPrWithColorObj['a:solidFill']).toBeDefined();

    // Run with explicit underline and strikethrough styles
    const styledRPr = serializeRunProperties({
      bold: true,
      italic: true,
      strikethrough: 'dblStrike',
      underline: 'dbl',
    });
    expect(styledRPr['@_u']).toBe('dbl');
    expect(styledRPr['@_strike']).toBe('dblStrike');

    // Run with string hyperlink with and without override
    expect(serializeHyperlink('https://example.com', 'rId99')).toEqual({ '@_r:id': 'rId99' });
    expect(serializeHyperlink('https://example.com')).toBeUndefined();
    expect(serializeHyperlink(undefined)).toBeUndefined();

    // Hyperlink with standard actions
    expect(serializeHyperlink({ action: 'firstSlide' })).toEqual({ '@_action': 'ppaction://hlinkshowjump?jump=firstslide' });
    expect(serializeHyperlink({ action: 'nextSlide' })).toEqual({ '@_action': 'ppaction://hlinkshowjump?jump=nextslide' });
    expect(serializeHyperlink({ action: 'endShow' })).toEqual({ '@_action': 'ppaction://hlinkshowjump?jump=endshow' });
    expect(serializeHyperlink({ action: 'lastSlide' })).toEqual({ '@_action': 'ppaction://hlinkshowjump?jump=lastslide' });
    expect(serializeHyperlink({ action: 'previousSlide' })).toEqual({ '@_action': 'ppaction://hlinkshowjump?jump=previousslide' });
    expect(serializeHyperlink({ slideIndex: 3 })).toEqual({ '@_action': 'ppaction://hlinksldjump' });

    // Hyperlink with rId and tooltip
    expect(serializeHyperlink({ rId: 'rId5', tooltip: 'My Tooltip' })).toEqual({
      '@_r:id': 'rId5',
      '@_tooltip': 'My Tooltip',
    });

    // Run properties with hyperlink
    const rPrWithHlink = serializeRunProperties({ hyperlink: { action: 'nextSlide', tooltip: 'Next' } });
    expect(rPrWithHlink['a:hlinkClick']).toBeDefined();

    // Empty run properties
    expect(serializeRunProperties(undefined)).toEqual({});
  });

  it('covers bullet startAt, text body padding insets, and empty text body', () => {
    // Bullet autoNum with startAt
    const numberedBullet = serializeBulletProperties({ autoNumType: 'romanUcPeriod', startAt: 5, type: 'autoNum' });
    expect(numberedBullet?.['a:buAutoNum']).toEqual({ '@_startAt': 5, '@_type': 'romanUcPeriod' });
    // @ts-expect-error Testing unsupported bullet type fallback
    expect(serializeBulletProperties({ type: 'other' })).toBeUndefined();

    // Body properties with all insets
    const bodyPr = serializeBodyProperties({
      bottomInset: emu(20000),
      leftInset: emu(10000),
      rightInset: emu(30000),
      topInset: emu(40000),
      wrap: 'square',
    });
    expect(bodyPr['@_lIns']).toBe(10000);
    expect(bodyPr['@_tIns']).toBe(40000);
    expect(bodyPr['@_rIns']).toBe(30000);
    expect(bodyPr['@_bIns']).toBe(20000);

    // Empty text body fallback
    const emptyBody = serializeTextBody({ bodyProperties: {}, paragraphs: [] });
    expect(emptyBody['a:p']).toHaveLength(1);

    // Paragraph with justify alignment and single char bullet margin
    const justifyPara = serializeParagraph({
      properties: {
        alignment: 'justify',
        bullet: { char: '-', type: 'char' },
      },
      runs: [{ properties: {}, text: 'Justified item' }],
    });
    const pPr = justifyPara['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_algn']).toBe('justify');
    expect(pPr['a:buChar']).toBeDefined();
    expect(pPr['@_marL']).toBeUndefined();
    expect(pPr['@_indent']).toBeUndefined();

    // Scheme color string
    const schemeNode = serializeColorNode('accent2');
    expect(schemeNode['a:schemeClr']).toBeDefined();

    // Degree angle <= 360 gradient
    const angleGrad = serializeFill({
      gradient: { angle: 90, stops: [] },
      type: 'gradient',
    });
    const linNode = (angleGrad?.['a:gradFill'] as Record<string, Record<string, number>>)?.['a:lin'];
    expect(linNode?.['@_ang']).toBe(5400000);
  });

  it('serializes non-bulleted hierarchical indentation ({ bullet: { type: "none" }, level: 2 }) in fast-path object mode', () => {
    const para = serializeParagraph({
      properties: {
        bullet: { type: 'none' },
        level: 2,
      },
      runs: [{ properties: {}, text: 'Level 2 code outline without bullet' }],
    });
    expect(typeof para).toBe('object');
    const pPr = (para as Record<string, unknown>)['a:pPr'] as Record<string, unknown>;
    expect(pPr['@_lvl']).toBe(2);
    expect(pPr['a:buNone']).toBeDefined();
  });

  it('serializes non-bulleted hierarchical indentation ({ bullet: { type: "none" }, level: 1 }) in raw-XML mode with line breaks', () => {
    const xml = serializeParagraph({
      properties: {
        bullet: { type: 'none' },
        level: 1,
      },
      runs: [
        { text: 'First line of block explanation' },
        { break: true },
        { text: 'Second line continuing block' },
      ],
    });
    expect(typeof xml).toBe('string');
    expect(xml).toContain('<a:pPr lvl="1"><a:buNone/></a:pPr>');
    expect(xml).toContain('<a:r><a:t>First line of block explanation</a:t></a:r>');
    expect(xml).toContain('<a:br/>');
    expect(xml).toContain('<a:r><a:t>Second line continuing block</a:t></a:r>');
  });

  it('omits @_algn="l" and @_lvl="0" on level 0 paragraphs for clean PowerPoint master inheritance', () => {
    const para = serializeParagraph({
      properties: {
        alignment: 'left',
        level: 0,
      },
      runs: [{ properties: {}, text: 'Clean level 0 run' }],
    });
    expect(typeof para).toBe('object');
    // If no overrides other than default left and level 0, a:pPr should be omitted entirely
    expect((para as Record<string, unknown>)['a:pPr']).toBeUndefined();
  });
});


