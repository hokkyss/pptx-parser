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
import { renderXml } from '../../lib/xml/xml-element';

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

    const xmlNode = serializeTextBody(textBody);
    expect(xmlNode).toBeDefined();
    const xml = renderXml(xmlNode);

    // bodyPr
    expect(xml).toContain('anchor="ctr"');
    expect(xml).toContain('wrap="square"');

    // paragraph alignment
    expect(xml).toContain('algn="ctr"');
    // buNone for bullet: none
    expect(xml).toContain('<a:buNone/>');

    // runs
    expect(xml).toContain('>Hello <');
    expect(xml).toContain('b="1"');
    expect(xml).toContain('sz="2400"');
    expect(xml).toContain('val="FF0000"');

    expect(xml).toContain('>World!</');
    expect(xml).toContain('i="1"');
    expect(xml).toContain('u="sng"');
    expect(xml).toContain('val="0000FF"');
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

    const xml = renderXml(serializeTextBody(textBody));
    expect(xml).toContain('CleanTextWithControlChars!');
    expect(xml).not.toContain('\x00');
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
    expect(radialFill).toBeDefined();
    const radialXml = renderXml(radialFill);
    expect(radialXml).toContain('flip="xy"');
    expect(radialXml).toContain('<a:gradFill');

    const pathFill = serializeFill({
      gradient: {
        pathBounds: { bottom: 0.8, left: 0.2, right: 0.8, top: 0.2 },
        stops: [{ color: { type: 'srgb', value: 'FF0000' }, position: thousandthsPercent(0) }],
        type: 'path',
      },
      type: 'gradient',
    });
    const pathXml = renderXml(pathFill);
    expect(pathXml).toContain('path="rect"');
  });

  it('serializes subscript, superscript, strikethrough, and baseline', () => {
    const rPr = serializeRunProperties({
      baseline: -25000,
      strikethrough: true,
      subscript: true,
    });
    expect(rPr).toBeDefined();
    const xml = renderXml(rPr);
    expect(xml).toContain('strike="sngStrike"');
    expect(xml).toContain('baseline="-25000"');
  });
});

describe('Paragraph serializer bullets and empty runs', () => {
  it('computes marL and indent for numbered and char bullets', () => {
    const autoNumPara = serializeParagraph({
      properties: {
        bullet: { autoNumType: 'arabicPeriod', type: 'autoNum' },
        level: 2,
      },
      runs: [{ properties: {}, text: 'Item' }],
    });
    const xml = renderXml(autoNumPara);
    const expectedMarL = (2 * 228600) + 203200;
    expect(xml).toContain(`marL="${expectedMarL}"`);

    const emptyPara = serializeParagraph({ properties: {}, runs: [] });
    const emptyXml = renderXml(emptyPara);
    expect(emptyXml).toContain('<a:endParaRPr/>');
  });
});

describe('Paragraph serializer margin and indent legacy fallbacks', () => {
  it('serializes margin and indent from legacy paragraph properties', () => {
    // @ts-expect-error Testing legacy margin/indent property fallbacks
    const p = serializeParagraph({ indent: -50000, margin: 150000, runs: [{ properties: {}, text: 'Legacy' }] });
    const xml = renderXml(p);
    expect(xml).toContain('marL="150000"');
    expect(xml).toContain('indent="-50000"');
  });
});

describe('Text Serializer color objects and paragraph insets', () => {
  it('serializes scheme color objects on text runs and explicit margins/indents on paragraphs', () => {
    const rPr = serializeRunProperties({
      color: 'accent1',
      superscript: true,
    });
    expect(rPr).toBeDefined();
    const rPrXml = renderXml(rPr);
    expect(rPrXml).toContain('baseline="30000"');
    expect(rPrXml).toContain('<a:solidFill>');

    const p = serializeParagraph({
      properties: {
        firstLineIndent: emu(-150000),
        leftMargin: emu(300000),
      },
      runs: [{ properties: {}, text: 'Indented' }],
    });
    const pXml = renderXml(p);
    expect(pXml).toContain('marL="300000"');
    expect(pXml).toContain('indent="-150000"');
  });
});

describe('Text Serializer subscript baseline', () => {
  it('sets baseline to -25000 when subscript is true without explicit baseline', () => {
    const rPr = serializeRunProperties({ subscript: true });
    expect(rPr).toBeDefined();
    const xml = renderXml(rPr);
    expect(xml).toContain('baseline="-25000"');
  });
});

describe('Text Serializer char bullet and fill fallbacks', () => {
  it('serializes char bullet with explicit character and handles empty bodyProperties/fill fallbacks', () => {
    const charBullet = serializeBulletProperties({ char: '•', type: 'char' });
    expect(charBullet).toBeDefined();
    const xml = renderXml(charBullet);
    expect(xml).toContain('char="•"');
    expect(xml).toContain('<a:buChar');

    // serializeBodyProperties(undefined) returns an empty <a:bodyPr/> element
    const emptyBodyPr = serializeBodyProperties(undefined);
    const emptyXml = renderXml(emptyBodyPr);
    expect(emptyXml).toBe('<a:bodyPr/>');

    // @ts-expect-error Testing unsupported fill type
    expect(serializeFill({ type: 'unsupported' })).toBeUndefined();
    expect(serializeFill(undefined)).toBeUndefined();
  });
});

describe('Text Serializer noFill and large angle gradients', () => {
  it('serializes noFill and raw angle values in gradient fills', () => {
    const noFill = serializeFill({ type: 'none' });
    expect(noFill).toBeDefined();
    expect(renderXml(noFill)).toBe('<a:noFill/>');

    const largeAngleFill = serializeFill({
      gradient: { angle: 5400000, stops: [] },
      type: 'gradient',
    });
    const linXml = renderXml(largeAngleFill);
    expect(linXml).toContain('ang="5400000"');
  });
});

describe('Text Serializer color node alpha and string fallbacks', () => {
  it('serializes color object with alpha and unrecognized color string fallbacks', () => {
    const clrWithAlpha = serializeColorNode({ alpha: thousandthsPercent(50000), type: 'srgb', value: 'FF0000' });
    const xml = renderXml(clrWithAlpha);
    expect(xml).toContain('<a:srgbClr');
    expect(xml).toContain('val="50000"');

    const fallbackClr = serializeColorNode('custom-named-color');
    expect(renderXml(fallbackClr)).toContain('<a:srgbClr');
  });

  it('covers rich text runs, line spacing in points and percentage', () => {
    const tBodyNode = serializeTextBody({
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
    expect(tBodyNode).toBeDefined();
    expect(renderXml(tBodyNode)).toContain('Styled Run');
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
    expect(autoStops).toBeDefined();
    const autoXml = renderXml(autoStops);
    expect(autoXml).toContain('<a:gradFill');
    // The 3 stops should have pos="0", pos="50000", pos="100000"
    expect(autoXml).toContain('pos="0"');
    expect(autoXml).toContain('pos="50000"');
    expect(autoXml).toContain('pos="100000"');

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
    const singleXml = renderXml(singleStop);
    expect(singleXml).toContain('pos="0"');

    // Radial gradient with default fillToRect (no pathBounds)
    const radialDefault = serializeFill({
      gradient: {
        stops: [{ color: '#111111', position: thousandthsPercent(0) }],
        type: 'radial',
      },
      type: 'gradient',
    });
    const radialXml = renderXml(radialDefault);
    expect(radialXml).toContain('l="50000"');

    // Radial gradient with pathBounds <= 1 scaling
    const radialBounded = serializeFill({
      gradient: {
        pathBounds: { bottom: 0.9, left: 0.1, right: 0.9, top: 0.1 },
        stops: [{ color: '#222222', position: thousandthsPercent(0) }],
        type: 'radial',
      },
      type: 'gradient',
    });
    const boundedXml = renderXml(radialBounded);
    expect(boundedXml).toContain('l="10000"');

    // Solid fill with opacity override (< 1 → percentage)
    const colorWithOpacity = serializeColorNode('#ABCDEF', 0.75);
    const opacXml = renderXml(colorWithOpacity);
    expect(opacXml).toContain('val="75000"');

    const colorWithAlphaVal = serializeColorNode('#ABCDEF', 80000);
    const alphaXml = renderXml(colorWithAlphaVal);
    expect(alphaXml).toContain('val="80000"');
  });

  it('covers run color object, underline/strikethrough styles, and hyperlink actions', () => {
    // Run with color object
    // @ts-expect-error Testing object color on text run
    const rPrWithColorObj = serializeRunProperties({ color: { type: 'srgb', value: '10B981' } });
    expect(rPrWithColorObj).toBeDefined();
    const colorXml = renderXml(rPrWithColorObj);
    expect(colorXml).toContain('<a:solidFill>');

    // Run with explicit underline and strikethrough styles
    const styledRPr = serializeRunProperties({
      bold: true,
      italic: true,
      strikethrough: 'dblStrike',
      underline: 'dbl',
    });
    const styledXml = renderXml(styledRPr);
    expect(styledXml).toContain('u="dbl"');
    expect(styledXml).toContain('strike="dblStrike"');

    // Run with string hyperlink with and without override
    const hlinkWithOverride = serializeHyperlink('https://example.com', 'rId99');
    expect(hlinkWithOverride).toBeDefined();
    expect(renderXml(hlinkWithOverride)).toContain('r:id="rId99"');
    expect(serializeHyperlink('https://example.com')).toBeUndefined();
    expect(serializeHyperlink(undefined)).toBeUndefined();

    // Hyperlink with standard actions
    const actions = [
      ['firstSlide', 'firstslide'],
      ['nextSlide', 'nextslide'],
      ['endShow', 'endshow'],
      ['lastSlide', 'lastslide'],
      ['previousSlide', 'previousslide'],
    ] as const;
    for (const [action, jumpName] of actions) {
      const node = serializeHyperlink({ action });
      expect(node).toBeDefined();
      expect(renderXml(node)).toContain(`jump=${jumpName}`);
    }
    const slideJumpNode = serializeHyperlink({ slideIndex: 3 });
    expect(slideJumpNode).toBeDefined();
    expect(renderXml(slideJumpNode)).toContain('hlinksldjump');

    // Hyperlink with rId and tooltip
    const hlinkWithTooltip = serializeHyperlink({ rId: 'rId5', tooltip: 'My Tooltip' });
    expect(hlinkWithTooltip).toBeDefined();
    const ttXml = renderXml(hlinkWithTooltip);
    expect(ttXml).toContain('r:id="rId5"');
    expect(ttXml).toContain('tooltip="My Tooltip"');

    // Run properties with hyperlink
    const rPrWithHlink = serializeRunProperties({ hyperlink: { action: 'nextSlide', tooltip: 'Next' } });
    expect(rPrWithHlink).toBeDefined();
    const hlinkRPrXml = renderXml(rPrWithHlink);
    expect(hlinkRPrXml).toContain('<a:hlinkClick');

    // Empty run properties
    const emptyRPr = serializeRunProperties(undefined);
    expect(emptyRPr).toBeUndefined();
  });

  it('covers bullet startAt, text body padding insets, and empty text body', () => {
    // Bullet autoNum with startAt
    const numberedBullet = serializeBulletProperties({ autoNumType: 'romanUcPeriod', startAt: 5, type: 'autoNum' });
    expect(numberedBullet).toBeDefined();
    const bulletXml = renderXml(numberedBullet);
    expect(bulletXml).toContain('startAt="5"');
    expect(bulletXml).toContain('type="romanUcPeriod"');
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
    const bodyPrXml = renderXml(bodyPr);
    expect(bodyPrXml).toContain('lIns="10000"');
    expect(bodyPrXml).toContain('tIns="40000"');
    expect(bodyPrXml).toContain('rIns="30000"');
    expect(bodyPrXml).toContain('bIns="20000"');

    // Empty text body fallback — should produce one empty paragraph
    const emptyBody = serializeTextBody({ bodyProperties: {}, paragraphs: [] });
    const emptyBodyXml = renderXml(emptyBody);
    expect(emptyBodyXml).toContain('<a:p>');

    // Paragraph with justify alignment and single char bullet margin
    const justifyPara = serializeParagraph({
      properties: {
        alignment: 'justify',
        bullet: { char: '-', type: 'char' },
      },
      runs: [{ properties: {}, text: 'Justified item' }],
    });
    const justifyXml = renderXml(justifyPara);
    expect(justifyXml).toContain('algn="justify"');
    expect(justifyXml).toContain('marL="152400"');
    expect(justifyXml).toContain('indent="-152400"');

    // Scheme color string
    const schemeNode = serializeColorNode('accent2');
    expect(renderXml(schemeNode)).toContain('<a:schemeClr');

    // Degree angle <= 360 gradient
    const angleGrad = serializeFill({
      gradient: { angle: 90, stops: [] },
      type: 'gradient',
    });
    const angleXml = renderXml(angleGrad);
    expect(angleXml).toContain('ang="5400000"');
  });
});

describe('Soft line breaks (<a:br>) and sequential interleaving', () => {
  it('serializes explicit break runs interleaved sequentially between text runs', () => {
    const p = serializeParagraph({
      properties: {},
      runs: [
        { properties: { bold: true }, text: 'Before break' },
        { break: true, properties: { fontSize: hundredthsPoint(1400) } },
        { properties: { italic: true }, text: 'After break' },
      ],
    });

    const xml = renderXml(p);
    // Strict sequential check
    const r1Idx = xml.indexOf('Before break');
    const brIdx = xml.indexOf('<a:br>');
    const r2Idx = xml.indexOf('After break');

    expect(r1Idx).toBeGreaterThan(-1);
    expect(brIdx).toBeGreaterThan(r1Idx);
    expect(r2Idx).toBeGreaterThan(brIdx);

    // Verify <a:br> contains the run properties
    expect(xml).toContain('<a:br><a:rPr sz="1400"');
  });

  it('splits newline in text runs into interleaved <a:r> and <a:br>', () => {
    const p = serializeParagraph({
      properties: {},
      runs: [
        { properties: { color: 'FF0000' }, text: 'Line 1\nLine 2\nLine 3' },
      ],
    });

    const xml = renderXml(p);
    expect(xml).toContain('Line 1');
    expect(xml).toContain('<a:br');
    expect(xml).toContain('Line 2');
    expect(xml).toContain('Line 3');

    // Count <a:br> nodes
    const brCount = (xml.match(/<a:br>/g) || []).length;
    expect(brCount).toBe(2);
  });

  it('uses custom indentSettings for levelIndent, char bulletGap, and autoNum bulletGap', () => {
    const customIndentSettings = {
      levelIndent: 365760, // 0.4"
      bulletGap: {
        char: 200000,
        autoNum: 300000,
      },
    };

    // Char bullet at level 2
    const charPara = serializeParagraph(
      {
        properties: {
          bullet: { type: 'char', char: '•' },
          level: 2,
        },
        runs: [{ text: 'Custom bullet' }],
      },
      customIndentSettings,
    );

    const xml = renderXml(charPara);
    // marL = (lvl * levelIndent) + bulletGap = (2 * 365760) + 200000 = 731520 + 200000 = 931520
    expect(xml).toContain('marL="931520"');
    expect(xml).toContain('indent="-200000"');

    // AutoNum bullet at level 1
    const autoNumPara = serializeParagraph(
      {
        properties: {
          bullet: { type: 'autoNum', autoNumType: 'arabicPeriod' },
          level: 1,
        },
        runs: [{ text: 'Numbered item' }],
      },
      customIndentSettings,
    );

    const numXml = renderXml(autoNumPara);
    // marL = (lvl * levelIndent) + bulletGap = (1 * 365760) + 300000 = 665760
    expect(numXml).toContain('marL="665760"');
    expect(numXml).toContain('indent="-300000"');
  });
});

