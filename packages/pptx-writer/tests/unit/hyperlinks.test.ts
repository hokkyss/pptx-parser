import type { PptxDocument } from '@hokkyss/pptx-core';
import { emu, emuDegree } from '@hokkyss/pptx-core';
import { createZipReader } from '@hokkyss/pptx-reader';
import { describe, expect, it } from 'vitest';
import { serializePicture } from '../../lib/serializers/picture-serializer';
import { serializeShape } from '../../lib/serializers/shape-serializer';
import {
  serializeBulletProperties,
  serializeHyperlink,
  serializeRunProperties,
} from '../../lib/serializers/text-serializer';
import { writePptx } from '../../lib/writer';
import { renderXml } from '../../lib/xml/xml-element';

describe('Hyperlink Serializer (@hokkyss/pptx-writer)', () => {
  it('should serialize external URL hyperlink with tooltip', () => {
    const hlinkNode = serializeHyperlink({
      rId: 'rId2',
      tooltip: 'Visit Website',
      url: 'https://example.com',
    });

    expect(hlinkNode).toBeDefined();
    const xml = renderXml(hlinkNode);
    expect(xml).toContain('r:id="rId2"');
    expect(xml).toContain('tooltip="Visit Website"');
  });

  it('should serialize slide jump actions properly', () => {
    const nextNode = serializeHyperlink({ action: 'nextSlide' });
    expect(nextNode).toBeDefined();
    expect(renderXml(nextNode)).toContain('jump=nextslide');

    const slideJumpNode = serializeHyperlink({ rId: 'rId3', slideIndex: 4 });
    expect(slideJumpNode).toBeDefined();
    const sjXml = renderXml(slideJumpNode);
    expect(sjXml).toContain('hlinksldjump');
    expect(sjXml).toContain('r:id="rId3"');
  });

  it('should serialize hyperlink inside run properties', () => {
    const rPr = serializeRunProperties({
      bold: true,
      hyperlink: {
        rId: 'rId5',
        tooltip: 'Click Me',
      },
    });

    expect(rPr).toBeDefined();
    const xml = renderXml(rPr);
    expect(xml).toContain('b="1"');
    expect(xml).toContain('<a:hlinkClick');
    expect(xml).toContain('r:id="rId5"');
    expect(xml).toContain('tooltip="Click Me"');
  });

  it('should serialize hyperlink inside shape and picture cNvPr', () => {
    const shape = serializeShape({
      elementType: 'shape',
      hyperlink: {
        rId: 'rId10',
        tooltip: 'Shape Link',
      },
      id: '2',
      isVisible: true,
      name: 'Linked Shape',
      position: { cx: emu(1000000), cy: emu(500000), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      shapeType: 'roundRect',
      type: 'shape',
      zIndex: 0,
    });

    const shapeXml = renderXml(shape);
    expect(shapeXml).toContain('<a:hlinkClick');
    expect(shapeXml).toContain('r:id="rId10"');
    expect(shapeXml).toContain('tooltip="Shape Link"');

    const pic = serializePicture({
      elementType: 'picture',
      hyperlink: {
        action: 'nextSlide',
      },
      id: '3',
      isVisible: true,
      name: 'Linked Pic',
      picture: { mediaId: 'img1' },
      position: { cx: emu(1000000), cy: emu(500000), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      type: 'picture',
      zIndex: 0,
    });

    const picXml = renderXml(pic);
    expect(picXml).toContain('<a:hlinkClick');
    expect(picXml).toContain('jump=nextslide');
  });

  it('should register relationships in slide.xml.rels when writing presentation', async () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        slideCount: 2,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000),
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [
            {
              elementType: 'shape',
              hyperlink: 'https://hokkyss.dev',
              id: '2',
              isVisible: true,
              name: 'Shape 2',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
              rotation: emuDegree(0),
              shapeType: 'rect',
              textBody: {
                bodyProperties: {},
                paragraphs: [
                  {
                    properties: {},
                    runs: [
                      {
                        properties: {
                          hyperlink: {
                            slideIndex: 2,
                            tooltip: 'Jump to slide 2',
                          },
                        },
                        text: 'Go to Slide 2',
                      },
                    ],
                  },
                ],
              },
              type: 'shape',
              zIndex: 0,
            },
          ],
          shapes: [],
          slideId: 'rId2',
          slideNumber: 1,
        },
        {
          animations: [],
          elements: [],
          shapes: [],
          slideId: 'rId3',
          slideNumber: 2,
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(doc);
    const zip = await createZipReader(buffer);

    expect(zip.hasFile('ppt/slides/_rels/slide1.xml.rels')).toBe(true);
    const relsXml = zip.getFileText('ppt/slides/_rels/slide1.xml.rels');
    expect(relsXml).toContain('Target="https://hokkyss.dev"');
    expect(relsXml).toContain('TargetMode="External"');
    expect(relsXml).toContain('Target="slide2.xml"');
  });

  it('neutralizes dangerous javascript: and vbscript: URIs from writer relationships', async () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        slideCount: 1,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000),
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [
            {
              elementType: 'shape',
              hyperlink: 'javascript:alert(1)',
              id: '2',
              isVisible: true,
              name: 'Shape 2',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
              rotation: emuDegree(0),
              shapeType: 'rect',
              type: 'shape',
              zIndex: 0,
            },
            {
              elementType: 'shape',
              hyperlink: {
                tooltip: 'Dangerous\r\nTooltip\0Breakout',
                url: 'file:///C:/Windows/System32/cmd.exe',
              },
              id: '3',
              isVisible: true,
              name: 'Shape 3',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
              rotation: emuDegree(0),
              shapeType: 'rect',
              type: 'shape',
              zIndex: 1,
            },
          ],
          shapes: [],
          slideId: 'rId2',
          slideNumber: 1,
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(doc);
    const zip = await createZipReader(buffer);

    const relsXml = zip.getFileText('ppt/slides/_rels/slide1.xml.rels') || '';
    expect(relsXml).not.toContain('javascript:');
    expect(relsXml).not.toContain('file:///');
  });
});

describe('Hyperlink Serializer action jumps', () => {
  it('serializes all predefined jump actions', () => {
    const cases = [
      ['endShow', 'endshow'],
      ['firstSlide', 'firstslide'],
      ['lastSlide', 'lastslide'],
      ['nextSlide', 'nextslide'],
      ['previousSlide', 'previousslide'],
    ] as const;
    for (const [action, jump] of cases) {
      const node = serializeHyperlink({ action });
      expect(node).toBeDefined();
      expect(renderXml(node)).toContain(`jump=${jump}`);
    }
    // Custom ppaction
    const customNode = serializeHyperlink({ action: 'ppaction://customAction' });
    expect(customNode).toBeDefined();
    expect(renderXml(customNode)).toContain('ppaction://customAction');
  });
});

describe('Hyperlink Serializer string target and bullet fallback', () => {
  it('serializes string hyperlink with relIdOverride and handles bullet fallback', () => {
    const node = serializeHyperlink('https://example.com', 'rId9');
    expect(node).toBeDefined();
    expect(renderXml(node)).toContain('r:id="rId9"');
    expect(serializeHyperlink('https://example.com')).toBeUndefined();
    // @ts-expect-error Testing unknown bullet type fallback
    expect(serializeBulletProperties({ type: 'unknown' })).toBeUndefined();
  });
});
