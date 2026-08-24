import type { PptxDocument } from '@hokkyss/pptx-core';
import { emu } from '@hokkyss/pptx-core';
import { createZipReader } from '@hokkyss/pptx-reader';
import { describe, expect, it } from 'vitest';
import { serializePicture } from '../../lib/serializers/picture-serializer';
import { serializeShape } from '../../lib/serializers/shape-serializer';
import { serializeHyperlink, serializeRunProperties } from '../../lib/serializers/text-serializer';
import { writePptx } from '../../lib/writer';

describe('Hyperlink Serializer (@hokkyss/pptx-writer)', () => {
  it('should serialize external URL hyperlink with tooltip', () => {
    const hlinkNode = serializeHyperlink({
      rId: 'rId2',
      tooltip: 'Visit Website',
      url: 'https://example.com',
    });

    expect(hlinkNode).toEqual({
      '@_r:id': 'rId2',
      '@_tooltip': 'Visit Website',
    });
  });

  it('should serialize slide jump actions properly', () => {
    const nextNode = serializeHyperlink({ action: 'nextSlide' });
    expect(nextNode).toEqual({
      '@_action': 'ppaction://hlinkshowjump?jump=nextslide',
    });

    const slideJumpNode = serializeHyperlink({ rId: 'rId3', slideIndex: 4 });
    expect(slideJumpNode).toEqual({
      '@_action': 'ppaction://hlinksldjump',
      '@_r:id': 'rId3',
    });
  });

  it('should serialize hyperlink inside run properties', () => {
    const rPr = serializeRunProperties({
      bold: true,
      hyperlink: {
        rId: 'rId5',
        tooltip: 'Click Me',
      },
    });

    expect(rPr['@_b']).toBe('1');
    expect(rPr['a:hlinkClick']).toEqual({
      '@_r:id': 'rId5',
      '@_tooltip': 'Click Me',
    });
  });

  it('should serialize hyperlink inside shape and picture cNvPr', () => {
    const shape = serializeShape({
      elementType: 'shape',
      hyperlink: {
        rId: 'rId10',
        tooltip: 'Shape Link',
      },
      id: '2',
      name: 'Linked Shape',
      position: { cx: emu(1000000), cy: emu(500000), x: emu(0), y: emu(0) },
      shapeType: 'roundRect',
      type: 'shape',
    });

    const nvSpPr = shape['p:nvSpPr'] as Record<string, unknown>;
    const cNvPr = nvSpPr['p:cNvPr'] as Record<string, unknown>;
    expect(cNvPr['a:hlinkClick']).toEqual({
      '@_r:id': 'rId10',
      '@_tooltip': 'Shape Link',
    });

    const pic = serializePicture({
      elementType: 'picture',
      hyperlink: {
        action: 'nextSlide',
      },
      id: '3',
      name: 'Linked Pic',
      picture: { mediaId: 'img1' },
      position: { cx: emu(1000000), cy: emu(500000), x: emu(0), y: emu(0) },
      type: 'picture',
    });

    const nvPicPr = pic['p:nvPicPr'] as Record<string, unknown>;
    const picCNvPr = nvPicPr['p:cNvPr'] as Record<string, unknown>;
    expect(picCNvPr['a:hlinkClick']).toEqual({
      '@_action': 'ppaction://hlinkshowjump?jump=nextslide',
    });
  });

  it('should register relationships in slide.xml.rels when writing presentation', async () => {
    const doc: PptxDocument = {
      media: [],
      metadata: {},
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          elements: [
            {
              elementType: 'shape',
              hyperlink: 'https://hokkyss.dev',
              id: '2',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
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
            },
          ],
          slideId: 'rId2',
          slideNumber: 1,
        },
        {
          elements: [],
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
      media: [],
      metadata: {},
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          elements: [
            {
              elementType: 'shape',
              hyperlink: 'javascript:alert(1)',
              id: '2',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
              shapeType: 'rect',
              type: 'shape',
            },
            {
              elementType: 'shape',
              hyperlink: {
                tooltip: 'Dangerous\r\nTooltip\0Breakout',
                url: 'file:///C:/Windows/System32/cmd.exe',
              },
              id: '3',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(0), y: emu(0) },
              shapeType: 'rect',
              type: 'shape',
            },
          ],
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

import { serializeHyperlink } from '../../lib/serializers/text-serializer';

describe('Hyperlink Serializer action jumps', () => {
  it('serializes all predefined jump actions', () => {
    expect(serializeHyperlink({ action: 'endShow' })?.['@_action']).toBe('ppaction://hlinkshowjump?jump=endshow');
    expect(serializeHyperlink({ action: 'firstSlide' })?.['@_action']).toBe('ppaction://hlinkshowjump?jump=firstslide');
    expect(serializeHyperlink({ action: 'lastSlide' })?.['@_action']).toBe('ppaction://hlinkshowjump?jump=lastslide');
    expect(serializeHyperlink({ action: 'nextSlide' })?.['@_action']).toBe('ppaction://hlinkshowjump?jump=nextslide');
    expect(serializeHyperlink({ action: 'previousSlide' })?.['@_action']).toBe('ppaction://hlinkshowjump?jump=previousslide');
    expect(serializeHyperlink({ action: 'ppaction://customAction' })?.['@_action']).toBe('ppaction://customAction');
  });
});

import { serializeBulletProperties } from '../../lib/serializers/text-serializer';

describe('Hyperlink Serializer string target and bullet fallback', () => {
  it('serializes string hyperlink with relIdOverride and handles bullet fallback', () => {
    expect(serializeHyperlink('https://example.com', 'rId9')?.['@_r:id']).toBe('rId9');
    expect(serializeHyperlink('https://example.com')).toBeUndefined();
    // @ts-expect-error Testing unknown bullet type fallback
    expect(serializeBulletProperties({ type: 'unknown' })).toBeUndefined();
  });
});
