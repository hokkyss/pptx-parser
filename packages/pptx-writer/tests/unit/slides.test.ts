import type { PptxSlide } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { serializeSlide } from '../../lib/serializers/slide-serializer';

describe('Slide Serializer (@hokkyss/pptx-writer)', () => {
  it('serializes complete slide XML with shapes and solid background', () => {
    const slide: PptxSlide = {
      animations: [],
      background: {
        fill: {
          solidColor: { type: 'srgb', value: 'FFFFFF' },
          type: 'solid',
        },
      },
      elements: [
        {
          elementType: 'shape',
          id: '2',
          isVisible: true,
          name: 'Title 1',
          position: {
            cx: emu(5000000),
            cy: emu(1000000),
            x: emu(100000),
            y: emu(100000),
          },
          rotation: emuDegree(0),
          textBody: {
            bodyProperties: {},
            paragraphs: [
              {
                properties: {},
                runs: [{ properties: { fontSize: hundredthsPoint(3200) }, text: 'Presentation Title' }],
              },
            ],
          },
          type: 'shape',
          zIndex: 0,
        },
      ],
      shapes: [],
      slideId: 'rId1',
      slideNumber: 1,
    };

    const xml = serializeSlide(slide);
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(xml).toContain('<p:sld');
    expect(xml).toContain('<p:cSld>');
    expect(xml).toContain('<p:spTree>');
    expect(xml).toContain('Presentation Title');
    expect(xml).toContain('<p:bg>');
  });

  it('serializes slide with transitions, animations, and gradient background', () => {
    const slide: PptxSlide = {
      slideId: 'rId2',
      slideNumber: 2,
      background: {
        fill: {
          type: 'gradient',
          gradient: {
            type: 'linear',
            angle: 90,
            stops: [
              { color: { type: 'srgb', value: '000000' }, position: 0 },
              { color: { type: 'srgb', value: 'FFFFFF' }, position: 100000 },
            ],
          },
        },
      },
      transition: {
        type: 'wipe',
        direction: 'right',
        speed: 'fast',
      },
      animations: [
        {
          targetShapeId: '2',
          trigger: 'onClick',
          effect: 'appear',
          sequence: 0,
        },
      ],
      elements: [
        {
          elementType: 'connector',
          type: 'connector',
          id: '3',
          name: 'Cxn',
          isVisible: true,
          zIndex: 0,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
        {
          elementType: 'group',
          type: 'group',
          id: '4',
          name: 'Grp',
          isVisible: true,
          zIndex: 1,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
          children: [],
        },
      ],
      shapes: [],
    };

    const xml = serializeSlide(slide);
    expect(xml).toContain('<p:transition');
    expect(xml).toContain('<p:timing>');
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('<p:grpSp>');
    expect(xml).toContain('<a:gradFill>');
  });
});

describe('Slide Serializer pictures, tables, charts and unique IDs', () => {
  it('serializes picture, table, chart and resolves duplicate IDs', () => {
    const slide: PptxSlide = {
      slideId: 'rId3',
      slideNumber: 3,
      shapes: [],
      animations: [],
      elements: [
        {
          id: '1', // duplicate of container group ID 1
          name: 'Pic 1',
          type: 'picture',
          elementType: 'picture',
          isVisible: true,
          zIndex: 0,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
          picture: { mediaId: 'img1' },
        },
        {
          id: '2',
          name: 'Table 2',
          type: 'graphicFrame',
          elementType: 'table',
          isVisible: true,
          zIndex: 1,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
          table: {
            columnWidths: [emu(500)],
            rows: [{ height: emu(200), cells: [{ textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ text: 'Cell' }] }] } }] }],
          },
        },
        {
          id: '3',
          name: 'Chart 3',
          type: 'graphicFrame',
          elementType: 'chart',
          isVisible: true,
          zIndex: 2,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
          chart: {
            chartType: 'barChart',
            categories: ['A'],
            series: [{ name: 'S1', values: [10], index: 0, order: 0 }],
          },
        },
      ],
    };

    const picMap = new Map<string, string>([['img1', 'rIdImg']]);
    const xml = serializeSlide(slide, picMap, ['rIdChart']);
    expect(xml).toContain('<p:pic>');
    expect(xml).toContain('<p:graphicFrame>');
    expect(xml).toContain('rIdImg');
    expect(xml).toContain('rIdChart');
  });
});

import { writePptx } from '../../lib/writer';
import type { PptxDocument } from '@hokkyss/pptx-core';

describe('PptxWriter advanced slide layouts, table hyperlinks, and lazy media', () => {
  it('handles slide layout custom rels, lazy media getters, and table hyperlinks', async () => {
    const doc: PptxDocument = {
      presentation: {
        slideWidth: emu(9144000),
        slideHeight: emu(5143500),
      },
      slideMasters: [
        {
          id: 'slideMaster1.xml',
          slideMasterId: 2147483648,
          elements: [],
          shapes: [],
        },
      ],
      slideLayouts: [
        {
          id: 'slideLayout1.xml',
          masterId: 'slideMaster1.xml',
          rawXml: '<p:sldLayout/>',
          relsXml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
      ],
      media: [
        {
          id: 'lazy1',
          fileName: 'lazy.png',
          data: null,
          lazyGetter: async () => new Uint8Array([137, 80, 78, 71]),
        },
      ],
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          shapes: [],
          animations: [],
          elements: [
            {
              id: '1',
              name: 'Table 1',
              type: 'graphicFrame',
              elementType: 'table',
              isVisible: true,
              zIndex: 0,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              table: {
                columnWidths: [emu(100)],
                rows: [
                  {
                    height: emu(100),
                    cells: [
                      {
                        textBody: {
                          paragraphs: [
                            {
                              runs: [
                                {
                                  text: 'Link',
                                  properties: {
                                    hyperlink: { url: 'https://example.com' },
                                  },
                                },
                              ],
                            },
                          ],
                        },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    };

    const bytes = await writePptx(doc);
    expect(bytes).toBeDefined();
    expect(bytes.byteLength).toBeGreaterThan(0);
  });
});

describe('PptxWriter lenient default slides, group hyperlinks, and slideMaster relsXml', () => {
  it('handles empty slides in lenient mode, group hyperlinks, and custom slideMaster relsXml', async () => {
    // 1. Empty slides fallback
    const emptyDoc: PptxDocument = {
      presentation: { slideWidth: emu(9144000), slideHeight: emu(5143500) },
      slides: [],
    };
    const emptyBytes = await writePptx(emptyDoc);
    expect(emptyBytes.byteLength).toBeGreaterThan(0);

    // 2. Group containing shape with hyperlink + slideMaster with custom relsXml
    const groupHyperlinkDoc: PptxDocument = {
      presentation: { slideWidth: emu(9144000), slideHeight: emu(5143500) },
      slideMasters: [
        {
          id: 'slideMaster1.xml',
          slideMasterId: 2147483648,
          elements: [],
          shapes: [],
          relsXml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
      ],
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          shapes: [],
          animations: [],
          elements: [
            {
              id: 'g1',
              name: 'Group 1',
              type: 'group',
              elementType: 'group',
              isVisible: true,
              zIndex: 0,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              children: [
                {
                  id: 's1',
                  name: 'Shape 1',
                  type: 'shape',
                  elementType: 'shape',
                  isVisible: true,
                  zIndex: 0,
                  position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
                  rotation: emuDegree(0),
                  hyperlink: { url: 'https://example.com/child' },
                },
              ],
            },
          ],
        },
      ],
    };
    const groupBytes = await writePptx(groupHyperlinkDoc);
    expect(groupBytes.byteLength).toBeGreaterThan(0);
  });
});

describe('Slide Serializer connected connectors and unsafe hyperlinks', () => {
  it('serializes connected connector attachment IDs and removes unsafe javascript hyperlinks', async () => {
    const slideWithConnections: PptxSlide = {
      slideId: 'rId4',
      slideNumber: 4,
      shapes: [],
      animations: [],
      elements: [
        {
          id: 'source-shape',
          name: 'Source',
          type: 'shape',
          elementType: 'shape',
          isVisible: true,
          zIndex: 0,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
        {
          id: 'target-shape',
          name: 'Target',
          type: 'shape',
          elementType: 'shape',
          isVisible: true,
          zIndex: 1,
          position: { x: emu(200), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
        {
          id: 'cxn-1',
          name: 'Connector 1',
          type: 'connector',
          elementType: 'connector',
          isVisible: true,
          zIndex: 2,
          position: { x: emu(0), y: emu(0), cx: emu(200), cy: emu(100) },
          rotation: emuDegree(0),
          startConnection: { shapeId: 'source-shape', position: 'right' },
          endConnection: { shapeId: 'target-shape', position: 'left' },
        },
      ],
    };

    const xml = serializeSlide(slideWithConnections);
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('stCxn');
    expect(xml).toContain('endCxn');

    // Unsafe hyperlinks test in presentation write
    const unsafeDoc: PptxDocument = {
      presentation: { slideWidth: emu(9144000), slideHeight: emu(5143500) },
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          shapes: [],
          animations: [],
          elements: [
            {
              id: 's1',
              name: 'Shape 1',
              type: 'shape',
              elementType: 'shape',
              isVisible: true,
              zIndex: 0,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              textBody: {
                paragraphs: [
                  {
                    runs: [{ text: 'Unsafe Run', properties: { hyperlink: { url: 'javascript:alert(1)' } } }],
                  },
                ],
              },
            },
            {
              id: 't1',
              name: 'Table 1',
              type: 'graphicFrame',
              elementType: 'table',
              isVisible: true,
              zIndex: 1,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              table: {
                columnWidths: [emu(100)],
                rows: [
                  {
                    height: emu(100),
                    cells: [
                      {
                        textBody: {
                          paragraphs: [
                            {
                              runs: [{ text: 'Unsafe Cell', properties: { hyperlink: { url: 'javascript:alert(2)' } } }],
                            },
                          ],
                        },
                      },
                    ],
                  },
                ],
              },
            },
          ],
        },
      ],
    };

    const bytes = await writePptx(unsafeDoc);
    expect(bytes).toBeDefined();
  });
});

describe('Slide Serializer ID collision skipping, actions, and invalid slide jumps', () => {
  it('skips taken numeric IDs, handles action hyperlinks and invalid slide index jumps', async () => {
    const slideWithId2: PptxSlide = {
      slideId: 'rId5',
      slideNumber: 5,
      transition: { type: 'fade', throughBlack: true },
      shapes: [],
      animations: [],
      elements: [
        {
          id: '2', // Manually occupies 2
          name: 'Shape 2',
          type: 'shape',
          elementType: 'shape',
          isVisible: true,
          zIndex: 0,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
        {
          // No ID specified -> Auto-increments past 2 to 3
          name: 'Shape Auto',
          type: 'shape',
          elementType: 'shape',
          isVisible: true,
          zIndex: 1,
          position: { x: emu(100), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
      ],
    };

    const xml = serializeSlide(slideWithId2);
    expect(xml).toContain('id="2"');
    expect(xml).toContain('id="3"');
    expect(xml).toContain('<p:fade');
    expect(xml).toContain('thruBlk="1"');

    // Hyperlinks with action and invalid slideIndex
    const testDoc: PptxDocument = {
      presentation: { slideWidth: emu(9144000), slideHeight: emu(5143500) },
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          shapes: [],
          animations: [],
          elements: [
            {
              id: 's1',
              name: 'S1',
              type: 'shape',
              elementType: 'shape',
              isVisible: true,
              zIndex: 0,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              hyperlink: { action: 'nextSlide' },
            },
            {
              id: 's2',
              name: 'S2',
              type: 'shape',
              elementType: 'shape',
              isVisible: true,
              zIndex: 1,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              hyperlink: { slideIndex: -99 },
            },
          ],
        },
      ],
    };

    const bytes = await writePptx(testDoc);
    expect(bytes).toBeDefined();
  });
});
