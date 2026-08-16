import type { PptxDocument } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { parsePptx } from '@hokkyss/pptx-reader';
import { describe, expect, it } from 'vitest';
import { writePptx } from '../../lib';

describe('Synthetic Presentation Generation', () => {
  it('generates a valid PPTX document from scratch that parsePptx can read', async () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        creator: 'Writer Test',
        slideCount: 2,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000), // 16:9 widescreen
        title: 'Quarterly Kickoff',
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [
            {
              elementType: 'shape',
              fill: {
                solidColor: { type: 'srgb', value: '4F81BD' },
                type: 'solid',
              },
              id: '2',
              isVisible: true,
              name: 'Title Box',
              position: {
                cx: emu(8000000),
                cy: emu(1500000),
                x: emu(1000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              textBody: {
                bodyProperties: {},
                paragraphs: [
                  {
                    properties: {
                      alignment: 'center',
                    },
                    runs: [
                      {
                        properties: {
                          bold: true,
                          color: 'FFFFFF',
                          fontSize: hundredthsPoint(2800),
                        },
                        text: 'Welcome to PPTX Writer',
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
          slideId: 'rId1',
          slideNumber: 1,
        },
        {
          animations: [],
          elements: [
            {
              elementType: 'table',
              id: '3',
              isVisible: true,
              name: 'Data Table',
              position: {
                cx: emu(6000000),
                cy: emu(3000000),
                x: emu(1000000),
                y: emu(1000000),
              },
              rotation: emuDegree(0),
              table: {
                columnWidths: [emu(3000000), emu(3000000)],
                rows: [
                  {
                    cells: [
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Metric' }] }] } },
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Value' }] }] } },
                    ],
                    height: emu(1000000),
                  },
                  {
                    cells: [
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: 'Growth' }] }] } },
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: {}, text: '120%' }] }] } },
                    ],
                    height: emu(1000000),
                  },
                ],
              },
              type: 'graphicFrame',
              zIndex: 0,
            },
          ],
          shapes: [],
          slideId: 'rId2',
          slideNumber: 2,
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(doc);
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    // Read back with pptx-reader
    const parsed = await parsePptx(buffer);
    expect(parsed.slides).toHaveLength(2);
    expect(parsed.metadata.title).toBe('Quarterly Kickoff');
    expect(parsed.metadata.slideWidth).toBe(12192000);
    expect(parsed.metadata.slideHeight).toBe(6858000);

    // Verify slide 1 shape & text
    const s1 = parsed.slides[0];
    expect(s1.elements.length).toBeGreaterThan(0);
    const shape = s1.elements[0];
    expect(shape.elementType).toBe('shape');
    expect(shape.textBody?.paragraphs[0].runs[0].text).toContain('Welcome to PPTX Writer');

    // Verify slide 2 table
    const s2 = parsed.slides[1];
    const tableElem = s2.elements[0];
    expect(tableElem.elementType).toBe('table');
    if (tableElem.elementType === 'table') {
      expect(tableElem.table.rows).toHaveLength(2);
    }
  });
});
