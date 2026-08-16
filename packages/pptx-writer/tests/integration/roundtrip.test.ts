import { parsePptx } from '@hokkyss/pptx-reader';
import { emu, emuDegree, hundredthsPoint, inchesToEmu } from '@hokkyss/pptx-core';
import type { PptxDocument, PptxTheme } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { writePptx } from '../../lib';

describe('Round-Trip Integration (Synthetic Presentation)', () => {
  it('writes a complete multi-slide AST, parses it back, and verifies 100% round-trip fidelity', async () => {
    const syntheticTheme: PptxTheme = {
      colorScheme: {
        accent1: '0284C7',
        accent2: '6366F1',
        accent3: '10B981',
        accent4: 'F59E0B',
        accent5: 'EF4444',
        accent6: '8B5CF6',
        dk1: '000000',
        dk2: '1E293B',
        folHlink: '7C3AED',
        hlink: '2563EB',
        lt1: 'FFFFFF',
        lt2: 'F8FAFC',
      },
      customColors: {},
      fontScheme: {
        majorFont: 'Inter',
        minorFont: 'Roboto',
        name: 'Synthetic Fonts',
      },
      formatScheme: {},
      id: 'theme1',
      name: 'Synthetic Theme',
    };

    const originalDoc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        creator: 'Synthetic Architect',
        lastModifiedBy: 'Synthetic Architect',
        revision: 2,
        slideCount: 2,
        slideHeight: inchesToEmu(7.5),
        slideWidth: inchesToEmu(13.333),
        title: 'High-Fidelity Synthetic Deck',
      },
      slideLayouts: [
        {
          elements: [],
          id: 'slideLayout1',
          masterId: 'slideMaster1',
          name: 'Title and Content',
          shapes: [],
          type: 'layout',
        },
      ],
      slideMasters: [
        {
          elements: [],
          id: 'slideMaster1',
          layoutIds: ['slideLayout1'],
          name: 'Synthetic Master',
          shapes: [],
          theme: syntheticTheme,
        },
      ],
      slides: [
        {
          animations: [],
          elements: [
            {
              elementType: 'shape',
              id: '2',
              isVisible: true,
              name: 'Header Shape',
              position: { cx: emu(8000000), cy: emu(1500000), x: emu(1000000), y: emu(1000000) },
              rotation: emuDegree(0),
              textBody: {
                bodyProperties: {},
                paragraphs: [
                  {
                    properties: { alignment: 'center', level: 0 },
                    runs: [
                      {
                        properties: { bold: true, color: '0284C7', fontSize: hundredthsPoint(3600) },
                        text: 'Slide 1 Title',
                      },
                    ],
                  },
                ],
              },
              type: 'shape',
              zIndex: 0,
            },
          ],
          layoutId: 'slideLayout1',
          shapes: [],
          slideId: 'rId2',
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
              position: { cx: emu(8000000), cy: emu(3000000), x: emu(1000000), y: emu(1000000) },
              rotation: emuDegree(0),
              table: {
                columnWidths: [emu(4000000), emu(4000000)],
                rows: [
                  {
                    cells: [
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: { bold: true }, text: 'Column A' }] }] } },
                      { textBody: { bodyProperties: {}, paragraphs: [{ properties: {}, runs: [{ properties: { bold: true }, text: 'Column B' }] }] } },
                    ],
                    height: emu(500000),
                  },
                ],
              },
              type: 'graphicFrame',
              zIndex: 0,
            },
          ],
          layoutId: 'slideLayout1',
          shapes: [],
          slideId: 'rId3',
          slideNumber: 2,
        },
      ],
      themes: [syntheticTheme],
    };

    // 1. Serialize
    const generatedBuffer = await writePptx(originalDoc);
    expect(generatedBuffer).toBeInstanceOf(Uint8Array);
    expect(generatedBuffer.length).toBeGreaterThan(0);

    // 2. Parse back
    const roundTripDoc = await parsePptx(generatedBuffer);

    // 3. Verify fidelity
    expect(roundTripDoc.slides.length).toBe(2);
    expect(roundTripDoc.metadata.title).toBe(originalDoc.metadata.title);
    expect(roundTripDoc.metadata.creator).toBe(originalDoc.metadata.creator);
    expect(roundTripDoc.themes[0].colorScheme.accent1).toBe('0284C7');
    expect(roundTripDoc.themes[0].fontScheme.majorFont).toBe('Inter');

    const shape = roundTripDoc.slides[0].elements[0];
    expect(shape.name).toBe('Header Shape');
    if (shape.elementType === 'shape') {
      expect(shape.textBody?.paragraphs[0]?.runs[0]?.text).toBe('Slide 1 Title');
    }

    const table = roundTripDoc.slides[1].elements[0];
    expect(table.elementType).toBe('table');
    if (table.elementType === 'table') {
      expect(table.table?.rows.length).toBe(1);
    }
  });
});
