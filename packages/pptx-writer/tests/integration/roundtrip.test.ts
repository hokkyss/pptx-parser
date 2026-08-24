import { parsePptx } from '@hokkyss/pptx-reader';
import { emu, emuDegree, hundredthsPoint, inchesToEmu } from '@hokkyss/pptx-core';
import type { PptxDocument, PptxTheme } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { writePptx } from '../../lib';

describe('Round-Trip Integration (Synthetic Presentation)', () => {
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

  it('writes a complete multi-slide AST, parses it back, and verifies 100% round-trip fidelity', async () => {
    const originalDoc: PptxDocument = {
      customXml: [
        { path: 'customXml/item1.xml', xmlString: '<item>meta</item>' },
        { path: 'ppt/commentAuthors.xml', xmlString: '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
        { path: 'ppt/handoutMasters/handoutMaster1.xml', xmlString: '<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
      ],
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
          notes: 'Speaker notes for slide 1',
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
          notesBody: {
            bodyProperties: {},
            paragraphs: [{ properties: {}, runs: [{ text: 'Structured speaker notes' }] }],
          },
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
    const roundTripDoc = await parsePptx(generatedBuffer, { customXml: true });

    // 3. Verify fidelity
    expect(roundTripDoc.slides.length).toBe(2);
    expect(roundTripDoc.metadata.title).toBe(originalDoc.metadata.title);
    expect(roundTripDoc.metadata.creator).toBe(originalDoc.metadata.creator);
    expect(roundTripDoc.themes[0].colorScheme.accent1).toBe('0284C7');
    expect(roundTripDoc.themes[0].fontScheme.majorFont).toBe('Inter');
    expect(roundTripDoc.slides[0].notes).toBe('Speaker notes for slide 1');
    expect(roundTripDoc.slides[1].notes).toBe('Structured speaker notes');

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

  it('handles lenient mode auto-generating defaults when masters or themes are missing', async () => {
    const minimalDoc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        slideCount: 1,
        slideHeight: inchesToEmu(7.5),
        slideWidth: inchesToEmu(13.333),
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          elements: [],
          shapes: [],
          slideId: 'rId1',
          slideNumber: 1,
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(minimalDoc, { mode: 'lenient' });
    expect(buffer.length).toBeGreaterThan(0);
    const parsed = await parsePptx(buffer);
    expect(parsed.slides).toHaveLength(1);
  });

  it('throws in strict mode when document has no slides', async () => {
    const invalidDoc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        slideCount: 0,
        slideHeight: inchesToEmu(7.5),
        slideWidth: inchesToEmu(13.333),
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [],
      themes: [],
    };

    await expect(writePptx(invalidDoc, { mode: 'strict' })).rejects.toThrow('Strict mode error');
  });
});
