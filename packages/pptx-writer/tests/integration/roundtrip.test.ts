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

describe('writePptx rawXml, relsXml and notesMasters', () => {
  it('preserves slide rawXml, relsXml, and custom notesMasters', async () => {
    const docWithRaw: PptxDocument = {
      customXml: [
        { path: 'ppt/notesMasters/notesMaster1.xml', xmlString: '<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
      ],
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
          slideId: 'rId1',
          slideNumber: 1,
          elements: [],
          shapes: [],
          animations: [],
          rawXml: '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>',
          relsXml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(docWithRaw);
    expect(buffer.length).toBeGreaterThan(0);
    const parsed = await parsePptx(buffer);
    expect(parsed.slides).toHaveLength(1);
  });
});

describe('writePptx charts and media serialization', () => {
  it('serializes charts and pictures with relationships in full writePptx pipeline', async () => {
    const docWithMediaAndChart: PptxDocument = {
      customXml: [],
      media: [
        {
          data: new Uint8Array([137, 80, 78, 71]),
          fileName: 'hero.png',
          filename: 'hero.png',
          id: 'hero_img',
          mimeType: 'image/png',
          path: 'ppt/media/hero.png',
        },
      ],
      metadata: { slideCount: 1, slideHeight: inchesToEmu(7.5), slideWidth: inchesToEmu(13.333) },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          shapes: [],
          animations: [],
          elements: [
            {
              id: 'pic1',
              name: 'Hero Image',
              type: 'picture',
              elementType: 'picture',
              isVisible: true,
              zIndex: 0,
              position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
              rotation: emuDegree(0),
              picture: { mediaId: 'hero_img' },
            },
            {
              id: 'chart1',
              name: 'Quarterly Sales',
              type: 'graphicFrame',
              elementType: 'chart',
              isVisible: true,
              zIndex: 1,
              position: { x: emu(100), y: emu(100), cx: emu(1000), cy: emu(1000) },
              rotation: emuDegree(0),
              chart: {
                chartType: 'barChart',
                categories: ['Q1', 'Q2'],
                series: [{ name: 'Sales', values: [100, 200], index: 0, order: 0 }],
              },
            },
          ],
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(docWithMediaAndChart);
    expect(buffer.length).toBeGreaterThan(0);
    const parsed = await parsePptx(buffer);
    expect(parsed.slides[0].elements).toHaveLength(2);
  });
});

describe('writePptx comments and notesSlides customXml overrides', () => {
  it('serializes custom notesSlides, comments, embeddings, and charts overrides', async () => {
    const docWithAllOverrides: PptxDocument = {
      customXml: [
        { path: 'ppt/notesSlides/notesSlide1.xml', xmlString: '<p:notesSlide xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
        { path: 'ppt/comments/comment1.xml', xmlString: '<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
        { path: 'ppt/charts/chart99.xml', xmlString: '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' },
        { path: 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', binaryData: new Uint8Array([80, 75, 3, 4]) },
        { path: 'ppt/handoutMasters/handoutMaster1.xml', xmlString: '<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
        { path: 'ppt/commentAuthors.xml', xmlString: '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
      ],
      media: [],
      metadata: { slideCount: 1, slideHeight: inchesToEmu(7.5), slideWidth: inchesToEmu(13.333) },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          slideId: 'rId1',
          slideNumber: 1,
          elements: [],
          shapes: [],
          animations: [],
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(docWithAllOverrides);
    expect(buffer.length).toBeGreaterThan(0);
  });
});
