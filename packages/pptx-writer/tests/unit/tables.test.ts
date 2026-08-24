import { describe, expect, it } from 'vitest';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeTable } from '../../lib/serializers/table-serializer';

describe('Table Serializer', () => {
  it('serializes table inside graphicFrame with grid columns, rows, and cells', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '3',
      name: 'Table 1',
      isVisible: true,
      zIndex: 1,
      position: {
        x: emu(500000),
        y: emu(500000),
        cx: emu(4000000),
        cy: emu(2000000),
      },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(2000000), emu(2000000)],
        rows: [
          {
            height: emu(1000000),
            cells: [
              {
                textBody: {
                  bodyProperties: {},
                  paragraphs: [
                    {
                      properties: {},
                      runs: [{ text: 'Header 1', properties: { bold: true, fontSize: hundredthsPoint(1400) } }],
                    },
                  ],
                },
              },
              {
                textBody: {
                  bodyProperties: {},
                  paragraphs: [
                    {
                      properties: {},
                      runs: [{ text: 'Header 2', properties: { bold: true, fontSize: hundredthsPoint(1400) } }],
                    },
                  ],
                },
              },
            ],
          },
          {
            height: emu(1000000),
            cells: [
              {
                textBody: {
                  bodyProperties: {},
                  paragraphs: [
                    {
                      properties: {},
                      runs: [{ text: 'Value 1', properties: { fontSize: hundredthsPoint(1200) } }],
                    },
                  ],
                },
              },
              {
                textBody: {
                  bodyProperties: {},
                  paragraphs: [
                    {
                      properties: {},
                      runs: [{ text: 'Value 2', properties: { fontSize: hundredthsPoint(1200) } }],
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement);
    expect(xmlObject).toBeDefined();

    const nvGFPr = xmlObject['p:nvGraphicFramePr'] as Record<string, Record<string, unknown>>;
    expect(nvGFPr['p:cNvPr']['@_id']).toBe('3');

    const graphic = xmlObject['a:graphic'] as Record<string, Record<string, Record<string, unknown>>>;
    const tbl = graphic['a:graphicData']['a:tbl'] as Record<string, unknown[]>;
    expect(tbl).toBeDefined();

    const tblGrid = tbl['a:tblGrid'] as unknown as Record<string, unknown[]>;
    expect(tblGrid['a:gridCol']).toHaveLength(2);
    expect(tbl['a:tr']).toHaveLength(2);

    const firstRow = tbl['a:tr'][0] as Record<string, unknown[]>;
    expect(firstRow['a:tc']).toHaveLength(2);
  });

  it('serializes table cell with colSpan > 1 adding @_gridSpan', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '4',
      name: 'Table 4',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(4000000), cy: emu(1000000) },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(2000000), emu(2000000)],
        rows: [
          {
            height: emu(500000),
            cells: [{ colSpan: 2, textBody: { bodyProperties: {}, paragraphs: [] } }],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement);
    const tbl = (xmlObject['a:graphic'] as Record<string, Record<string, unknown>>)['a:graphicData']['a:tbl'] as Record<string, unknown[]>;
    const firstRow = tbl['a:tr'][0] as Record<string, unknown[]>;
    const cell = firstRow['a:tc'][0] as Record<string, unknown>;
    expect(cell['@_gridSpan']).toBe(2);
  });

  it('serializes table cell with rowSpan > 1 adding @_rowSpan', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '5',
      name: 'Table 5',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(4000000), cy: emu(2000000) },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(4000000)],
        rows: [
          {
            height: emu(1000000),
            cells: [{ rowSpan: 2, textBody: { bodyProperties: {}, paragraphs: [] } }],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement);
    const tbl = (xmlObject['a:graphic'] as Record<string, Record<string, unknown>>)['a:graphicData']['a:tbl'] as Record<string, unknown[]>;
    const firstRow = tbl['a:tr'][0] as Record<string, unknown[]>;
    const cell = firstRow['a:tc'][0] as Record<string, unknown>;
    expect(cell['@_rowSpan']).toBe(2);
  });

  it('serializes a cell without textBody using the empty paragraph fallback', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '6',
      name: 'Table 6',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(2000000), cy: emu(1000000) },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(2000000)],
        rows: [{ height: emu(500000), cells: [{}] }],
      },
    };

    const xmlObject = serializeTable(tableElement);
    const tbl = (xmlObject['a:graphic'] as Record<string, Record<string, unknown>>)['a:graphicData']['a:tbl'] as Record<string, unknown[]>;
    const cell = (tbl['a:tr'][0] as Record<string, unknown[]>)['a:tc'][0] as Record<string, unknown>;
    const txBody = cell['a:txBody'] as Record<string, unknown>;
    expect(txBody).toHaveProperty('a:bodyPr');
    expect(txBody).toHaveProperty('a:lstStyle');
  });

  it('serializes cell properties: insets are added to a:tcPr', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '7',
      name: 'Table 7',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(2000000), cy: emu(1000000) },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(2000000)],
        rows: [
          {
            height: emu(500000),
            cells: [
              {
                properties: {
                  leftInset: emu(91440),
                  rightInset: emu(91440),
                  topInset: emu(45720),
                  bottomInset: emu(45720),
                },
              },
            ],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement);
    const tbl = (xmlObject['a:graphic'] as Record<string, Record<string, unknown>>)['a:graphicData']['a:tbl'] as Record<string, unknown[]>;
    const cell = (tbl['a:tr'][0] as Record<string, unknown[]>)['a:tc'][0] as Record<string, unknown>;
    const tcPr = cell['a:tcPr'] as Record<string, unknown>;
    expect(tcPr['@_marL']).toBe(91440);
    expect(tcPr['@_marR']).toBe(91440);
    expect(tcPr['@_marT']).toBe(45720);
    expect(tcPr['@_marB']).toBe(45720);
  });

  it('falls back to default gridCol when columnWidths is empty', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '8',
      name: 'Table 8',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(4000000), cy: emu(2000000) },
      rotation: emuDegree(0),
      table: { columnWidths: [], rows: [] },
    };

    const xmlObject = serializeTable(tableElement);
    const tbl = (xmlObject['a:graphic'] as Record<string, Record<string, unknown>>)['a:graphicData']['a:tbl'] as Record<string, unknown>;
    const tblGrid = tbl['a:tblGrid'] as Record<string, unknown[]>;
    expect(tblGrid['a:gridCol']).toHaveLength(1);
    expect((tblGrid['a:gridCol'][0] as Record<string, unknown>)['@_w']).toBe(2000000);
  });
});
