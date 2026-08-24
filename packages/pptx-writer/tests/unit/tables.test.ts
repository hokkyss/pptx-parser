import { describe, expect, it } from 'vitest';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeTable } from '../../lib/serializers/table-serializer';

type XmlNode = Record<string, XmlNode | XmlNode[] | string | number | boolean | undefined>;

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
                      runs: [{ text: 'Data 1', properties: { color: '003366' } }],
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
                      runs: [{ text: 'Data 2', properties: { color: '003366' } }],
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement) as XmlNode;
    expect(xmlObject).toBeDefined();

    const nvGFPr = xmlObject['p:nvGraphicFramePr'] as XmlNode;
    expect((nvGFPr['p:cNvPr'] as XmlNode)['@_id']).toBe('3');

    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    expect(tbl).toBeDefined();

    const tblGrid = tbl['a:tblGrid'] as XmlNode;
    expect(tblGrid['a:gridCol'] as XmlNode[]).toHaveLength(2);
    expect(tbl['a:tr'] as XmlNode[]).toHaveLength(2);

    const firstRow = (tbl['a:tr'] as XmlNode[])[0];
    expect(firstRow['a:tc'] as XmlNode[]).toHaveLength(2);
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

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const firstRow = (tbl['a:tr'] as XmlNode[])[0];
    const cell = (firstRow['a:tc'] as XmlNode[])[0];
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

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const firstRow = (tbl['a:tr'] as XmlNode[])[0];
    const cell = (firstRow['a:tc'] as XmlNode[])[0];
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

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const cell = ((tbl['a:tr'] as XmlNode[])[0]['a:tc'] as XmlNode[])[0];
    const txBody = cell['a:txBody'] as XmlNode;
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

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const cell = ((tbl['a:tr'] as XmlNode[])[0]['a:tc'] as XmlNode[])[0];
    const tcPr = cell['a:tcPr'] as XmlNode;
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

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const tblGrid = tbl['a:tblGrid'] as XmlNode;
    expect(tblGrid['a:gridCol'] as XmlNode[]).toHaveLength(1);
    expect(((tblGrid['a:gridCol'] as XmlNode[])[0])['@_w']).toBe(2000000);
  });
});

describe('Table Serializer cell fill styling', () => {
  it('serializes table cell fill into a:tcPr', () => {
    const tableElement: PptxTableElement = {
      elementType: 'table',
      type: 'graphicFrame',
      id: '9',
      name: 'Table 9',
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
                  fill: { type: 'solid', solidColor: { type: 'srgb', value: 'E2E8F0' } },
                },
              },
            ],
          },
        ],
      },
    };

    const xmlObject = serializeTable(tableElement) as XmlNode;
    const graphic = xmlObject['a:graphic'] as XmlNode;
    const graphicData = graphic['a:graphicData'] as XmlNode;
    const tbl = graphicData['a:tbl'] as XmlNode;
    const cell = ((tbl['a:tr'] as XmlNode[])[0]['a:tc'] as XmlNode[])[0];
    const tcPr = cell['a:tcPr'] as XmlNode;
    expect(tcPr).toHaveProperty('a:solidFill');
  });
});
