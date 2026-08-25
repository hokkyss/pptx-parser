import { describe, expect, it } from 'vitest';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeTable } from '../../lib/serializers/table-serializer';

interface TableXmlCellProperties {
  '@_marB'?: number;
  '@_marL'?: number;
  '@_marR'?: number;
  '@_marT'?: number;
  'a:solidFill'?: {
    'a:srgbClr'?: {
      '@_val'?: string;
    };
  };
}

interface TableXmlCell {
  '@_gridSpan'?: number;
  '@_rowSpan'?: number;
  'a:tcPr'?: TableXmlCellProperties;
  'a:txBody'?: {
    'a:bodyPr'?: Record<string, boolean | number | string>;
    'a:lstStyle'?: Record<string, boolean | number | string>;
    'a:p'?: Array<Record<string, boolean | number | string>>;
  };
}

interface TableXmlRow {
  'a:tc'?: TableXmlCell[];
}

interface TableXmlGridCol {
  '@_w'?: number;
}

interface TableXml {
  'a:tblGrid'?: {
    'a:gridCol'?: TableXmlGridCol[];
  };
  'a:tr'?: TableXmlRow[];
}

interface SerializedTableGraphicFrame {
  'a:graphic'?: {
    'a:graphicData'?: {
      'a:tbl'?: TableXml;
    };
  };
  'p:nvGraphicFramePr'?: {
    'p:cNvPr'?: {
      '@_id'?: string;
      '@_name'?: string;
    };
  };
}

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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    expect(xmlObject).toBeDefined();

    const nvGFPr = xmlObject['p:nvGraphicFramePr'];
    expect(nvGFPr?.['p:cNvPr']?.['@_id']).toBe('3');

    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    expect(tbl).toBeDefined();

    const tblGrid = tbl?.['a:tblGrid'];
    expect(tblGrid?.['a:gridCol']).toHaveLength(2);
    expect(tbl?.['a:tr']).toHaveLength(2);

    const firstRow = tbl?.['a:tr']?.[0];
    expect(firstRow?.['a:tc']).toHaveLength(2);
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const firstRow = tbl?.['a:tr']?.[0];
    const cell = firstRow?.['a:tc']?.[0];
    expect(cell?.['@_gridSpan']).toBe(2);
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const firstRow = tbl?.['a:tr']?.[0];
    const cell = firstRow?.['a:tc']?.[0];
    expect(cell?.['@_rowSpan']).toBe(2);
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const cell = tbl?.['a:tr']?.[0]?.['a:tc']?.[0];
    const txBody = cell?.['a:txBody'];
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const cell = tbl?.['a:tr']?.[0]?.['a:tc']?.[0];
    const tcPr = cell?.['a:tcPr'];
    expect(tcPr?.['@_marL']).toBe(91440);
    expect(tcPr?.['@_marR']).toBe(91440);
    expect(tcPr?.['@_marT']).toBe(45720);
    expect(tcPr?.['@_marB']).toBe(45720);
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const tblGrid = tbl?.['a:tblGrid'];
    expect(tblGrid?.['a:gridCol']).toHaveLength(1);
    expect(tblGrid?.['a:gridCol']?.[0]?.['@_w']).toBe(2000000);
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

    const xmlObject = serializeTable(tableElement) as SerializedTableGraphicFrame;
    const graphic = xmlObject['a:graphic'];
    const graphicData = graphic?.['a:graphicData'];
    const tbl = graphicData?.['a:tbl'];
    const cell = tbl?.['a:tr']?.[0]?.['a:tc']?.[0];
    const tcPr = cell?.['a:tcPr'];
    expect(tcPr).toHaveProperty('a:solidFill');
  });

  it('covers table cell alignment, borders and vertical alignment', () => {
    const tblEl: PptxTableElement = {
      elementType: 'table',
      id: '20',
      isVisible: true,
      name: 'Table 20',
      position: { cx: emu(1000), cy: emu(1000), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      table: {
        columnWidths: [emu(500), emu(500)],
        rows: [
          {
            cells: [
              {
                properties: {
                  verticalAlignment: 'middle',
                },
              },
            ],
            height: emu(500),
          },
        ],
      },
      type: 'graphicFrame',
      zIndex: 0,
    };
    const tblXml = serializeTable(tblEl);
    expect(tblXml).toBeDefined();
  });

  it('covers table spanning, empty textBody, insets, fill, empty columns, and position fallbacks', () => {
    const fallbackTable: PptxTableElement = {
      elementType: 'table',
      id: '',
      isVisible: true,
      name: '',
      rotation: emuDegree(0),
      table: {
        columnWidths: [],
        columns: [{ width: emu(1500000) }, {}],
        rows: [
          {
            cells: [
              {
                colSpan: 2,
                rowSpan: 3,
                properties: {
                  bottomInset: emu(10000),
                  fill: { solidColor: { type: 'srgb', value: '10B981' }, type: 'solid' },
                  leftInset: emu(20000),
                  rightInset: emu(30000),
                  topInset: emu(40000),
                },
              },
              {
                // @ts-expect-error Testing legacy gridSpan property
                gridSpan: 2,
              },
            ],
            // @ts-expect-error Testing undefined height fallback
            height: undefined,
          },
          {
            // @ts-expect-error Testing undefined cells fallback
            cells: undefined,
            height: emu(300000),
          },
        ],
      },
      type: 'graphicFrame',
      zIndex: 0,
    };

    const xml = serializeTable(fallbackTable);
    expect(xml).toBeDefined();
    const nvPr = xml['p:nvGraphicFramePr'] as Record<string, Record<string, string>>;
    expect(nvPr['p:cNvPr']['@_id']).toBe('3');
    expect(nvPr['p:cNvPr']['@_name']).toBe('Table 3');

    const xfrm = xml['p:xfrm'] as Record<string, Record<string, number>>;
    expect(xfrm['a:off']['@_x']).toBe(0);
    expect(xfrm['a:ext']['@_cx']).toBe(4000000);

    // Empty gridCols fallback
    const emptyGridTable: PptxTableElement = {
      elementType: 'table',
      id: '5',
      isVisible: true,
      name: 'T5',
      position: { cx: emu(100), cy: emu(100), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      table: { columnWidths: [], rows: [] },
      type: 'graphicFrame',
      zIndex: 0,
    };
    const emptyGridXml = serializeTable(emptyGridTable);
    const tbl = (emptyGridXml['a:graphic'] as Record<string, Record<string, Record<string, { 'a:gridCol': Record<string, number>[] }>>>)['a:graphicData']['a:tbl'];
    expect(tbl['a:tblGrid']['a:gridCol']).toHaveLength(1);
  });
});
