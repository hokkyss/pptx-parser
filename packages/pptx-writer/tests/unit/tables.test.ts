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
});
