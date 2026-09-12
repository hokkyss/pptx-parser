import { describe, expect, it } from 'vitest';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeTable } from '../../lib/serializers/table-serializer';
import { renderXml, type XmlElement } from '../../lib/xml/xml-element';

/** Finds all descendants (and self) matching the given tag. */
function findAllEl(node: XmlElement, tag: string): XmlElement[] {
  const results: XmlElement[] = [];
  if (node.tag === tag) results.push(node);
  for (const child of node.children ?? []) {
    if (typeof child === 'object' && child !== null && 'tag' in child) {
      results.push(...findAllEl(child, tag));
    }
  }
  return results;
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

    const xmlNode = serializeTable(tableElement);
    expect(xmlNode).toBeDefined();
    const xml = renderXml(xmlNode);

    // nvGraphicFramePr
    expect(xml).toContain('id="3"');
    // graphic structure
    expect(xml).toContain('<a:tbl>');
    // 2 gridCol columns
    const gridCols = findAllEl(xmlNode, 'a:gridCol');
    expect(gridCols).toHaveLength(2);
    // 2 rows
    const rows = findAllEl(xmlNode, 'a:tr');
    expect(rows).toHaveLength(2);
    // First row has 2 cells
    const firstRowCells = findAllEl(rows[0], 'a:tc');
    expect(firstRowCells).toHaveLength(2);
  });

  it('serializes table cell with colSpan > 1 adding gridSpan', () => {
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

    const xml = renderXml(serializeTable(tableElement));
    expect(xml).toContain('gridSpan="2"');
  });

  it('serializes table cell with rowSpan > 1 adding rowSpan', () => {
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

    const xml = renderXml(serializeTable(tableElement));
    expect(xml).toContain('rowSpan="2"');
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

    const xml = renderXml(serializeTable(tableElement));
    // txBody should contain bodyPr and lstStyle
    expect(xml).toContain('<a:bodyPr/>');
    expect(xml).toContain('<a:lstStyle/>');
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

    const xml = renderXml(serializeTable(tableElement));
    expect(xml).toContain('marL="91440"');
    expect(xml).toContain('marR="91440"');
    expect(xml).toContain('marT="45720"');
    expect(xml).toContain('marB="45720"');
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

    const xmlNode = serializeTable(tableElement);
    const gridCols = findAllEl(xmlNode, 'a:gridCol');
    expect(gridCols).toHaveLength(1);
    expect(gridCols[0].attrs?.w).toBe(2000000);
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

    const xml = renderXml(serializeTable(tableElement));
    expect(xml).toContain('<a:solidFill>');
    expect(xml).toContain('val="E2E8F0"');
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
    const xml = renderXml(tblXml);
    expect(xml).toContain('anchor="ctr"');
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

    const xml = renderXml(serializeTable(fallbackTable));
    expect(xml).toBeDefined();
    expect(xml).toContain('id="3"');
    expect(xml).toContain('name="Table 3"');
    expect(xml).toContain('x="0"');
    expect(xml).toContain('cx="4000000"');

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
    const emptyGridNode = serializeTable(emptyGridTable);
    const gridCols = findAllEl(emptyGridNode, 'a:gridCol');
    expect(gridCols).toHaveLength(1);
  });
});
