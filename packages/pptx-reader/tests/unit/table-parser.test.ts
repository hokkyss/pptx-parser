import { describe, expect, it } from 'vitest';
import { parseTable } from '../../lib/parsers/table-parser';

describe('Table Parser (@hokkyss/pptx-reader)', () => {
  it('returns undefined for null or undefined input', () => {
    expect(parseTable(undefined as unknown as string)).toBeUndefined();
    expect(parseTable('' as unknown as string)).toBeUndefined();
  });

  it('parses a table from raw XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:tblGrid>
    <a:gridCol w="2000000"/>
    <a:gridCol w="3000000"/>
  </a:tblGrid>
  <a:tr h="1000000">
    <a:tc marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr">
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Cell 1</a:t></a:r></a:p>
      </a:txBody>
    </a:tc>
    <a:tc rowSpan="2" gridSpan="2" anchor="top">
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Cell 2 (Merged)</a:t></a:r></a:p>
      </a:txBody>
    </a:tc>
  </a:tr>
</a:tbl>`;

    const result = parseTable(xml);
    expect(result).toBeDefined();
    expect(result!.columnWidths).toEqual([2000000, 3000000]);
    expect(result!.rows).toHaveLength(1);
    expect(result!.rows[0].height).toBe(1000000);
    expect(result!.rows[0].cells).toHaveLength(2);

    const cell1 = result!.rows[0].cells[0];
    expect(cell1.properties?.verticalAlignment).toBe('middle');
    expect(cell1.properties?.leftInset).toBe(91440);
    expect(cell1.properties?.rightInset).toBe(91440);
    expect(cell1.properties?.topInset).toBe(45720);
    expect(cell1.properties?.bottomInset).toBe(45720);

    const cell2 = result!.rows[0].cells[1];
    expect(cell2.rowSpan).toBe(2);
    expect(cell2.colSpan).toBe(2);
    expect(cell2.properties?.verticalAlignment).toBe('top');
  });

  it('finds nested a:tbl within a graphicFrame object node (findTblNode)', () => {
    const graphicFrame = {
      'p:graphicFrame': {
        'a:graphic': {
          'a:graphicData': {
            'a:tbl': {
              'a:tblGrid': {
                'a:gridCol': { '@_w': '1500000' },
              },
              'a:tr': {
                '@_h': '500000',
                'a:tc': {
                  '@_anchor': 'b',
                  'a:txBody': {
                    'a:p': { 'a:r': { 'a:t': 'Bottom aligned' } },
                  },
                },
              },
            },
          },
        },
      },
    };

    const result = parseTable(graphicFrame);
    expect(result).toBeDefined();
    expect(result!.columnWidths).toEqual([1500000]);
    expect(result!.rows).toHaveLength(1);
    expect(result!.rows[0].cells[0].properties?.verticalAlignment).toBe('bottom');
  });

  it('maps all vertical alignment anchor variants', () => {
    const anchors: Array<[string, 'bottom' | 'middle' | 'top' | undefined]> = [
      ['t', 'top'],
      ['top', 'top'],
      ['ctr', 'middle'],
      ['center', 'middle'],
      ['b', 'bottom'],
      ['bottom', 'bottom'],
      ['unknown', undefined],
    ];

    for (const [anchor, expected] of anchors) {
      const node = {
        'a:tbl': {
          'a:tr': {
            'a:tc': { '@_anchor': anchor },
          },
        },
      };
      const result = parseTable(node);
      expect(result!.rows[0].cells[0].properties?.verticalAlignment).toBe(expected);
    }
  });

  it('handles missing gridCol, tr, or tc gracefully', () => {
    const emptyTbl = { 'a:tbl': {} };
    const result = parseTable(emptyTbl);
    expect(result).toBeDefined();
    expect(result!.columnWidths).toEqual([]);
    expect(result!.rows).toEqual([]);
  });
});
