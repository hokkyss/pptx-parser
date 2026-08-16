import type { PptxTableElement } from '@hokkyss/pptx-core';
import type { PptxTableCell, PptxTableRow } from '@hokkyss/pptx-core';
import type { Emu } from '@hokkyss/pptx-core';
import { serializeFill, serializeTextBody } from './text-serializer';

/**
 * Serializes a table cell `<a:tc>`.
 * Follows schema order: a:txBody -> a:tcPr
 */
export function serializeTableCell(cell: PptxTableCell): Record<string, unknown> {
  const tc: Record<string, unknown> = {};

  const colSpan = cell.colSpan || (cell as unknown as { gridSpan?: number }).gridSpan;
  if (colSpan && colSpan > 1) {
    tc['@_gridSpan'] = colSpan;
  }
  if (cell.rowSpan && cell.rowSpan > 1) {
    tc['@_rowSpan'] = cell.rowSpan;
  }

  // 1. Text Body (MUST BE FIRST)
  if (cell.textBody) {
    tc['a:txBody'] = serializeTextBody(cell.textBody);
  } else {
    tc['a:txBody'] = {
      'a:bodyPr': {},
      'a:lstStyle': {},
      'a:p': [{ 'a:pPr': {}, 'a:endParaRPr': {} }],
    };
  }

  // 2. Cell Properties (MUST BE SECOND)
  const props = cell.properties;
  const tcPr: Record<string, unknown> = {};

  const lIns = props?.leftInset;
  const rIns = props?.rightInset;
  const tIns = props?.topInset;
  const bIns = props?.bottomInset;

  if (lIns !== undefined) tcPr['@_marL'] = Math.round(Number(lIns));
  if (rIns !== undefined) tcPr['@_marR'] = Math.round(Number(rIns));
  if (tIns !== undefined) tcPr['@_marT'] = Math.round(Number(tIns));
  if (bIns !== undefined) tcPr['@_marB'] = Math.round(Number(bIns));

  if (props?.fill) {
    const fillNode = serializeFill(props.fill);
    if (fillNode) Object.assign(tcPr, fillNode);
  }

  tc['a:tcPr'] = tcPr;

  return tc;
}

/**
 * Serializes a table row `<a:tr>`.
 */
export function serializeTableRow(row: PptxTableRow): Record<string, unknown> {
  return {
    '@_h': Math.round(Number(row.height ?? 500000)),
    'a:tc': (row.cells || []).map(serializeTableCell),
  };
}

/**
 * Serializes a table element into an OpenXML `<p:graphicFrame>` table.
 * Follows schema order: p:nvGraphicFramePr -> p:xfrm -> a:graphic
 */
export function serializeTable(tableElement: PptxTableElement): Record<string, unknown> {
  const table = tableElement.table;

  const rawColumns = (table as unknown as { columns?: Array<{ width?: Emu }> }).columns;
  const columnWidths = table.columnWidths || (rawColumns ? rawColumns.map((c) => c.width ?? ((1000000 as unknown) as Emu)) : []);

  const gridCols = columnWidths.map((w) => ({
    '@_w': Math.round(Number(w ?? 1000000)),
  }));

  const rows = (table.rows || []).map(serializeTableRow);

  return {
    'p:nvGraphicFramePr': {
      'p:cNvPr': {
        '@_id': tableElement.id || '3',
        '@_name': tableElement.name || `Table ${tableElement.id || '3'}`,
      },
      'p:cNvGraphicFramePr': {
        'a:graphicFrameLocks': { '@_noGrp': '1' },
      },
      'p:nvPr': {},
    },
    'p:xfrm': {
      'a:off': {
        '@_x': Math.round(Number(tableElement.position?.x ?? 0)),
        '@_y': Math.round(Number(tableElement.position?.y ?? 0)),
      },
      'a:ext': {
        '@_cx': Math.round(Number(tableElement.position?.cx ?? 4000000)),
        '@_cy': Math.round(Number(tableElement.position?.cy ?? 2000000)),
      },
    },
    'a:graphic': {
      'a:graphicData': {
        '@_uri': 'http://schemas.openxmlformats.org/drawingml/2006/table',
        'a:tbl': {
          'a:tblPr': {},
          'a:tblGrid': {
            'a:gridCol': gridCols.length > 0 ? gridCols : [{ '@_w': 2000000 }],
          },
          'a:tr': rows,
        },
      },
    },
  };
}
