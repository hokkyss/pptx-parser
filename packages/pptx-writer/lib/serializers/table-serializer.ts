import { emu } from '@hokkyss/pptx-core';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import type { PptxTableCell, PptxTableRow } from '@hokkyss/pptx-core';
import type { Emu } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';
import { serializeFill, serializeTextBody } from './text-serializer';

/**
 * Serializes a table cell `<a:tc>`.
 * Follows schema order: a:txBody -> a:tcPr
 */
export function serializeTableCell(cell: PptxTableCell): XmlElement {
  const attrs: Record<string, number | undefined> = {};

  const colSpan = cell.colSpan || ('gridSpan' in cell ? (cell as { gridSpan?: number }).gridSpan : undefined);
  if (colSpan && colSpan > 1) {
    attrs.gridSpan = colSpan;
  }
  if (cell.rowSpan && cell.rowSpan > 1) {
    attrs.rowSpan = cell.rowSpan;
  }

  // 1. Text Body (MUST BE FIRST)
  const txBody = cell.textBody
    ? serializeTextBody(cell.textBody, 'a:txBody')
    : el('a:txBody', [
        el('a:bodyPr'),
        el('a:lstStyle'),
        el('a:p', [el('a:pPr'), el('a:endParaRPr')]),
      ]);

  // 2. Cell Properties (MUST BE SECOND)
  const props = cell.properties;
  const CELL_VERTICAL_ALIGNMENT_MAP: Record<string, string> = {
    bottom: 'b',
    middle: 'ctr',
    top: 't',
  };
  const tcPrAttrs: Record<string, number | string | undefined> = {};

  if (props?.verticalAlignment) {
    tcPrAttrs.anchor = CELL_VERTICAL_ALIGNMENT_MAP[props.verticalAlignment] ?? props.verticalAlignment;
  }

  const lIns = props?.leftInset;
  const rIns = props?.rightInset;
  const tIns = props?.topInset;
  const bIns = props?.bottomInset;

  if (lIns !== undefined) tcPrAttrs.marL = Math.round(Number(lIns));
  if (rIns !== undefined) tcPrAttrs.marR = Math.round(Number(rIns));
  if (tIns !== undefined) tcPrAttrs.marT = Math.round(Number(tIns));
  if (bIns !== undefined) tcPrAttrs.marB = Math.round(Number(bIns));

  const tcPrChildren: XmlElement[] = [];

  if (props?.fill) {
    const fillNode = serializeFill(props.fill);
    if (fillNode) tcPrChildren.push(fillNode);
  }

  const tcPr = el('a:tcPr', tcPrAttrs, tcPrChildren);

  return el('a:tc', attrs, [txBody, tcPr]);
}

/**
 * Serializes a table row `<a:tr>`.
 */
export function serializeTableRow(row: PptxTableRow): XmlElement {
  const cells = (row.cells || []).map(serializeTableCell);
  return el('a:tr', { h: Math.round(Number(row.height ?? 500000)) }, cells);
}

/**
 * Serializes a table element into an OpenXML `<p:graphicFrame>` table.
 * Follows schema order: p:nvGraphicFramePr -> p:xfrm -> a:graphic
 */
export function serializeTable(tableElement: PptxTableElement): XmlElement {
  const table = tableElement.table;

  const rawColumns = 'columns' in table ? (table as { columns?: Array<{ width?: Emu }> }).columns : undefined;
  const columnWidths = table.columnWidths || (rawColumns ? rawColumns.map((c) => c.width ?? emu(1000000)) : []);

  const gridCols = (columnWidths.length > 0 ? columnWidths : [emu(2000000)]).map((w) =>
    el('a:gridCol', { w: Math.round(Number(w ?? 1000000)) }),
  );

  const rows = (table.rows || []).map(serializeTableRow);

  return el('p:graphicFrame', [
    el('p:nvGraphicFramePr', [
      el('p:cNvPr', {
        id: tableElement.id || '3',
        name: tableElement.name || `Table ${tableElement.id || '3'}`,
      }),
      el('p:cNvGraphicFramePr', [
        el('a:graphicFrameLocks', { noGrp: '1' }),
      ]),
      el('p:nvPr'),
    ]),
    el('p:xfrm', [
      el('a:off', {
        x: Math.round(Number(tableElement.position?.x ?? 0)),
        y: Math.round(Number(tableElement.position?.y ?? 0)),
      }),
      el('a:ext', {
        cx: Math.round(Number(tableElement.position?.cx ?? 4000000)),
        cy: Math.round(Number(tableElement.position?.cy ?? 2000000)),
      }),
    ]),
    el('a:graphic', [
      el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' }, [
        el('a:tbl', [
          el('a:tblPr'),
          el('a:tblGrid', gridCols),
          ...rows,
        ]),
      ]),
    ]),
  ]);
}
