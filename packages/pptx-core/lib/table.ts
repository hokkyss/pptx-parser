import type { PptxFill } from './color';
import type { PptxTextBody } from './text';
import type { Emu } from './units';

/** Strongly-typed table cell properties */
export interface PptxTableCellProperties {
  /** Bottom padding inset in EMU. OpenXML: `<a:tc @_marB>` */
  bottomInset?: Emu;
  /** Cell background fill. OpenXML: `<a:tcPr><a:solidFill>` */
  fill?: PptxFill;
  /** Left padding inset in EMU. OpenXML: `<a:tc @_marL>` */
  leftInset?: Emu;
  /** Right padding inset in EMU. OpenXML: `<a:tc @_marR>` */
  rightInset?: Emu;
  /** Top padding inset in EMU. OpenXML: `<a:tc @_marT>` */
  topInset?: Emu;
  /** Cell vertical alignment ('top' | 'middle' | 'bottom'). OpenXML: `<a:tc @_anchor>` */
  verticalAlignment?: 'bottom' | 'middle' | 'top';
}

/** Represents a cell in a table */
export interface PptxTableCell {
  /** Column span. OpenXML: `<a:tc @_gridSpan>` */
  colSpan?: number;
  /** Cell properties. OpenXML: `<a:tcPr>` & attributes */
  properties?: PptxTableCellProperties;
  /** Row span. OpenXML: `<a:tc @_rowSpan>` */
  rowSpan?: number;
  /** Text content of the cell. OpenXML: `<a:tc><a:txBody>` */
  textBody?: PptxTextBody;
}

/** Represents a row in a table */
export interface PptxTableRow {
  /** Cells in the row. OpenXML: `<a:tc>` */
  cells: PptxTableCell[];
  /** Row height in EMU. OpenXML: `<a:tr @_h>` */
  height: Emu;
}

/** Represents a table */
export interface PptxTable {
  /** Column widths in EMU. OpenXML: `<a:tblGrid><a:gridCol @_w>` */
  columnWidths: Emu[];
  /** Table rows. OpenXML: `<a:tr>` */
  rows: PptxTableRow[];
}
