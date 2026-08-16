import { PptxTable, PptxTableRow, PptxTableCell, PptxTableCellProperties, Emu } from '../types/ast';
import { parseTextBody } from './text-parser';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

/**
 * Recursively searches an XML node tree for an OpenXML table node (`a:tbl`).
 * @param obj Root XML node to search.
 * @returns Table node object if found, undefined otherwise.
 */
function findTblNode(obj: unknown): Record<string, unknown> | undefined {
  if (!obj || typeof obj !== 'object') return undefined;
  const record = obj as Record<string, unknown>;
  if (record['a:tbl'] || record['tbl']) {
    return (record['a:tbl'] || record['tbl']) as Record<string, unknown>;
  }
  for (const val of Object.values(record)) {
    if (typeof val === 'object' && val !== null) {
      const found = findTblNode(val);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Parses OpenXML tables (`<a:tbl>`) into structured, strongly-typed `PptxTable` AST elements.
 *
 * Extracts column widths, row heights, cell text bodies, row/column spans, padding insets, and vertical alignment.
 * @param tblNode Raw XML string or object node representing a table (`<a:tbl>`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Parsed `PptxTable` object or `undefined` if node is invalid.
 * @example
 * ```ts
 * const table = parseTable(graphicFrameNode);
 * console.log(`Table grid columns: ${table?.columnWidths.length}, rows: ${table?.rows.length}`);
 * ```
 */
export function parseTable(tblNode: Record<string, unknown> | string, parser: XmlParser = defaultXmlParser): PptxTable | undefined {
  if (!tblNode) return undefined;

  let parsed = tblNode;
  if (typeof tblNode === 'string') {
    parsed = parser.parse<Record<string, unknown>>(tblNode);
  }

  const parsedObj = parsed as Record<string, unknown>;
  const tbl = (parsedObj['a:tbl'] || parsedObj['tbl'] || findTblNode(parsedObj) || parsedObj) as Record<string, unknown>;

  // Parse column widths (<a:tblGrid><a:gridCol w="..."/></a:tblGrid>)
  const tblGrid = (tbl['a:tblGrid'] || tbl['tblGrid'] || {}) as Record<string, unknown>;
  let gridCols = tblGrid['a:gridCol'] || tblGrid['gridCol'];
  const columnWidths: Emu[] = [];

  if (gridCols) {
    if (!Array.isArray(gridCols)) gridCols = [gridCols];
    for (const col of gridCols as Record<string, unknown>[]) {
      const w = col['@_w'] !== undefined ? ((Number(col['@_w']) as unknown) as Emu) : ((0 as unknown) as Emu);
      columnWidths.push(w);
    }
  }

  // Parse rows (<a:tr h="...">)
  let trNodes = tbl['a:tr'] || tbl['tr'];
  const rows: PptxTableRow[] = [];

  if (trNodes) {
    if (!Array.isArray(trNodes)) trNodes = [trNodes];
    for (const trNode of trNodes as Record<string, unknown>[]) {
      const height = (trNode['@_h'] !== undefined ? Number(trNode['@_h']) : 0) as unknown as Emu;

      let tcNodes = trNode['a:tc'] || trNode['tc'];
      const cells: PptxTableCell[] = [];

      if (tcNodes) {
        if (!Array.isArray(tcNodes)) tcNodes = [tcNodes];
        for (const tcNode of tcNodes as Record<string, unknown>[]) {
          const rowSpan = tcNode['@_rowSpan'] !== undefined ? Number(tcNode['@_rowSpan']) : 1;
          const colSpan = tcNode['@_gridSpan'] !== undefined ? Number(tcNode['@_gridSpan']) : 1;
          const textBody = parseTextBody(tcNode['a:txBody'] as Record<string, unknown>);

          const cellProperties = parseCellProperties(tcNode);

          cells.push({
            textBody,
            rowSpan,
            colSpan,
            properties: cellProperties,
          });
        }
      }

      rows.push({
        cells,
        height,
      });
    }
  }

  return {
    rows,
    columnWidths,
  };
}

/**
 * Helper converting raw `<a:tc>` XML attributes into clean camelCase `PptxTableCellProperties`.
 * @param tcNode Raw table cell XML node (`<a:tc>`).
 * @returns Strongly-typed camelCase cell properties (insets, vertical alignment, etc.).
 */
function parseCellProperties(tcNode: Record<string, unknown>): PptxTableCellProperties {
  const anchorMap: Record<string, 'bottom' | 'middle' | 'top'> = {
    t: 'top',
    top: 'top',
    ctr: 'middle',
    center: 'middle',
    b: 'bottom',
    bottom: 'bottom',
  };

  const anchorRaw = (tcNode['@_anchor'] as string) || '';
  const verticalAlignment = anchorMap[anchorRaw];

  const leftInset = tcNode['@_marL'] !== undefined ? ((Number(tcNode['@_marL']) as unknown) as Emu) : undefined;
  const topInset = tcNode['@_marT'] !== undefined ? ((Number(tcNode['@_marT']) as unknown) as Emu) : undefined;
  const rightInset = tcNode['@_marR'] !== undefined ? ((Number(tcNode['@_marR']) as unknown) as Emu) : undefined;
  const bottomInset = tcNode['@_marB'] !== undefined ? ((Number(tcNode['@_marB']) as unknown) as Emu) : undefined;

  return {
    verticalAlignment,
    leftInset,
    topInset,
    rightInset,
    bottomInset,
  };
}
