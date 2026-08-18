import type { PptxFill } from '@hokkyss/pptx-core';
import type { PptxTableCell, PptxTableRow } from '@hokkyss/pptx-core';
import type { PptxTableElement } from '@hokkyss/pptx-core';
import {
  emu,
  emuDegree,
  type Emu,
  type Inches,
  inchesToEmu,
  type Points,
} from '@hokkyss/pptx-core';
import { normalizeFill } from './shape-builder';
import {
  buildTextBody,
  type ParagraphConfig,
  type TextRunConfig,
} from './text-builder';

export interface CellConfig {
  align?: 'center' | 'left' | 'right';
  bold?: boolean;
  color?: string;
  colSpan?: number;
  fill?: PptxFill | string;
  font?: string;
  fontSize?: Points;
  italic?: boolean;
  rowSpan?: number;
  text: ParagraphConfig[] | string | TextRunConfig[];
  verticalAlign?: 'bottom' | 'middle' | 'top';
}

export interface RowConfig {
  cells?: (CellConfig | string)[];
  h?: Inches;
}

export interface AddTableOptions {
  colWidths?: Inches[];
  columns?: Array<{ w?: Inches } | Inches>;
  h?: Inches;
  header?: Partial<CellConfig>;
  id?: string;
  name?: string;
  placeholder?: number | string;
  w?: Inches;
  x?: Inches;
  y?: Inches;
  zIndex?: number;
}

export type TableMatrix = (CellConfig | string)[][];

/**
 * Fluent builder for creating highly-styled OpenXML tables.
 */
export class TableBuilder {
  private colWidths: Inches[] = [];
  private options: AddTableOptions;
  private rows: PptxTableRow[] = [];
  private currentRow: null | PptxTableRow = null;

  constructor(options: AddTableOptions = {}) {
    this.options = options;
    if (options.colWidths) {
      this.colWidths = options.colWidths;
    } else if (options.columns) {
      this.colWidths = options.columns.map((c) =>
        typeof c === 'object' && c !== null && 'w' in c ? c.w! : (c as Inches),
      );
    }
  }

  /**
   * Sets explicit column widths for the table.
   */
  setColWidths(widths: Inches[]): this {
    this.colWidths = widths;
    return this;
  }

  /**
   * Appends a new table row.
   */
  addRow(rowConfig?: RowConfig): this {
    const heightEmu = rowConfig?.h
      ? inchesToEmu(rowConfig.h)
      : emu(360000); // ~0.4 inch default row height

    const row: PptxTableRow = {
      cells: [],
      height: heightEmu,
    };
    this.rows.push(row);
    this.currentRow = row;

    if (rowConfig?.cells && Array.isArray(rowConfig.cells)) {
      for (const cell of rowConfig.cells) {
        this.addCell(cell);
      }
    }

    return this;
  }

  /**
   * Appends a cell to the current table row.
   */
  addCell(cellInput: CellConfig | string): this {
    if (!this.currentRow) {
      this.addRow();
    }

    const config: CellConfig = typeof cellInput === 'string'
      ? { text: cellInput }
      : cellInput;

    const fill = normalizeFill(config.fill);
    const textBody = buildTextBody(config.text, {
      align: config.align,
      bold: config.bold,
      color: config.color,
      font: config.font,
      fontSize: config.fontSize,
      italic: config.italic,
      verticalAlignment: config.verticalAlign || 'middle',
    });

    const tableCell: PptxTableCell = {
      colSpan: config.colSpan,
      properties: fill || config.verticalAlign
        ? {
            fill,
            verticalAlignment: config.verticalAlign || 'middle',
          }
        : undefined,
      rowSpan: config.rowSpan,
      textBody,
    };

    this.currentRow!.cells.push(tableCell);
    return this;
  }

  /**
   * Compiles the table configuration into a strongly-typed `PptxTableElement` AST node.
   */
  build(counter: number | string = 1): PptxTableElement {
    const totalColWidthsInches = this.colWidths.reduce((acc, w) => acc + (w || 0), 0);
    const widthEmu = this.options.w
      ? inchesToEmu(this.options.w)
      : totalColWidthsInches > 0
        ? inchesToEmu(totalColWidthsInches)
        : emu(9144000);
    const heightEmu = this.options.h ? inchesToEmu(this.options.h) : emu(2743200);
    const xEmu = this.options.x ? inchesToEmu(this.options.x) : emu(914400);
    const yEmu = this.options.y ? inchesToEmu(this.options.y) : emu(1828800);

    const totalCols = this.rows.reduce(
      (max, r) => Math.max(max, r.cells.length),
      this.colWidths.length,
    );

    let columnWidthsEmu: Emu[] = [];
    if (this.colWidths.length > 0) {
      columnWidthsEmu = this.colWidths.map((w) => inchesToEmu(w));
    } else if (totalCols > 0) {
      const autoWidth = widthEmu / totalCols;
      columnWidthsEmu = Array.from({ length: totalCols }, () => emu(Math.round(autoWidth)));
    }

    const id = this.options.id || String(counter);
    const name = this.options.name || `Table ${id}`;

    return {
      elementType: 'table',
      id,
      isVisible: true,
      name,
      position: {
        cx: widthEmu,
        cy: heightEmu,
        x: xEmu,
        y: yEmu,
      },
      rotation: emuDegree(0),
      table: {
        columnWidths: columnWidthsEmu,
        rows: this.rows,
      },
      type: 'graphicFrame',
      zIndex: this.options.zIndex ?? 0,
    };
  }

  /**
   * Static helper constructing a `PptxTableElement` directly from a 2D matrix.
   */
  static fromMatrix(matrix: TableMatrix, options: AddTableOptions = {}, counter: number | string = 1): PptxTableElement {
    const builder = new TableBuilder(options);

    for (let rIdx = 0; rIdx < matrix.length; rIdx++) {
      const row = matrix[rIdx];
      const isHeaderRow = rIdx === 0 && options.header;

      builder.addRow();
      for (const cell of row) {
        if (typeof cell === 'string') {
          if (isHeaderRow) {
            builder.addCell({
              ...options.header,
              text: cell,
            });
          } else {
            builder.addCell(cell);
          }
        } else {
          if (isHeaderRow) {
            builder.addCell({
              ...options.header,
              ...cell,
            });
          } else {
            builder.addCell(cell);
          }
        }
      }
    }

    return builder.build(counter);
  }
}
