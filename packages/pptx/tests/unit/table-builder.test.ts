import { describe, expect, it } from 'vitest';
import { inches, points } from '@hokkyss/pptx-core';
import { TableBuilder } from '../../lib/builders/table-builder';
import { Presentation } from '../../lib/presentation';

describe('TableBuilder (Unit Tests)', () => {
  it('builds a table from simple 2D matrix data', () => {
    const tableEl = TableBuilder.fromMatrix(
      [
        ['Feature', 'SPA', 'RSC'],
        ['Bundle Size', 'Heavy', 'Zero Client'],
        ['Data Fetching', 'Client Waterfall', 'Server Direct'],
      ],
      {
        colWidths: [inches(3), inches(3), inches(3)],
        h: inches(2),
        header: {
          bold: true,
          fill: '1E293B',
          fontSize: points(14),
        },
        w: inches(9),
        x: inches(1),
        y: inches(1.5),
      },
    );

    expect(tableEl.elementType).toBe('table');
    expect(tableEl.table.rows.length).toBe(3);
    expect(tableEl.table.columnWidths.length).toBe(3);
    expect(tableEl.table.columnWidths[0]).toBe(2743200); // 3 inches in EMU

    // Verify header cell
    const headerCell = tableEl.table.rows[0].cells[0];
    expect(headerCell.textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('Feature');
    expect(headerCell.textBody?.paragraphs?.[0]?.runs?.[0]?.properties?.bold).toBe(true);
    expect(headerCell.properties?.fill?.type).toBe('solid');

    // Verify body cell
    const bodyCell = tableEl.table.rows[1].cells[1];
    expect(bodyCell.textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('Heavy');
  });

  it('builds a table using fluent method chaining', () => {
    const builder = new TableBuilder({
      colWidths: [inches(2.5), inches(4.5)],
      h: inches(2),
      w: inches(7),
      x: inches(1),
      y: inches(1),
    });

    builder
      .addRow({ h: inches(0.5) })
      .addCell({ bold: true, fill: '3B82F6', text: 'Metric' })
      .addCell({ bold: true, fill: '3B82F6', text: 'Description' })
      .addRow({ h: inches(0.5) })
      .addCell('FCP')
      .addCell('First Contentful Paint speed')
      .addRow({ h: inches(0.5) })
      .addCell('TTI')
      .addCell('Time to Interactive');

    const tableEl = builder.build();
    expect(tableEl.table.rows.length).toBe(3);
    expect(tableEl.table.rows[0].cells.length).toBe(2);
    expect(tableEl.table.rows[0].cells[0].textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('Metric');
    expect(tableEl.table.rows[1].cells[0].textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('FCP');
  });

  it('can be added directly to slide via slide.addTable()', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addTable(
      [
        ['A', 'B'],
        ['1', '2'],
      ],
      {
        h: inches(1.5),
        w: inches(4),
        x: inches(1),
        y: inches(1),
      },
    );

    expect(slide.getElements().length).toBe(1);
    expect(slide.getElements()[0].elementType).toBe('table');
  });
});

describe('TableBuilder extended configurations', () => {
  it('supports columns option array, setColWidths, addRow with cells array, and direct addCell', () => {
    const builder = new TableBuilder({
      columns: [{ w: inches(2) }, inches(3)],
    });

    builder.setColWidths([inches(2.5), inches(3.5)]);

    // Call addCell directly without addRow
    builder.addCell('Auto-row Cell');

    // addRow with cells array
    builder.addRow({
      cells: [
        'Cell 1',
        { text: 'Cell 2', bold: true, fill: 'EEEEEE', verticalAlign: 'top' },
      ],
      h: inches(0.5),
    });

    const el = builder.build();
    expect(el.table.rows).toHaveLength(2);
    expect(el.table.columnWidths).toEqual([2286000, 3200400]);
  });
});

describe('TableBuilder.fromMatrix with CellConfig objects', () => {
  it('supports object cells in header and body rows', () => {
    const table = TableBuilder.fromMatrix(
      [
        [{ text: 'Header Cell', fill: '003366', bold: true }],
        [{ text: 'Body Cell', fill: 'FFFFFF', italic: true }],
      ],
      {
        header: { color: 'FFFFFF' },
      },
    );

    expect(table.table.rows).toHaveLength(2);
    expect(table.table.rows[0].cells[0].properties?.fill?.solidColor?.value).toBe('003366');
    expect(table.table.rows[1].cells[0].properties?.fill?.solidColor?.value).toBe('FFFFFF');
  });

  it('handles 0 and empty column widths fallback', () => {
    const builder = new TableBuilder({ colWidths: [inches(0), inches(0)] });
    builder.addRow(['A', 'B']);
    const table = builder.build();
    expect(table.table.columnWidths).toHaveLength(2);
  });
});
