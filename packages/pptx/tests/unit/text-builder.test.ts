import { describe, expect, it } from 'vitest';
import { points } from '@hokkyss/pptx-core';
import { buildTextBody, buildTextRun, normalizeBullet } from '../../lib/builders/text-builder';

describe('TextBuilder Helper Functions (@hokkyss/pptx)', () => {
  it('normalizes bullet options (number, custom char, object)', () => {
    const bodyNum = buildTextBody('Numbered item', { bullet: 'number' });
    expect(bodyNum.paragraphs[0].properties.bullet).toEqual({
      autoNumType: 'arabicPeriod',
      type: 'autoNum',
    });

    const bodyChar = buildTextBody('Dash item', { bullet: '-' });
    expect(bodyChar.paragraphs[0].properties.bullet).toEqual({
      char: '-',
      type: 'char',
    });

    const bodyObj = buildTextBody('Custom bullet', { bullet: { char: '★', type: 'char' } });
    expect(bodyObj.paragraphs[0].properties.bullet).toEqual({
      char: '★',
      type: 'char',
    });
  });

  it('normalizes text runs with strikethrough, subscript, superscript, baseline', () => {
    const run1 = buildTextRun('Subscript', { baseline: -25000, strikethrough: true, subscript: true });
    expect(run1.properties.subscript).toBe(true);
    expect(run1.properties.strikethrough).toBe(true);
    expect(run1.properties.baseline).toBe(-25000);

    const run2 = buildTextRun({
      baseline: 30000,
      superscript: true,
      text: 'Super',
    });
    expect(run2.properties.superscript).toBe(true);
    expect(run2.properties.baseline).toBe(30000);
  });
});

describe('buildTextBody paragraph arrays and multiline runs', () => {
  it('handles mixed paragraph config array with strings and run arrays', () => {
    const body = buildTextBody([
      'String Paragraph',
      { level: 1, spaceAfter: points(6), spaceBefore: points(12), text: ['Run A', 'Run B'] },
    ]);
    expect(body.paragraphs).toHaveLength(2);
    expect(body.paragraphs[0].runs[0].text).toBe('String Paragraph');
    expect(body.paragraphs[1].runs).toHaveLength(2);
    expect(body.paragraphs[1].properties.level).toBe(1);
  });

  it('handles multi-line text runs creating new paragraphs', () => {
    const body = buildTextBody([
      { bold: true, text: 'First Line\nSecond Line' },
    ]);
    expect(body.paragraphs).toHaveLength(2);
    expect(body.paragraphs[0].runs[0].text).toBe('First Line');
    expect(body.paragraphs[1].runs[0].text).toBe('Second Line');
  });
});

describe('buildTextBody nested text run config', () => {
  it('handles item with align and text inside paragraph config list', () => {
    const body = buildTextBody([
      { align: 'center', text: 'Centered Paragraph' },
      { level: 1, text: { italic: true, text: 'Single TextRunConfig Object' } },
    ]);
    expect(body.paragraphs[0].runs[0].text).toBe('Centered Paragraph');
    expect(body.paragraphs[0].properties.alignment).toBe('center');
    expect(body.paragraphs[1].runs[0].text).toBe('Single TextRunConfig Object');
    expect(body.paragraphs[1].runs[0].properties.italic).toBe(true);
  });
});

describe('normalizeBullet invalid input fallback', () => {
  it('returns undefined for invalid numeric bullet type', () => {
    // @ts-expect-error testing runtime fallback
    expect(normalizeBullet(123)).toBeUndefined();
  });

  it('covers default font size, spacing, array of mixed configs and multiline runs', () => {
    const run1 = buildTextRun({ text: 'Hello' }, { fontSize: points(16) });
    expect(run1.properties.fontSize).toBe(1600);

    const run2 = buildTextRun({ text: 'No size' }, {});
    expect(run2.properties?.fontSize).toBeUndefined();

    const body1 = buildTextBody('Line 1\nLine 2', { spaceAfter: points(10), spaceBefore: points(5) });
    expect(body1.paragraphs[0].properties.spaceAfter).toBe(1000);
    expect(body1.paragraphs[0].properties.spaceBefore).toBe(500);

    const body2 = buildTextBody([
      { bold: true, text: 'First paragraph\n\nSecond paragraph' },
    ]);
    expect(body2.paragraphs.length).toBeGreaterThanOrEqual(2);

    const body3 = buildTextBody([{ text: 'Multi\nLine' }, { text: 'Single' }]);
    expect(body3.paragraphs).toHaveLength(2);

    // ParagraphConfig with explicit runs array and options merging
    const bodyWithRuns = buildTextBody([
      {
        align: 'right',
        bullet: true,
        runs: [{ bold: true, text: 'R1' }, { italic: true, text: 'R2' }],
      },
      'Simple paragraph string',
    ], { align: 'center', bullet: 'number', level: 2 });
    expect(bodyWithRuns.paragraphs[0].runs).toHaveLength(2);
    expect(bodyWithRuns.paragraphs[0].properties.alignment).toBe('right');
    expect(bodyWithRuns.paragraphs[1].runs[0].text).toBe('Simple paragraph string');
    expect(bodyWithRuns.paragraphs[1].properties.alignment).toBe('center');

    // Content with array of strings containing empty lines
    const bodyWithEmptyLines = buildTextBody(['Line 1', '', 'Line 3']);
    expect(bodyWithEmptyLines.paragraphs).toHaveLength(1);
    expect(bodyWithEmptyLines.paragraphs[0].runs).toHaveLength(3);
  });
});
