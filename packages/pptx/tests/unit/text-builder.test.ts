import { describe, expect, it } from 'vitest';
import { buildTextBody, buildTextRun } from '../../lib/builders/text-builder';

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
    const run1 = buildTextRun('Subscript', { subscript: true, strikethrough: true, baseline: -25000 });
    expect(run1.properties.subscript).toBe(true);
    expect(run1.properties.strikethrough).toBe(true);
    expect(run1.properties.baseline).toBe(-25000);

    const run2 = buildTextRun({
      text: 'Super',
      superscript: true,
      baseline: 30000,
    });
    expect(run2.properties.superscript).toBe(true);
    expect(run2.properties.baseline).toBe(30000);
  });
});

describe('buildTextBody paragraph arrays and multiline runs', () => {
  it('handles mixed paragraph config array with strings and run arrays', () => {
    const body = buildTextBody([
      'String Paragraph',
      { text: ['Run A', 'Run B'], level: 1, spaceBefore: 12, spaceAfter: 6 },
    ]);
    expect(body.paragraphs).toHaveLength(2);
    expect(body.paragraphs[0].runs[0].text).toBe('String Paragraph');
    expect(body.paragraphs[1].runs).toHaveLength(2);
    expect(body.paragraphs[1].properties.level).toBe(1);
  });

  it('handles multi-line text runs creating new paragraphs', () => {
    const body = buildTextBody([
      { text: 'First Line\nSecond Line', bold: true },
    ]);
    expect(body.paragraphs).toHaveLength(2);
    expect(body.paragraphs[0].runs[0].text).toBe('First Line');
    expect(body.paragraphs[1].runs[0].text).toBe('Second Line');
  });
});
