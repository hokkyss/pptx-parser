import { describe, expect, it } from 'vitest';
import { points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Multilevel Text Bullets API', () => {
  it('supports explicit multilevel paragraph configurations', () => {
    const pres = Presentation.create({ title: 'Multilevel Bullets Test' });
    const slide = pres.addSlide();

    slide.addText([
      { text: 'Root topic', level: 0 },
      { text: 'First subtopic', level: 1 },
      { text: 'Deep detail A', level: 2 },
      { text: 'Deep detail B', level: 2 },
      { text: 'Second subtopic', level: 1 },
    ]);

    const elements = slide.getElements();
    expect(elements.length).toBe(1);

    const shape = elements[0];
    if (shape.elementType === 'shape') {
      const paragraphs = shape.textBody?.paragraphs;
      expect(paragraphs?.length).toBe(5);
      expect(paragraphs?.[0]?.properties.level).toBe(0);
      expect(paragraphs?.[1]?.properties.level).toBe(1);
      expect(paragraphs?.[2]?.properties.level).toBe(2);
      expect(paragraphs?.[3]?.properties.level).toBe(2);
      expect(paragraphs?.[4]?.properties.level).toBe(1);
      expect(paragraphs?.[0]?.runs[0]?.text).toBe('Root topic');
      expect(paragraphs?.[2]?.runs[0]?.text).toBe('Deep detail A');
    }
  });

  it('supports mixed runs and custom bullet types per level', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      {
        bullet: { char: '•', type: 'char' },
        level: 0,
        runs: [
          { bold: true, text: 'Main Highlight: ' },
          { text: 'High throughput' },
        ],
      },
      {
        bullet: { autoNumType: 'arabicPeriod', type: 'autoNum' },
        level: 1,
        text: 'Step 1: Ingest PPTX package',
      },
      {
        bullet: { autoNumType: 'arabicPeriod', type: 'autoNum' },
        level: 1,
        text: 'Step 2: Parse AST representation',
      },
      {
        bullet: false,
        level: 2,
        text: '(Note: Raw XML fallback preserved)',
      },
    ]);

    const shape = slide.getElements()[0];
    if (shape.elementType === 'shape') {
      const p = shape.textBody?.paragraphs;
      expect(p?.[0]?.properties.bullet?.type).toBe('char');
      expect(p?.[0]?.properties.bullet?.char).toBe('•');
      expect(p?.[1]?.properties.bullet?.type).toBe('autoNum');
      expect(p?.[3]?.properties.bullet?.type).toBe('none');
    }
  });

  it('automatically detects indentation levels from tab characters in string input', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    const multilineText = 'Top Level\n\tSub Item 1\n\t\tSub Item 1.1\n\tSub Item 2';
    slide.addText(multilineText);

    const shape = slide.getElements()[0];
    if (shape.elementType === 'shape') {
      const p = shape.textBody?.paragraphs;
      expect(p?.length).toBe(4);
      expect(p?.[0]?.properties.level).toBe(0);
      expect(p?.[0]?.runs[0]?.text).toBe('Top Level');
      expect(p?.[1]?.properties.level).toBe(1);
      expect(p?.[1]?.runs[0]?.text).toBe('Sub Item 1');
      expect(p?.[2]?.properties.level).toBe(2);
      expect(p?.[2]?.runs[0]?.text).toBe('Sub Item 1.1');
      expect(p?.[3]?.properties.level).toBe(1);
      expect(p?.[3]?.runs[0]?.text).toBe('Sub Item 2');
    }
  });

  it('populates multilevel bullets directly into layout placeholders and round-trips cleanly', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      { bold: true, level: 0, text: '• Executive Summary' },
      { level: 1, text: '– Q1 Revenue: +24% YoY' },
      { level: 1, text: '– Operating Margin: 18.5%' },
      { level: 2, text: '▸ Cloud division leading growth' },
    ], {
      fontSize: points(14),
    });

    const buffer = await pres.toBuffer();
    const loaded = await Presentation.load(buffer);

    const reloadedSlide = loaded.getSlide(1);
    const elements = reloadedSlide?.getElements();
    expect(elements?.length).toBe(1);

    const shape = elements?.[0];
    if (shape?.elementType === 'shape') {
      const p = shape.textBody?.paragraphs;
      expect(p?.length).toBe(4);
      expect(p?.[0]?.properties.level).toBe(0);
      expect(p?.[1]?.properties.level).toBe(1);
      expect(p?.[3]?.properties.level).toBe(2);
    }
  });

  it('preserves clean master styling inheritance when no overrides are passed', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      { level: 0, text: 'Level 0 without overrides' },
      { level: 1, text: 'Level 1 without overrides' },
      { level: 2, text: 'Level 2 without overrides' },
    ]);

    const shape = slide.getElements()[0];
    if (shape.elementType === 'shape') {
      const p = shape.textBody?.paragraphs;
      expect(p?.length).toBe(3);
      // No forced bullet or fontSize override — inherits layout master style!
      expect(p?.[0]?.properties.bullet).toBeUndefined();
      expect(p?.[0]?.runs[0]?.properties.fontSize).toBeUndefined();
      expect(p?.[1]?.properties.level).toBe(1);
      expect(p?.[1]?.properties.bullet).toBeUndefined();
      expect(p?.[2]?.properties.level).toBe(2);
      expect(p?.[2]?.properties.bullet).toBeUndefined();
    }
  });

  it('supports inline modifiers: bold, italic, underline, strikethrough, superscript, subscript', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      { bold: true, level: 0, text: 'Bold Level 0' },
      { italic: true, level: 1, text: 'Italic Level 1' },
      {
        level: 2,
        runs: [
          { strikethrough: true, text: 'Strikethrough Part' },
          { text: ' and ' },
          { underline: true, text: 'Underlined Part' },
        ],
      },
      {
        level: 3,
        runs: [
          { text: 'E=mc' },
          { superscript: true, text: '2' },
          { text: ' and H' },
          { subscript: true, text: '2' },
          { text: 'O' },
        ],
      },
    ]);

    const buffer = await pres.toBuffer();
    const reloaded = await Presentation.load(buffer);
    const shape = reloaded.getSlide(1)?.getElements()[0];

    if (shape?.elementType === 'shape') {
      const p = shape.textBody?.paragraphs;
      expect(p?.length).toBe(4);
      expect(p?.[0]?.runs[0]?.properties.bold).toBe(true);
      expect(p?.[1]?.runs[0]?.properties.italic).toBe(true);
      expect(p?.[2]?.runs[0]?.properties.strikethrough).toBeTruthy();
      expect(p?.[2]?.runs[2]?.properties.underline).toBeTruthy();
      expect(p?.[3]?.runs[1]?.properties.superscript).toBe(true);
      expect(p?.[3]?.runs[3]?.properties.subscript).toBe(true);
    }
  });
});
