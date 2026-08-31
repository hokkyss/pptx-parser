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

  it('supports Shift+Enter line breaks ({ break: true }) inside bulleted paragraphs and round-trips cleanly', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      {
        level: 0,
        runs: [
          { text: 'A1' },
          { break: true },
          { text: 'a1 level, no bullet' },
        ],
      },
      {
        level: 1,
        text: 'B1',
      },
      {
        level: 2,
        runs: [
          { text: 'C1' },
          { break: true },
          { text: 'c1 level, no bullet' },
        ],
      },
    ]);

    const buffer = await pres.toBuffer();
    const loaded = await Presentation.load(buffer);
    const shape = loaded.getSlide(1)?.getElements()[0];

    if (shape?.elementType === 'shape') {
      const paragraphs = shape.textBody?.paragraphs;
      expect(paragraphs?.length).toBe(3);

      // Paragraph 0 (level 0)
      expect(paragraphs?.[0]?.properties.level).toBe(0);
      expect(paragraphs?.[0]?.runs.some((r) => r.text === 'A1')).toBe(true);
      expect(paragraphs?.[0]?.runs.some((r) => r.break === true)).toBe(true);
      expect(paragraphs?.[0]?.runs.some((r) => r.text === 'a1 level, no bullet')).toBe(true);

      // Paragraph 1 (level 1 - Tab)
      expect(paragraphs?.[1]?.properties.level).toBe(1);
      expect(paragraphs?.[1]?.runs[0]?.text).toBe('B1');

      // Paragraph 2 (level 2 - Tab Tab)
      expect(paragraphs?.[2]?.properties.level).toBe(2);
      expect(paragraphs?.[2]?.runs.some((r) => r.text === 'C1')).toBe(true);
      expect(paragraphs?.[2]?.runs.some((r) => r.break === true)).toBe(true);
      expect(paragraphs?.[2]?.runs.some((r) => r.text === 'c1 level, no bullet')).toBe(true);
    }
  });

  it('loads bullets.pptx correctly with all line breaks and hierarchical levels', async () => {
    const fs = await import('node:fs');
    const path = await import('node:path');
    const filePath = path.resolve(__dirname, '../../../../bullets.pptx');

    if (fs.existsSync(filePath)) {
      const fileBuffer = fs.readFileSync(filePath);
      const pres = await Presentation.load(fileBuffer);
      const slide = pres.getSlide(1);
      const elements = slide?.getElements();
      expect(elements).toBeDefined();

      const placeholder = elements?.find((el) => el.elementType === 'shape' && el.name.includes('Placeholder'));
      if (placeholder?.elementType === 'shape') {
        const paragraphs = placeholder.textBody?.paragraphs;
        expect(paragraphs?.length).toBeGreaterThanOrEqual(8);

        // First paragraph should have A1, a break, and continuation
        const firstP = paragraphs?.[0];
        expect(firstP?.runs.some((r) => r.text === 'A1')).toBe(true);
        expect(firstP?.runs.some((r) => r.break === true)).toBe(true);
      }
    }
  });

  it('supports non-bulleted hierarchical indentation ({ level: N, bullet: false }) with roundtrip preservation', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText([
      {
        bullet: true,
        level: 0,
        text: 'Root Bullet (Level 0)',
      },
      {
        bullet: false,
        level: 1,
        runs: [
          { text: 'Indented block commentary (Level 1, no bullet glyph)' },
          { break: true },
          { text: 'Second line under same block indent' },
        ],
      },
      {
        bullet: false,
        level: 2,
        text: 'Deeply indented technical spec (Level 2, no bullet glyph)',
      },
      {
        bullet: true,
        level: 1,
        text: 'Sub-bullet (Level 1)',
      },
    ]);

    const buffer = await pres.toBuffer();
    const loaded = await Presentation.load(buffer);
    const shape = loaded.getSlide(1)?.getElements()[0];

    if (shape?.elementType === 'shape') {
      const paragraphs = shape.textBody?.paragraphs;
      expect(paragraphs?.length).toBe(4);

      // Paragraph 0: Level 0 bullet
      expect(paragraphs?.[0]?.properties.level).toBe(0);
      expect(paragraphs?.[0]?.properties.bullet?.type).toBe('char');

      // Paragraph 1: Level 1 non-bulleted with soft break
      expect(paragraphs?.[1]?.properties.level).toBe(1);
      expect(paragraphs?.[1]?.properties.bullet?.type).toBe('none');
      expect(paragraphs?.[1]?.runs.some((r) => r.text === 'Indented block commentary (Level 1, no bullet glyph)')).toBe(true);
      expect(paragraphs?.[1]?.runs.some((r) => r.break === true)).toBe(true);
      expect(paragraphs?.[1]?.runs.some((r) => r.text === 'Second line under same block indent')).toBe(true);

      // Paragraph 2: Level 2 non-bulleted
      expect(paragraphs?.[2]?.properties.level).toBe(2);
      expect(paragraphs?.[2]?.properties.bullet?.type).toBe('none');

      // Paragraph 3: Level 1 bullet
      expect(paragraphs?.[3]?.properties.level).toBe(1);
      expect(paragraphs?.[3]?.properties.bullet?.type).toBe('char');
    }
  });

  it('serializes presentation with full 9-level <p:defaultTextStyle> in ppt/presentation.xml for desktop tab parity', async () => {
    const { createZipReader } = await import('@hokkyss/pptx-reader');
    const pres = Presentation.create();
    pres.addSlide();

    const buffer = await pres.toBuffer();
    const zipReader = await createZipReader(buffer);
    const presentationXml = zipReader.getFileText('ppt/presentation.xml') ?? '';

    expect(presentationXml).toContain('<p:defaultTextStyle>');
    for (let i = 1; i <= 9; i++) {
      expect(presentationXml).toContain(`<a:lvl${i}pPr`);
    }
    expect(presentationXml).toContain('defTabSz="914400"');
    expect(presentationXml).toContain('marL="0"');
    expect(presentationXml).toContain('marL="457200"');
    expect(presentationXml).toContain('marL="914400"');
    expect(presentationXml).toContain('marL="3657600"');
  });
});

