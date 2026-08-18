import { describe, expect, it } from 'vitest';
import type { PptxConnectorElement } from '@hokkyss/pptx-core';
import { inches, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Slide Class (Unit Tests)', () => {
  it('adds a simple text box to the slide', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText('Hello World', {
      align: 'center',
      bold: true,
      color: '2563EB',
      font: 'Arial',
      fontSize: points(24),
      h: inches(1),
      w: inches(6),
      x: inches(2),
      y: inches(1.5),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);
    expect(elements[0].elementType).toBe('shape');
    expect(elements[0].position?.x).toBe(1828800); // 2 inches in EMU
    expect(elements[0].position?.y).toBe(1371600); // 1.5 inches in EMU
    expect(elements[0].position?.cx).toBe(5486400); // 6 inches in EMU
    expect(elements[0].position?.cy).toBe(914400); // 1 inch in EMU

    const shape = elements[0];
    if (shape.elementType === 'shape') {
      const run = shape.textBody?.paragraphs?.[0]?.runs?.[0];
      expect(run?.text).toBe('Hello World');
      expect(run?.properties?.bold).toBe(true);
      expect(run?.properties?.color).toBe('2563EB');
      expect(run?.properties?.fontSize).toBe(2400); // 24 pt in hundredths
    }
  });

  it('adds geometric shapes with fills and outlines', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('roundRect', {
      fill: '10B981',
      h: inches(2),
      line: {
        color: '047857',
        width: inches(0.02),
      },
      text: 'Callout Box',
      w: inches(4),
      x: inches(1),
      y: inches(1),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);
    const shape = elements[0];
    expect(shape.elementType).toBe('shape');
    if (shape.elementType === 'shape') {
      expect(shape.shapeType).toBe('roundRect');
      expect(shape.fill?.type).toBe('solid');
      expect(shape.textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('Callout Box');
    }
  });

  it('adds an embedded image to the slide', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    const mockPngData = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    slide.addImage(mockPngData, {
      fileName: 'logo.png',
      h: inches(2),
      w: inches(3),
      x: inches(1),
      y: inches(1),
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(1);
    expect(elements[0].elementType).toBe('picture');
    expect(pres.ast.media.length).toBe(1);
    expect(pres.ast.media[0].fileName).toBe('logo.png');
  });

  it('sets slide background and speaker notes and preserves them across save/load', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.setBackground('0F172A');
    slide.setNotes('Speaker note: emphasize slide key points.\nSecond line of speaker remarks.');

    expect(slide.ast.background?.fill?.type).toBe('solid');
    expect(slide.notes).toBe('Speaker note: emphasize slide key points.\nSecond line of speaker remarks.');

    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    expect(reloaded.slides[0].notes).toBe('Speaker note: emphasize slide key points.\nSecond line of speaker remarks.');
  });

  it('supports customized rich speaker notes with underline, bold, italic and color', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.setNotes([
      {
        runs: [
          { text: 'Speaker remarks: ', bold: true },
          { text: 'underline this key objective', underline: true },
          { text: ', ' },
          { text: 'italic note', italic: true },
          { text: ', and ' },
          { text: 'urgent alert in red', color: 'EF4444', bold: true },
        ],
      },
    ]);

    expect(slide.notes).toBe('Speaker remarks: underline this key objective, italic note, and urgent alert in red');
    expect(slide.notesBody?.paragraphs[0].runs.length).toBe(6);
    expect(slide.notesBody?.paragraphs[0].runs[1].properties.underline).toBe(true);

    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    expect(reloaded.slides[0].notes).toBe('Speaker remarks: underline this key objective, italic note, and urgent alert in red');
    expect(reloaded.slides[0].notesBody?.paragraphs[0].runs[1].properties.underline).toBe(true);
    expect(reloaded.slides[0].notesBody?.paragraphs[0].runs[1].text).toBe('underline this key objective');
  });

  it('supports bulleted speaker notes across save/load', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.setNotes([
      { text: 'Talking points:' },
      { bullet: true, text: 'First bulleted remark' },
      { bullet: true, text: 'Second bulleted remark' },
    ]);

    expect(slide.notes).toBe('Talking points:\nFirst bulleted remark\nSecond bulleted remark');
    expect(slide.notesBody?.paragraphs[1].properties.bullet?.type).toBe('char');

    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    expect(reloaded.slides[0].notes).toBe('Talking points:\nFirst bulleted remark\nSecond bulleted remark');
    expect(reloaded.slides[0].notesBody?.paragraphs[1].properties.bullet?.type).toBe('char');
  });

  it('removes elements by ID', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText('First Item', { h: inches(1), w: inches(2), x: inches(1), y: inches(1) });
    slide.addText('Second Item', { h: inches(1), w: inches(2), x: inches(1), y: inches(2) });

    expect(slide.getElements().length).toBe(2);
    const firstId = slide.getElements()[0].id;
    const removed = slide.removeElement(firstId);

    expect(removed).toBe(true);
    expect(slide.getElements().length).toBe(1);
    expect(slide.getElements()[0].name).toContain('2');
  });

  it('supports addConnector and addGroup across save/load', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addConnector({
      color: '0284C7',
      dashStyle: 'solid',
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(4), y: inches(1) },
      width: inches(0.02),
    });

    slide.addGroup({
      h: inches(3),
      w: inches(3),
      x: inches(5),
      y: inches(1),
    }, (g) => {
      g.addShape('roundRect', { fill: 'FFFFFF', h: inches(3), w: inches(3), x: inches(5), y: inches(1) });
      g.addText('Grouped Text', { h: inches(1), w: inches(2), x: inches(5.5), y: inches(1.5) });
    });

    expect(slide.getElements().length).toBe(2);
    expect(slide.getElements()[0].elementType).toBe('connector');
    expect(slide.getElements()[1].elementType).toBe('group');

    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    expect(reloaded.slides[0].getElements().length).toBe(2);
  });

  it('supports customizable arrowheads (endArrow, startArrow, headEnd, tailEnd) across save/load', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    // 1. Shorthand arrow type
    slide.addConnector({
      color: '0284C7',
      endArrow: 'triangle',
      from: { x: inches(1), y: inches(1) },
      startArrow: 'oval',
      to: { x: inches(4), y: inches(1) },
    });

    // 2. Granular arrow configuration with custom width and length
    slide.addConnector({
      color: '10B981',
      endArrow: { length: 'lg', type: 'stealth', width: 'lg' },
      from: { x: inches(1), y: inches(3) },
      startArrow: { length: 'sm', type: 'diamond', width: 'sm' },
      to: { x: inches(4), y: inches(3) },
    });

    // 3. Aliases headEnd / tailEnd
    slide.addConnector({
      color: '6366F1',
      from: { x: inches(1), y: inches(5) },
      headEnd: { length: 'med', type: 'arrow', width: 'med' },
      tailEnd: { length: 'med', type: 'none', width: 'med' },
      to: { x: inches(4), y: inches(5) },
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(3);

    const c1 = elements[0] as PptxConnectorElement;
    expect(c1.line?.headEnd?.type).toBe('triangle');
    expect(c1.line?.tailEnd?.type).toBe('oval');

    const c2 = elements[1] as PptxConnectorElement;
    expect(c2.line?.headEnd?.type).toBe('stealth');
    expect(c2.line?.headEnd?.width).toBe('lg');
    expect(c2.line?.headEnd?.length).toBe('lg');
    expect(c2.line?.tailEnd?.type).toBe('diamond');
    expect(c2.line?.tailEnd?.width).toBe('sm');
    expect(c2.line?.tailEnd?.length).toBe('sm');

    const c3 = elements[2] as PptxConnectorElement;
    expect(c3.line?.headEnd?.type).toBe('arrow');
    expect(c3.line?.tailEnd?.type).toBe('none');

    // Verify full write and parse roundtrip fidelity
    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    const reloadedElements = reloaded.slides[0].getElements();
    expect(reloadedElements.length).toBe(3);

    const rc1 = reloadedElements[0] as PptxConnectorElement;
    expect(rc1.line?.headEnd?.type).toBe('triangle');
    expect(rc1.line?.tailEnd?.type).toBe('oval');

    const rc2 = reloadedElements[1] as PptxConnectorElement;
    expect(rc2.line?.headEnd?.type).toBe('stealth');
    expect(rc2.line?.headEnd?.width).toBe('lg');
    expect(rc2.line?.headEnd?.length).toBe('lg');
    expect(rc2.line?.tailEnd?.type).toBe('diamond');
    expect(rc2.line?.tailEnd?.width).toBe('sm');
    expect(rc2.line?.tailEnd?.length).toBe('sm');

    const rc3 = reloadedElements[2] as PptxConnectorElement;
    expect(rc3.line?.headEnd?.type).toBe('arrow');
    expect(rc3.line?.tailEnd?.type).toBe('none');
  });
});
