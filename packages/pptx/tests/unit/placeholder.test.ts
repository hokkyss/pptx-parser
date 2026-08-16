import { describe, expect, it } from 'vitest';
import { emu, emuDegree, points } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Placeholder Replacement API (Unit Tests)', () => {
  it('adds text targeting a placeholder name defined on the layout', () => {
    const pres = Presentation.create();

    // Setup layout with named placeholders
    pres.ast.slideLayouts[0].elements = [
      {
        elementType: 'shape',
        id: 'ph_title_1',
        isVisible: true,
        name: 'placeholder:slide-title',
        placeholder: { idx: 0, type: 'title' },
        position: { cx: emu(9144000), cy: emu(914400), x: emu(914400), y: emu(457200) }, // 10in x 1in at (1in, 0.5in)
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 0,
      },
      {
        elementType: 'shape',
        id: 'ph_content_1',
        isVisible: true,
        name: 'placeholder:slide-content',
        placeholder: { idx: 10, type: 'body' },
        position: { cx: emu(9144000), cy: emu(4572000), x: emu(914400), y: emu(1828800) }, // 10in x 5in at (1in, 2in)
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 1,
      },
    ];

    const slide = pres.addSlide();

    // Replace slide-title via slide.addText
    slide.addText('Custom Slide Title Here', {
      bold: true,
      color: '2563EB',
      fontSize: points(28),
      placeholder: 'placeholder:slide-title',
    });

    // Replace slide-content via slide.addText
    slide.addText('First bullet item\nSecond bullet item', {
      placeholder: 'placeholder:slide-content',
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(2);

    // Verify title element
    const titleEl = elements.find((e) => e.name === 'placeholder:slide-title');
    expect(titleEl).toBeDefined();
    expect(titleEl?.elementType).toBe('shape');
    expect(titleEl?.position.x).toBe(914400); // 1 in EMU
    expect(titleEl?.position.y).toBe(457200); // 0.5 in EMU
    expect(titleEl?.position.cx).toBe(9144000); // 10 in EMU
    expect(titleEl?.placeholder?.type).toBe('title');
    if (titleEl?.elementType === 'shape') {
      const run = titleEl.textBody?.paragraphs?.[0]?.runs?.[0];
      expect(run?.text).toBe('Custom Slide Title Here');
      expect(run?.properties.bold).toBe(true);
      expect(run?.properties.color).toBe('2563EB');
    }

    // Verify content element
    const contentEl = elements.find((e) => e.name === 'placeholder:slide-content');
    expect(contentEl).toBeDefined();
    expect(contentEl?.position.y).toBe(1828800); // 2 in EMU
    expect(contentEl?.position.cy).toBe(4572000); // 5 in EMU
    expect(contentEl?.placeholder?.type).toBe('body');
  });

  it('updates placeholder in-place when called multiple times on the same slide', () => {
    const pres = Presentation.create();
    pres.ast.slideLayouts[0].elements = [
      {
        elementType: 'shape',
        id: 'ph_title',
        isVisible: true,
        name: 'placeholder:slide-title',
        placeholder: { idx: 0, type: 'title' },
        position: { cx: emu(5000000), cy: emu(1000000), x: emu(1000000), y: emu(1000000) },
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 0,
      },
    ];

    const slide = pres.addSlide();
    slide.addText('Initial Title', { placeholder: 'placeholder:slide-title' });
    expect(slide.getElements().length).toBe(1);

    // Overwrite placeholder content
    slide.addText('Updated Final Title', { placeholder: 'placeholder:slide-title' });
    expect(slide.getElements().length).toBe(1);

    const titleEl = slide.getElements()[0];
    if (titleEl.elementType === 'shape') {
      expect(titleEl.textBody?.paragraphs?.[0]?.runs?.[0]?.text).toBe('Updated Final Title');
    }
  });

  it('discovers placeholders with getPlaceholders()', () => {
    const pres = Presentation.create();
    pres.ast.slideLayouts[0].elements = [
      {
        elementType: 'shape',
        id: 'ph1',
        isVisible: true,
        name: 'placeholder:slide-title',
        placeholder: { idx: 0, type: 'title' },
        position: { cx: emu(5000000), cy: emu(1000000), x: emu(1000000), y: emu(1000000) },
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 0,
      },
      {
        elementType: 'shape',
        id: 'ph2',
        isVisible: true,
        name: 'placeholder:slide-content',
        placeholder: { idx: 10, type: 'body' },
        position: { cx: emu(5000000), cy: emu(4000000), x: emu(1000000), y: emu(2200000) },
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 1,
      },
    ];

    const slide = pres.addSlide();
    const placeholders = slide.getPlaceholders();

    expect(placeholders.length).toBe(2);
    expect(placeholders.map((p) => p.name)).toEqual([
      'placeholder:slide-title',
      'placeholder:slide-content',
    ]);
  });

  it('allows adding tables directly to a placeholder position', () => {
    const pres = Presentation.create();
    pres.ast.slideLayouts[0].elements = [
      {
        elementType: 'shape',
        id: 'ph_body',
        isVisible: true,
        name: 'placeholder:slide-content',
        placeholder: { idx: 10, type: 'body' },
        position: { cx: emu(7000000), cy: emu(3000000), x: emu(1500000), y: emu(2000000) },
        rotation: emuDegree(0),
        shapeType: 'rect',
        type: 'shape',
        zIndex: 0,
      },
    ];

    const slide = pres.addSlide();
    slide.addTable(
      [
        ['Col A', 'Col B'],
        ['Val 1', 'Val 2'],
      ],
      {
        placeholder: 'placeholder:slide-content',
      },
    );

    const elements = slide.getElements();
    expect(elements.length).toBe(1);
    expect(elements[0].elementType).toBe('table');
    expect(elements[0].position.x).toBe(1500000);
    expect(elements[0].position.y).toBe(2000000);
  });
});
