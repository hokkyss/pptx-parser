import { describe, expect, it } from 'vitest';
import type { PptxConnectorElement, PptxShapeElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, inches, points } from '@hokkyss/pptx-core';
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

  it('supports attaching connectors to shapes via shapeId and position (top, bottom, left, right)', async () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    // Add Card 1: x: 1", y: 2", w: 3", h: 2" -> Center-right is (4", 3")
    slide.addShape('roundRect', {
      id: 'step-1',
      fill: '0284C7',
      h: inches(2),
      w: inches(3),
      x: inches(1),
      y: inches(2),
    });

    // Add Card 2: x: 6", y: 2", w: 3", h: 2" -> Center-left is (6", 3")
    slide.addShape('roundRect', {
      id: 'step-2',
      fill: '6366F1',
      h: inches(2),
      w: inches(3),
      x: inches(6),
      y: inches(2),
    });

    // Attach connector between step-1 (right) and step-2 (left)
    slide.addConnector({
      color: '0284C7',
      endArrow: 'triangle',
      from: { position: 'right', shapeId: 'step-1' },
      to: { position: 'left', shapeId: 'step-2' },
    });

    const elements = slide.getElements();
    expect(elements.length).toBe(3);

    const connector = elements[2] as PptxConnectorElement;
    expect(connector.elementType).toBe('connector');
    expect(connector.startConnection).toEqual({
      position: 'right',
      shapeId: 'step-1',
    });
    expect(connector.endConnection).toEqual({
      position: 'left',
      shapeId: 'step-2',
    });

    // Position check: (minX=4", minY=3", cx=2", cy=0")
    expect(connector.position.x).toBe(3657600); // 4 inches in EMU
    expect(connector.position.y).toBe(2743200); // 3 inches in EMU
    expect(connector.position.cx).toBe(1828800); // 2 inches in EMU
    expect(connector.position.cy).toBe(0); // Exact horizontal line

    // Full roundtrip write and load
    const bytes = await pres.toArrayBuffer();
    const reloaded = await Presentation.load(bytes);
    const reloadedSlide = reloaded.slides[0];
    const shape1 = reloadedSlide.getElements()[0];
    const shape2 = reloadedSlide.getElements()[1];
    const reloadedConnector = reloadedSlide.getElements()[2] as PptxConnectorElement;

    expect(reloadedConnector.elementType).toBe('connector');
    expect(reloadedConnector.startConnection).toEqual({
      position: 'right',
      shapeId: shape1.id,
    });
    expect(reloadedConnector.endConnection).toEqual({
      position: 'left',
      shapeId: shape2.id,
    });
  });

  it('throws an informative error if attaching a connector to a non-existent shapeId', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    expect(() => {
      slide.addConnector({
        from: { position: 'right', shapeId: 'non-existent-shape' },
        to: { x: inches(5), y: inches(5) },
      });
    }).toThrow('Shape with id "non-existent-shape" was not found on this slide');
  });

  it('throws an informative error if attaching a connector to a group', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addGroup({ id: 'my-group', x: inches(1), y: inches(1), w: inches(3), h: inches(2) }, (g) => {
      g.addShape('rect', { x: inches(1), y: inches(1), w: inches(3), h: inches(2) });
    });

    expect(() => {
      slide.addConnector({
        from: { position: 'right', shapeId: 'my-group' },
        to: { x: inches(5), y: inches(5) },
      });
    }).toThrow('Cannot attach connector to a group ("my-group")');
  });

  it('provides O(1) shape lookup with getElementById and tracks deletions', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('roundRect', { id: 'card-alpha', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
    expect(slide.getElementById('card-alpha')).toBeDefined();
    expect(slide.getElementById('card-alpha')?.id).toBe('card-alpha');

    // Remove element
    const removed = slide.removeElement('card-alpha');
    expect(removed).toBe(true);
    expect(slide.getElementById('card-alpha')).toBeUndefined();
  });

  it('throws an error early when adding an element with duplicate ID on the same slide', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('roundRect', { id: 'stage-1', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });

    expect(() => {
      slide.addShape('ellipse', { id: 'stage-1', x: inches(4), y: inches(1), w: inches(2), h: inches(2) });
    }).toThrow('Duplicate element ID "stage-1" detected on Slide 1');
  });

  it('detects duplicate IDs across mixed element types (text, table, chart, connector, image, group)', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addText('Title Text', { id: 'elem-1', x: inches(1), y: inches(1) });

    // Colliding with addTable
    expect(() => {
      slide.addTable([['Header'], ['Value']], { id: 'elem-1', x: inches(1), y: inches(2) });
    }).toThrow('Duplicate element ID "elem-1" detected on Slide 1');

    // Colliding with addChart
    expect(() => {
      slide.addChart({ categories: ['Q1'], id: 'elem-1', series: [{ data: [100], name: 'Revenue' }] });
    }).toThrow('Duplicate element ID "elem-1" detected on Slide 1');

    // Colliding with addConnector
    expect(() => {
      slide.addConnector({ from: { x: inches(1), y: inches(1) }, id: 'elem-1', to: { x: inches(3), y: inches(1) } });
    }).toThrow('Duplicate element ID "elem-1" detected on Slide 1');

    // Colliding with addGroup
    expect(() => {
      slide.addGroup({ h: inches(2), id: 'elem-1', w: inches(2), x: inches(1), y: inches(1) }, (g) => {
        g.addShape('rect', { h: inches(1), w: inches(1), x: inches(1), y: inches(1) });
      });
    }).toThrow('Duplicate element ID "elem-1" detected on Slide 1');
  });

  it('detects duplicate IDs between top-level shapes and shapes nested inside groups', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('roundRect', { id: 'nested-card', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });

    expect(() => {
      slide.addGroup({ h: inches(3), id: 'my-group-1', w: inches(3), x: inches(1), y: inches(1) }, (g) => {
        g.addShape('rect', { id: 'nested-card', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
      });
    }).toThrow('Duplicate element ID "nested-card" detected on Slide 1');
  });

  it('allows ID reuse after an element is removed with removeElement', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('roundRect', { id: 'temp-id', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
    expect(slide.removeElement('temp-id')).toBe(true);

    // Re-adding with same ID should now succeed
    expect(() => {
      slide.addShape('ellipse', { id: 'temp-id', x: inches(3), y: inches(1), w: inches(2), h: inches(2) });
    }).not.toThrow();

    expect(slide.getElementById('temp-id')?.shapeType).toBe('ellipse');
  });

  it('auto-increment element counter avoids colliding with manual numeric IDs', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    // User manually specifies id '1' and '2'
    slide.addShape('roundRect', { id: '1', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
    slide.addShape('roundRect', { id: '2', x: inches(3), y: inches(1), w: inches(2), h: inches(2) });

    // Auto-generated shape should skip 1 and 2 and take 3 without throwing
    expect(() => {
      slide.addShape('roundRect', { x: inches(5), y: inches(1), w: inches(2), h: inches(2) });
    }).not.toThrow();

    const elements = slide.getElements();
    expect(elements[2].id).toBe('3');
  });

  it('allows the same element ID across different slides', () => {
    const pres = Presentation.create();
    const slide1 = pres.addSlide();
    const slide2 = pres.addSlide();

    slide1.addShape('roundRect', { id: 'card-1', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });
    slide2.addShape('roundRect', { id: 'card-1', x: inches(1), y: inches(1), w: inches(2), h: inches(2) });

    expect(slide1.getElementById('card-1')).toBeDefined();
    expect(slide2.getElementById('card-1')).toBeDefined();
  });
});

import { TableBuilder } from '../../lib/builders/table-builder';

describe('Slide Class extended methods and placeholder resolution', () => {
  it('covers removeElement false return and transition object input', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    expect(slide.removeElement('non-existent')).toBe(false);

    slide.setTransition({ type: 'fade', durationMs: 500 });
    expect(slide.getTransition()?.type).toBe('fade');
    expect(slide.getTransition()?.durationMs).toBe(500);
  });

  it('supports addTable with TableBuilder instance and with builder callback', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    // With callback
    slide.addTable((builder) => {
      builder.addRow().addCell('Callback Cell');
    });

    // With TableBuilder instance
    const customBuilder = new TableBuilder();
    customBuilder.addRow().addCell('Instance Cell');
    slide.addTable(customBuilder);

    const elements = slide.getElements();
    expect(elements).toHaveLength(2);
    expect(elements[0].elementType).toBe('table');
    expect(elements[1].elementType).toBe('table');
  });

  it('resolves placeholders from master when not in layout and updates existing placeholder shape', () => {
    const pres = Presentation.create();
    // Add master element placeholder
    pres.ast.slideMasters[0].elements = [
      {
        elementType: 'shape',
        id: 'ph1',
        name: 'Master Header Placeholder',
        position: { x: emu(0), y: emu(0), cx: emu(1000), cy: emu(1000) },
        rotation: emuDegree(0),
        type: 'shape',
        zIndex: 0,
        placeholder: { type: 'header', idx: '10' },
      },
    ];

    const slide = pres.addSlide();
    expect(slide.resolvePlaceholder('header')).toBeDefined();
    expect(slide.resolvePlaceholder(10)).toBeDefined();

    // Add image targeting placeholder
    const mockPngData = new Uint8Array([137, 80, 78, 71]);
    slide.addImage(mockPngData, { placeholder: 'header' });
    expect(slide.getElements()[0].name).toBe('Master Header Placeholder');

    // Add text modifying existing slide placeholder
    slide.addText('Header Updated Text', {
      placeholder: 'header',
      fill: 'FF0000',
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
    });
  });
});

describe('ShapeBuilder shadow and styling', () => {
  it('adds shape with drop shadow configuration', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    slide.addShape('rect', {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      fill: 'FFFFFF',
      shadow: {
        blur: inches(0.1),
        color: '#333333',
        direction: 45,
        distance: inches(0.05),
        opacity: 0.5,
        rotateWithShape: true,
      },
    });

    const elements = slide.getElements();
    expect(elements[0].elementType).toBe('shape');
    const shape = elements[0];
    if (shape.elementType === 'shape') {
      expect(shape.shadow).toBeDefined();
      expect(shape.shadow?.color).toBe('333333');
      expect(shape.shadow?.opacity).toBe(0.5);
    }
  });
});
