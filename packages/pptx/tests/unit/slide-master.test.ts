import { describe, expect, it } from 'vitest';
import { Presentation } from '../../lib/presentation';
import { SlideMaster } from '../../lib/slide-master';

describe('Slide Master API (Unit Tests)', () => {
  it('discovers masters and layouts in the presentation', () => {
    const pres = Presentation.create();
    const masters = pres.getMasters();

    expect(masters.length).toBeGreaterThan(0);
    const master = masters[0];
    expect(master).toBeInstanceOf(SlideMaster);
    expect(master.id).toBe('slideMaster1');
    expect(master.name).toBe('Office Theme');

    const layouts = master.getLayouts();
    expect(layouts.length).toBeGreaterThan(0);
    expect(layouts[0].name).toBe('Blank');
  });

  it('retrieves master by name, ID, or index', () => {
    const pres = Presentation.create();

    const byId = pres.getMaster('slideMaster1');
    expect(byId).toBeDefined();
    expect(byId?.id).toBe('slideMaster1');

    const byName = pres.getMaster('Office Theme');
    expect(byName).toBeDefined();
    expect(byName?.id).toBe('slideMaster1');

    const byIndex = pres.getMaster(1);
    expect(byIndex).toBeDefined();
    expect(byIndex?.id).toBe('slideMaster1');
  });

  it('adds slide based on a slide master and layout name', () => {
    const pres = Presentation.create();

    // Add extra layout for testing
    pres.ast.slideLayouts.push({
      elements: [],
      id: 'slideLayout2',
      masterId: 'slideMaster1',
      name: 'Title and Content',
      shapes: [],
      type: 'title',
    });
    pres.ast.slideMasters[0].layoutIds.push('slideLayout2');

    // Add slide using master and layout names
    const slide1 = pres.addSlide({
      layout: 'Title and Content',
      master: 'Office Theme',
    });

    expect(slide1.layoutId).toBe('slideLayout2');
  });

  it('adds slide directly from SlideMaster instance', () => {
    const pres = Presentation.create();
    const master = pres.getMasters()[0];

    const slide = master.addSlide({
      notes: 'Slide generated directly from master',
    });

    expect(slide.slideNumber).toBe(1);
    expect(slide.layoutId).toBe('slideLayout1');
    expect(pres.slides.length).toBe(1);
    expect(slide.ast.notes).toBe('Slide generated directly from master');
  });

  it('defaults to first layout of master when layout is omitted', () => {
    const pres = Presentation.create();

    // Create a second master with custom layouts
    pres.ast.slideLayouts.push({
      elements: [],
      id: 'slideLayout10',
      masterId: 'slideMaster2',
      name: 'Custom Dark Layout',
      shapes: [],
      type: 'custom',
    });

    pres.ast.slideMasters.push({
      elements: [],
      id: 'slideMaster2',
      layoutIds: ['slideLayout10'],
      name: 'Dark Template Master',
      shapes: [],
    });

    const slide = pres.addSlide({
      master: 'Dark Template Master',
    });

    expect(slide.layoutId).toBe('slideLayout10');
  });
});
