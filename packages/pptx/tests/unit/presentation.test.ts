import { describe, expect, it } from 'vitest';
import type { PptxDocument, PptxSlide } from '@hokkyss/pptx-core';
import { emu, inches } from '@hokkyss/pptx-core';
import { Presentation } from '../../lib/presentation';

describe('Presentation Class (Unit Tests)', () => {
  describe('Presentation.create()', () => {
    it('creates a default 16:9 widescreen presentation', () => {
      const pres = Presentation.create();
      expect(pres.slides.length).toBe(0);
      expect(pres.metadata.slideWidth).toBe(12192000); // 13.333333 inches in EMU
      expect(pres.metadata.slideHeight).toBe(6858000); // 7.5 inches in EMU
      expect(pres.ast.slideMasters.length).toBeGreaterThan(0);
      expect(pres.ast.slideLayouts.length).toBeGreaterThan(0);
    });

    it('creates a presentation with custom dimensions in Inches', () => {
      const pres = Presentation.create({
        title: 'Custom Deck',
        width: inches(10),
        height: inches(7.5),
      });
      expect(pres.metadata.title).toBe('Custom Deck');
      expect(pres.metadata.slideWidth).toBe(9144000); // 10 inches in EMU
      expect(pres.metadata.slideHeight).toBe(6858000); // 7.5 inches in EMU
    });
  });

  describe('Slide CRUD Operations', () => {
    it('adds slides to the presentation', () => {
      const pres = Presentation.create();
      const slide1 = pres.addSlide();
      const slide2 = pres.addSlide();

      expect(pres.slides.length).toBe(2);
      expect(slide1.slideNumber).toBe(1);
      expect(slide2.slideNumber).toBe(2);
      expect(pres.getSlide(1)).toBe(slide1);
      expect(pres.getSlide(2)).toBe(slide2);
    });

    it('retrieves slide by 1-based index or rId', () => {
      const pres = Presentation.create();
      const slide = pres.addSlide();

      expect(pres.getSlide(1)).toBe(slide);
      expect(pres.getSlide(slide.slideId)).toBe(slide);
      expect(pres.getSlide(999)).toBeUndefined();
    });

    it('removes slide by index or id and re-indexes remaining slides', () => {
      const pres = Presentation.create();
      const s1 = pres.addSlide();
      pres.addSlide();
      const s3 = pres.addSlide();

      const removed = pres.removeSlide(2);
      expect(removed).toBe(true);
      expect(pres.slides.length).toBe(2);
      expect(pres.slides[0]).toBe(s1);
      expect(pres.slides[1]).toBe(s3);
      expect(pres.slides[1].slideNumber).toBe(2); // reindexed
    });

    it('duplicates an existing slide with deep clone of elements', () => {
      const pres = Presentation.create();
      const s1 = pres.addSlide();
      s1.addText('Original Slide Header', {
        x: inches(1),
        y: inches(1),
        w: inches(5),
        h: inches(1),
      });

      const s2 = pres.duplicateSlide(1);
      expect(pres.slides.length).toBe(2);
      expect(s2.slideNumber).toBe(2);
      expect(s2.getElements().length).toBe(1);
      expect(s2.getElements()[0].name).toBe(s1.getElements()[0].name);

      // Verify deep copy mutation isolation
      s2.addText('Second Element on Copy', {
        x: inches(1),
        y: inches(2),
        w: inches(5),
        h: inches(1),
      });
      expect(s1.getElements().length).toBe(1);
      expect(s2.getElements().length).toBe(2);
    });

    it('moves slide to a new position', () => {
      const pres = Presentation.create();
      const s1 = pres.addSlide();
      const s2 = pres.addSlide();
      const s3 = pres.addSlide();

      pres.moveSlide(1, 3); // Move s1 from position 1 to position 3
      expect(pres.slides[0]).toBe(s2);
      expect(pres.slides[1]).toBe(s3);
      expect(pres.slides[2]).toBe(s1);
      expect(pres.slides[0].slideNumber).toBe(1);
      expect(pres.slides[1].slideNumber).toBe(2);
      expect(pres.slides[2].slideNumber).toBe(3);
    });
  });

  describe('Theme Modification API', () => {
    it('modifies theme colors with partial merge and # stripping', () => {
      const pres = Presentation.create();
      const initialAccent2 = pres.ast.themes[0].colorScheme.accent2;

      pres.setThemeColors({
        accent1: '#FF0000',
        accent3: '38BDF8',
      });

      expect(pres.ast.themes[0].colorScheme.accent1).toBe('FF0000');
      expect(pres.ast.themes[0].colorScheme.accent3).toBe('38BDF8');
      expect(pres.ast.themes[0].colorScheme.accent2).toBe(initialAccent2); // unchanged
    });

    it('modifies theme fonts including major, minor, and font scheme name', () => {
      const pres = Presentation.create();
      pres.setThemeFonts({
        major: 'Inter',
        minor: 'Roboto',
        name: 'Modern Corporate Fonts',
      });

      expect(pres.ast.themes[0].fontScheme.majorFont).toBe('Inter');
      expect(pres.ast.themes[0].fontScheme.minorFont).toBe('Roboto');
      expect(pres.ast.themes[0].fontScheme.name).toBe('Modern Corporate Fonts');
    });

    it('modifies theme name', () => {
      const pres = Presentation.create();
      pres.setThemeName('Ajinomoto Custom Theme');

      expect(pres.ast.themes[0].name).toBe('Ajinomoto Custom Theme');
    });

    it('chains theme methods fluently', () => {
      const pres = Presentation.create();
      pres
        .setThemeName('Fluent Theme')
        .setThemeColors({ accent1: '#0055FF' })
        .setThemeFonts({ major: 'Arial' });

      expect(pres.ast.themes[0].name).toBe('Fluent Theme');
      expect(pres.ast.themes[0].colorScheme.accent1).toBe('0055FF');
      expect(pres.ast.themes[0].fontScheme.majorFont).toBe('Arial');
    });
  });

  describe('First Slide Numbering (Starting Slide Number)', () => {
    it('initializes firstSlideNumber from CreatePresentationOptions', () => {
      const pres = Presentation.create({ firstSlideNumber: 0 });
      expect(pres.firstSlideNumber).toBe(0);
      expect(pres.metadata.firstSlideNumber).toBe(0);
    });

    it('sets firstSlideNumber via fluent method', () => {
      const pres = Presentation.create();
      expect(pres.firstSlideNumber).toBeUndefined();

      pres.setFirstSlideNumber(0);
      expect(pres.firstSlideNumber).toBe(0);

      pres.setFirstSlideNumber(5);
      expect(pres.firstSlideNumber).toBe(5);
    });

    it('roundtrips firstSlideNumber cleanly through binary write and parse', async () => {
      const pres = Presentation.create({ firstSlideNumber: 0 });
      pres.addSlide();
      const buffer = await pres.toBuffer();

      const loaded = await Presentation.load(buffer);
      expect(loaded.firstSlideNumber).toBe(0);
    });
  });
});

describe('Edge cases & Error Handling', () => {
  it('returns false when removing non-existent slide', () => {
    const pres = Presentation.create();
    expect(pres.removeSlide(99)).toBe(false);
    expect(pres.removeSlide('rId99')).toBe(false);
  });

  it('throws error when duplicating non-existent slide', () => {
    const pres = Presentation.create();
    expect(() => pres.duplicateSlide(99)).toThrow('not found');
  });

  it('throws error when moving slide with out-of-bounds indices', () => {
    const pres = Presentation.create();
    pres.addSlide();
    expect(() => pres.moveSlide(0, 1)).toThrow('Invalid slide index');
    expect(() => pres.moveSlide(1, 5)).toThrow('Invalid slide index');
  });

  it('sets theme colors with custom palette name', () => {
    const pres = Presentation.create();
    pres.setThemeColors({ accent1: '#123456' }, 'MyPalette');
    expect(pres.ast.themes[0].colorScheme.name).toBe('MyPalette');
  });

  it('resolves layout by type or partial matchingName on addSlide', () => {
    const pres = Presentation.create();
    pres.ast.slideLayouts.push({
      elements: [],
      id: 'slideLayout3',
      masterId: 'slideMaster1',
      name: 'Comparison',
      matchingName: 'TwoColumnLayout',
      shapes: [],
      type: 'twoColumn',
    });

    const byName = pres.getMaster('Office Theme');
    expect(byName).toBeDefined();
    expect(byName?.id).toBe('slideMaster1');

    const byPartialName = pres.getMaster('Office');
    expect(byPartialName).toBeDefined();
    expect(byPartialName?.id).toBe('slideMaster1');

    const byIndex = pres.getMaster(1);
    expect(byIndex).toBeDefined();
    const s1 = pres.addSlide({ layout: 'twoColumn' });
    expect(s1.layoutId).toBe('slideLayout3');

    const s2 = pres.addSlide({ layout: 'TwoColumnLayout' });
    expect(s2.layoutId).toBe('slideLayout3');
  });
});

describe('Presentation Class theme guards and rich slide duplication', () => {
  it('handles empty themes in setThemeColors, setThemeFonts, and setThemeName', () => {
    const pres = Presentation.create();
    pres.ast.themes = [];

    expect(pres.setThemeColors({ accent1: '#FF0000' })).toBe(pres);
    expect(pres.setThemeFonts({ major: 'Calibri' })).toBe(pres);
    expect(pres.setThemeName('Custom')).toBe(pres);
  });

  it('duplicates slide with background, animations, shapes, and notes', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();
    slide.setBackground('0F172A');
    slide.setNotes('Speaker notes remark');
    slide.addText('Slide 1 Content', { x: inches(1), y: inches(1), w: inches(4), h: inches(1) });

    const duplicated = pres.duplicateSlide(1);
    expect(duplicated.notes).toBe('Speaker notes remark');
    expect(duplicated.ast.background?.fill?.type).toBe('solid');
    expect(duplicated.getElements().length).toBe(1);
    expect(pres.slides.length).toBe(2);
  });

  it('handles missing collections, fallback indices and empty layout lists', () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: { created: new Date(), modified: new Date(), revision: 1, slideCount: 0, slideHeight: emu(6858000), slideWidth: emu(12192000) },
      slideLayouts: [],
      slideMasters: [],
      slides: [],
      themes: [],
    };
    const pres = new Presentation(doc);
    expect(pres.slides).toEqual([]);
    expect(pres.getMasters()).toEqual([]);
    expect(pres.getMaster(0)).toBeUndefined();
    expect(pres.getMaster(1)).toBeUndefined();

    // duplicateSlide on minimal slide
    const minimalSlide: PptxSlide = {
      animations: [],
      elements: [],
      shapes: [],
      slideId: 'rId1',
      slideNumber: 1,
    };
    pres.ast.slides = [minimalSlide];
    pres.ast.slideLayouts = [];
    const presWithSlide = new Presentation(pres.ast);
    const duplicated = presWithSlide.duplicateSlide(1);
    expect(duplicated.slideNumber).toBe(2);

    // resolveLayoutId fallback when slideLayouts is empty
    const pres2 = Presentation.create();
    pres2.ast.slideLayouts = [];
    const s = pres2.addSlide();
    expect(s.layoutId).toBe('slideLayout1');

    // Presentation with undefined slides/masters/layouts
    // @ts-expect-error Testing undefined collections at runtime
    const presWithUndefinedArrays = new Presentation({ ...doc, slideLayouts: undefined, slideMasters: undefined, slides: undefined });
    expect(presWithUndefinedArrays.slides).toHaveLength(0);
    expect(presWithUndefinedArrays.getMasters()).toHaveLength(0);

    // Duplicate slide with undefined animations/elements/shapes
    const slideWithUndefinedArrays: PptxSlide = {
      animations: [],
      elements: [],
      shapes: [],
      slideId: 'rId1',
      slideNumber: 1,
    };
    // @ts-expect-error Testing undefined animations at runtime
    delete slideWithUndefinedArrays.animations;
    // @ts-expect-error Testing undefined elements at runtime
    delete slideWithUndefinedArrays.elements;
    // @ts-expect-error Testing undefined shapes at runtime
    delete slideWithUndefinedArrays.shapes;
    presWithUndefinedArrays.ast.slides = [slideWithUndefinedArrays];
    const pres3 = new Presentation(presWithUndefinedArrays.ast);
    const dupSlide = pres3.duplicateSlide(1);
    expect(dupSlide.slideNumber).toBe(2);
  });
});


