import { describe, expect, it } from 'vitest';
import { Presentation } from '../../lib/presentation';

describe('Slide Transitions (Fluent SDK & Roundtrip)', () => {
  it('configures, updates, and clears transitions on slides', () => {
    const pres = Presentation.create();
    const slide = pres.addSlide();

    expect(slide.getTransition()).toBeUndefined();

    // 1. Configure fade transition
    slide.setTransition('fade', { durationMs: 500, throughBlack: true });
    expect(slide.getTransition()).toEqual({
      type: 'fade',
      durationMs: 500,
      throughBlack: true,
    });

    // 2. Update to directional wipe
    slide.setTransition('wipe', { direction: 'right', speed: 'fast' });
    expect(slide.getTransition()).toEqual({
      type: 'wipe',
      direction: 'right',
      speed: 'fast',
    });

    // 3. Clear transition
    slide.clearTransition();
    expect(slide.getTransition()).toBeUndefined();
  });

  it('preserves slide transitions across full write and parse roundtrip', async () => {
    const pres = Presentation.create();

    // Slide 1: Fade through black
    const slide1 = pres.addSlide();
    slide1.addText('Slide 1: Intro', { x: 1, y: 1 });
    slide1.setTransition('fade', { throughBlack: true, speed: 'fast' });

    // Slide 2: Push up with auto-advance
    const slide2 = pres.addSlide();
    slide2.addText('Slide 2: Data', { x: 1, y: 1 });
    slide2.setTransition('push', { direction: 'up', advanceAfterMs: 4000, advanceOnClick: false });

    // Slide 3: Wheel with 4 spokes
    const slide3 = pres.addSlide();
    slide3.addText('Slide 3: Conclusion', { x: 1, y: 1 });
    slide3.setTransition('wheel', { spokes: 4 });

    const buffer = await pres.toBuffer();
    expect(buffer).toBeDefined();

    // Read back and verify transitions
    const loadedPres = await Presentation.load(buffer);
    const slides = loadedPres.slides;
    expect(slides.length).toBe(3);

    expect(slides[0].getTransition()?.type).toBe('fade');
    expect(slides[0].getTransition()?.throughBlack).toBe(true);
    expect(slides[0].getTransition()?.speed).toBe('fast');

    expect(slides[1].getTransition()?.type).toBe('push');
    expect(slides[1].getTransition()?.direction).toBe('up');
    expect(slides[1].getTransition()?.advanceAfterMs).toBe(4000);
    expect(slides[1].getTransition()?.advanceOnClick).toBe(false);

    expect(slides[2].getTransition()?.type).toBe('wheel');
    expect(slides[2].getTransition()?.spokes).toBe(4);
  });
});
