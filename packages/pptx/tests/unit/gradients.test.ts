import { describe, expect, it } from 'vitest';
import { degrees, inches, Presentation } from '../../lib/index';

describe('Fluent Gradients (@hokkyss/pptx)', () => {
  it('should create and roundtrip presentation with shorthand linear and radial gradients', async () => {
    const pres = Presentation.create();

    // Slide 1: Gradient Background & Linear Gradient Shape
    const slide1 = pres.addSlide();
    slide1.setBackground({
      angle: degrees(135),
      stops: ['#0F172A', '#1E293B'],
      type: 'linear',
    });

    slide1.addShape('roundRect', {
      fill: {
        angle: degrees(90),
        stops: ['#0284C7', '#6366F1'],
      },
      h: inches(2),
      w: inches(4),
      x: inches(1),
      y: inches(1),
    });

    // Multi-stop gradient with alpha opacity
    slide1.addShape('rect', {
      fill: {
        angle: degrees(45),
        stops: [
          { color: '#38BDF8', opacity: 0.9, position: 0 },
          { color: '#818CF8', opacity: 0.5, position: 0.5 },
          { color: '#C084FC', opacity: 0.2, position: 1 },
        ],
      },
      h: inches(2),
      w: inches(4),
      x: inches(6),
      y: inches(1),
    });

    // Slide 2: Radial Gradient Card
    const slide2 = pres.addSlide();
    slide2.addShape('ellipse', {
      fill: {
        stops: [
          { color: '#FFFFFF', position: 0 },
          { color: '#3B82F6', position: 1 },
        ],
        type: 'radial',
      },
      h: inches(3),
      w: inches(3),
      x: inches(2),
      y: inches(2),
    });

    const buffer = await pres.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Reload and verify
    const reloaded = await Presentation.load(buffer);
    expect(reloaded.slides).toHaveLength(2);

    // Verify Slide 1 Background
    const s1 = reloaded.slides[0];
    expect(s1.ast.background?.fill?.type).toBe('gradient');
    expect(s1.ast.background?.fill?.gradient?.type).toBe('linear');
    expect(s1.ast.background?.fill?.gradient?.stops).toHaveLength(2);

    // Verify Slide 1 Shapes
    const elements = s1.getElements();
    expect(elements).toHaveLength(2);

    // Shape 1 (Linear 2-stop)
    expect(elements[0].fill?.type).toBe('gradient');
    expect(elements[0].fill?.gradient?.stops).toHaveLength(2);
    expect(elements[0].fill?.gradient?.stops[0].color).toEqual({
      alpha: undefined,
      type: 'srgb',
      value: '0284C7',
    });
    expect(elements[0].fill?.gradient?.stops[1].color).toEqual({
      alpha: undefined,
      type: 'srgb',
      value: '6366F1',
    });

    // Shape 2 (Linear 3-stop with alpha)
    expect(elements[1].fill?.type).toBe('gradient');
    const stops = elements[1].fill?.gradient?.stops || [];
    expect(stops).toHaveLength(3);
    expect(stops[0].opacity).toBe(0.9);
    expect(stops[1].opacity).toBe(0.5);
    expect(stops[2].opacity).toBe(0.2);

    // Verify Slide 2 Radial Shape
    const s2Elements = reloaded.slides[1].getElements();
    expect(s2Elements[0].fill?.type).toBe('gradient');
    expect(s2Elements[0].fill?.gradient?.type).toBe('radial');
    expect(s2Elements[0].fill?.gradient?.stops).toHaveLength(2);
  });
});
