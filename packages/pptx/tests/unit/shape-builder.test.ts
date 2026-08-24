import { describe, expect, it } from 'vitest';
import { buildShapeElement } from '../../lib/builders/shape-builder';
import { inches } from '../../lib/units';

describe('ShapeBuilder (Unit Tests)', () => {
  it('handles single stop gradient, empty lines, and default shadow', () => {
    const s1 = buildShapeElement('rect', {
      fill: { stops: ['#FF0000'] },
      h: inches(1),
      line: { width: inches(0.1) },
      shadow: {},
      w: inches(1),
      x: inches(0),
      y: inches(0),
    });
    expect(s1.fill?.type).toBe('gradient');
    if (s1.fill?.type === 'gradient') {
      expect(s1.fill.gradient.stops[0].position).toBe(0);
    }
    expect(s1.line?.fill).toBeUndefined();
    expect(s1.shadow?.color).toBe('000000');

    const s2 = buildShapeElement('rect', {
      fill: { stops: [{ color: '#00FF00' }, { color: '#0000FF' }] },
      h: inches(1),
      line: { color: '0000FF' },
      rotation: 90,
      shadow: { blur: inches(0.2), color: 'FF0000', direction: 45, distance: inches(0.1), opacity: 0.8 },
      w: inches(1),
      x: inches(0),
      y: inches(0),
    });
    expect(s2.rotation).toBe(5400000);
    expect(s2.line?.width).toBeDefined();
    if (s2.fill?.type === 'gradient') {
      expect(s2.fill.gradient.stops[1].position).toBe(1);
    }
  });
});
