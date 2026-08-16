import { describe, expect, it } from 'vitest';
import { serializeTransition } from '../../lib/serializers/transition-serializer';

describe('Slide Transition Serializer', () => {
  it('returns undefined for undefined or "none" transition', () => {
    expect(serializeTransition(undefined)).toBeUndefined();
    expect(serializeTransition({ type: 'none' })).toBeUndefined();
  });

  it('serializes fade transition with default settings', () => {
    const node = serializeTransition({ type: 'fade' });
    expect(node).toBeDefined();
    expect(node?.['p:fade']).toEqual({});
  });

  it('serializes fade through black transition', () => {
    const node = serializeTransition({ type: 'fade', throughBlack: true });
    expect(node?.['p:fade']).toEqual({ '@_thruBlk': '1' });
  });

  it('serializes wipe transition with direction and speed', () => {
    const node = serializeTransition({
      type: 'wipe',
      direction: 'right',
      speed: 'fast',
    });
    expect(node?.['@_spd']).toBe('fast');
    expect(node?.['p:wipe']).toEqual({ '@_dir': 'r' });
  });

  it('serializes push transition with direction and auto-advance delay', () => {
    const node = serializeTransition({
      type: 'push',
      direction: 'up',
      advanceAfterMs: 4500,
      advanceOnClick: false,
    });
    expect(node?.['@_advTm']).toBe(4500);
    expect(node?.['@_advClick']).toBe('0');
    expect(node?.['p:push']).toEqual({ '@_dir': 'u' });
  });

  it('serializes split transition with orientation and in/out direction', () => {
    const node = serializeTransition({
      type: 'split',
      direction: 'in',
    });
    expect(node?.['p:split']).toEqual({ '@_dir': 'in' });
  });

  it('serializes wheel transition with spoke count', () => {
    const node = serializeTransition({
      type: 'wheel',
      spokes: 4,
      durationMs: 2500,
    });
    expect(node?.['@_spd']).toBe('slow');
    expect(node?.['p:wheel']).toEqual({ '@_spokes': '4' });
  });

  it('maps durationMs to speed correctly', () => {
    expect(serializeTransition({ type: 'cut', durationMs: 400 })?.['@_spd']).toBe('fast');
    expect(serializeTransition({ type: 'cut', durationMs: 1200 })?.['@_spd']).toBe('med');
    expect(serializeTransition({ type: 'cut', durationMs: 3000 })?.['@_spd']).toBe('slow');
  });
});
