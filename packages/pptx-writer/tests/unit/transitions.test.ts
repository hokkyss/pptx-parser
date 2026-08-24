import type { PptxTransitionType } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { serializeTransition } from '../../lib/serializers/transition-serializer';

describe('Slide Transition Serializer (@hokkyss/pptx-writer)', () => {
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

  it('serializes blinds transition with horz and vert direction', () => {
    const horz = serializeTransition({ type: 'blinds', direction: 'horz' });
    expect(horz?.['p:blinds']).toEqual({ '@_dir': 'horz' });
    const vert = serializeTransition({ type: 'blinds', direction: 'vert' });
    expect(vert?.['p:blinds']).toEqual({ '@_dir': 'vert' });
  });

  it('serializes checker and comb transitions with direction', () => {
    const checker = serializeTransition({ type: 'checker', direction: 'vert' });
    expect(checker?.['p:checker']).toEqual({ '@_dir': 'vert' });
    const comb = serializeTransition({ type: 'comb', direction: 'horz' });
    expect(comb?.['p:comb']).toEqual({ '@_dir': 'horz' });
  });

  it('serializes cover and pull transitions with direction', () => {
    const cover = serializeTransition({ type: 'cover', direction: 'down' });
    expect(cover?.['p:cover']).toEqual({ '@_dir': 'd' });
    const pull = serializeTransition({ type: 'pull', direction: 'left' });
    expect(pull?.['p:pull']).toEqual({ '@_dir': 'l' });
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
      duration: 1000,
    });
    expect(node?.['@_advTm']).toBe(4500);
    expect(node?.['@_advClick']).toBe('0');
    expect(node?.['@_dur']).toBe(1000);
    expect(node?.['p:push']).toEqual({ '@_dir': 'u' });
  });

  it('serializes randomBar transition', () => {
    const rb = serializeTransition({ type: 'randomBar', direction: 'horz' });
    expect(rb?.['p:randomBar']).toEqual({ '@_dir': 'horz' });
  });

  it('serializes split transition with orientation and in/out direction', () => {
    const splitIn = serializeTransition({ type: 'split', direction: 'in' });
    expect(splitIn?.['p:split']).toEqual({ '@_dir': 'in' });
    const splitHorz = serializeTransition({ type: 'split', direction: 'horz' });
    expect(splitHorz?.['p:split']).toEqual({ '@_orient': 'horz' });
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

  it('serializes zoom transition with in/out direction', () => {
    const zoomIn = serializeTransition({ type: 'zoom', direction: 'in' });
    expect(zoomIn?.['p:zoom']).toEqual({ '@_dir': 'in' });
    const zoomOut = serializeTransition({ type: 'zoom', direction: 'out' });
    expect(zoomOut?.['p:zoom']).toEqual({ '@_dir': 'out' });
  });

  it('serializes custom or unknown transition type via fallback', () => {
    const custom = serializeTransition({ type: 'newsflash' as PptxTransitionType });
    expect(custom?.['p:newsflash']).toEqual({});
  });

  it('maps durationMs and speed strings correctly', () => {
    expect(serializeTransition({ type: 'cut', speed: 'medium' })?.['@_spd']).toBe('med');
    expect(serializeTransition({ type: 'cut', speed: 'fast' })?.['@_spd']).toBe('fast');
    expect(serializeTransition({ type: 'cut', durationMs: 400 })?.['@_spd']).toBe('fast');
    expect(serializeTransition({ durationMs: 1200, type: 'cut' })?.['@_spd']).toBe('med');
    expect(serializeTransition({ durationMs: 3000, type: 'cut' })?.['@_spd']).toBe('slow');
  });

  it('covers transition directions and speeds', () => {
    const t1 = serializeTransition({ direction: 'left', speed: 'fast', type: 'push' });
    expect(t1).toBeDefined();
    const t2 = serializeTransition({ direction: 'down', speed: 'slow', type: 'wipe' });
    expect(t2).toBeDefined();
    const t3 = serializeTransition({ duration: 1000, type: 'fade' });
    expect(t3).toBeDefined();
  });
});
