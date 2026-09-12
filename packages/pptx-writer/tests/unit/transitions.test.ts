import type { PptxTransitionType } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { renderXml } from '../../lib/xml/xml-element';
import { serializeTransition } from '../../lib/serializers/transition-serializer';

describe('Slide Transition Serializer (@hokkyss/pptx-writer)', () => {
  it('returns undefined for undefined or "none" transition', () => {
    expect(serializeTransition(undefined)).toBeUndefined();
    expect(serializeTransition({ type: 'none' })).toBeUndefined();
  });

  it('serializes fade transition with default settings', () => {
    const node = serializeTransition({ type: 'fade' });
    expect(node).toBeDefined();
    expect(node?.tag).toBe('p:transition');
    expect(renderXml(node)).toBe('<p:transition><p:fade/></p:transition>');
  });

  it('serializes fade through black transition', () => {
    const node = serializeTransition({ type: 'fade', throughBlack: true });
    expect(renderXml(node)).toBe('<p:transition><p:fade thruBlk="1"/></p:transition>');
  });

  it('serializes blinds transition with horz and vert direction', () => {
    const horz = serializeTransition({ type: 'blinds', direction: 'horz' });
    expect(renderXml(horz)).toBe('<p:transition><p:blinds dir="horz"/></p:transition>');
    const vert = serializeTransition({ type: 'blinds', direction: 'vert' });
    expect(renderXml(vert)).toBe('<p:transition><p:blinds dir="vert"/></p:transition>');
  });

  it('serializes checker and comb transitions with direction', () => {
    const checker = serializeTransition({ type: 'checker', direction: 'vert' });
    expect(renderXml(checker)).toBe('<p:transition><p:checker dir="vert"/></p:transition>');
    const comb = serializeTransition({ type: 'comb', direction: 'horz' });
    expect(renderXml(comb)).toBe('<p:transition><p:comb dir="horz"/></p:transition>');
  });

  it('serializes cover and pull transitions with direction', () => {
    const cover = serializeTransition({ type: 'cover', direction: 'down' });
    expect(renderXml(cover)).toBe('<p:transition><p:cover dir="d"/></p:transition>');
    const pull = serializeTransition({ type: 'pull', direction: 'left' });
    expect(renderXml(pull)).toBe('<p:transition><p:pull dir="l"/></p:transition>');
  });

  it('serializes wipe transition with direction and speed', () => {
    const node = serializeTransition({
      direction: 'right',
      speed: 'fast',
      type: 'wipe',
    });
    expect(node?.attrs?.spd).toBe('fast');
    expect(renderXml(node)).toBe('<p:transition spd="fast"><p:wipe dir="r"/></p:transition>');
  });

  it('serializes push transition with direction and auto-advance delay', () => {
    const node = serializeTransition({
      advanceAfterMs: 4500,
      advanceOnClick: false,
      direction: 'up',
      duration: 1000,
      type: 'push',
    });
    expect(node?.attrs?.advTm).toBe(4500);
    expect(node?.attrs?.advClick).toBe('0');
    expect(node?.attrs?.dur).toBe(1000);
    expect(renderXml(node)).toBe('<p:transition spd="med" dur="1000" advClick="0" advTm="4500"><p:push dir="u"/></p:transition>');
  });

  it('serializes randomBar transition', () => {
    const rb = serializeTransition({ direction: 'horz', type: 'randomBar' });
    expect(renderXml(rb)).toBe('<p:transition><p:randomBar dir="horz"/></p:transition>');
  });

  it('serializes split transition with orientation and in/out direction', () => {
    const splitIn = serializeTransition({ direction: 'in', type: 'split' });
    expect(renderXml(splitIn)).toBe('<p:transition><p:split dir="in"/></p:transition>');
    const splitHorz = serializeTransition({ direction: 'horz', type: 'split' });
    expect(renderXml(splitHorz)).toBe('<p:transition><p:split orient="horz"/></p:transition>');
  });

  it('serializes wheel transition with spoke count', () => {
    const node = serializeTransition({
      durationMs: 2500,
      spokes: 4,
      type: 'wheel',
    });
    expect(node?.attrs?.spd).toBe('slow');
    expect(renderXml(node)).toBe('<p:transition spd="slow"><p:wheel spokes="4"/></p:transition>');
  });

  it('serializes zoom transition with in/out direction', () => {
    const zoomIn = serializeTransition({ direction: 'in', type: 'zoom' });
    expect(renderXml(zoomIn)).toBe('<p:transition><p:zoom dir="in"/></p:transition>');
    const zoomOut = serializeTransition({ direction: 'out', type: 'zoom' });
    expect(renderXml(zoomOut)).toBe('<p:transition><p:zoom dir="out"/></p:transition>');
  });

  it('serializes custom or unknown transition type via fallback', () => {
    const custom = serializeTransition({ type: 'newsflash' as PptxTransitionType });
    expect(renderXml(custom)).toBe('<p:transition><p:newsflash/></p:transition>');
  });

  it('maps durationMs and speed strings correctly', () => {
    expect(serializeTransition({ speed: 'medium', type: 'cut' })?.attrs?.spd).toBe('med');
    expect(serializeTransition({ speed: 'fast', type: 'cut' })?.attrs?.spd).toBe('fast');
    expect(serializeTransition({ durationMs: 400, type: 'cut' })?.attrs?.spd).toBe('fast');
    expect(serializeTransition({ durationMs: 1200, type: 'cut' })?.attrs?.spd).toBe('med');
    expect(serializeTransition({ durationMs: 3000, type: 'cut' })?.attrs?.spd).toBe('slow');
  });

  it('covers transition directions and speeds', () => {
    const t1 = serializeTransition({ direction: 'left', speed: 'fast', type: 'push' });
    expect(t1).toBeDefined();
    const t3 = serializeTransition({ duration: 1000, type: 'fade' });
    expect(t3).toBeDefined();

    // Empty transition object (default to fade)
    // @ts-expect-error Testing empty transition object
    const defaultTrans = serializeTransition({});
    expect(renderXml(defaultTrans)).toBe('<p:transition><p:fade/></p:transition>');

    // advanceOnClick true
    const advClickTrue = serializeTransition({ advanceOnClick: true, type: 'fade' });
    expect(advClickTrue?.attrs?.advClick).toBe('1');

    // Custom direction and comb/randomBar vert
    const customDir = serializeTransition({ direction: 'customDiagonal', type: 'cover' });
    expect(renderXml(customDir)).toBe('<p:transition><p:cover dir="customDiagonal"/></p:transition>');

    const combVert = serializeTransition({ direction: 'vert', type: 'comb' });
    expect(renderXml(combVert)).toBe('<p:transition><p:comb dir="vert"/></p:transition>');

    const rbVert = serializeTransition({ direction: 'vert', type: 'randomBar' });
    expect(renderXml(rbVert)).toBe('<p:transition><p:randomBar dir="vert"/></p:transition>');

    const splitVert = serializeTransition({ direction: 'vert', type: 'split' });
    expect(renderXml(splitVert)).toBe('<p:transition><p:split orient="vert"/></p:transition>');
  });
});
