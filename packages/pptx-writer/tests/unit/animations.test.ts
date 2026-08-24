import { describe, expect, it } from 'vitest';
import type { PptxAnimation } from '@hokkyss/pptx-core';
import { serializeAnimations } from '../../lib/serializers/animation-serializer';

function makeAnim(overrides: Partial<PptxAnimation> = {}): PptxAnimation {
  return {
    effect: 'appear',
    sequence: 0,
    targetShapeId: '10',
    trigger: 'onClick',
    ...overrides,
  };
}

function getAnimNodes(result: Record<string, unknown>): unknown[] {
  const tnLst = result['p:tnLst'] as Record<string, unknown>;
  const par = tnLst['p:par'] as Record<string, unknown>;
  const outerCTn = par['p:cTn'] as Record<string, unknown>;
  const outerChildTn = outerCTn['p:childTnLst'] as Record<string, unknown>;
  const seq = outerChildTn['p:seq'] as Record<string, unknown>;
  const innerCTn = seq['p:cTn'] as Record<string, unknown>;
  return innerCTn['p:childTnLst'] as unknown[];
}

function getFirstCTn(nodes: unknown[]): Record<string, unknown> {
  const first = nodes[0] as Record<string, unknown>;
  const cMediaNode = first['p:cMediaNode'] as Record<string, unknown>;
  return cMediaNode['p:cTn'] as Record<string, unknown>;
}

describe('serializeAnimations', () => {
  it('returns undefined for undefined input', () => {
    expect(serializeAnimations(undefined)).toBeUndefined();
  });

  it('returns undefined for an empty array', () => {
    expect(serializeAnimations([])).toBeUndefined();
  });

  it('returns a timing tree object for a non-empty array', () => {
    const result = serializeAnimations([makeAnim()]);
    expect(result).toBeDefined();
    expect(result).toHaveProperty('p:tnLst');
  });

  it('produces the standard tnLst > par > cTn wrapper structure', () => {
    const result = serializeAnimations([makeAnim()]) as Record<string, unknown>;
    const tnLst = result['p:tnLst'] as Record<string, unknown>;
    expect(tnLst).toHaveProperty('p:par');
    const par = tnLst['p:par'] as Record<string, unknown>;
    const cTn = par['p:cTn'] as Record<string, unknown>;
    expect(cTn['@_dur']).toBe('indefinite');
    expect(cTn['@_id']).toBe('1');
    expect(cTn['@_nodeType']).toBe('tmRoot');
    expect(cTn['@_restart']).toBe('never');
  });

  it('maps onClick trigger to clickEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'onClick' })]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('clickEffect');
  });

  it('maps withPrevious trigger to withEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'withPrevious' })]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('withEffect');
  });

  it('maps afterPrevious trigger to afterEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'afterPrevious' })]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('afterEffect');
  });

  it('maps any other trigger value to clickEffect', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'somethingElse' })]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('clickEffect');
  });

  it('defaults duration to 500 when not specified', () => {
    const result = serializeAnimations([makeAnim()]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_dur']).toBe(500);
  });

  it('uses custom duration when specified', () => {
    const result = serializeAnimations([makeAnim({ duration: 1200 })]) as Record<string, unknown>;
    expect(getFirstCTn(getAnimNodes(result))['@_dur']).toBe(1200);
  });

  it('defaults delay to 0 when not specified', () => {
    const result = serializeAnimations([makeAnim()]) as Record<string, unknown>;
    const cTn = getFirstCTn(getAnimNodes(result));
    const stCond = (cTn['p:stCondLst'] as Record<string, unknown>)['p:cond'] as Record<string, unknown>;
    expect(stCond['@_delay']).toBe(0);
  });

  it('uses custom delay when specified', () => {
    const result = serializeAnimations([makeAnim({ delay: 750 })]) as Record<string, unknown>;
    const cTn = getFirstCTn(getAnimNodes(result));
    const stCond = (cTn['p:stCondLst'] as Record<string, unknown>)['p:cond'] as Record<string, unknown>;
    expect(stCond['@_delay']).toBe(750);
  });

  it('propagates targetShapeId into the spTgt node', () => {
    const result = serializeAnimations([makeAnim({ targetShapeId: 'sp-42' })]) as Record<string, unknown>;
    const nodes = getAnimNodes(result);
    const cMediaNode = (nodes[0] as Record<string, unknown>)['p:cMediaNode'] as Record<string, unknown>;
    const cTn = cMediaNode['p:cTn'] as Record<string, unknown>;
    const set = (cTn['p:childTnLst'] as Record<string, unknown>)['p:set'] as Record<string, unknown>;
    const spTgt = ((set['p:cBhvr'] as Record<string, unknown>)['p:tgtEl'] as Record<string, unknown>)['p:spTgt'] as Record<string, unknown>;
    expect(spTgt['@_spid']).toBe('sp-42');
  });

  it('assigns sequential @_id (1-based) to each animation node', () => {
    const result = serializeAnimations([
      makeAnim({ targetShapeId: 'a' }),
      makeAnim({ targetShapeId: 'b' }),
      makeAnim({ targetShapeId: 'c' }),
    ]) as Record<string, unknown>;
    const nodes = getAnimNodes(result);
    expect(nodes).toHaveLength(3);
    nodes.forEach((node, i) => {
      const cTn = getFirstCTn([node]);
      expect(cTn['@_id']).toBe(i + 1);
    });
  });

  it('uses (idx+1)*10 as inner cTn @_id for shape behavior', () => {
    const result = serializeAnimations([makeAnim(), makeAnim()]) as Record<string, unknown>;
    const nodes = getAnimNodes(result);
    [10, 20].forEach((expectedId, i) => {
      const cMediaNode = (nodes[i] as Record<string, unknown>)['p:cMediaNode'] as Record<string, unknown>;
      const outerCTn = cMediaNode['p:cTn'] as Record<string, unknown>;
      const set = (outerCTn['p:childTnLst'] as Record<string, unknown>)['p:set'] as Record<string, unknown>;
      const innerCTn = (set['p:cBhvr'] as Record<string, unknown>)['p:cTn'] as Record<string, unknown>;
      expect(innerCTn['@_id']).toBe(expectedId);
    });
  });
});
