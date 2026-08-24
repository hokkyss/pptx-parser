import { describe, expect, it } from 'vitest';
import type { PptxAnimation } from '@hokkyss/pptx-core';
import { serializeAnimations } from '../../lib/serializers/animation-serializer';

/**
 * Helper to construct a test PptxAnimation AST element.
 */
function makeAnim(overrides: Partial<PptxAnimation> = {}): PptxAnimation {
  return {
    effect: 'appear',
    sequence: 0,
    targetShapeId: '10',
    trigger: 'onClick',
    ...overrides,
  };
}

interface AnimationInnerBehavior {
  'p:cBhvr'?: {
    'p:cTn'?: {
      '@_id'?: number;
    };
    'p:tgtEl'?: {
      'p:spTgt'?: {
        '@_spid'?: string;
      };
    };
  };
}

interface AnimationCTnNode {
  '@_dur'?: number | string;
  '@_id'?: number | string;
  '@_nodeType'?: string;
  '@_restart'?: string;
  'p:childTnLst'?: {
    'p:set'?: AnimationInnerBehavior;
  };
  'p:stCondLst'?: {
    'p:cond'?: {
      '@_delay'?: number | string;
      '@_evt'?: string;
    };
  };
}

interface AnimationNodeItem {
  'p:cMediaNode'?: {
    'p:cTn'?: AnimationCTnNode;
  };
}

interface AnimationTimingTree {
  'p:tnLst'?: {
    'p:par'?: {
      'p:cTn'?: {
        '@_dur'?: string;
        '@_id'?: string;
        '@_nodeType'?: string;
        '@_restart'?: string;
        'p:childTnLst'?: {
          'p:seq'?: {
            'p:cTn'?: {
              'p:childTnLst'?: AnimationNodeItem[];
            };
          };
        };
      };
    };
  };
}

/**
 * Extracts animation node array from serialized timing tree.
 */
function getAnimNodes(result: AnimationTimingTree | undefined): AnimationNodeItem[] {
  return result?.['p:tnLst']?.['p:par']?.['p:cTn']?.['p:childTnLst']?.['p:seq']?.['p:cTn']?.['p:childTnLst'] || [];
}

/**
 * Retrieves the first cTn node from an animation node list.
 */
function getFirstCTn(nodes: AnimationNodeItem[]): AnimationCTnNode {
  return nodes[0]?.['p:cMediaNode']?.['p:cTn'] || {};
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
    const result = serializeAnimations([makeAnim()]) as AnimationTimingTree;
    const tnLst = result['p:tnLst'];
    expect(tnLst).toHaveProperty('p:par');
    const cTn = tnLst?.['p:par']?.['p:cTn'];
    expect(cTn?.['@_dur']).toBe('indefinite');
    expect(cTn?.['@_id']).toBe('1');
    expect(cTn?.['@_nodeType']).toBe('tmRoot');
    expect(cTn?.['@_restart']).toBe('never');
  });

  it('maps onClick trigger to clickEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'onClick' })]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('clickEffect');
  });

  it('maps withPrevious trigger to withEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'withPrevious' })]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('withEffect');
  });

  it('maps afterPrevious trigger to afterEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'afterPrevious' })]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('afterEffect');
  });

  it('maps any other trigger value to clickEffect', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'somethingElse' })]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_nodeType']).toBe('clickEffect');
  });

  it('defaults duration to 500 when not specified', () => {
    const result = serializeAnimations([makeAnim()]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_dur']).toBe(500);
  });

  it('uses custom duration when specified', () => {
    const result = serializeAnimations([makeAnim({ duration: 1200 })]) as AnimationTimingTree;
    expect(getFirstCTn(getAnimNodes(result))['@_dur']).toBe(1200);
  });

  it('defaults delay to 0 when not specified', () => {
    const result = serializeAnimations([makeAnim()]) as AnimationTimingTree;
    const cTn = getFirstCTn(getAnimNodes(result));
    const stCond = cTn['p:stCondLst']?.['p:cond'];
    expect(stCond?.['@_delay']).toBe(0);
  });

  it('uses custom delay when specified', () => {
    const result = serializeAnimations([makeAnim({ delay: 750 })]) as AnimationTimingTree;
    const cTn = getFirstCTn(getAnimNodes(result));
    const stCond = cTn['p:stCondLst']?.['p:cond'];
    expect(stCond?.['@_delay']).toBe(750);
  });

  it('propagates targetShapeId into the spTgt node', () => {
    const result = serializeAnimations([makeAnim({ targetShapeId: 'sp-42' })]) as AnimationTimingTree;
    const nodes = getAnimNodes(result);
    const set = nodes[0]?.['p:cMediaNode']?.['p:cTn']?.['p:childTnLst']?.['p:set'];
    const spTgt = set?.['p:cBhvr']?.['p:tgtEl']?.['p:spTgt'];
    expect(spTgt?.['@_spid']).toBe('sp-42');
  });

  it('assigns sequential @_id (1-based) to each animation node', () => {
    const result = serializeAnimations([
      makeAnim({ targetShapeId: 'a' }),
      makeAnim({ targetShapeId: 'b' }),
      makeAnim({ targetShapeId: 'c' }),
    ]) as AnimationTimingTree;
    const nodes = getAnimNodes(result);
    expect(nodes).toHaveLength(3);
    nodes.forEach((node, i) => {
      const cTn = getFirstCTn([node]);
      expect(cTn['@_id']).toBe(i + 1);
    });
  });

  it('uses (idx+1)*10 as inner cTn @_id for shape behavior', () => {
    const result = serializeAnimations([makeAnim(), makeAnim()]) as AnimationTimingTree;
    const nodes = getAnimNodes(result);
    [10, 20].forEach((expectedId, i) => {
      const set = nodes[i]?.['p:cMediaNode']?.['p:cTn']?.['p:childTnLst']?.['p:set'];
      const innerCTn = set?.['p:cBhvr']?.['p:cTn'];
      expect(innerCTn?.['@_id']).toBe(expectedId);
    });
  });
});
