import { describe, expect, it } from 'vitest';
import type { PptxAnimation } from '@hokkyss/pptx-core';
import type { XmlElement } from '../../lib/xml/xml-element';
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

/**
 *
 */
function findElement(root: undefined | XmlElement, tag: string): undefined | XmlElement {
  if (!root) return undefined;
  if (root.tag === tag) return root;
  if (root.children) {
    for (const child of root.children) {
      if (typeof child === 'object' && 'tag' in child) {
        const found = findElement(child, tag);
        if (found) return found;
      }
    }
  }
  return undefined;
}

/**
 *
 */
function findElements(root: undefined | XmlElement, tag: string): XmlElement[] {
  const result: XmlElement[] = [];
  if (!root) return result;
  if (root.tag === tag) result.push(root);
  if (root.children) {
    for (const child of root.children) {
      if (typeof child === 'object' && 'tag' in child) {
        result.push(...findElements(child, tag));
      }
    }
  }
  return result;
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
    expect(result?.tag).toBe('p:timing');
    expect(findElement(result, 'p:tnLst')).toBeDefined();
  });

  it('produces the standard tnLst > par > cTn wrapper structure', () => {
    const result = serializeAnimations([makeAnim()]);
    const par = findElement(result, 'p:par');
    expect(par).toBeDefined();
    const cTn = findElement(par, 'p:cTn');
    expect(cTn?.attrs?.dur).toBe('indefinite');
    expect(cTn?.attrs?.id).toBe('1');
    expect(cTn?.attrs?.nodeType).toBe('tmRoot');
    expect(cTn?.attrs?.restart).toBe('never');
  });

  it('maps onClick trigger to clickEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'onClick' })]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.nodeType).toBe('clickEffect');
  });

  it('maps withPrevious trigger to withEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'withPrevious' })]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.nodeType).toBe('withEffect');
  });

  it('maps afterPrevious trigger to afterEffect nodeType', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'afterPrevious' })]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.nodeType).toBe('afterEffect');
  });

  it('maps any other trigger value to clickEffect', () => {
    const result = serializeAnimations([makeAnim({ trigger: 'somethingElse' })]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.nodeType).toBe('clickEffect');
  });

  it('defaults duration to 500 when not specified', () => {
    const result = serializeAnimations([makeAnim()]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.dur).toBe(500);
  });

  it('uses custom duration when specified', () => {
    const result = serializeAnimations([makeAnim({ duration: 1200 })]);
    const mediaNode = findElement(result, 'p:cMediaNode');
    const cTn = findElement(mediaNode, 'p:cTn');
    expect(cTn?.attrs?.dur).toBe(1200);
  });

  it('defaults delay to 0 when not specified', () => {
    const result = serializeAnimations([makeAnim()]);
    const cond = findElement(result, 'p:cond');
    expect(cond?.attrs?.delay).toBe(0);
  });

  it('uses custom delay when specified', () => {
    const result = serializeAnimations([makeAnim({ delay: 750 })]);
    const cond = findElement(result, 'p:cond');
    expect(cond?.attrs?.delay).toBe(750);
  });

  it('propagates targetShapeId into the spTgt node', () => {
    const result = serializeAnimations([makeAnim({ targetShapeId: 'sp-42' })]);
    const spTgt = findElement(result, 'p:spTgt');
    expect(spTgt?.attrs?.spid).toBe('sp-42');
  });

  it('assigns sequential @_id (1-based) to each animation node', () => {
    const result = serializeAnimations([
      makeAnim({ targetShapeId: 'a' }),
      makeAnim({ targetShapeId: 'b' }),
      makeAnim({ targetShapeId: 'c' }),
    ]);
    const mediaNodes = findElements(result, 'p:cMediaNode');
    expect(mediaNodes).toHaveLength(3);
    mediaNodes.forEach((node, i) => {
      const cTn = findElement(node, 'p:cTn');
      expect(cTn?.attrs?.id).toBe(i + 1);
    });
  });

  it('uses (idx+1)*10 as inner cTn @_id for shape behavior', () => {
    const result = serializeAnimations([makeAnim(), makeAnim()]);
    const mediaNodes = findElements(result, 'p:cMediaNode');
    [10, 20].forEach((expectedId, i) => {
      const cBhvr = findElement(mediaNodes[i], 'p:cBhvr');
      const innerCTn = findElement(cBhvr, 'p:cTn');
      expect(innerCTn?.attrs?.id).toBe(expectedId);
    });
  });
});
