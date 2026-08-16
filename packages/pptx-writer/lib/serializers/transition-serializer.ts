import type { PptxTransition } from '@hokkyss/pptx-core';

/**
 * Serializes slide `<p:transition>`.
 * @param transition
 */
export function serializeTransition(transition?: PptxTransition): Record<string, unknown> | undefined {
  if (!transition) return undefined;

  const transNode: Record<string, unknown> = {};
  if (transition.speed) {
    transNode['@_spd'] = transition.speed;
  }
  if (transition.duration !== undefined) {
    transNode['@_dur'] = transition.duration;
  }
  if (transition.advanceOnClick !== undefined) {
    transNode['@_advClick'] = transition.advanceOnClick ? '1' : '0';
  }
  if (transition.advanceAfterMs !== undefined) {
    transNode['@_advTm'] = transition.advanceAfterMs;
  }

  const type = transition.type || 'fade';
  transNode[`p:${type}`] = {};

  return transNode;
}
