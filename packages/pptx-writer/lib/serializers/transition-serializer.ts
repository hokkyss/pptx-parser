import type { PptxTransition, PptxTransitionSpeed } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';

const DIRECTION_MAP: Record<string, string> = {
  down: 'd',
  horz: 'horz',
  in: 'in',
  left: 'l',
  out: 'out',
  right: 'r',
  up: 'u',
  vert: 'vert',
};

/**
 * Maps duration in milliseconds to OpenXML transition speed enum if not explicitly given.
 */
function resolveSpeed(speed?: PptxTransitionSpeed, durationMs?: number): string | undefined {
  if (speed) {
    if (speed === 'medium') return 'med';
    return speed;
  }
  if (durationMs !== undefined) {
    if (durationMs <= 750) return 'fast';
    if (durationMs <= 1500) return 'med';
    return 'slow';
  }
  return undefined;
}

/**
 * Serializes slide `<p:transition>` conforming strictly to ECMA-376 PresentationML schema.
 * @param transition Slide transition configuration.
 * @returns Transition XmlElement node or undefined.
 */
export function serializeTransition(transition?: PptxTransition): undefined | XmlElement {
  if (!transition) return undefined;

  const type = (transition.type || 'fade');
  if (type === 'none') {
    return undefined;
  }

  const transAttrs: Record<string, number | string | undefined> = {};

  const speed = resolveSpeed(transition.speed, transition.durationMs ?? transition.duration);
  if (speed) {
    transAttrs.spd = speed;
  }

  if (transition.duration !== undefined) {
    transAttrs.dur = transition.duration;
  }

  if (transition.advanceOnClick !== undefined) {
    transAttrs.advClick = transition.advanceOnClick ? '1' : '0';
  }

  if (transition.advanceAfterMs !== undefined) {
    transAttrs.advTm = Math.round(transition.advanceAfterMs);
  }

  const dirAttr = transition.direction ? (DIRECTION_MAP[transition.direction] ?? transition.direction) : undefined;
  const childAttrs: Record<string, string | undefined> = {};

  switch (type) {
    case 'blinds': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childAttrs.dir = dirAttr;
      }
      break;
    }
    case 'checker': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childAttrs.dir = dirAttr;
      }
      break;
    }
    case 'comb': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childAttrs.dir = dirAttr;
      }
      break;
    }
    case 'cover': {
      if (dirAttr) childAttrs.dir = dirAttr;
      break;
    }
    case 'fade': {
      if (transition.throughBlack) {
        childAttrs.thruBlk = '1';
      }
      break;
    }
    case 'pull': {
      if (dirAttr) childAttrs.dir = dirAttr;
      break;
    }
    case 'push': {
      if (dirAttr) childAttrs.dir = dirAttr;
      break;
    }
    case 'randomBar': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childAttrs.dir = dirAttr;
      }
      break;
    }
    case 'split': {
      if (dirAttr === 'in' || dirAttr === 'out') {
        childAttrs.dir = dirAttr;
      }
      if (transition.direction === 'horz' || transition.direction === 'vert') {
        childAttrs.orient = transition.direction;
      }
      break;
    }
    case 'wheel': {
      if (transition.spokes !== undefined) {
        childAttrs.spokes = String(transition.spokes);
      }
      break;
    }
    case 'wipe': {
      if (dirAttr) childAttrs.dir = dirAttr;
      break;
    }
    case 'zoom': {
      if (dirAttr === 'in' || dirAttr === 'out') {
        childAttrs.dir = dirAttr;
      }
      break;
    }
  }

  const childNode = el(`p:${type}`, childAttrs);
  return el('p:transition', transAttrs, [childNode]);
}
