import type { PptxTransition, PptxTransitionSpeed } from '@hokkyss/pptx-core';

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
 * @returns Transition XML node object or undefined.
 */
export function serializeTransition(transition?: PptxTransition): Record<string, unknown> | undefined {
  if (!transition) return undefined;

  const type = (transition.type || 'fade');
  if (type === 'none') {
    return undefined;
  }

  const transNode: Record<string, unknown> = {};

  const speed = resolveSpeed(transition.speed, transition.durationMs ?? transition.duration);
  if (speed) {
    transNode['@_spd'] = speed;
  }

  if (transition.duration !== undefined) {
    transNode['@_dur'] = transition.duration;
  }

  if (transition.advanceOnClick !== undefined) {
    transNode['@_advClick'] = transition.advanceOnClick ? '1' : '0';
  }

  if (transition.advanceAfterMs !== undefined) {
    transNode['@_advTm'] = Math.round(transition.advanceAfterMs);
  }

  const dirAttr = transition.direction ? (DIRECTION_MAP[transition.direction] ?? transition.direction) : undefined;
  const childNode: Record<string, unknown> = {};

  switch (type) {
    case 'blinds': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childNode['@_dir'] = dirAttr;
      }
      transNode['p:blinds'] = childNode;
      break;
    }
    case 'checker': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childNode['@_dir'] = dirAttr;
      }
      transNode['p:checker'] = childNode;
      break;
    }
    case 'comb': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childNode['@_dir'] = dirAttr;
      }
      transNode['p:comb'] = childNode;
      break;
    }
    case 'cover': {
      if (dirAttr) childNode['@_dir'] = dirAttr;
      transNode['p:cover'] = childNode;
      break;
    }
    case 'fade': {
      if (transition.throughBlack) {
        childNode['@_thruBlk'] = '1';
      }
      transNode['p:fade'] = childNode;
      break;
    }
    case 'pull': {
      if (dirAttr) childNode['@_dir'] = dirAttr;
      transNode['p:pull'] = childNode;
      break;
    }
    case 'push': {
      if (dirAttr) childNode['@_dir'] = dirAttr;
      transNode['p:push'] = childNode;
      break;
    }
    case 'randomBar': {
      if (dirAttr === 'horz' || dirAttr === 'vert') {
        childNode['@_dir'] = dirAttr;
      }
      transNode['p:randomBar'] = childNode;
      break;
    }
    case 'split': {
      if (dirAttr === 'in' || dirAttr === 'out') {
        childNode['@_dir'] = dirAttr;
      }
      if (transition.direction === 'horz' || transition.direction === 'vert') {
        childNode['@_orient'] = transition.direction;
      }
      transNode['p:split'] = childNode;
      break;
    }
    case 'wheel': {
      if (transition.spokes !== undefined) {
        childNode['@_spokes'] = String(transition.spokes);
      }
      transNode['p:wheel'] = childNode;
      break;
    }
    case 'wipe': {
      if (dirAttr) childNode['@_dir'] = dirAttr;
      transNode['p:wipe'] = childNode;
      break;
    }
    case 'zoom': {
      if (dirAttr === 'in' || dirAttr === 'out') {
        childNode['@_dir'] = dirAttr;
      }
      transNode['p:zoom'] = childNode;
      break;
    }
    default: {
      transNode[`p:${type}`] = childNode;
      break;
    }
  }

  return transNode;
}
