import type { PptxAnimation } from '@hokkyss/pptx-core';

/**
 * Serializes slide `<p:timing>` element tree from animations.
 * @param animations
 */
export function serializeAnimations(animations?: PptxAnimation[]): Record<string, unknown> | undefined {
  if (!animations || animations.length === 0) return undefined;

  const animNodes = animations.map((anim, idx) => ({
    'p:cMediaNode': {
      'p:cTn': {
        '@_dur': anim.duration ?? 500,
        '@_id': idx + 1,
        '@_nodeType': anim.trigger === 'afterPrevious' ? 'afterEffect' : anim.trigger === 'withPrevious' ? 'withEffect' : 'clickEffect',
        'p:stCondLst': {
          'p:cond': {
            '@_delay': anim.delay ?? 0,
          },
        },
        'p:childTnLst': {
          'p:set': {
            'p:cBhvr': {
              'p:cTn': {
                '@_dur': 1,
                '@_id': (idx + 1) * 10,
              },
              'p:tgtEl': {
                'p:spTgt': {
                  '@_spid': anim.targetShapeId,
                },
              },
            },
          },
        },
      },
    },
  }));

  return {
    'p:tnLst': {
      'p:par': {
        'p:cTn': {
          '@_dur': 'indefinite',
          '@_id': '1',
          '@_nodeType': 'tmRoot',
          '@_restart': 'never',
          'p:childTnLst': {
            'p:seq': {
              '@_concurrent': '1',
              '@_nextAc': 'seek',
              'p:cTn': {
                '@_dur': 'indefinite',
                '@_id': '2',
                '@_nodeType': 'mainSeq',
                'p:childTnLst': animNodes,
              },
            },
          },
        },
      },
    },
  };
}
