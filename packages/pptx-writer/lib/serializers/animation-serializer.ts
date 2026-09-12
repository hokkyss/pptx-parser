import type { PptxAnimation } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';

/**
 * Serializes slide `<p:timing>` element tree from animations.
 * @param animations
 */
export function serializeAnimations(animations?: PptxAnimation[]): undefined | XmlElement {
  if (!animations || animations.length === 0) return undefined;

  const animNodes = animations.map((anim, idx) =>
    el('p:cMediaNode', [
      el('p:cTn', {
        dur: anim.duration ?? 500,
        id: idx + 1,
        nodeType: anim.trigger === 'afterPrevious' ? 'afterEffect' : anim.trigger === 'withPrevious' ? 'withEffect' : 'clickEffect',
      }, [
        el('p:stCondLst', [
          el('p:cond', { delay: anim.delay ?? 0 }),
        ]),
        el('p:childTnLst', [
          el('p:set', [
            el('p:cBhvr', [
              el('p:cTn', { dur: 1, id: (idx + 1) * 10 }),
              el('p:tgtEl', [
                el('p:spTgt', { spid: anim.targetShapeId }),
              ]),
            ]),
          ]),
        ]),
      ]),
    ]),
  );

  return el('p:timing', [
    el('p:tnLst', [
      el('p:par', [
        el('p:cTn', {
          dur: 'indefinite',
          id: '1',
          nodeType: 'tmRoot',
          restart: 'never',
        }, [
          el('p:childTnLst', [
            el('p:seq', {
              concurrent: '1',
              nextAc: 'seek',
            }, [
              el('p:cTn', {
                dur: 'indefinite',
                id: '2',
                nodeType: 'mainSeq',
              }, [
                el('p:childTnLst', animNodes),
              ]),
            ]),
          ]),
        ]),
      ]),
    ]),
  ]);
}
