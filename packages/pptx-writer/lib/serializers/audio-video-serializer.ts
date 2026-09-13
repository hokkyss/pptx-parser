import type { PptxAnimation, PptxAudioElement, PptxMediaPlayback, PptxVideoElement } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';

/**
 * P14 namespace URI required for media and trim/seek extensions.
 * @see https://docs.microsoft.com/office/open-xml/presentation#media-extensions
 */
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/**
 * Well-known Microsoft Office extension GUID for embedded media elements.
 */
const MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';

/**
 * Converts a playback volume (0–1) to OpenXML thousandths (0–100000).
 */
function volumeToOpenXml(v: number): number {
  return Math.round(Math.max(0, Math.min(1, v)) * 100000);
}

export interface MediaTimingEntry {
  id: string;
  mediaType: 'audio' | 'video';
  muted?: boolean;
  playback?: PptxMediaPlayback;
}

/**
 * Builds the `<p:timing>` element tree combining slide animations and media elements.
 * Follows strict ECMA-376 Part 1 schema element ordering for CT_TLCommonTimeNodeData:
 * sequence: (stCondLst?, endCondLst?, endSync?, iterate?, childTnLst?, subTnLst?)
 * @param animations - Optional slide animation definitions.
 * @param mediaElements - Optional audio/video media elements on the slide.
 */
export function buildSlideTiming(
  animations?: PptxAnimation[],
  mediaElements?: MediaTimingEntry[],
): undefined | XmlElement {
  const hasAnimations = animations && animations.length > 0;
  const hasMedia = mediaElements && mediaElements.length > 0;

  if (!hasAnimations && !hasMedia) return undefined;

  let nextId = 1;
  const rootId = nextId++;
  const mainSeqId = nextId++;

  const mainSeqChildren: XmlElement[] = [];
  const interactiveSeqList: XmlElement[] = [];

  // 1. Standard shape animations (clickEffect, withEffect, afterEffect)
  if (hasAnimations) {
    for (const anim of animations) {
      const animCtnId = nextId++;
      mainSeqChildren.push(
        el('p:cMediaNode', [
          el('p:cTn', {
            dur: anim.duration ?? 500,
            id: animCtnId,
            nodeType: anim.trigger === 'afterPrevious' ? 'afterEffect' : anim.trigger === 'withPrevious' ? 'withEffect' : 'clickEffect',
          }, [
            el('p:stCondLst', [
              el('p:cond', { delay: anim.delay ?? 0 }),
            ]),
            el('p:childTnLst', [
              el('p:set', [
                el('p:cBhvr', [
                  el('p:cTn', { dur: 1, id: animCtnId * 10 }),
                  el('p:tgtEl', [
                    el('p:spTgt', { spid: anim.targetShapeId }),
                  ]),
                ]),
              ]),
            ]),
          ]),
        ]),
      );
    }
  }

  // 2. Media elements (audio and video)
  if (hasMedia) {
    for (const media of mediaElements) {
      const trigger = media.playback?.trigger ?? 'onClick';
      const loop = media.playback?.loop ?? false;
      const vol = media.muted ? 0 : (media.playback?.volume !== undefined ? volumeToOpenXml(media.playback.volume) : 80000);

      const activeMediaCtnId = nextId++;
      const endCondLst = el('p:endCondLst', [
        el('p:cond', { delay: '0', evt: 'onStopAudio' }),
      ]);

      const cTnAttrs: Record<string, number | string> = {
        dur: 'indefinite',
        id: activeMediaCtnId,
        nodeType: trigger === 'automatic' ? 'withEffect' : 'clickEffect',
        restart: 'never',
      };
      if (loop) {
        cTnAttrs.repeatCount = 'indefinite';
      }

      const cMediaNodeChildren: XmlElement[] = [
        el('p:cTn', cTnAttrs, [
          el('p:stCondLst', [
            el('p:cond', { delay: '0' }),
          ]),
          endCondLst,
        ]),
        el('p:tgtEl', [
          el('p:spTgt', { spid: String(media.id) }),
        ]),
      ];

      if (media.playback?.startTime !== undefined || media.playback?.endTime !== undefined) {
        const trimAttrs: Record<string, number | string> = {
          'xmlns:p14': P14_NS,
        };
        if (media.playback?.startTime !== undefined) trimAttrs.st = media.playback.startTime;
        if (media.playback?.endTime !== undefined) trimAttrs.end = media.playback.endTime;
        cMediaNodeChildren.push(
          el('p:extLst', [
            el('p:ext', { uri: MEDIA_EXT_URI }, [
              el('p14:trim', trimAttrs),
            ]),
          ]),
        );
      }

      const cMediaNode = el('p:cMediaNode', { vol }, cMediaNodeChildren);
      const mediaTag = media.mediaType === 'video' ? 'p:video' : 'p:audio';

      if (trigger === 'automatic') {
        mainSeqChildren.push(
          el(mediaTag, [cMediaNode]),
        );
      } else {
        const interactiveSeqId = nextId++;
        interactiveSeqList.push(
          el('p:seq', {
            concurrent: '1',
            nextAc: 'seek',
          }, [
            el('p:cTn', {
              evtFilter: 'cancelBubble',
              fill: 'hold',
              id: interactiveSeqId,
              nodeType: 'interactiveSeq',
              restart: 'whenNotActive',
            }, [
              el('p:stCondLst', [
                el('p:cond', {
                  delay: '0',
                  evt: 'onClick',
                }, [
                  el('p:tgtEl', [
                    el('p:spTgt', { spid: String(media.id) }),
                  ]),
                ]),
              ]),
              el('p:childTnLst', [
                el(mediaTag, [cMediaNode]),
              ]),
            ]),
          ]),
        );
      }
    }
  }

  const seqList: XmlElement[] = [];
  if (mainSeqChildren.length > 0) {
    seqList.push(
      el('p:seq', {
        concurrent: '1',
        nextAc: 'seek',
      }, [
        el('p:cTn', {
          dur: 'indefinite',
          id: mainSeqId,
          nodeType: 'mainSeq',
        }, [
          el('p:childTnLst', mainSeqChildren),
        ]),
      ]),
    );
  }
  for (const interactiveSeq of interactiveSeqList) {
    seqList.push(interactiveSeq);
  }

  return el('p:timing', [
    el('p:tnLst', [
      el('p:par', [
        el('p:cTn', {
          dur: 'indefinite',
          id: rootId,
          nodeType: 'tmRoot',
          restart: 'never',
        }, [
          el('p:childTnLst', seqList),
        ]),
      ]),
    ]),
  ]);
}

/**
 * Serializes a `PptxAudioElement` into an OpenXML `<p:pic>` element.
 * The `<p:nvPr>` block includes `<a:audioFile r:link="rIdN"/>` so PowerPoint
 * recognises the element as embedded audio and renders its native speaker icon.
 * @param element - The audio element AST node.
 * @param linkRelId - The `r:link` relationship ID pointing to the audio file.
 * @param embedRelId - The `r:embed` relationship ID for the `p14:media` extension.
 * @param imageRelId - Optional image relationship ID for the poster frame blip.
 * @returns XmlElement representing `<p:pic>`.
 */
export function serializeAudio(
  element: PptxAudioElement,
  linkRelId: string,
  embedRelId: string,
  imageRelId?: string,
): XmlElement {
  const xfrm = el('a:xfrm', [
    el('a:off', {
      x: Math.round(Number(element.position?.x ?? 0)),
      y: Math.round(Number(element.position?.y ?? 0)),
    }),
    el('a:ext', {
      cx: Math.round(Number(element.position?.cx ?? 914400)),
      cy: Math.round(Number(element.position?.cy ?? 914400)),
    }),
  ]);

  const blipAttrs: Record<string, string> = {};
  if (imageRelId) {
    blipAttrs['r:embed'] = imageRelId;
  }

  const nvPicPr = el('p:nvPicPr', [
    el('p:cNvPr', {
      id: element.id,
      name: element.name,
    }),
    el('p:cNvPicPr', [
      el('a:picLocks', { noChangeAspect: '1' }),
    ]),
    el('p:nvPr', [
      el('a:audioFile', { 'r:link': linkRelId }),
      el('p:extLst', [
        el('p:ext', { uri: MEDIA_EXT_URI }, [
          el('p14:media', {
            'xmlns:p14': P14_NS,
            'r:embed': embedRelId,
          }),
        ]),
      ]),
    ]),
  ]);

  const blipFill = el('p:blipFill', [
    el('a:blip', blipAttrs),
    el('a:stretch', [el('a:fillRect')]),
  ]);

  const spPr = el('p:spPr', [
    xfrm,
    el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]),
    el('a:noFill'),
    el('a:ln', [el('a:noFill')]),
  ]);

  return el('p:pic', [nvPicPr, blipFill, spPr]);
}

/**
 * Serializes a `PptxVideoElement` into an OpenXML `<p:pic>` element.
 * The `<p:nvPr>` block includes `<a:videoFile r:link="rIdN"/>` so PowerPoint
 * recognises the element as embedded video and renders its native film strip placeholder.
 * @param element - The video element AST node.
 * @param linkRelId - The `r:link` relationship ID pointing to the video file.
 * @param embedRelId - The `r:embed` relationship ID for the `p14:media` extension.
 * @param imageRelId - Optional image relationship ID for the placeholder blip.
 * @returns XmlElement representing `<p:pic>`.
 */
export function serializeVideo(
  element: PptxVideoElement,
  linkRelId: string,
  embedRelId: string,
  imageRelId?: string,
): XmlElement {
  const xfrm = el('a:xfrm', [
    el('a:off', {
      x: Math.round(Number(element.position?.x ?? 0)),
      y: Math.round(Number(element.position?.y ?? 0)),
    }),
    el('a:ext', {
      cx: Math.round(Number(element.position?.cx ?? 3657600)),
      cy: Math.round(Number(element.position?.cy ?? 2743200)),
    }),
  ]);

  const blipAttrs: Record<string, string> = {};
  if (imageRelId) {
    blipAttrs['r:embed'] = imageRelId;
  }

  const nvPicPr = el('p:nvPicPr', [
    el('p:cNvPr', {
      id: element.id,
      name: element.name,
    }),
    el('p:cNvPicPr', [
      el('a:picLocks', { noChangeAspect: '1' }),
    ]),
    el('p:nvPr', [
      el('a:videoFile', { 'r:link': linkRelId }),
      el('p:extLst', [
        el('p:ext', { uri: MEDIA_EXT_URI }, [
          el('p14:media', {
            'xmlns:p14': P14_NS,
            'r:embed': embedRelId,
          }),
        ]),
      ]),
    ]),
  ]);

  const blipFill = el('p:blipFill', [
    el('a:blip', blipAttrs),
    el('a:stretch', [el('a:fillRect')]),
  ]);

  const spPr = el('p:spPr', [
    xfrm,
    el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]),
    el('a:noFill'),
    el('a:ln', [el('a:noFill')]),
  ]);

  return el('p:pic', [nvPicPr, blipFill, spPr]);
}
