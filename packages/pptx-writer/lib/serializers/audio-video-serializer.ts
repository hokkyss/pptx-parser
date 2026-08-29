import type { PptxAnimation, PptxAudioElement, PptxMediaPlayback, PptxVideoElement } from '@hokkyss/pptx-core';

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

/**
 * Builds the `<p:timing>` element tree combining slide animations and media elements.
 * Follows strict ECMA-376 Part 1 schema element ordering for CT_TLCommonTimeNodeData:
 * sequence: (stCondLst?, endCondLst?, endSync?, iterate?, childTnLst?, subTnLst?)
 * @param animations - Optional slide animation definitions.
 * @param mediaElements - Optional audio/video media elements on the slide.
 */
export function buildSlideTiming(
  animations?: PptxAnimation[],
  mediaElements?: Array<{ id: string; mediaType: 'audio' | 'video'; muted?: boolean; playback?: PptxMediaPlayback }>,
): Record<string, unknown> | undefined {
  const hasAnimations = animations && animations.length > 0;
  const hasMedia = mediaElements && mediaElements.length > 0;

  if (!hasAnimations && !hasMedia) return undefined;

  let nextId = 1;
  const rootId = String(nextId++);
  const mainSeqId = String(nextId++);

  const mainSeqChildren: Record<string, unknown>[] = [];
  const interactiveSeqList: Record<string, unknown>[] = [];

  // 1. Standard shape animations (clickEffect, withEffect, afterEffect)
  if (hasAnimations) {
    for (const anim of animations) {
      const animCtnId = nextId++;
      mainSeqChildren.push({
        'p:cMediaNode': {
          'p:cTn': {
            '@_dur': anim.duration ?? 500,
            '@_id': animCtnId,
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
                    '@_id': animCtnId * 10,
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
      });
    }
  }

  // 2. Media elements (audio and video)
  if (hasMedia) {
    for (const media of mediaElements) {
      const trigger = media.playback?.trigger ?? 'onClick';
      const loop = media.playback?.loop ?? false;
      const vol = media.muted ? 0 : (media.playback?.volume !== undefined ? volumeToOpenXml(media.playback.volume) : 80000);

      const activeMediaCtnId = nextId++;
      const endCondLst: Record<string, unknown> = {
        'p:cond': { '@_delay': '0', '@_evt': 'onStopAudio' },
      };

      const cMediaNode: Record<string, unknown> = {
        '@_vol': vol,
        'p:cTn': {
          '@_dur': 'indefinite',
          '@_id': String(activeMediaCtnId),
          '@_nodeType': trigger === 'automatic' ? 'withEffect' : 'clickEffect',
          '@_restart': 'never',
          'p:stCondLst': {
            'p:cond': { '@_delay': '0' },
          },
          'p:endCondLst': endCondLst,
          ...(loop ? { '@_repeatCount': 'indefinite' } : {}),
        },
        'p:tgtEl': {
          'p:spTgt': { '@_spid': String(media.id) },
        },
      };

      if (media.playback?.startTime !== undefined || media.playback?.endTime !== undefined) {
        const trimAttrs: Record<string, unknown> = {};
        if (media.playback?.startTime !== undefined) trimAttrs['@_st'] = media.playback.startTime;
        if (media.playback?.endTime !== undefined) trimAttrs['@_end'] = media.playback.endTime;
        cMediaNode['p:extLst'] = {
          'p:ext': {
            '@_uri': MEDIA_EXT_URI,
            [`p14:trim`]: {
              [`@_xmlns:p14`]: P14_NS,
              ...trimAttrs,
            },
          },
        };
      }

      const mediaTag = media.mediaType === 'video' ? 'p:video' : 'p:audio';

      if (trigger === 'automatic') {
        const autoSeqId = nextId++;
        mainSeqChildren.push({
          [mediaTag]: {
            'p:cMediaNode': {
              ...cMediaNode,
              'p:cTn': {
                '@_dur': 'indefinite',
                '@_id': String(autoSeqId),
                '@_nodeType': 'withEffect',
                'p:stCondLst': {
                  'p:cond': { '@_delay': '0' },
                },
                'p:endCondLst': endCondLst,
                ...(loop ? { '@_repeatCount': 'indefinite' } : {}),
              },
            },
          },
        });
      } else {
        const interactiveSeqId = nextId++;
        interactiveSeqList.push({
          '@_concurrent': '1',
          '@_nextAc': 'seek',
          'p:cTn': {
            '@_evtFilter': 'cancelBubble',
            '@_fill': 'hold',
            '@_id': String(interactiveSeqId),
            '@_nodeType': 'interactiveSeq',
            '@_restart': 'whenNotActive',
            'p:stCondLst': {
              'p:cond': {
                '@_delay': '0',
                '@_evt': 'onClick',
                'p:tgtEl': {
                  'p:spTgt': { '@_spid': String(media.id) },
                },
              },
            },
            'p:childTnLst': {
              [mediaTag]: { 'p:cMediaNode': cMediaNode },
            },
          },
        });
      }
    }
  }

  const seqList: Record<string, unknown>[] = [];
  if (mainSeqChildren.length > 0) {
    seqList.push({
      '@_concurrent': '1',
      '@_nextAc': 'seek',
      'p:cTn': {
        '@_dur': 'indefinite',
        '@_id': mainSeqId,
        '@_nodeType': 'mainSeq',
        'p:childTnLst': mainSeqChildren,
      },
    });
  }
  for (const interactiveSeq of interactiveSeqList) {
    seqList.push(interactiveSeq);
  }

  return {
    'p:tnLst': {
      'p:par': {
        'p:cTn': {
          '@_dur': 'indefinite',
          '@_id': rootId,
          '@_nodeType': 'tmRoot',
          '@_restart': 'never',
          'p:childTnLst': {
            'p:seq': seqList.length === 1 ? seqList[0] : seqList,
          },
        },
      },
    },
  };
}


/**
 * Serializes a `PptxAudioElement` into an OpenXML `<p:pic>` object.
 * The `<p:nvPr>` block includes `<p:audioFile r:link="rIdN"/>` so PowerPoint
 * recognises the element as embedded audio and renders its native speaker icon.
 * @param element - The audio element AST node.
 * @param linkRelId - The `r:link` relationship ID pointing to the audio file.
 * @param embedRelId - The `r:embed` relationship ID for the `p14:media` extension.
 * @returns JSON object ready to be passed to the XML serializer.
 */
export function serializeAudio(
  element: PptxAudioElement,
  linkRelId: string,
  embedRelId: string,
  imageRelId?: string,
): Record<string, unknown> {
  const xfrm: Record<string, unknown> = {
    'a:off': {
      '@_x': Math.round(Number(element.position?.x ?? 0)),
      '@_y': Math.round(Number(element.position?.y ?? 0)),
    },
    'a:ext': {
      '@_cx': Math.round(Number(element.position?.cx ?? 914400)),
      '@_cy': Math.round(Number(element.position?.cy ?? 914400)),
    },
  };

  return {
    'p:nvPicPr': {
      'p:cNvPr': {
        '@_id': element.id,
        '@_name': element.name,
      },
      'p:cNvPicPr': {
        'a:picLocks': { '@_noChangeAspect': '1' },
      },
      'p:nvPr': {
        'a:audioFile': { '@_r:link': linkRelId },
        'p:extLst': {
          'p:ext': {
            '@_uri': MEDIA_EXT_URI,
            [`p14:media`]: {
              [`@_xmlns:p14`]: P14_NS,
              '@_r:embed': embedRelId,
            },
          },
        },
      },
    },
    'p:blipFill': {
      'a:blip': imageRelId ? { '@_r:embed': imageRelId } : {},
      'a:stretch': { 'a:fillRect': {} },
    },
    'p:spPr': {
      'a:xfrm': xfrm,
      'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
      'a:noFill': {},
      'a:ln': { 'a:noFill': {} },
    },
  };
}

/**
 * Serializes a `PptxVideoElement` into an OpenXML `<p:pic>` object.
 * The `<p:nvPr>` block includes `<a:videoFile r:link="rIdN"/>` so PowerPoint
 * recognises the element as embedded video and renders its native film strip placeholder.
 * @param element - The video element AST node.
 * @param linkRelId - The `r:link` relationship ID pointing to the video file.
 * @param embedRelId - The `r:embed` relationship ID for the `p14:media` extension.
 * @param imageRelId - Optional image relationship ID for the placeholder blip.
 * @returns JSON object ready to be passed to the XML serializer.
 */
export function serializeVideo(
  element: PptxVideoElement,
  linkRelId: string,
  embedRelId: string,
  imageRelId?: string,
): Record<string, unknown> {
  const xfrm: Record<string, unknown> = {
    'a:off': {
      '@_x': Math.round(Number(element.position?.x ?? 0)),
      '@_y': Math.round(Number(element.position?.y ?? 0)),
    },
    'a:ext': {
      '@_cx': Math.round(Number(element.position?.cx ?? 3657600)),
      '@_cy': Math.round(Number(element.position?.cy ?? 2743200)),
    },
  };

  return {
    'p:nvPicPr': {
      'p:cNvPr': {
        '@_id': element.id,
        '@_name': element.name,
      },
      'p:cNvPicPr': {
        'a:picLocks': { '@_noChangeAspect': '1' },
      },
      'p:nvPr': {
        'a:videoFile': { '@_r:link': linkRelId },
        'p:extLst': {
          'p:ext': {
            '@_uri': MEDIA_EXT_URI,
            [`p14:media`]: {
              [`@_xmlns:p14`]: P14_NS,
              '@_r:embed': embedRelId,
            },
          },
        },
      },
    },
    'p:blipFill': {
      'a:blip': imageRelId ? { '@_r:embed': imageRelId } : {},
      'a:stretch': { 'a:fillRect': {} },
    },
    'p:spPr': {
      'a:xfrm': xfrm,
      'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
      'a:noFill': {},
      'a:ln': { 'a:noFill': {} },
    },
  };
}

export { buildMediaTiming };
