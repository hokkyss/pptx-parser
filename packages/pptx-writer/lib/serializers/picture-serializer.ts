import type { PptxPictureElement } from '@hokkyss/pptx-core';
import { serializeHyperlink } from './text-serializer';

/**
 * Serializes a picture element into OpenXML `<p:pic>`.
 * Follows schema order: p:nvPicPr -> p:blipFill -> p:spPr
 */
export function serializePicture(pictureElement: PptxPictureElement, overrideEmbedId?: string): Record<string, unknown> {
  const pic = pictureElement.picture;
  const embedId = overrideEmbedId ?? (pic?.mediaId || undefined) ?? pictureElement.blipEmbedId ?? 'rId2';

  const blip: Record<string, unknown> = {
    '@_r:embed': embedId,
  };

  if (pic?.alpha !== undefined) {
    blip['a:alphaModFix'] = { '@_amt': Math.round(Number(pic.alpha)) };
  }

  const blipFill: Record<string, unknown> = {
    'a:blip': blip,
  };

  if (pic?.crop) {
    const srcRect: Record<string, unknown> = {};
    if (pic.crop.left !== undefined) srcRect['@_l'] = Math.round(Number(pic.crop.left));
    if (pic.crop.right !== undefined) srcRect['@_r'] = Math.round(Number(pic.crop.right));
    if (pic.crop.top !== undefined) srcRect['@_t'] = Math.round(Number(pic.crop.top));
    if (pic.crop.bottom !== undefined) srcRect['@_b'] = Math.round(Number(pic.crop.bottom));
    blipFill['a:srcRect'] = srcRect;
  }

  blipFill['a:stretch'] = { 'a:fillRect': {} };

  const xfrm: Record<string, unknown> = {
    'a:off': {
      '@_x': Math.round(Number(pictureElement.position?.x ?? 0)),
      '@_y': Math.round(Number(pictureElement.position?.y ?? 0)),
    },
    'a:ext': {
      '@_cx': Math.round(Number(pictureElement.position?.cx ?? 2000000)),
      '@_cy': Math.round(Number(pictureElement.position?.cy ?? 2000000)),
    },
  };
  if (pictureElement.rotation) {
    xfrm['@_rot'] = Math.round(Number(pictureElement.rotation));
  }

  const cNvPr: Record<string, unknown> = {
    '@_id': pictureElement.id,
    '@_name': pictureElement.name || '',
  };
  if (pictureElement.hyperlink) {
    const hlinkNode = serializeHyperlink(pictureElement.hyperlink);
    if (hlinkNode) {
      cNvPr['a:hlinkClick'] = hlinkNode;
    }
  }

  return {
    'p:nvPicPr': {
      'p:cNvPr': cNvPr,
      'p:cNvPicPr': {
        'a:picLocks': { '@_noChangeAspect': '1' },
      },
      'p:nvPr': {},
    },
    'p:blipFill': blipFill,
    'p:spPr': {
      'a:xfrm': xfrm,
      'a:prstGeom': {
        '@_prst': 'rect',
        'a:avLst': {},
      },
    },
  };
}
