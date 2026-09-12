import type { PptxPictureElement } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';
import { serializeHyperlink } from './text-serializer';

/**
 * Serializes a picture element into OpenXML `<p:pic>`.
 * Follows schema order: p:nvPicPr -> p:blipFill -> p:spPr
 */
export function serializePicture(pictureElement: PptxPictureElement, overrideEmbedId?: string): XmlElement {
  const pic = pictureElement.picture;
  const embedId = overrideEmbedId ?? (pic?.mediaId || undefined) ?? pictureElement.blipEmbedId ?? 'rId2';

  const blipAttrs: Record<string, string> = {
    'r:embed': embedId,
  };

  const blipChildren: XmlElement[] = [];
  if (pic?.alpha !== undefined) {
    blipChildren.push(el('a:alphaModFix', { amt: Math.round(Number(pic.alpha)) }));
  }

  const blip = el('a:blip', blipAttrs, blipChildren);

  const blipFillChildren: XmlElement[] = [blip];

  if (pic?.crop) {
    const srcRectAttrs: Record<string, number> = {};
    if (pic.crop.left !== undefined) srcRectAttrs.l = Math.round(Number(pic.crop.left));
    if (pic.crop.right !== undefined) srcRectAttrs.r = Math.round(Number(pic.crop.right));
    if (pic.crop.top !== undefined) srcRectAttrs.t = Math.round(Number(pic.crop.top));
    if (pic.crop.bottom !== undefined) srcRectAttrs.b = Math.round(Number(pic.crop.bottom));
    blipFillChildren.push(el('a:srcRect', srcRectAttrs));
  }

  blipFillChildren.push(el('a:stretch', [el('a:fillRect')]));
  const blipFill = el('p:blipFill', blipFillChildren);

  const xfrmAttrs: Record<string, number | undefined> = {};
  if (pictureElement.rotation) {
    xfrmAttrs.rot = Math.round(Number(pictureElement.rotation));
  }
  const xfrm = el('a:xfrm', xfrmAttrs, [
    el('a:off', {
      x: Math.round(Number(pictureElement.position?.x ?? 0)),
      y: Math.round(Number(pictureElement.position?.y ?? 0)),
    }),
    el('a:ext', {
      cx: Math.round(Number(pictureElement.position?.cx ?? 2000000)),
      cy: Math.round(Number(pictureElement.position?.cy ?? 2000000)),
    }),
  ]);

  const cNvPrAttrs: Record<string, string> = {
    id: pictureElement.id || '4',
    name: pictureElement.name || `Picture ${pictureElement.id || '4'}`,
  };
  const cNvPrChildren: XmlElement[] = [];
  if (pictureElement.hyperlink) {
    const hlinkNode = serializeHyperlink(pictureElement.hyperlink);
    if (hlinkNode) {
      cNvPrChildren.push(hlinkNode);
    }
  }

  const nvPicPr = el('p:nvPicPr', [
    el('p:cNvPr', cNvPrAttrs, cNvPrChildren),
    el('p:cNvPicPr', [
      el('a:picLocks', { noChangeAspect: '1' }),
    ]),
    el('p:nvPr'),
  ]);

  const spPr = el('p:spPr', [
    xfrm,
    el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]),
  ]);

  return el('p:pic', [nvPicPr, blipFill, spPr]);
}
