import { describe, expect, it } from 'vitest';
import type { PptxPictureElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, thousandthsPercent } from '@hokkyss/pptx-core';
import { serializePicture } from '../../lib/serializers/picture-serializer';
import { renderXml } from '../../lib/xml/xml-element';

describe('Picture Serializer', () => {
  it('serializes picture element with blip embed, crop, alpha, and transforms', () => {
    const pic: PptxPictureElement = {
      elementType: 'picture',
      id: '4',
      isVisible: true,
      name: 'Picture 1',
      picture: {
        alpha: thousandthsPercent(90000),
        crop: {
          bottom: thousandthsPercent(5000),
          left: thousandthsPercent(10000),
          right: thousandthsPercent(10000),
          top: thousandthsPercent(5000),
        },
        mediaId: 'rId2',
      },
      position: {
        cx: emu(2000000),
        cy: emu(2000000),
        x: emu(1000000),
        y: emu(1000000),
      },
      rotation: emuDegree(0),
      type: 'picture',
      zIndex: 2,
    };

    const xmlNode = serializePicture(pic);
    expect(xmlNode).toBeDefined();
    const xml = renderXml(xmlNode);

    // cNvPr id
    expect(xml).toContain('id="4"');
    // blipFill embed
    expect(xml).toContain('r:embed="rId2"');
    // alpha
    expect(xml).toContain('<a:alphaModFix');
    expect(xml).toContain('amt="90000"');
    // crop — left=10000
    expect(xml).toContain('<a:srcRect');
    expect(xml).toContain('l="10000"');
  });
});

describe('Picture Serializer rotation', () => {
  it('serializes picture rotation attribute in xfrm', () => {
    const pic: PptxPictureElement = {
      elementType: 'picture',
      id: '5',
      isVisible: true,
      name: 'Rotated Pic',
      picture: { mediaId: 'rId5' },
      position: { cx: emu(100), cy: emu(100), x: emu(0), y: emu(0) },
      rotation: emuDegree(5400000),
      type: 'picture',
      zIndex: 0,
    };
    const xml = renderXml(serializePicture(pic));
    expect(xml).toContain('rot="5400000"');
  });

  it('covers blipEmbedId fallback, undefined positions and name fallbacks', () => {
    const picMinimal: PptxPictureElement = {
      blipEmbedId: 'rId8',
      elementType: 'picture',
      id: '',
      isVisible: true,
      name: '',
      picture: { mediaId: '' },
      position: { cx: emu(0), cy: emu(0), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      type: 'picture',
      zIndex: 0,
    };
    const xml = renderXml(serializePicture(picMinimal));
    // blipEmbedId fallback used when mediaId is empty
    expect(xml).toContain('r:embed="rId8"');
    // id fallback
    expect(xml).toContain('id="4"');

    // Override embedId, hyperlink, partial crops, undefined position
    // @ts-expect-error Testing undefined position
    const picWithOverride: PptxPictureElement = {
      elementType: 'picture',
      hyperlink: { rId: 'rIdHlink', url: 'https://example.com' },
      id: '9',
      isVisible: true,
      name: 'P9',
      picture: {
        crop: {
          bottom: thousandthsPercent(1000),
          left: thousandthsPercent(2000),
          right: thousandthsPercent(3000),
          top: thousandthsPercent(4000),
        },
        mediaId: 'med-override',
      },
      rotation: emuDegree(0),
      type: 'picture',
      zIndex: 0,
    };
    const overriddenXml = renderXml(serializePicture(picWithOverride, 'rIdOverride'));
    expect(overriddenXml).toContain('r:embed="rIdOverride"');
    expect(overriddenXml).toContain('<a:hlinkClick');
  });
});
