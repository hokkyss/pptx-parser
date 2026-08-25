import { describe, expect, it } from 'vitest';
import type { PptxPictureElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, thousandthsPercent } from '@hokkyss/pptx-core';
import { serializePicture } from '../../lib/serializers/picture-serializer';

interface PicNvPicPr {
  'p:cNvPr'?: {
    '@_id'?: string;
    '@_name'?: string;
  };
}

interface PicBlipFill {
  'a:blip'?: {
    '@_r:embed'?: string;
    'a:alphaModFix'?: {
      '@_amt'?: number;
    };
  };
  'a:srcRect'?: {
    '@_l'?: number;
  };
}

interface PicSpPr {
  'a:xfrm'?: {
    '@_rot'?: number;
  };
}

interface SerializedPicNode {
  'p:blipFill'?: PicBlipFill;
  'p:nvPicPr'?: PicNvPicPr;
  'p:spPr'?: PicSpPr;
}

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

    const xmlObject = serializePicture(pic) as SerializedPicNode;
    expect(xmlObject).toBeDefined();

    const nvPicPr = xmlObject['p:nvPicPr'];
    expect(nvPicPr?.['p:cNvPr']?.['@_id']).toBe('4');

    const blipFill = xmlObject['p:blipFill'];
    expect(blipFill?.['a:blip']?.['@_r:embed']).toBe('rId2');
    expect(blipFill?.['a:blip']?.['a:alphaModFix']?.['@_amt']).toBe(90000);
    expect(blipFill?.['a:srcRect']?.['@_l']).toBe(10000);
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
    const xml = serializePicture(pic) as SerializedPicNode;
    const spPr = xml['p:spPr'];
    expect(spPr?.['a:xfrm']?.['@_rot']).toBe(5400000);
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
    const obj = serializePicture(picMinimal) as SerializedPicNode;
    const blipFill = obj['p:blipFill'];
    expect(blipFill?.['a:blip']?.['@_r:embed']).toBe('rId8');
    const nvPicPr = obj['p:nvPicPr'];
    expect(nvPicPr?.['p:cNvPr']?.['@_id']).toBe('4');
  });
});
