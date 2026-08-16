import { describe, expect, it } from 'vitest';
import type { PptxPictureElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, thousandthsPercent } from '@hokkyss/pptx-core';
import { serializePicture } from '../../lib/serializers/picture-serializer';

describe('Picture Serializer', () => {
  it('serializes picture element with blip embed, crop, alpha, and transforms', () => {
    const pic: PptxPictureElement = {
      elementType: 'picture',
      type: 'picture',
      id: '4',
      name: 'Picture 1',
      isVisible: true,
      zIndex: 2,
      position: {
        x: emu(1000000),
        y: emu(1000000),
        cx: emu(2000000),
        cy: emu(2000000),
      },
      rotation: emuDegree(0),
      picture: {
        mediaId: 'rId2',
        alpha: thousandthsPercent(90000),
        crop: {
          left: thousandthsPercent(10000),
          right: thousandthsPercent(10000),
          top: thousandthsPercent(5000),
          bottom: thousandthsPercent(5000),
        },
      },
    };

    const xmlObject = serializePicture(pic);
    expect(xmlObject).toBeDefined();

    const nvPicPr = xmlObject['p:nvPicPr'] as Record<string, Record<string, unknown>>;
    expect(nvPicPr['p:cNvPr']['@_id']).toBe('4');

    const blipFill = xmlObject['p:blipFill'] as Record<string, Record<string, unknown>>;
    expect(blipFill['a:blip']['@_r:embed']).toBe('rId2');
    expect((blipFill['a:blip']['a:alphaModFix'] as Record<string, unknown>)['@_amt']).toBe(90000);
    expect(blipFill['a:srcRect']['@_l']).toBe(10000);
  });
});
