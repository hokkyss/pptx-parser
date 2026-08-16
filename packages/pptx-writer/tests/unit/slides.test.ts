import type { PptxSlide } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { serializeSlide } from '../../lib/serializers/slide-serializer';

describe('Slide Serializer', () => {
  it('serializes complete slide XML with shapes and background', () => {
    const slide: PptxSlide = {
      animations: [],
      background: {
        fill: {
          solidColor: { type: 'srgb', value: 'FFFFFF' },
          type: 'solid',
        },
      },
      elements: [
        {
          elementType: 'shape',
          id: '2',
          isVisible: true,
          name: 'Title 1',
          position: {
            cx: emu(5000000),
            cy: emu(1000000),
            x: emu(100000),
            y: emu(100000),
          },
          rotation: emuDegree(0),
          textBody: {
            bodyProperties: {},
            paragraphs: [
              {
                properties: {},
                runs: [{ properties: { fontSize: hundredthsPoint(3200) }, text: 'Presentation Title' }],
              },
            ],
          },
          type: 'shape',
          zIndex: 0,
        },
      ],
      shapes: [],
      slideId: 'rId1',
      slideNumber: 1,
    };

    const xml = serializeSlide(slide);
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    expect(xml).toContain('<p:sld');
    expect(xml).toContain('<p:cSld>');
    expect(xml).toContain('<p:spTree>');
    expect(xml).toContain('Presentation Title');
    expect(xml).toContain('<p:bg>');
  });
});
