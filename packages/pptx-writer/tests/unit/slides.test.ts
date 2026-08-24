import type { PptxSlide } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { serializeSlide } from '../../lib/serializers/slide-serializer';

describe('Slide Serializer (@hokkyss/pptx-writer)', () => {
  it('serializes complete slide XML with shapes and solid background', () => {
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

  it('serializes slide with transitions, animations, and gradient background', () => {
    const slide: PptxSlide = {
      slideId: 'rId2',
      slideNumber: 2,
      background: {
        fill: {
          type: 'gradient',
          gradient: {
            type: 'linear',
            angle: 90,
            stops: [
              { color: { type: 'srgb', value: '000000' }, position: 0 },
              { color: { type: 'srgb', value: 'FFFFFF' }, position: 100000 },
            ],
          },
        },
      },
      transition: {
        type: 'wipe',
        direction: 'right',
        speed: 'fast',
      },
      animations: [
        {
          targetShapeId: '2',
          trigger: 'onClick',
          effect: 'appear',
          sequence: 0,
        },
      ],
      elements: [
        {
          elementType: 'connector',
          type: 'connector',
          id: '3',
          name: 'Cxn',
          isVisible: true,
          zIndex: 0,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
        },
        {
          elementType: 'group',
          type: 'group',
          id: '4',
          name: 'Grp',
          isVisible: true,
          zIndex: 1,
          position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
          rotation: emuDegree(0),
          children: [],
        },
      ],
      shapes: [],
    };

    const xml = serializeSlide(slide);
    expect(xml).toContain('<p:transition');
    expect(xml).toContain('<p:timing>');
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('<p:grpSp>');
    expect(xml).toContain('<a:gradFill>');
  });
});
