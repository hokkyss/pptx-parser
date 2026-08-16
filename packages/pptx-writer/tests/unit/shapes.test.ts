import { describe, expect, it } from 'vitest';
import type { PptxShapeElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeShape } from '../../lib/serializers/shape-serializer';

describe('Shape Serializer', () => {
  it('serializes shape element with transforms, geometry, fills, and outline', () => {
    const shape: PptxShapeElement = {
      elementType: 'shape',
      type: 'shape',
      id: '2',
      name: 'Rectangle 1',
      isVisible: true,
      zIndex: 0,
      position: {
        x: emu(100000),
        y: emu(200000),
        cx: emu(3000000),
        cy: emu(1500000),
      },
      rotation: emuDegree(5400000), // 90 deg
      geometry: {
        presetGeometry: 'rect',
      },
      fill: {
        type: 'solid',
        solidColor: { type: 'srgb', value: '007ACC' },
      },
      line: {
        width: emu(12700),
        fill: { type: 'solid', solidColor: { type: 'srgb', value: '000000' } },
      },
      textBody: {
        bodyProperties: {},
        paragraphs: [
          {
            properties: {},
            runs: [{ text: 'Box Content', properties: { fontSize: hundredthsPoint(1800) } }],
          },
        ],
      },
    };

    const xmlObject = serializeShape(shape);
    expect(xmlObject).toBeDefined();

    const nvSpPr = xmlObject['p:nvSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvSpPr['p:cNvPr']['@_id']).toBe('2');
    expect(nvSpPr['p:cNvPr']['@_name']).toBe('Rectangle 1');

    const spPr = xmlObject['p:spPr'] as Record<string, Record<string, unknown>>;
    const xfrm = spPr['a:xfrm'] as Record<string, Record<string, unknown>>;
    expect(xfrm['a:off']['@_x']).toBe(100000);
    expect(xfrm['a:off']['@_y']).toBe(200000);
    expect(xfrm['a:ext']['@_cx']).toBe(3000000);
    expect(xfrm['a:ext']['@_cy']).toBe(1500000);
    expect(xfrm['@_rot']).toBe(5400000);

    expect(spPr['a:prstGeom']['@_prst']).toBe('rect');
    expect((spPr['a:solidFill']['a:srgbClr'] as Record<string, unknown>)['@_val']).toBe('007ACC');
    expect(spPr['a:ln']['@_w']).toBe(12700);
    expect(((spPr['a:ln']['a:solidFill'] as Record<string, unknown>)['a:srgbClr'] as Record<string, unknown>)['@_val']).toBe('000000');

    expect(xmlObject['p:txBody']).toBeDefined();
  });
});
