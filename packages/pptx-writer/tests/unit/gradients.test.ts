import { describe, expect, it } from 'vitest';
import type { PptxDocument, PptxFill } from '@hokkyss/pptx-core';
import { emu, emuDegree, thousandthsPercent } from '@hokkyss/pptx-core';
import { createZipReader } from '@hokkyss/pptx-reader';
import { serializeFill } from '../../lib/serializers/text-serializer';
import { writePptx } from '../../lib/writer';

describe('Gradient Fill Serialization (@hokkyss/pptx-writer)', () => {
  it('serializes standard 2-stop linear gradient fill', () => {
    const fill: PptxFill = {
      gradient: {
        angle: 90,
        stops: [
          { color: '#0284C7', position: 0 },
          { color: '#6366F1', position: 1 },
        ],
        type: 'linear',
      },
      type: 'gradient',
    };

    const node = serializeFill(fill);
    expect(node).toBeDefined();
    expect(node?.['a:gradFill']).toBeDefined();

    const grad = node?.['a:gradFill'] as Record<string, unknown>;
    expect(grad['a:lin']).toEqual({
      '@_ang': 5400000,
      '@_scaled': '1',
    });

    const gsLst = grad['a:gsLst'] as { 'a:gs': Array<Record<string, unknown>> };
    expect(gsLst['a:gs']).toHaveLength(2);
    expect(gsLst['a:gs'][0]).toEqual({
      '@_pos': 0,
      'a:srgbClr': { '@_val': '0284C7' },
    });
    expect(gsLst['a:gs'][1]).toEqual({
      '@_pos': 100000,
      'a:srgbClr': { '@_val': '6366F1' },
    });
  });

  it('serializes multi-stop gradient with alpha transparency and custom angles', () => {
    const fill: PptxFill = {
      gradient: {
        angle: 135,
        stops: [
          { color: 'accent1', opacity: 0.8, position: 0 },
          { color: '#38BDF8', position: 0.5 },
          { color: '#0F172A', opacity: 0.2, position: 100000 },
        ],
        type: 'linear',
      },
      type: 'gradient',
    };

    const node = serializeFill(fill);
    const grad = node?.['a:gradFill'] as Record<string, unknown>;
    expect(grad['a:lin']).toEqual({
      '@_ang': 8100000, // 135 * 60000
      '@_scaled': '1',
    });

    const gsLst = grad['a:gsLst'] as { 'a:gs': Array<Record<string, unknown>> };
    expect(gsLst['a:gs']).toHaveLength(3);
    expect(gsLst['a:gs'][0]['a:schemeClr']).toEqual({
      '@_val': 'accent1',
      'a:alpha': { '@_val': 80000 },
    });
    expect(gsLst['a:gs'][1]['@_pos']).toBe(50000);
    expect(gsLst['a:gs'][2]['a:srgbClr']).toEqual({
      '@_val': '0F172A',
      'a:alpha': { '@_val': 20000 },
    });
  });

  it('serializes radial / path gradients with center bounds', () => {
    const fill: PptxFill = {
      gradient: {
        pathBounds: { bottom: 0.5, left: 0.5, right: 0.5, top: 0.5 },
        stops: [
          { color: '#FFFFFF', position: 0 },
          { color: '#000000', position: 1 },
        ],
        type: 'radial',
      },
      type: 'gradient',
    };

    const node = serializeFill(fill);
    const grad = node?.['a:gradFill'] as Record<string, unknown>;
    expect(grad['a:path']).toEqual({
      '@_path': 'circle',
      'a:fillToRect': {
        '@_b': 50000,
        '@_l': 50000,
        '@_r': 50000,
        '@_t': 50000,
      },
    });
  });

  it('writes PPTX package with gradient shape fill and slide background', async () => {
    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: {
        slideCount: 1,
        slideHeight: emu(6858000),
        slideWidth: emu(12192000),
      },
      slideLayouts: [],
      slideMasters: [],
      slides: [
        {
          animations: [],
          background: {
            fill: {
              gradient: {
                angle: 45,
                stops: [
                  { color: '#0F172A', position: thousandthsPercent(0) },
                  { color: '#1E293B', position: thousandthsPercent(100000) },
                ],
                type: 'linear',
              },
              type: 'gradient',
            },
          },
          elements: [
            {
              elementType: 'shape',
              fill: {
                gradient: {
                  angle: 180,
                  stops: [
                    { color: '#38BDF8', position: thousandthsPercent(0) },
                    { color: '#6366F1', position: thousandthsPercent(100000) },
                  ],
                  type: 'linear',
                },
                type: 'gradient',
              },
              id: '2',
              isVisible: true,
              name: 'Gradient Card',
              position: { cx: emu(2000000), cy: emu(1000000), x: emu(1000000), y: emu(1000000) },
              rotation: emuDegree(0),
              shapeType: 'roundRect',
              type: 'shape',
              zIndex: 0,
            },
          ],
          shapes: [],
          slideId: 'rId2',
          slideNumber: 1,
        },
      ],
      themes: [],
    };

    const buffer = await writePptx(doc);
    const zip = await createZipReader(buffer);

    expect(zip.hasFile('ppt/slides/slide1.xml')).toBe(true);
    const slideXml = zip.getFileText('ppt/slides/slide1.xml');
    expect(slideXml).toContain('<a:gradFill');
    expect(slideXml).toContain('<a:gsLst>');
    expect(slideXml).toContain('val="0F172A"');
    expect(slideXml).toContain('val="1E293B"');
    expect(slideXml).toContain('val="38BDF8"');
    expect(slideXml).toContain('val="6366F1"');
    expect(slideXml).toContain('<a:lin');
  });
});
