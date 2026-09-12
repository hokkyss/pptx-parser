import { describe, expect, it } from 'vitest';
import type { PptxDocument, PptxFill } from '@hokkyss/pptx-core';
import { emu, emuDegree, thousandthsPercent } from '@hokkyss/pptx-core';
import { createZipReader } from '@hokkyss/pptx-reader';
import { serializeFill } from '../../lib/serializers/text-serializer';
import { writePptx } from '../../lib/writer';
import { renderXml, type XmlElement } from '../../lib/xml/xml-element';

/** Finds the first descendant (or self) matching the given tag. */
function findEl(node: XmlElement, tag: string): undefined | XmlElement {
  if (node.tag === tag) return node;
  for (const child of node.children ?? []) {
    if (typeof child === 'object' && child !== null && 'tag' in child) {
      const found = findEl(child, tag);
      if (found) return found;
    }
  }
  return undefined;
}

/** Finds all descendants (and self) matching the given tag. */
function findAllEl(node: XmlElement, tag: string): XmlElement[] {
  const results: XmlElement[] = [];
  if (node.tag === tag) results.push(node);
  for (const child of node.children ?? []) {
    if (typeof child === 'object' && child !== null && 'tag' in child) {
      results.push(...findAllEl(child, tag));
    }
  }
  return results;
}

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
    const xml = renderXml(node);

    expect(xml).toContain('<a:gradFill');
    // linear node
    expect(xml).toContain('ang="5400000"');
    expect(xml).toContain('scaled="1"');

    // gradient stops
    const gsList = findAllEl(node!, 'a:gs');
    expect(gsList).toHaveLength(2);
    expect(gsList[0].attrs?.pos).toBe(0);
    expect(gsList[1].attrs?.pos).toBe(100000);
    expect(findEl(gsList[0], 'a:srgbClr')?.attrs?.val).toBe('0284C7');
    expect(findEl(gsList[1], 'a:srgbClr')?.attrs?.val).toBe('6366F1');
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
    expect(node).toBeDefined();
    const xml = renderXml(node);

    expect(xml).toContain('ang="8100000"'); // 135 * 60000
    expect(xml).toContain('scaled="1"');

    const gsList = findAllEl(node!, 'a:gs');
    expect(gsList).toHaveLength(3);

    // stop 0: accent1 with alpha=80000
    const stop0Scheme = findEl(gsList[0], 'a:schemeClr');
    expect(stop0Scheme?.attrs?.val).toBe('accent1');
    expect(findEl(stop0Scheme!, 'a:alpha')?.attrs?.val).toBe(80000);

    // stop 1: pos=50000
    expect(gsList[1].attrs?.pos).toBe(50000);

    // stop 2: srgbClr with alpha=20000
    const stop2Srgb = findEl(gsList[2], 'a:srgbClr');
    expect(stop2Srgb?.attrs?.val).toBe('0F172A');
    expect(findEl(stop2Srgb!, 'a:alpha')?.attrs?.val).toBe(20000);
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
    expect(node).toBeDefined();
    const xml = renderXml(node);

    expect(xml).toContain('path="circle"');
    const fillToRect = findEl(node!, 'a:fillToRect');
    expect(fillToRect).toBeDefined();
    expect(fillToRect?.attrs?.l).toBe(50000);
    expect(fillToRect?.attrs?.t).toBe(50000);
    expect(fillToRect?.attrs?.r).toBe(50000);
    expect(fillToRect?.attrs?.b).toBe(50000);
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
