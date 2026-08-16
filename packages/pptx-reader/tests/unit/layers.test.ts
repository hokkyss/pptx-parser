import { describe, it, expect } from 'vitest';
import { parseShapes } from '../../lib/parsers/shape-parser';
import { resolveSlideLayers } from '../../lib/resolvers/layer-resolver';
import type { Emu, EmuDegree, PptxDocument } from '../../lib/types/ast';

describe('Layer Parsing & Composition', () => {
  it('should parse 0-based zIndex in shape order', () => {
    const xml = `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:sp><p:nvSpPr><p:cNvPr id="1" name="Shape 1"/></p:nvSpPr></p:sp>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Shape 2"/></p:nvSpPr></p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>
    `;

    const shapes = parseShapes(xml);
    expect(shapes.length).toBe(2);
    expect(shapes[0].zIndex).toBe(0);
    expect(shapes[1].zIndex).toBe(1);
  });

  it('should resolve slide layers in Master -> Layout -> Slide order', () => {
    const mockDoc: PptxDocument = {
      customXml: [],
      media: [],
      metadata: { slideCount: 1, slideHeight: 100 as Emu, slideWidth: 100 as Emu },
      slideLayouts: [
        {
          elements: [],
          id: 'layout1',
          masterId: 'master1',
          name: 'Title Layout',
          shapes: [
            {
              elementType: 'shape',
              id: 'l1',
              isVisible: true,
              name: 'Layout Placeholder',
              position: { cx: 10 as Emu, cy: 10 as Emu, x: 0 as Emu, y: 0 as Emu },
              rotation: 0 as EmuDegree,
              type: 'shape',
              zIndex: 0,
            },
          ],
          type: 'title',
        },
      ],
      slideMasters: [
        {
          elements: [],
          id: 'master1',
          layoutIds: ['layout1'],
          shapes: [
            {
              elementType: 'shape',
              id: 'm1',
              isVisible: true,
              name: 'Master Background',
              position: { cx: 10 as Emu, cy: 10 as Emu, x: 0 as Emu, y: 0 as Emu },
              rotation: 0 as EmuDegree,
              type: 'shape',
              zIndex: 0,
            },
          ],
        },
      ],
      slides: [
        {
          animations: [],
          elements: [],
          layoutId: 'layout1',
          shapes: [
            {
              elementType: 'shape',
              id: 's1',
              isVisible: true,
              name: 'Slide Shape 1',
              position: { cx: 10 as Emu, cy: 10 as Emu, x: 0 as Emu, y: 0 as Emu },
              rotation: 0 as EmuDegree,
              type: 'shape',
              zIndex: 0,
            },
          ],
          slideId: 'rId1',
          slideNumber: 1,
        },
      ],
      themes: [],
    };

    // Ensure elements array is synced with shapes for test mock
    mockDoc.slides[0].elements = mockDoc.slides[0].shapes;
    mockDoc.slideMasters[0].elements = mockDoc.slideMasters[0].shapes;
    mockDoc.slideLayouts[0].elements = mockDoc.slideLayouts[0].shapes;

    const layers = resolveSlideLayers(mockDoc, 1);
    expect(layers).toBeDefined();
    expect(layers?.masterShapes.length).toBe(1);
    expect(layers?.layoutShapes.length).toBe(1);
    expect(layers?.slideShapes.length).toBe(1);

    expect(layers?.allShapesInRenderOrder.length).toBe(3);
    expect(layers?.allShapesInRenderOrder[0].layerSource).toBe('master');
    expect(layers?.allShapesInRenderOrder[1].layerSource).toBe('layout');
    expect(layers?.allShapesInRenderOrder[2].layerSource).toBe('slide');
  });
});
