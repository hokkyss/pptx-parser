import { describe, expect, it } from 'vitest';
import { parseBackground, parseFill } from '../../lib/parsers/fill-parser';
import { parseShapes } from '../../lib/parsers/shape-parser';
import { defaultXmlParser } from '../../lib/xml/xml-parser';

describe('Gradient Fill Parser (@hokkyss/pptx-reader)', () => {
  it('parses linear gradient fill with multiple stops and angle', () => {
    const xml = `
      <p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:srgbClr val="0284C7"/>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="accent1">
                <a:alpha val="80000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:srgbClr val="6366F1"/>
            </a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="1"/>
        </a:gradFill>
      </p:spPr>
    `;

    const parsed = defaultXmlParser.parse<Record<string, unknown>>(xml);
    const spPrNode = (parsed['p:spPr'] || parsed['spPr']) as Record<string, unknown>;

    const fill = parseFill(spPrNode);
    expect(fill).toBeDefined();
    expect(fill?.type).toBe('gradient');
    expect(fill?.gradient?.type).toBe('linear');
    expect(fill?.gradient?.angle).toBe(5400000);
    expect(fill?.gradient?.rotateWithShape).toBe(true);

    const stops = fill?.gradient?.stops || [];
    expect(stops).toHaveLength(3);
    expect(stops[0].position).toBe(0);
    expect(stops[0].color).toEqual({ alpha: undefined, type: 'srgb', value: '0284C7' });

    expect(stops[1].position).toBe(50000);
    expect(stops[1].opacity).toBe(0.8);
    expect(stops[1].color).toEqual({ alpha: 80000, type: 'scheme', value: 'accent1' });

    expect(stops[2].position).toBe(100000);
    expect(stops[2].color).toEqual({ alpha: undefined, type: 'srgb', value: '6366F1' });
  });

  it('parses radial gradient fill with circle path and center bounds', () => {
    const xml = `
      <p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:gradFill flip="xy">
          <a:gsLst>
            <a:gs pos="0">
              <a:srgbClr val="FFFFFF"/>
            </a:gs>
            <a:gs pos="100000">
              <a:srgbClr val="000000"/>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect b="50000" l="50000" r="50000" t="50000"/>
          </a:path>
        </a:gradFill>
      </p:spPr>
    `;

    const parsed = defaultXmlParser.parse<Record<string, unknown>>(xml);
    const spPrNode = (parsed['p:spPr'] || parsed['spPr']) as Record<string, unknown>;

    const fill = parseFill(spPrNode);
    expect(fill?.type).toBe('gradient');
    expect(fill?.gradient?.type).toBe('radial');
    expect(fill?.gradient?.flip).toBe('xy');
    expect(fill?.gradient?.pathBounds).toEqual({
      bottom: 0.5,
      left: 0.5,
      right: 0.5,
      top: 0.5,
    });
  });

  it('parses slide background gradient fill from <p:bg>', () => {
    const xml = `
      <p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:bgPr>
          <a:gradFill>
            <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr val="0F172A"/>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr val="1E293B"/>
              </a:gs>
            </a:gsLst>
            <a:lin ang="2700000"/>
          </a:gradFill>
          <a:effectLst/>
        </p:bgPr>
      </p:bg>
    `;

    const parsed = defaultXmlParser.parse<Record<string, unknown>>(xml);
    const bgNode = (parsed['p:bg'] || parsed['bg']) as Record<string, unknown>;

    const bg = parseBackground(bgNode);
    expect(bg).toBeDefined();
    expect(bg?.fill?.type).toBe('gradient');
    expect(bg?.fill?.gradient?.angle).toBe(2700000);
    expect(bg?.fill?.gradient?.stops).toHaveLength(2);
  });

  it('extracts shape fill and line fill in parseShapes', () => {
    const slideXml = `
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
          <p:spTree>
            <p:nvGrpSpPr>
              <p:cNvPr id="1" name=""/>
              <p:cNvGrpSpPr/>
              <p:nvPr/>
            </p:nvGrpSpPr>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="GradientCard"/>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="1000" y="1000"/>
                  <a:ext cx="2000" cy="2000"/>
                </a:xfrm>
                <a:prstGeom prst="roundRect">
                  <a:avLst/>
                </a:prstGeom>
                <a:gradFill>
                  <a:gsLst>
                    <a:gs pos="0">
                      <a:srgbClr val="38BDF8"/>
                    </a:gs>
                    <a:gs pos="100000">
                      <a:srgbClr val="6366F1"/>
                    </a:gs>
                  </a:gsLst>
                  <a:lin ang="8100000"/>
                </a:gradFill>
                <a:ln w="12700">
                  <a:solidFill>
                    <a:srgbClr val="FFFFFF"/>
                  </a:solidFill>
                </a:ln>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>
    `;

    const shapes = parseShapes(slideXml);
    expect(shapes).toHaveLength(1);
    const shape = shapes[0];
    expect(shape.shapeType).toBe('roundRect');
    expect(shape.fill?.type).toBe('gradient');
    expect(shape.fill?.gradient?.stops).toHaveLength(2);
    expect(shape.line?.width).toBe(12700);
    expect(shape.line?.fill?.type).toBe('solid');
    expect(shape.line?.fill?.solidColor?.value).toBe('FFFFFF');
  });
});

describe('parseFill and parseBackground edge cases', () => {
  it('parses noFill and handles empty background node', () => {
    expect(parseFill({})).toBeUndefined();
    expect(parseFill(undefined)).toBeUndefined();
    expect(parseFill({ 'a:noFill': {} })?.type).toBe('none');

    expect(parseBackground(undefined)).toBeUndefined();
    expect(parseBackground({})).toBeUndefined();
  });
});
