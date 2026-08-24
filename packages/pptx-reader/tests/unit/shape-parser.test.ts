import { describe, expect, it } from 'vitest';
import { parseShapes } from '../../lib/parsers/shape-parser';
import { createRelationshipResolver } from '../../lib/resolvers/relationship-resolver';
import { defaultXmlParser } from '../../lib/xml/xml-parser';

describe('Shape Parser (@hokkyss/pptx-reader)', () => {
  const dummyResolver = createRelationshipResolver('', 'ppt/slides/slide1.xml');

  it('returns empty array when spTree is missing or empty', () => {
    const xml = `<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>`;
    expect(parseShapes(xml, dummyResolver)).toEqual([]);
  });

  it('parses connectors (<p:cxnSp>)', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="5" name="Connector 5"/>
          <p:cNvCxnSpPr/>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="100" y="200"/><a:ext cx="5000" cy="3000"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="19050">
            <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          </a:ln>
        </p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const shapes = parseShapes(xml, dummyResolver);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('connector');
    expect(shapes[0].elementType).toBe('connector');
    expect(shapes[0].id).toBe('5');
    expect(shapes[0].name).toBe('Connector 5');
    expect(shapes[0].position.x).toBe(100);
    expect(shapes[0].position.y).toBe(200);
    expect(shapes[0].line?.width).toBe(19050);
  });

  it('parses graphicFrames containing tables or charts', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="6" name="Chart 6"/>
          <p:cNvGraphicFramePr/>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="500" y="500"/><a:ext cx="4000000" cy="3000000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="rId2"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const shapes = parseShapes(xml, dummyResolver);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('graphicFrame');
    expect(shapes[0].id).toBe('6');
    expect((shapes[0] as unknown as { _chartRelId: string })._chartRelId).toBe('rId2');
  });

  it('parses shape locks and placeholders', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="7" name="Locked Shape"/>
          <p:cNvSpPr txBox="1">
            <a:spLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeShapeType="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title" idx="0"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const shapes = parseShapes(xml, dummyResolver);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].isLocked).toBe(true);
    expect(shapes[0].locks?.noGrp).toBe(true);
    expect(shapes[0].locks?.noRot).toBe(true);
    expect(shapes[0].locks?.noChangeAspect).toBe(true);
    expect(shapes[0].locks?.noMove).toBe(true);
    expect(shapes[0].locks?.noResize).toBe(true);
    expect(shapes[0].placeholder?.type).toBe('title');
    expect(shapes[0].placeholder?.idx).toBe('0');
    expect(shapes[0].rotation).toBe(5400000);
  });
});
