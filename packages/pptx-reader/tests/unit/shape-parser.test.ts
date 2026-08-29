import type { PptxElement } from '@hokkyss/pptx-core';
import { describe, expect, it } from 'vitest';
import { parseShapes } from '../../lib/parsers/shape-parser';
import { createRelationshipResolver } from '../../lib/resolvers/relationship-resolver';

const dummyResolver = createRelationshipResolver('', 'ppt/slides/slide1.xml');

describe('Shape Parser (@hokkyss/pptx-reader)', () => {
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
    expect((shapes[0] as { _chartRelId?: string } & PptxElement)._chartRelId).toBe('rId2');
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

describe('Shape Parser group and single shape parsing', () => {
  it('parses group shapes (<p:grpSp>) with nested child shapes', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="8" name="Group 8"/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="2000"/></a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="9" name="Child Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const resolver = createRelationshipResolver('', 'ppt/slides/slide1.xml');
    const shapes = parseShapes(xml, resolver);
    const group = shapes[0];
    expect(group.type).toBe('group');
    expect(group.elementType).toBe('group');
    if (group.elementType === 'group') {
      expect(group.children).toHaveLength(1);
      expect(group.children[0].name).toBe('Child Shape');
    }
  });
});

describe('Shape Parser JS object input and empty text body', () => {
  it('parses shape directly from JS object and handles empty txBody', () => {
    const jsObj = {
      'p:spTree': {
        'p:sp': {
          'p:nvSpPr': {
            'p:cNvPr': { '@_id': '20', '@_name': 'Obj Shape' },
            'p:cNvSpPr': {},
            'p:nvPr': {},
          },
          'p:spPr': {},
          'p:txBody': {
            'a:bodyPr': {},
          },
        },
      },
    };

    const shapes = parseShapes(jsObj);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].name).toBe('Obj Shape');
    expect(shapes[0].textBody).toBeUndefined();
  });
});

describe('Shape Parser shape without nvPr node', () => {
  it('handles shape without nvSpPr or nvPrNode gracefully', () => {
    const rawShapeTree = {
      'p:spTree': {
        'p:sp': {
          'p:spPr': {},
        },
      },
    };
    const shapes = parseShapes(rawShapeTree);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].id).toBe('0');
    expect(shapes[0].name).toBe('');
    expect(shapes[0].isVisible).toBe(true);
  });
});

describe('Shape Parser connector arrowheads and attachment parsing', () => {
  it('parses headEnd, tailEnd, and attachment connection points on connectors', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="10" name="Arrow Connector"/>
          <p:cNvCxnSpPr>
            <a:stCxn id="2" idx="3"/>
            <a:endCxn id="3" idx="1"/>
          </p:cNvCxnSpPr>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>
          <a:ln w="12700">
            <a:headEnd type="triangle" w="lg" len="lg"/>
            <a:tailEnd type="oval" w="sm" len="sm"/>
          </a:ln>
        </p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const shapes = parseShapes(xml, dummyResolver);
    expect(shapes).toHaveLength(1);
    const conn = shapes[0];
    expect(conn.elementType).toBe('connector');
    if (conn.elementType === 'connector') {
      expect(conn.line?.headEnd?.type).toBe('triangle');
      expect(conn.line?.headEnd?.width).toBe('lg');
      expect(conn.line?.headEnd?.length).toBe('lg');
      expect(conn.line?.tailEnd?.type).toBe('oval');
      expect(conn.startConnection?.shapeId).toBe('2');
      expect(conn.startConnection?.position).toBe('right');
      expect(conn.endConnection?.shapeId).toBe('3');
      expect(conn.endConnection?.position).toBe('left');
    }
  });

  it('covers custom geometry, line arrow heads, and dash styles', () => {
    const xml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="1" name="Shape 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm rot="2700000" flipH="1" flipV="1"><a:off x="100" y="200"/><a:ext cx="1000" cy="1000"/></a:xfrm>
          <a:custGeom><a:avLst/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path></a:pathLst></a:custGeom>
          <a:ln w="25400" cap="rnd" cmpd="s">
            <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
            <a:prstDash val="dash"/>
            <a:headEnd type="oval" w="med" len="med"/>
            <a:tailEnd type="stealth" w="sm" len="lg"/>
          </a:ln>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const shapes = parseShapes(xml, dummyResolver);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].rotation).toBe(2700000);
    expect(shapes[0].line?.dashStyle).toBe('dash');

    // graphicFrame with table and chart variations
    const gfXml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="20" name="Table 20"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl><a:tblGrid/><a:tr/></a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="21" name="Chart 21"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart id="rIdChart"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="30" name="Empty Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
    const gfShapes = parseShapes(gfXml, dummyResolver);
    expect(gfShapes).toHaveLength(3);
    expect((gfShapes[0] as { _tblNode?: Record<string, Record<string, string>> } & PptxElement)._tblNode).toBeDefined();
    expect((gfShapes[1] as { _chartRelId?: string } & PptxElement)._chartRelId).toBe('rIdChart');
    expect(gfShapes[2].type).toBe('group');
  });

  it('parses embedded audio and video picture elements with custom posters', () => {
    const mediaXml = `<?xml version="1.0"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="50" name="Audio Shape"/>
          <p:cNvPicPr/>
          <p:nvPr>
            <a:audioFile r:link="rIdAudio"/>
          </p:nvPr>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rIdPosterAudio"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="100" y="200"/><a:ext cx="914400" cy="914400"/></a:xfrm>
        </p:spPr>
      </p:pic>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="51" name="Video Shape"/>
          <p:cNvPicPr/>
          <p:nvPr>
            <a:videoFile r:link="rIdVideo"/>
          </p:nvPr>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rIdPosterVideo"/>
        </p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="500" y="600"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    const relsXml = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdAudio" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/track.mp3"/>
  <Relationship Id="rIdPosterAudio" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/audio_cover.png"/>
  <Relationship Id="rIdVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/demo.mp4"/>
  <Relationship Id="rIdPosterVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/video_cover.jpg"/>
</Relationships>`;

    const resolver = createRelationshipResolver(relsXml, 'ppt/slides/slide1.xml');
    const mediaShapes = parseShapes(mediaXml, resolver);

    expect(mediaShapes).toHaveLength(2);
    expect(mediaShapes[0].elementType).toBe('audio');
    expect((mediaShapes[0] as any).audio.mediaId).toBe('ppt/media/track.mp3');
    expect((mediaShapes[0] as any).audio.mimeType).toBe('audio/mpeg');
    expect((mediaShapes[0] as any).audio.posterImageId).toBe('ppt/media/audio_cover.png');

    expect(mediaShapes[1].elementType).toBe('video');
    expect((mediaShapes[1] as any).video.mediaId).toBe('ppt/media/demo.mp4');
    expect((mediaShapes[1] as any).video.mimeType).toBe('video/mp4');
    expect((mediaShapes[1] as any).video.posterImageId).toBe('ppt/media/video_cover.jpg');
  });
});
