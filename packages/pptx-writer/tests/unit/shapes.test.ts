import { describe, expect, it } from 'vitest';
import type { PptxConnectorElement, PptxShapeElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import {
  serializeConnector,
  serializeGeometry,
  serializeLine,
  serializeShadow,
  serializeShape,
} from '../../lib/serializers/shape-serializer';
import { renderXml } from '../../lib/xml/xml-element';

describe('Shape Serializer (@hokkyss/pptx-writer)', () => {
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
      rotation: emuDegree(5400000),
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

    const xmlNode = serializeShape(shape);
    expect(xmlNode).toBeDefined();
    const xml = renderXml(xmlNode);

    // nvSpPr / cNvPr
    expect(xml).toContain('id="2"');
    expect(xml).toContain('name="Rectangle 1"');

    // xfrm
    expect(xml).toContain('x="100000"');
    expect(xml).toContain('y="200000"');
    expect(xml).toContain('cx="3000000"');
    expect(xml).toContain('cy="1500000"');
    expect(xml).toContain('rot="5400000"');

    // geometry
    expect(xml).toContain('prst="rect"');
    // fill
    expect(xml).toContain('val="007ACC"');
    // line
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('w="12700"');
    expect(xml).toContain('val="000000"');
    // txBody
    expect(xml).toContain('Box Content');
  });

  it('serializes connector element with customizable start and end arrows', () => {
    const connector = {
      elementType: 'connector' as const,
      id: '5',
      isVisible: true,
      line: {
        fill: { solidColor: { type: 'srgb' as const, value: '0284C7' }, type: 'solid' as const },
        headEnd: { length: 'lg' as const, type: 'triangle' as const, width: 'lg' as const },
        tailEnd: { length: 'med' as const, type: 'oval' as const, width: 'sm' as const },
        width: emu(25400),
      },
      name: 'Flow Arrow 1',
      position: { cx: emu(3000000), cy: emu(1000000), x: emu(500000), y: emu(500000) },
      rotation: emuDegree(0),
      shapeType: 'bentConnector2',
      type: 'connector' as const,
      zIndex: 1,
    };

    const xmlNode = serializeConnector(connector);
    expect(xmlNode).toBeDefined();
    const xml = renderXml(xmlNode);

    expect(xml).toContain('id="5"');
    expect(xml).toContain('name="Flow Arrow 1"');
    expect(xml).toContain('prst="bentConnector2"');
    expect(xml).toContain('w="25400"');

    // headEnd
    expect(xml).toContain('<a:headEnd');
    expect(xml).toContain('type="triangle"');
    expect(xml).toContain('w="lg"');
    expect(xml).toContain('len="lg"');

    // tailEnd
    expect(xml).toContain('<a:tailEnd');
    expect(xml).toContain('type="oval"');
    expect(xml).toContain('w="sm"');
    expect(xml).toContain('len="med"');
  });

  it('serializes connector with startConnection and endConnection shape attachments', () => {
    const connector = {
      elementType: 'connector' as const,
      endConnection: { position: 'left' as const, shapeId: 'card-2' },
      id: '10',
      isVisible: true,
      name: 'Attached Connector',
      position: { cx: emu(2000000), cy: emu(0), x: emu(1000000), y: emu(1000000) },
      rotation: emuDegree(0),
      shapeType: 'line',
      startConnection: { position: 'right' as const, shapeId: 'card-1' },
      type: 'connector' as const,
      zIndex: 2,
    };

    const xmlNode = serializeConnector(connector);
    const xml = renderXml(xmlNode);

    // stCxn — 'right' → idx 3
    expect(xml).toContain('<a:stCxn');
    expect(xml).toContain('id="card-1"');
    expect(xml).toContain('idx="3"');

    // endCxn — 'left' → idx 1
    expect(xml).toContain('<a:endCxn');
    expect(xml).toContain('id="card-2"');
    expect(xml).toContain('idx="1"');

    // cxnSpLocks
    expect(xml).toContain('<a:cxnSpLocks/>');
  });

  it('serializes shape locks, placeholders, shadows, and text box attributes', () => {
    const shape: PptxShapeElement = {
      elementType: 'shape',
      type: 'shape',
      id: '3',
      name: 'Locked Box',
      isTextBox: true,
      isVisible: false,
      zIndex: 1,
      position: { x: emu(0), y: emu(0), cx: emu(1000000), cy: emu(1000000) },
      rotation: emuDegree(0),
      locks: {
        noGrp: true,
        noRot: true,
        noChangeAspect: true,
        noMove: true,
        noResize: true,
        noEditPoints: true,
        noAdjustHandles: true,
        noChangeShapeType: true,
      },
      placeholder: { type: 'body', idx: 1 },
      shadow: {
        blurRadius: emu(50000),
        distance: emu(30000),
        direction: emuDegree(5400000),
        alignment: 'ctr',
        rotateWithShape: true,
        opacity: 0.5,
        color: '#333333',
      },
      hyperlink: { rId: 'rId5' },
    };

    const xml = renderXml(serializeShape(shape));
    expect(xml).toContain('hidden="1"');
    expect(xml).toContain('<a:hlinkClick');
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain('<a:spLocks');
    expect(xml).toContain('type="body"');
    expect(xml).toContain('<a:effectLst');
  });

  it('serializes geometry adjustments into preset geometry', () => {
    const shape: PptxShapeElement = {
      elementType: 'shape',
      type: 'shape',
      id: '4',
      name: 'Adjusted Shape',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(2000000), cy: emu(2000000) },
      rotation: emuDegree(0),
      geometry: {
        presetGeometry: 'roundRect',
        adjustments: { adj1: 50000 },
      },
    };

    const xml = renderXml(serializeShape(shape));
    expect(xml).toContain('prst="roundRect"');
    expect(xml).toContain('<a:avLst>');
  });

  it('serializes connector element (<p:cxnSp>)', () => {
    const connector: PptxConnectorElement = {
      elementType: 'connector',
      type: 'connector',
      id: '10',
      name: 'Arrow Connector',
      isVisible: true,
      zIndex: 0,
      shapeType: 'straightConnector1',
      position: { x: emu(100), y: emu(200), cx: emu(500000), cy: emu(500000) },
      rotation: emuDegree(0),
      line: {
        width: emu(25400),
        fill: { type: 'solid', solidColor: { type: 'srgb', value: 'FF5500' } },
      },
      hyperlink: { rId: 'rId8' },
    };

    const xmlNode = serializeConnector(connector);
    const xml = renderXml(xmlNode);
    expect(xml).toContain('<p:nvCxnSpPr>');
    expect(xml).toContain('<p:spPr>');
    expect(xml).toContain('id="10"');
    expect(xml).toContain('name="Arrow Connector"');
    expect(xml).toContain('<a:hlinkClick');
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('prst="straightConnector1"');
  });
});

describe('Shape and Connector edge cases', () => {
  it('serializes hidden connector and shape with fallback xfrm bounds', () => {
    const hiddenConnector: PptxConnectorElement = {
      elementType: 'connector',
      type: 'connector',
      id: '11',
      name: 'Hidden Connector',
      isVisible: false,
      zIndex: 0,
      position: { x: emu(0), y: emu(0), cx: emu(100), cy: emu(100) },
      rotation: emuDegree(0),
    };

    const connXml = renderXml(serializeConnector(hiddenConnector));
    expect(connXml).toContain('hidden="1"');

    const shapeWithoutExt: PptxShapeElement = {
      elementType: 'shape',
      type: 'shape',
      id: '12',
      name: 'No Ext Shape',
      isVisible: true,
      zIndex: 0,
      position: { x: emu(100), y: emu(200), cx: emu(0), cy: emu(0) },
      rotation: emuDegree(0),
    };
    // @ts-expect-error test shape without cx / cy dimensions
    delete shapeWithoutExt.position.cx;
    // @ts-expect-error test shape without cx / cy dimensions
    delete shapeWithoutExt.position.cy;

    const shapeXml = renderXml(serializeShape(shapeWithoutExt));
    expect(shapeXml).toContain('cx="2000000"');
    expect(shapeXml).toContain('cy="1000000"');
  });
});

describe('Shape Serializer helper direct exports', () => {
  it('covers serializeLine, serializeGeometry, and serializeShadow with empty inputs', () => {
    expect(serializeLine(undefined)).toBeUndefined();
    const lineWithDash = serializeLine({ dashStyle: 'dash' });
    expect(lineWithDash).toBeDefined();
    expect(renderXml(lineWithDash)).toContain('val="dash"');
    expect(renderXml(lineWithDash)).toContain('<a:prstDash');

    const defaultGeom = serializeGeometry(undefined);
    const defaultGeomXml = renderXml(defaultGeom);
    expect(defaultGeomXml).toContain('prst="rect"');
    expect(defaultGeomXml).toContain('<a:avLst/>');

    const emptyGeom = serializeGeometry({});
    expect(renderXml(emptyGeom)).toContain('prst="rect"');

    expect(serializeShadow(undefined)).toBeUndefined();
  });

  it('covers various dashStyles, lines, shadows and shape locks', () => {
    const lineObj = serializeLine({
      dashStyle: 'dashDot',
      fill: { solidColor: { type: 'srgb', value: '123456' }, type: 'solid' },
      width: emu(12700),
    });
    expect(lineObj).toBeDefined();
    const lineXml = renderXml(lineObj);
    expect(lineXml).toContain('w="12700"');
    expect(lineXml).toContain('val="dashDot"');

    const geom1 = serializeGeometry({ adjustments: { adj: 50000 }, presetGeometry: 'rect' });
    expect(renderXml(geom1)).toContain('prst="rect"');

    const shadowObj = serializeShadow({
      blurRadius: emu(50800),
      color: '333333',
      direction: emuDegree(5400000),
      distance: emu(38100),
      opacity: 0.5,
    });
    expect(shadowObj).toBeDefined();

    const shape: PptxShapeElement = {
      elementType: 'shape',
      id: '50',
      isLocked: true,
      isVisible: false,
      locks: {
        noAdjustHandles: true,
        noChangeAspect: true,
        noChangeShapeType: true,
        noCrop: true,
        noEditPoints: true,
        noGrp: true,
        noMove: true,
        noResize: true,
        noRot: true,
        noSelect: true,
        noUngrp: true,
      },
      name: 'Locked Shape',
      placeholder: { idx: '0', type: 'title' },
      position: { cx: emu(100), cy: emu(100), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      shapeType: 'ellipse',
      type: 'shape',
      zIndex: 0,
    };
    const shapeXml = serializeShape(shape);
    expect(shapeXml).toBeDefined();
  });

  it('covers PRESET_GEOMETRY_MAP, placeholder without size, and connector position/attachment fallbacks', () => {
    // PRESET_GEOMETRY_MAP types
    const geomTypes = ['box', 'circle', 'cylinder', 'oval', 'square', 'star', 'wedgeRoundRect'];
    for (const shapeType of geomTypes) {
      const s: PptxShapeElement = {
        elementType: 'shape',
        id: '1',
        isVisible: true,
        name: 'S',
        position: { cx: emu(100), cy: emu(100), x: emu(0), y: emu(0) },
        rotation: emuDegree(0),
        shapeType,
        type: 'shape',
        zIndex: 0,
      };
      const xml = renderXml(serializeShape(s));
      expect(xml).toContain('<a:prstGeom');
    }

    // Placeholder shape with 0 size (inherits geometry from layout)
    const phShape: PptxShapeElement = {
      elementType: 'shape',
      id: '2',
      isVisible: true,
      name: 'PH',
      placeholder: { type: 'body' },
      position: { cx: emu(0), cy: emu(0), x: emu(0), y: emu(0) },
      rotation: emuDegree(0),
      type: 'shape',
      zIndex: 0,
    };
    const phXml = renderXml(serializeShape(phShape));
    // No xfrm when cx/cy are 0 and there's a placeholder
    expect(phXml).not.toContain('<a:xfrm');

    // Connector with empty name/id, undefined position, and top/bottom connection points
    // @ts-expect-error Testing undefined position
    const connMinimal: PptxConnectorElement = {
      elementType: 'connector',
      endConnection: { position: 'bottom', shapeId: 's2' },
      id: '',
      isVisible: true,
      name: '',
      rotation: emuDegree(0),
      startConnection: { position: 'top', shapeId: 's1' },
      type: 'connector',
      zIndex: 0,
    };
    const connXml = renderXml(serializeConnector(connMinimal));
    expect(connXml).toContain('id="2"');
    expect(connXml).toContain('name="Connector 2"');
    expect(connXml).toContain('x="0"');
    expect(connXml).toContain('cx="100000"');
  });
});
