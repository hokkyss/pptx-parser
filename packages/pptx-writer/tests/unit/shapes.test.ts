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

interface SerializedShapeSpPr {
  'a:xfrm'?: {
    'a:ext'?: {
      '@_cx'?: number;
      '@_cy'?: number;
    };
  };
}

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

    const xmlObject = serializeConnector(connector);
    expect(xmlObject).toBeDefined();

    const nvCxnSpPr = xmlObject['p:nvCxnSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvCxnSpPr['p:cNvPr']['@_id']).toBe('5');
    expect(nvCxnSpPr['p:cNvPr']['@_name']).toBe('Flow Arrow 1');

    const spPr = xmlObject['p:spPr'] as Record<string, Record<string, unknown>>;
    expect(spPr['a:prstGeom']['@_prst']).toBe('bentConnector2');

    const ln = spPr['a:ln'] as Record<string, Record<string, unknown>>;
    expect(ln['@_w']).toBe(25400);

    const headEnd = ln['a:headEnd'];
    expect(headEnd['@_type']).toBe('triangle');
    expect(headEnd['@_w']).toBe('lg');
    expect(headEnd['@_len']).toBe('lg');

    const tailEnd = ln['a:tailEnd'];
    expect(tailEnd['@_type']).toBe('oval');
    expect(tailEnd['@_w']).toBe('sm');
    expect(tailEnd['@_len']).toBe('med');
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

    const xmlObject = serializeConnector(connector);
    expect(xmlObject).toBeDefined();

    const nvCxnSpPr = xmlObject['p:nvCxnSpPr'] as Record<string, Record<string, unknown>>;
    const cNvCxnSpPr = nvCxnSpPr['p:cNvCxnSpPr'] as Record<string, Record<string, unknown>>;
    expect(cNvCxnSpPr['a:cxnSpLocks']).toEqual({});
    expect(cNvCxnSpPr['a:stCxn']).toEqual({
      '@_id': 'card-1',
      '@_idx': 3, // 'right' -> 3
    });
    expect(cNvCxnSpPr['a:endCxn']).toEqual({
      '@_id': 'card-2',
      '@_idx': 1, // 'left' -> 1
    });
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

    const xmlObject = serializeShape(shape);
    const nvSpPr = xmlObject['p:nvSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvSpPr['p:cNvPr']['@_hidden']).toBe('1');
    expect(nvSpPr['p:cNvPr']['a:hlinkClick']).toBeDefined();
    expect(nvSpPr['p:cNvSpPr']['@_txBox']).toBe('1');
    expect(nvSpPr['p:cNvSpPr']['a:spLocks']).toBeDefined();
    expect((nvSpPr['p:nvPr']['p:ph'] as Record<string, unknown>)['@_type']).toBe('body');

    const spPr = xmlObject['p:spPr'] as Record<string, Record<string, unknown>>;
    expect(spPr['a:effectLst']).toBeDefined();
    expect(xmlObject['p:txBody']).toBeDefined();
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

    const xmlObject = serializeShape(shape);
    const spPr = xmlObject['p:spPr'] as Record<string, Record<string, unknown>>;
    expect(spPr['a:prstGeom']).toBeDefined();
    const prstGeom = spPr['a:prstGeom'] as Record<string, Record<string, unknown>>;
    expect(prstGeom['@_prst']).toBe('roundRect');
    expect(prstGeom['a:avLst']).toBeDefined();
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

    const xmlObject = serializeConnector(connector);
    expect(xmlObject).toHaveProperty('p:nvCxnSpPr');
    expect(xmlObject).toHaveProperty('p:spPr');

    const nvCxnSpPr = xmlObject['p:nvCxnSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvCxnSpPr['p:cNvPr']['@_id']).toBe('10');
    expect(nvCxnSpPr['p:cNvPr']['@_name']).toBe('Arrow Connector');
    expect(nvCxnSpPr['p:cNvPr']['a:hlinkClick']).toBeDefined();

    const spPr = xmlObject['p:spPr'] as Record<string, Record<string, unknown>>;
    expect(spPr['a:ln']).toBeDefined();
    expect(spPr['a:prstGeom']['@_prst']).toBe('straightConnector1');
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

    const xmlConn = serializeConnector(hiddenConnector);
    const nvCxnSpPr = xmlConn['p:nvCxnSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvCxnSpPr['p:cNvPr']['@_hidden']).toBe('1');

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

    const xmlShape = serializeShape(shapeWithoutExt);
    const spPr = xmlShape['p:spPr'] as SerializedShapeSpPr;
    expect(spPr['a:xfrm']?.['a:ext']?.['@_cx']).toBe(2000000);
    expect(spPr['a:xfrm']?.['a:ext']?.['@_cy']).toBe(1000000);
  });
});

describe('Shape Serializer helper direct exports', () => {
  it('covers serializeLine, serializeGeometry, and serializeShadow with empty inputs', () => {
    expect(serializeLine(undefined)).toBeUndefined();
    expect(serializeLine({ dashStyle: 'dash' })?.['a:prstDash']).toEqual({ '@_val': 'dash' });
    expect(serializeGeometry(undefined)).toEqual({ 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} } });
    expect(serializeGeometry({})).toEqual({ 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} } });
    expect(serializeShadow(undefined)).toBeUndefined();
  });

  it('covers various dashStyles, lines, shadows and shape locks', () => {
    const lineObj = serializeLine({
      dashStyle: 'dashDot',
      fill: { solidColor: { type: 'srgb', value: '123456' }, type: 'solid' },
      width: emu(12700),
    });
    expect(lineObj?.['@_w']).toBe(12700);
    expect(lineObj?.['a:prstDash']).toEqual({ '@_val': 'dashDot' });

    const geom1 = serializeGeometry({ adjustments: { adj: 50000 }, presetGeometry: 'rect' });
    expect(geom1['a:prstGeom']).toBeDefined();

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
      const xml = serializeShape(s);
      expect((xml['p:spPr'] as Record<string, Record<string, string>>)['a:prstGeom']).toBeDefined();
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
    const phXml = serializeShape(phShape);
    expect((phXml['p:spPr'] as Record<string, unknown>)['a:xfrm']).toBeUndefined();

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
    const connXml = serializeConnector(connMinimal);
    const nv = connXml['p:nvCxnSpPr'] as Record<string, Record<string, string>>;
    expect(nv['p:cNvPr']['@_id']).toBe('');
    expect(nv['p:cNvPr']['@_name']).toBe('');
    const xfrm = (connXml['p:spPr'] as Record<string, Record<string, Record<string, number>>>)['a:xfrm'];
    expect(xfrm['a:off']['@_x']).toBe(0);
    expect(xfrm['a:ext']['@_cx']).toBe(100000);
  });
});
