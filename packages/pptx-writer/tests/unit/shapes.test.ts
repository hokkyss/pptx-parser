import { describe, expect, it } from 'vitest';
import type { PptxConnectorElement, PptxShapeElement } from '@hokkyss/pptx-core';
import { emu, emuDegree, hundredthsPoint } from '@hokkyss/pptx-core';
import { serializeConnector, serializeShape } from '../../lib/serializers/shape-serializer';

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
