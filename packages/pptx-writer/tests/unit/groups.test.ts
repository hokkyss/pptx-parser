import { describe, expect, it } from 'vitest';
import type { PptxGroupElement, PptxElement } from '@hokkyss/pptx-core';
import { emu, emuDegree } from '@hokkyss/pptx-core';
import { serializeGroup } from '../../lib/serializers/group-serializer';

function makeGroup(overrides: Partial<PptxGroupElement> = {}): PptxGroupElement {
  return {
    elementType: 'group',
    type: 'group',
    id: '5',
    name: 'Group 5',
    isVisible: true,
    zIndex: 0,
    position: { x: emu(0), y: emu(0), cx: emu(1000000), cy: emu(1000000) },
    rotation: emuDegree(0),
    children: [],
    ...overrides,
  };
}

function makeShape(id = '1'): PptxElement {
  return {
    elementType: 'shape',
    type: 'shape',
    id,
    name: `Shape ${id}`,
    isVisible: true,
    zIndex: 0,
    position: { x: emu(0), y: emu(0), cx: emu(500000), cy: emu(500000) },
    rotation: emuDegree(0),
    textBody: { bodyProperties: {}, paragraphs: [] },
  };
}

function makeConnector(id = '1'): PptxElement {
  return {
    elementType: 'connector',
    type: 'connector',
    id,
    name: `Connector ${id}`,
    isVisible: true,
    zIndex: 0,
    position: { x: emu(0), y: emu(0), cx: emu(100000), cy: emu(100000) },
    rotation: emuDegree(0),
  };
}

function makePicture(id = '1'): PptxElement {
  return {
    elementType: 'picture',
    type: 'picture',
    id,
    name: `Pic ${id}`,
    isVisible: true,
    zIndex: 0,
    position: { x: emu(0), y: emu(0), cx: emu(500000), cy: emu(500000) },
    rotation: emuDegree(0),
    picture: { mediaId: 'rId5' },
  };
}

function makeTable(id = '1'): PptxElement {
  return {
    elementType: 'table',
    type: 'graphicFrame',
    id,
    name: `Table ${id}`,
    isVisible: true,
    zIndex: 0,
    position: { x: emu(0), y: emu(0), cx: emu(2000000), cy: emu(1000000) },
    rotation: emuDegree(0),
    table: { columnWidths: [emu(2000000)], rows: [] },
  };
}

describe('serializeGroup', () => {
  // ── Basic structure ───────────────────────────────────────────────────────────

  it('produces a node with p:nvGrpSpPr and p:grpSpPr', () => {
    const result = serializeGroup(makeGroup());
    expect(result).toHaveProperty('p:nvGrpSpPr');
    expect(result).toHaveProperty('p:grpSpPr');
  });

  it('uses group id and name in p:cNvPr', () => {
    const result = serializeGroup(makeGroup({ id: '42', name: 'My Group' }));
    const nvGrpSpPr = result['p:nvGrpSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvGrpSpPr['p:cNvPr']['@_id']).toBe('42');
    expect(nvGrpSpPr['p:cNvPr']['@_name']).toBe('My Group');
  });

  it('falls back to id "5" and default name when id and name are empty strings', () => {
    const result = serializeGroup(makeGroup({ id: '', name: '' }));
    const nvGrpSpPr = result['p:nvGrpSpPr'] as Record<string, Record<string, unknown>>;
    expect(nvGrpSpPr['p:cNvPr']['@_id']).toBe('5');
    expect(nvGrpSpPr['p:cNvPr']['@_name']).toBe('Group 5');
  });

  it('encodes custom position values into a:off, a:ext, a:chOff, a:chExt', () => {
    const group = makeGroup({ position: { x: emu(100), y: emu(200), cx: emu(3000000), cy: emu(2000000) } });
    const result = serializeGroup(group);
    const xfrm = (result['p:grpSpPr'] as Record<string, unknown>)['a:xfrm'] as Record<string, Record<string, unknown>>;
    expect(xfrm['a:off']['@_x']).toBe(100);
    expect(xfrm['a:off']['@_y']).toBe(200);
    expect(xfrm['a:ext']['@_cx']).toBe(3000000);
    expect(xfrm['a:ext']['@_cy']).toBe(2000000);
    expect(xfrm['a:chOff']['@_x']).toBe(100);
    expect(xfrm['a:chExt']['@_cx']).toBe(3000000);
  });

  it('defaults position values to 0/1000000 when position is undefined', () => {
    // @ts-expect-error Testing undefined position resilience
    const group = makeGroup({ position: undefined });
    const result = serializeGroup(group);
    const xfrm = (result['p:grpSpPr'] as Record<string, unknown>)['a:xfrm'] as Record<string, Record<string, unknown>>;
    expect(xfrm['a:off']['@_x']).toBe(0);
    expect(xfrm['a:off']['@_y']).toBe(0);
    expect(xfrm['a:ext']['@_cx']).toBe(1000000);
    expect(xfrm['a:ext']['@_cy']).toBe(1000000);
  });

  // ── Empty / no children ───────────────────────────────────────────────────────

  it('does not emit any child list keys when group has no children', () => {
    const result = serializeGroup(makeGroup({ children: [] }));
    expect(result['p:sp']).toBeUndefined();
    expect(result['p:graphicFrame']).toBeUndefined();
    expect(result['p:pic']).toBeUndefined();
    expect(result['p:grpSp']).toBeUndefined();
    expect(result['p:cxnSp']).toBeUndefined();
  });

  it('handles undefined children gracefully', () => {
    // @ts-expect-error Testing undefined children resilience
    const group = makeGroup({ children: undefined });
    expect(() => serializeGroup(group)).not.toThrow();
    const result = serializeGroup(group);
    expect(result['p:sp']).toBeUndefined();
  });

  // ── Individual child types ────────────────────────────────────────────────────

  it('emits p:sp array for shape children', () => {
    const result = serializeGroup(makeGroup({ children: [makeShape('1'), makeShape('2')] }));
    const list = result['p:sp'];
    expect(Array.isArray(list)).toBe(true);
    if (Array.isArray(list)) {
      expect(list).toHaveLength(2);
    }
  });

  it('emits p:graphicFrame array for table children', () => {
    const result = serializeGroup(makeGroup({ children: [makeTable('10')] }));
    const list = result['p:graphicFrame'];
    expect(Array.isArray(list)).toBe(true);
    if (Array.isArray(list)) {
      expect(list).toHaveLength(1);
    }
  });

  it('emits p:pic array for picture children', () => {
    const result = serializeGroup(makeGroup({ children: [makePicture('20')] }));
    const list = result['p:pic'];
    expect(Array.isArray(list)).toBe(true);
    if (Array.isArray(list)) {
      expect(list).toHaveLength(1);
    }
  });

  it('emits p:cxnSp array for connector children', () => {
    const result = serializeGroup(makeGroup({ children: [makeConnector('30')] }));
    const list = result['p:cxnSp'];
    expect(Array.isArray(list)).toBe(true);
    if (Array.isArray(list)) {
      expect(list).toHaveLength(1);
    }
  });

  // ── Nested groups (recursive) ─────────────────────────────────────────────────

  it('emits p:grpSp for nested group children and recurses into them', () => {
    const innerGroup = makeGroup({ id: '99', name: 'Inner', children: [makeShape('1')] });
    const outerGroup = makeGroup({ id: '100', children: [innerGroup] });
    const result = serializeGroup(outerGroup);
    const nestedGrpSp = result['p:grpSp'];
    expect(Array.isArray(nestedGrpSp)).toBe(true);
    if (Array.isArray(nestedGrpSp)) {
      expect(nestedGrpSp).toHaveLength(1);
      // Inner result should itself have a p:sp from its shape child
      const innerResult = nestedGrpSp[0] as Record<string, unknown>;
      expect(innerResult['p:sp']).toBeDefined();
    }
  });

  // ── Mixed children ────────────────────────────────────────────────────────────

  it('populates the correct lists for mixed child types', () => {
    const group = makeGroup({
      children: [makeShape('1'), makeConnector('2'), makePicture('3'), makeTable('4')],
    });
    const result = serializeGroup(group);
    const sp = result['p:sp'];
    const cxnSp = result['p:cxnSp'];
    const pic = result['p:pic'];
    const gf = result['p:graphicFrame'];
    expect(Array.isArray(sp) && sp.length).toBe(1);
    expect(Array.isArray(cxnSp) && cxnSp.length).toBe(1);
    expect(Array.isArray(pic) && pic.length).toBe(1);
    expect(Array.isArray(gf) && gf.length).toBe(1);
    expect(result['p:grpSp']).toBeUndefined();
  });
});
