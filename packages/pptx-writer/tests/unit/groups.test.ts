import { describe, expect, it } from 'vitest';
import type { PptxGroupElement, PptxElement } from '@hokkyss/pptx-core';
import { emu, emuDegree } from '@hokkyss/pptx-core';
import { serializeGroup } from '../../lib/serializers/group-serializer';
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

/** Returns direct (non-recursive) children of a node. */
function directChildren(node: XmlElement): XmlElement[] {
  return (node.children ?? []).filter(
    (c): c is XmlElement => typeof c === 'object' && c !== null && 'tag' in c,
  );
}

/** Returns direct children with a specific tag. */
function directChildrenByTag(node: XmlElement, tag: string): XmlElement[] {
  return directChildren(node).filter((c) => c.tag === tag);
}

/**
 *
 */
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

/**
 *
 */
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

/**
 *
 */
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

/**
 *
 */
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

/**
 *
 */
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
    // The result is a p:grpSp node; its direct children include nvGrpSpPr and grpSpPr
    expect(findEl(result, 'p:nvGrpSpPr')).toBeDefined();
    expect(findEl(result, 'p:grpSpPr')).toBeDefined();
  });

  it('uses group id and name in p:cNvPr', () => {
    const result = serializeGroup(makeGroup({ id: '42', name: 'My Group' }));
    const xml = renderXml(result);
    expect(xml).toContain('id="42"');
    expect(xml).toContain('name="My Group"');
  });

  it('falls back to id "5" and default name when id and name are empty strings', () => {
    const result = serializeGroup(makeGroup({ id: '', name: '' }));
    const xml = renderXml(result);
    expect(xml).toContain('id="5"');
    expect(xml).toContain('name="Group 5"');
  });

  it('encodes custom position values into a:off, a:ext, a:chOff, a:chExt', () => {
    const group = makeGroup({ position: { x: emu(100), y: emu(200), cx: emu(3000000), cy: emu(2000000) } });
    const result = serializeGroup(group);
    const xml = renderXml(result);
    // a:off
    expect(xml).toContain('<a:off');
    expect(xml).toContain('x="100"');
    expect(xml).toContain('y="200"');
    // a:ext
    expect(xml).toContain('<a:ext');
    expect(xml).toContain('cx="3000000"');
    expect(xml).toContain('cy="2000000"');
    // a:chOff and a:chExt
    expect(xml).toContain('<a:chOff');
    expect(xml).toContain('<a:chExt');
  });

  it('defaults position values to 0/1000000 when position is undefined', () => {
    const group = makeGroup({ position: undefined });
    const result = serializeGroup(group);
    const xml = renderXml(result);
    expect(xml).toContain('x="0"');
    expect(xml).toContain('y="0"');
    expect(xml).toContain('cx="1000000"');
    expect(xml).toContain('cy="1000000"');
  });

  // ── Empty / no children ───────────────────────────────────────────────────────

  it('does not emit any child element tags when group has no children', () => {
    const result = serializeGroup(makeGroup({ children: [] }));
    const xml = renderXml(result);
    // Only nvGrpSpPr and grpSpPr should be direct children — no p:sp, p:pic, etc.
    expect(xml).not.toContain('<p:sp>');
    expect(xml).not.toContain('<p:graphicFrame>');
    expect(xml).not.toContain('<p:pic>');
    expect(xml).not.toContain('<p:cxnSp>');
    // Also should not contain nested grpSp beyond the root
    const topChildren = directChildren(result);
    // Should only have nvGrpSpPr and grpSpPr (2 children, no extra)
    expect(topChildren).toHaveLength(2);
  });

  it('handles undefined children gracefully', () => {
    const group = makeGroup({ children: undefined });
    expect(() => serializeGroup(group)).not.toThrow();
    const result = serializeGroup(group);
    // No shape elements emitted
    expect(renderXml(result)).not.toContain('<p:sp>');
  });

  // ── Individual child types ────────────────────────────────────────────────────

  it('emits p:sp for shape children', () => {
    const result = serializeGroup(makeGroup({ children: [makeShape('1'), makeShape('2')] }));
    const shapes = directChildrenByTag(result, 'p:sp');
    expect(shapes).toHaveLength(2);
  });

  it('emits p:graphicFrame for table children', () => {
    const result = serializeGroup(makeGroup({ children: [makeTable('10')] }));
    const frames = directChildrenByTag(result, 'p:graphicFrame');
    expect(frames).toHaveLength(1);
  });

  it('emits p:pic for picture children', () => {
    const result = serializeGroup(makeGroup({ children: [makePicture('20')] }));
    const pics = directChildrenByTag(result, 'p:pic');
    expect(pics).toHaveLength(1);
  });

  it('emits p:cxnSp for connector children', () => {
    const result = serializeGroup(makeGroup({ children: [makeConnector('30')] }));
    const cxns = directChildrenByTag(result, 'p:cxnSp');
    expect(cxns).toHaveLength(1);
  });

  // ── Nested groups (recursive) ─────────────────────────────────────────────────

  it('emits p:grpSp for nested group children and recurses into them', () => {
    const innerGroup = makeGroup({ id: '99', name: 'Inner', children: [makeShape('1')] });
    const outerGroup = makeGroup({ id: '100', children: [innerGroup] });
    const result = serializeGroup(outerGroup);

    // The outer grpSp should have exactly 1 p:grpSp as a direct child
    const nestedGrpSp = directChildrenByTag(result, 'p:grpSp');
    expect(nestedGrpSp).toHaveLength(1);
    // The inner grpSp should contain a p:sp (from its shape child)
    const innerResult = nestedGrpSp[0];
    const innerShapes = directChildrenByTag(innerResult, 'p:sp');
    expect(innerShapes).toHaveLength(1);
  });

  // ── Mixed children ────────────────────────────────────────────────────────────

  it('populates the correct lists for mixed child types', () => {
    const group = makeGroup({
      children: [makeShape('1'), makeConnector('2'), makePicture('3'), makeTable('4')],
    });
    const result = serializeGroup(group);
    expect(directChildrenByTag(result, 'p:sp')).toHaveLength(1);
    expect(directChildrenByTag(result, 'p:cxnSp')).toHaveLength(1);
    expect(directChildrenByTag(result, 'p:pic')).toHaveLength(1);
    expect(directChildrenByTag(result, 'p:graphicFrame')).toHaveLength(1);
    expect(directChildrenByTag(result, 'p:grpSp')).toHaveLength(0);
  });
});

describe('Group Serializer partial position fallback', () => {
  it('falls back cx and cy to 1000000 when only x and y are provided', () => {
    // @ts-expect-error Testing partial position resilience
    const group = makeGroup({ position: { x: emu(500), y: emu(600) } });
    const result = serializeGroup(group);
    const xml = renderXml(result);
    expect(xml).toContain('x="500"');
    expect(xml).toContain('cx="1000000"');
  });
});
