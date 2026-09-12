import type { PptxGroupElement } from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';
import { serializePicture } from './picture-serializer';
import { serializeConnector, serializeShape } from './shape-serializer';
import { serializeTable } from './table-serializer';

/**
 * Serializes a group shape `<p:grpSp>` containing nested children strictly conforming to ECMA-376 schema.
 * Preserves the exact sequential z-order of child elements.
 */
export function serializeGroup(group: PptxGroupElement): XmlElement {
  const childNodes: XmlElement[] = [];

  for (const child of group.children || []) {
    if (child.elementType === 'shape') {
      childNodes.push(serializeShape(child));
    } else if (child.elementType === 'table') {
      childNodes.push(serializeTable(child));
    } else if (child.elementType === 'picture') {
      childNodes.push(serializePicture(child));
    } else if (child.elementType === 'group') {
      childNodes.push(serializeGroup(child));
    } else if (child.elementType === 'connector') {
      childNodes.push(serializeConnector(child));
    }
  }

  const nvGrpSpPr = el('p:nvGrpSpPr', [
    el('p:cNvPr', {
      id: group.id || '5',
      name: group.name || `Group ${group.id || '5'}`,
    }),
    el('p:cNvGrpSpPr'),
    el('p:nvPr'),
  ]);

  const grpSpPr = el('p:grpSpPr', [
    el('a:xfrm', [
      el('a:off', {
        x: Math.round(Number(group.position?.x ?? 0)),
        y: Math.round(Number(group.position?.y ?? 0)),
      }),
      el('a:ext', {
        cx: Math.round(Number(group.position?.cx ?? 1000000)),
        cy: Math.round(Number(group.position?.cy ?? 1000000)),
      }),
      el('a:chOff', {
        x: Math.round(Number(group.position?.x ?? 0)),
        y: Math.round(Number(group.position?.y ?? 0)),
      }),
      el('a:chExt', {
        cx: Math.round(Number(group.position?.cx ?? 1000000)),
        cy: Math.round(Number(group.position?.cy ?? 1000000)),
      }),
    ]),
  ]);

  return el('p:grpSp', [
    nvGrpSpPr,
    grpSpPr,
    ...childNodes,
  ]);
}
