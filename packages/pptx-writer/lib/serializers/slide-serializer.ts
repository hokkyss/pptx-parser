import type { PptxElement, PptxSlide } from '@hokkyss/pptx-core';
import { el, serializeXml, type XmlElement } from '../xml/xml-element';
import { serializeAnimations } from './animation-serializer';
import { serializeGroup } from './group-serializer';
import { serializePicture } from './picture-serializer';
import { serializeConnector, serializeShape } from './shape-serializer';
import { serializeTable } from './table-serializer';
import { serializeFill } from './text-serializer';
import { serializeTransition } from './transition-serializer';

/**
 * Serializes slide background properties `<p:bg>`.
 */
export function serializeSlideBackground(slide: PptxSlide): undefined | XmlElement {
  if (!slide.background?.fill) return undefined;

  const fillNode = serializeFill(slide.background.fill);
  if (!fillNode) return undefined;

  return el('p:bg', [
    el('p:bgPr', [
      fillNode,
      el('a:effectLst'),
    ]),
  ]);
}

/**
 * Checks if a string represents a positive unsigned integer (>= 1).
 */
function isUnsignedInt(str?: string): boolean {
  return typeof str === 'string' && /^[1-9]\d*$/.test(str);
}

/**
 * Normalizes element and all its children with valid numeric OpenXML drawing IDs.
 */
function normalizeElementWithUniqueIds(
  element: PptxElement,
  getUniqueId: (preferredId?: string) => string,
): PptxElement {
  const uniqueId = getUniqueId(element.id);
  if (element.elementType === 'group') {
    const updatedChildren = (element.children || []).map((child) => normalizeElementWithUniqueIds(child, getUniqueId));
    return {
      ...element,
      children: updatedChildren,
      id: uniqueId,
    };
  }
  return {
    ...element,
    id: uniqueId,
  };
}

/**
 * Serializes a chart graphicFrame element.
 */
export function serializeChartGraphicFrame(
  elem: PptxElement,
  chartRelId: string,
): XmlElement {
  const x = Math.round(Number(elem.position?.x ?? 0));
  const y = Math.round(Number(elem.position?.y ?? 0));
  const cx = Math.round(Number(elem.position?.cx ?? 9144000));
  const cy = Math.round(Number(elem.position?.cy ?? 4572000));

  return el('p:graphicFrame', [
    el('p:nvGraphicFramePr', [
      el('p:cNvPr', {
        id: elem.id || '2',
        name: elem.name || `Chart ${elem.id || '2'}`,
      }),
      el('p:cNvGraphicFramePr'),
      el('p:nvPr'),
    ]),
    el('p:xfrm', [
      el('a:off', { x, y }),
      el('a:ext', { cx, cy }),
    ]),
    el('a:graphic', [
      el('a:graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' }, [
        el('c:chart', {
          'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'r:id': chartRelId,
        }),
      ]),
    ]),
  ]);
}

/**
 * Serializes a slide AST node into complete OpenXML `<p:sld>` format.
 * Preserves exact sequential visual z-order of child elements in `<p:spTree>`.
 */
export function serializeSlide(
  slide: PptxSlide,
  pictureEmbedMap?: Map<string, string>,
  chartRelIds?: string[],
): string {
  const elements = (slide.elements && slide.elements.length > 0) ? slide.elements : (slide.shapes || []);

  const usedIds = new Set<string>(['1']); // 1 is reserved for the root container group
  const idMap = new Map<string, string>();
  let nextId = 2;

  const getUniqueId = (preferredId?: string): string => {
    if (preferredId && isUnsignedInt(preferredId) && preferredId !== '1' && !usedIds.has(preferredId)) {
      usedIds.add(preferredId);
      idMap.set(preferredId, preferredId);
      return preferredId;
    }
    while (usedIds.has(String(nextId))) {
      nextId++;
    }
    const id = String(nextId++);
    usedIds.add(id);
    if (preferredId) {
      idMap.set(preferredId, id);
    }
    return id;
  };

  let chartIdx = 0;
  const spTreeElements: XmlElement[] = [];

  for (const rawEl of elements) {
    const elWithId = normalizeElementWithUniqueIds(rawEl, getUniqueId);

    if (elWithId.elementType === 'shape') {
      spTreeElements.push(serializeShape(elWithId));
    } else if (elWithId.elementType === 'table') {
      spTreeElements.push(serializeTable(elWithId));
    } else if (elWithId.elementType === 'chart') {
      const chartRelId = chartRelIds ? (chartRelIds[chartIdx++] || 'rId2') : 'rId2';
      spTreeElements.push(serializeChartGraphicFrame(elWithId, chartRelId));
    } else if (elWithId.elementType === 'picture') {
      const overrideEmbedId = pictureEmbedMap?.get(elWithId.picture.mediaId) ?? pictureEmbedMap?.get(elWithId.id);
      spTreeElements.push(serializePicture(elWithId, overrideEmbedId));
    } else if (elWithId.elementType === 'group') {
      spTreeElements.push(serializeGroup(elWithId));
    } else if (elWithId.elementType === 'connector') {
      // Map attached shape IDs to their normalized numeric IDs
      const mappedConnector = {
        ...elWithId,
        endConnection: elWithId.endConnection
          ? {
              ...elWithId.endConnection,
              shapeId: idMap.get(elWithId.endConnection.shapeId) ?? elWithId.endConnection.shapeId,
            }
          : undefined,
        startConnection: elWithId.startConnection
          ? {
              ...elWithId.startConnection,
              shapeId: idMap.get(elWithId.startConnection.shapeId) ?? elWithId.startConnection.shapeId,
            }
          : undefined,
      };
      spTreeElements.push(serializeConnector(mappedConnector));
    }
  }

  const spTree = el('p:spTree', [
    el('p:nvGrpSpPr', [
      el('p:cNvPr', { id: '1', name: '' }),
      el('p:cNvGrpSpPr'),
      el('p:nvPr'),
    ]),
    el('p:grpSpPr', [
      el('a:xfrm', [
        el('a:off', { x: 0, y: 0 }),
        el('a:ext', { cx: 0, cy: 0 }),
        el('a:chOff', { x: 0, y: 0 }),
        el('a:chExt', { cx: 0, cy: 0 }),
      ]),
    ]),
    ...spTreeElements,
  ]);

  const cSldChildren: XmlElement[] = [];
  const bg = serializeSlideBackground(slide);
  if (bg) {
    cSldChildren.push(bg);
  }
  cSldChildren.push(spTree);

  const cSld = el('p:cSld', cSldChildren);

  const sldChildren: XmlElement[] = [
    cSld,
    el('p:clrMapOvr', [el('a:masterClrMapping')]),
  ];

  const transitionNode = serializeTransition(slide.transition);
  if (transitionNode) {
    sldChildren.push(transitionNode);
  }

  const timingNode = serializeAnimations(slide.animations);
  if (timingNode) {
    sldChildren.push(timingNode);
  }

  const sldRoot = el('p:sld', {
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  }, sldChildren);

  return serializeXml(sldRoot);
}
