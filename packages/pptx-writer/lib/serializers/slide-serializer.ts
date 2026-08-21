import type { PptxElement, PptxSlide } from '@hokkyss/pptx-core';
import { serializeXml } from '../xml/xml-builder';
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
export function serializeSlideBackground(slide: PptxSlide): Record<string, unknown> | undefined {
  if (!slide.background?.fill) return undefined;

  const fillNode = serializeFill(slide.background.fill);
  if (!fillNode) return undefined;

  return {
    'p:bgPr': {
      ...fillNode,
      'a:effectLst': {},
    },
  };
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
  el: PptxElement,
  getUniqueId: (preferredId?: string) => string,
): PptxElement {
  const uniqueId = getUniqueId(el.id);
  if (el.elementType === 'group') {
    const updatedChildren = (el.children || []).map((child) => normalizeElementWithUniqueIds(child, getUniqueId));
    return {
      ...el,
      children: updatedChildren,
      id: uniqueId,
    };
  }
  return {
    ...el,
    id: uniqueId,
  };
}

/**
 * Serializes a chart graphicFrame element.
 */
export function serializeChartGraphicFrame(
  elem: PptxElement,
  chartRelId: string,
): Record<string, unknown> {
  const x = Math.round(Number(elem.position?.x ?? 0));
  const y = Math.round(Number(elem.position?.y ?? 0));
  const cx = Math.round(Number(elem.position?.cx ?? 9144000));
  const cy = Math.round(Number(elem.position?.cy ?? 4572000));

  return {
    'p:nvGraphicFramePr': {
      'p:cNvPr': {
        '@_id': elem.id || '2',
        '@_name': elem.name || `Chart ${elem.id || '2'}`,
      },
      'p:cNvGraphicFramePr': {},
      'p:nvPr': {},
    },
    'p:xfrm': {
      'a:off': { '@_x': String(x), '@_y': String(y) },
      'a:ext': { '@_cx': String(cx), '@_cy': String(cy) },
    },
    'a:graphic': {
      'a:graphicData': {
        '@_uri': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        'c:chart': {
          '@_xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
          '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          '@_r:id': chartRelId,
        },
      },
    },
  };
}

/**
 * Serializes a slide AST node into complete OpenXML `<p:sld>` format.
 */
export function serializeSlide(
  slide: PptxSlide,
  pictureEmbedMap?: Map<string, string>,
  chartRelIds?: string[],
): string {
  const elements = (slide.elements && slide.elements.length > 0) ? slide.elements : (slide.shapes || []);

  const shapeList: Record<string, unknown>[] = [];
  const graphicFrameList: Record<string, unknown>[] = [];
  const picList: Record<string, unknown>[] = [];
  const grpSpList: Record<string, unknown>[] = [];
  const cxnSpList: Record<string, unknown>[] = [];

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

  for (const rawEl of elements) {
    const elWithId = normalizeElementWithUniqueIds(rawEl, getUniqueId);

    if (elWithId.elementType === 'shape') {
      shapeList.push(serializeShape(elWithId));
    } else if (elWithId.elementType === 'table') {
      graphicFrameList.push(serializeTable(elWithId));
    } else if (elWithId.elementType === 'chart') {
      const chartRelId = chartRelIds ? (chartRelIds[chartIdx++] || 'rId2') : 'rId2';
      graphicFrameList.push(serializeChartGraphicFrame(elWithId, chartRelId));
    } else if (elWithId.elementType === 'picture') {
      const overrideEmbedId = pictureEmbedMap?.get(elWithId.picture.mediaId) ?? pictureEmbedMap?.get(elWithId.id);
      picList.push(serializePicture(elWithId, overrideEmbedId));
    } else if (elWithId.elementType === 'group') {
      grpSpList.push(serializeGroup(elWithId));
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
      cxnSpList.push(serializeConnector(mappedConnector));
    }
  }

  const spTree: Record<string, unknown> = {
    'p:nvGrpSpPr': {
      'p:cNvPr': { '@_id': '1', '@_name': '' },
      'p:cNvGrpSpPr': {},
      'p:nvPr': {},
    },
    'p:grpSpPr': {
      'a:xfrm': {
        'a:off': { '@_x': '0', '@_y': '0' },
        'a:ext': { '@_cx': '0', '@_cy': '0' },
        'a:chOff': { '@_x': '0', '@_y': '0' },
        'a:chExt': { '@_cx': '0', '@_cy': '0' },
      },
    },
  };

  if (shapeList.length > 0) {
    spTree['p:sp'] = shapeList;
  }
  if (graphicFrameList.length > 0) {
    spTree['p:graphicFrame'] = graphicFrameList;
  }
  if (picList.length > 0) {
    spTree['p:pic'] = picList;
  }
  if (grpSpList.length > 0) {
    spTree['p:grpSp'] = grpSpList;
  }
  if (cxnSpList.length > 0) {
    spTree['p:cxnSp'] = cxnSpList;
  }

  const cSld: Record<string, unknown> = {};
  const bg = serializeSlideBackground(slide);
  if (bg) {
    cSld['p:bg'] = bg;
  }
  cSld['p:spTree'] = spTree;

  const sldRoot: Record<string, unknown> = {
    'p:sld': {
      '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'p:cSld': cSld,
      'p:clrMapOvr': {
        'a:masterClrMapping': {},
      },
    },
  };

  const transitionNode = serializeTransition(slide.transition);
  if (transitionNode) {
    (sldRoot['p:sld'] as Record<string, unknown>)['p:transition'] = transitionNode;
  }

  const timingNode = serializeAnimations(slide.animations);
  if (timingNode) {
    (sldRoot['p:sld'] as Record<string, unknown>)['p:timing'] = timingNode;
  }

  return serializeXml(sldRoot);
}
