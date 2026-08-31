import type { PptxGroupElement } from '@hokkyss/pptx-core';
import { createXmlBuilder } from '../xml/xml-builder';
import { serializePicture } from './picture-serializer';
import { serializeConnector, serializeShape } from './shape-serializer';
import { serializeTable } from './table-serializer';

/**
 * Serializes a group shape `<p:grpSp>` containing nested children strictly conforming to ECMA-376 schema.
 */
export function serializeGroup(group: PptxGroupElement): Record<string, unknown> | string {
  const shapeList: (Record<string, unknown> | string)[] = [];
  const graphicFrameList: Record<string, unknown>[] = [];
  const picList: Record<string, unknown>[] = [];
  const grpSpList: (Record<string, unknown> | string)[] = [];
  const cxnSpList: Record<string, unknown>[] = [];

  for (const child of group.children || []) {
    if (child.elementType === 'shape') {
      shapeList.push(serializeShape(child));
    } else if (child.elementType === 'table') {
      graphicFrameList.push(serializeTable(child));
    } else if (child.elementType === 'picture') {
      picList.push(serializePicture(child));
    } else if (child.elementType === 'group') {
      grpSpList.push(serializeGroup(child));
    } else if (child.elementType === 'connector') {
      cxnSpList.push(serializeConnector(child));
    }
  }

  const hasRawChildren = shapeList.some((s) => typeof s === 'string') || grpSpList.some((g) => typeof g === 'string');

  if (hasRawChildren) {
    const elementBuilder = createXmlBuilder();
    const nvGrpPr = elementBuilder.build({
      'p:nvGrpSpPr': {
        'p:cNvPr': {
          '@_id': group.id,
          '@_name': group.name || '',
        },
        'p:cNvGrpSpPr': {},
        'p:nvPr': {},
      },
    });
    const grpPr = elementBuilder.build({
      'p:grpSpPr': {
        'a:xfrm': {
          'a:off': {
            '@_x': Math.round(Number(group.position?.x ?? 0)),
            '@_y': Math.round(Number(group.position?.y ?? 0)),
          },
          'a:ext': {
            '@_cx': Math.round(Number(group.position?.cx ?? 1000000)),
            '@_cy': Math.round(Number(group.position?.cy ?? 1000000)),
          },
          'a:chOff': {
            '@_x': Math.round(Number(group.position?.x ?? 0)),
            '@_y': Math.round(Number(group.position?.y ?? 0)),
          },
          'a:chExt': {
            '@_cx': Math.round(Number(group.position?.cx ?? 1000000)),
            '@_cy': Math.round(Number(group.position?.cy ?? 1000000)),
          },
        },
      },
    });

    const shapesXml = shapeList.map((s) => (typeof s === 'string' ? s : (elementBuilder.build({ 'p:sp': s })))).join('');
    const tablesXml = graphicFrameList.map((gf) => elementBuilder.build({ 'p:graphicFrame': gf })).join('');
    const picsXml = picList.map((pic) => elementBuilder.build({ 'p:pic': pic })).join('');
    const grpSpXml = grpSpList.map((gs) => (typeof gs === 'string' ? gs : (elementBuilder.build({ 'p:grpSp': gs })))).join('');
    const cxnSpXml = cxnSpList.map((cs) => elementBuilder.build({ 'p:cxnSp': cs })).join('');

    return `<p:grpSp>${nvGrpPr}${grpPr}${shapesXml}${tablesXml}${picsXml}${grpSpXml}${cxnSpXml}</p:grpSp>`;
  }

  const grpSp: Record<string, unknown> = {
    'p:nvGrpSpPr': {
      'p:cNvPr': {
        '@_id': group.id,
        '@_name': group.name || '',
      },
      'p:cNvGrpSpPr': {},
      'p:nvPr': {},
    },
    'p:grpSpPr': {
      'a:xfrm': {
        'a:off': {
          '@_x': Math.round(Number(group.position?.x ?? 0)),
          '@_y': Math.round(Number(group.position?.y ?? 0)),
        },
        'a:ext': {
          '@_cx': Math.round(Number(group.position?.cx ?? 1000000)),
          '@_cy': Math.round(Number(group.position?.cy ?? 1000000)),
        },
        'a:chOff': {
          '@_x': Math.round(Number(group.position?.x ?? 0)),
          '@_y': Math.round(Number(group.position?.y ?? 0)),
        },
        'a:chExt': {
          '@_cx': Math.round(Number(group.position?.cx ?? 1000000)),
          '@_cy': Math.round(Number(group.position?.cy ?? 1000000)),
        },
      },
    },
  };

  if (shapeList.length > 0) grpSp['p:sp'] = shapeList;
  if (graphicFrameList.length > 0) grpSp['p:graphicFrame'] = graphicFrameList;
  if (picList.length > 0) grpSp['p:pic'] = picList;
  if (grpSpList.length > 0) grpSp['p:grpSp'] = grpSpList;
  if (cxnSpList.length > 0) grpSp['p:cxnSp'] = cxnSpList;

  return grpSp;
}
