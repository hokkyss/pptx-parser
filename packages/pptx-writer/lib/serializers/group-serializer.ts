import type { PptxGroupElement } from '@hokkyss/pptx-core';
import { serializePicture } from './picture-serializer';
import { serializeConnector, serializeShape } from './shape-serializer';
import { serializeTable } from './table-serializer';

/**
 * Serializes a group shape `<p:grpSp>` containing nested children strictly conforming to ECMA-376 schema.
 */
export function serializeGroup(group: PptxGroupElement): Record<string, unknown> {
  const shapeList: Record<string, unknown>[] = [];
  const graphicFrameList: Record<string, unknown>[] = [];
  const picList: Record<string, unknown>[] = [];
  const grpSpList: Record<string, unknown>[] = [];
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

  const grpSp: Record<string, unknown> = {
    'p:nvGrpSpPr': {
      'p:cNvPr': {
        '@_id': group.id || '5',
        '@_name': group.name || `Group ${group.id || '5'}`,
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
