import type {
  PptxConnectionPosition,
  PptxConnectorElement,
  PptxGeometry,
  PptxLine,
  PptxLineEnd,
  PptxShapeElement,
  PptxShapeLocks,
} from '@hokkyss/pptx-core';
import { _rawXmlWrapBuilder, serializeFill, serializeHyperlink, serializeTextBody } from './text-serializer';

/**
 * Serializes shape locks `<a:spLocks>`.
 */
export function serializeShapeLocks(locks?: PptxShapeLocks): Record<string, unknown> | undefined {
  if (!locks) return undefined;

  const spLocks: Record<string, unknown> = {};
  if (locks.noAdjustHandles) spLocks['@_noAdjustHandles'] = '1';
  if (locks.noChangeAspect) spLocks['@_noChangeAspect'] = '1';
  if (locks.noChangeShapeType) spLocks['@_noChangeShapeType'] = '1';
  if (locks.noCrop) spLocks['@_noCrop'] = '1';
  if (locks.noEditPoints) spLocks['@_noEditPoints'] = '1';
  if (locks.noGrp) spLocks['@_noGrp'] = '1';
  if (locks.noMove) spLocks['@_noMove'] = '1';
  if (locks.noResize) spLocks['@_noResize'] = '1';
  if (locks.noRot) spLocks['@_noRot'] = '1';
  if (locks.noSelect) spLocks['@_noSelect'] = '1';
  if (locks.noUngrp) spLocks['@_noUngrp'] = '1';

  return Object.keys(spLocks).length > 0 ? spLocks : undefined;
}

/**
 * Serializes line end arrowhead properties `<a:headEnd>` or `<a:tailEnd>`.
 */
export function serializeLineEnd(lineEnd?: PptxLineEnd): Record<string, unknown> | undefined {
  if (!lineEnd) return undefined;
  const node: Record<string, unknown> = {};
  if (lineEnd.type !== undefined) node['@_type'] = lineEnd.type;
  if (lineEnd.width !== undefined) node['@_w'] = lineEnd.width;
  if (lineEnd.length !== undefined) node['@_len'] = lineEnd.length;
  return Object.keys(node).length > 0 ? node : undefined;
}

/**
 * Serializes line/outline properties `<a:ln>`.
 */
export function serializeLine(line?: PptxLine): Record<string, unknown> | undefined {
  if (!line) return undefined;

  const ln: Record<string, unknown> = {};
  if (line.width !== undefined) {
    ln['@_w'] = Math.round(Number(line.width));
  }
  if (line.fill) {
    const fillNode = serializeFill(line.fill);
    if (fillNode) {
      Object.assign(ln, fillNode);
    }
  }
  if (line.dashStyle) {
    ln['a:prstDash'] = { '@_val': line.dashStyle };
  }
  if (line.headEnd) {
    const head = serializeLineEnd(line.headEnd);
    if (head) ln['a:headEnd'] = head;
  }
  if (line.tailEnd) {
    const tail = serializeLineEnd(line.tailEnd);
    if (tail) ln['a:tailEnd'] = tail;
  }

  return ln;
}

/**
 * Serializes geometry `<a:prstGeom>` or `<a:custGeom>`.
 */
export function serializeGeometry(geometry?: PptxGeometry): Record<string, unknown> {
  if (!geometry) {
    return { 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} } };
  }

  if (geometry.presetGeometry) {
    const avLst: Record<string, unknown> = {};
    if (geometry.adjustments) {
      const gdList = Object.entries(geometry.adjustments).map(([name, val]) => ({
        '@_fmla': `val ${val}`,
        '@_name': name,
      }));
      avLst['a:gd'] = gdList;
    }

    return {
      'a:prstGeom': {
        '@_prst': geometry.presetGeometry,
        'a:avLst': avLst,
      },
    };
  }

  return { 'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} } };
}

/**
 * Serializes outer shadow effects `<a:effectLst><a:outerShdw>`.
 */
export function serializeShadow(shadow?: import('@hokkyss/pptx-core').PptxShadow): Record<string, unknown> | undefined {
  if (!shadow) return undefined;

  const outerShdw: Record<string, unknown> = {};
  if (shadow.blurRadius !== undefined) {
    outerShdw['@_blurRad'] = Math.round(Number(shadow.blurRadius));
  }
  if (shadow.distance !== undefined) {
    outerShdw['@_dist'] = Math.round(Number(shadow.distance));
  }
  if (shadow.direction !== undefined) {
    outerShdw['@_dir'] = Math.round(Number(shadow.direction));
  }
  if (shadow.alignment) {
    outerShdw['@_algn'] = shadow.alignment;
  }
  if (shadow.rotateWithShape !== undefined) {
    outerShdw['@_rotWithShape'] = shadow.rotateWithShape ? '1' : '0';
  }

  const srgbClr: Record<string, unknown> = {
    '@_val': (shadow.color || '000000').replace(/^#/, ''),
  };
  if (shadow.opacity !== undefined && shadow.opacity < 1) {
    srgbClr['a:alpha'] = {
      '@_val': Math.round(shadow.opacity * 100000),
    };
  }
  outerShdw['a:srgbClr'] = srgbClr;

  return {
    'a:effectLst': {
      'a:outerShdw': outerShdw,
    },
  };
}

/**
 * Serializes a shape element into OpenXML `<p:sp>` strictly following DrawingML schema sequence.
 */
export function serializeShape(shape: PptxShapeElement): Record<string, unknown> | string {
  const cNvPr: Record<string, unknown> = {
    '@_id': shape.id,
    '@_name': shape.name || '',
  };
  if (shape.isVisible === false) {
    cNvPr['@_hidden'] = '1';
  }
  if (shape.hyperlink) {
    const hlinkNode = serializeHyperlink(shape.hyperlink);
    if (hlinkNode) {
      cNvPr['a:hlinkClick'] = hlinkNode;
    }
  }

  const cNvSpPr: Record<string, unknown> = {};
  if (shape.isTextBox) {
    cNvSpPr['@_txBox'] = '1';
  }
  if (shape.locks) {
    const locks = serializeShapeLocks(shape.locks);
    if (locks) cNvSpPr['a:spLocks'] = locks;
  }

  const nvPr: Record<string, unknown> = {};
  if (shape.placeholder) {
    const ph: Record<string, unknown> = {};
    if (shape.placeholder.type) {
      ph['@_type'] = shape.placeholder.type;
    }
    if (shape.placeholder.idx !== undefined) {
      ph['@_idx'] = shape.placeholder.idx;
    }
    nvPr['p:ph'] = ph;
  }

  const nvSpPr = {
    'p:cNvPr': cNvPr,
    'p:cNvSpPr': cNvSpPr,
    'p:nvPr': nvPr,
  };

  const spPr: Record<string, unknown> = {};

  const hasExplicitSize = shape.position && (Number(shape.position.cx) > 0 || Number(shape.position.cy) > 0);
  if (hasExplicitSize) {
    const xfrm: Record<string, unknown> = {
      'a:off': {
        '@_x': Math.round(Number(shape.position?.x ?? 0)),
        '@_y': Math.round(Number(shape.position?.y ?? 0)),
      },
      'a:ext': {
        '@_cx': Math.round(Number(shape.position?.cx ?? 1000000)),
        '@_cy': Math.round(Number(shape.position?.cy ?? 1000000)),
      },
    };
    if (shape.rotation) {
      xfrm['@_rot'] = Math.round(Number(shape.rotation));
    }
    spPr['a:xfrm'] = xfrm;
  } else if (!shape.placeholder) {
    spPr['a:xfrm'] = {
      'a:off': { '@_x': Math.round(Number(shape.position?.x ?? 0)), '@_y': Math.round(Number(shape.position?.y ?? 0)) },
      'a:ext': { '@_cx': 2000000, '@_cy': 1000000 },
    };
    spPr['a:prstGeom'] = { '@_prst': 'rect', 'a:avLst': {} };
  }

  const PRESET_GEOMETRY_MAP: Record<string, string> = {
    box: 'rect',
    circle: 'ellipse',
    cylinder: 'can',
    oval: 'ellipse',
    square: 'rect',
    star: 'star5',
    wedgeRoundRect: 'wedgeRoundRectCallout',
  };

  // Geometry
  if (shape.geometry) {
    const geomNode = serializeGeometry(shape.geometry);
    Object.assign(spPr, geomNode);
  } else if (!hasExplicitSize && shape.placeholder) {
    // Inherits geometry from layout
  } else {
    const rawType = shape.shapeType || 'rect';
    const mappedType = PRESET_GEOMETRY_MAP[rawType] || rawType;
    spPr['a:prstGeom'] = { '@_prst': mappedType, 'a:avLst': {} };
  }

  // Fill
  if (shape.fill) {
    const fillNode = serializeFill(shape.fill);
    if (fillNode) Object.assign(spPr, fillNode);
  }

  // Line
  if (shape.line) {
    const lnNode = serializeLine(shape.line);
    if (lnNode) spPr['a:ln'] = lnNode;
  }

  // Effects (Shadows)
  if (shape.shadow) {
    const effectNode = serializeShadow(shape.shadow);
    if (effectNode) Object.assign(spPr, effectNode);
  }

  const txBody = shape.textBody
    ? serializeTextBody(shape.textBody)
    : {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': { 'a:endParaRPr': {} },
      };

  if (typeof txBody === 'string') {
    // Raw-XML path: the text body contains line breaks.
    // Build the entire shape as a raw XML string using _rawXmlWrapBuilder which won't
    // double-encode the pre-built XML strings.
    return _rawXmlWrapBuilder.build({
      'p:sp': {
        'p:nvSpPr': nvSpPr,
        'p:spPr': spPr,
        'p:txBody': txBody,
      },
    });
  }

  return {
    'p:nvSpPr': nvSpPr,
    'p:spPr': spPr,
    'p:txBody': txBody,
  };
}

const POSITION_TO_INDEX_MAP: Record<PptxConnectionPosition, number> = {
  top: 0,
  left: 1,
  bottom: 2,
  right: 3,
};

/**
 * Serializes a connector element into OpenXML `<p:cxnSp>`.
 */
export function serializeConnector(connector: PptxConnectorElement): Record<string, unknown> {
  const cNvPr: Record<string, unknown> = {
    '@_id': connector.id,
    '@_name': connector.name || '',
  };
  if (connector.isVisible === false) {
    cNvPr['@_hidden'] = '1';
  }
  if (connector.hyperlink) {
    const hlinkNode = serializeHyperlink(connector.hyperlink);
    if (hlinkNode) {
      cNvPr['a:hlinkClick'] = hlinkNode;
    }
  }

  const cNvCxnSpPr: Record<string, unknown> = {
    'a:cxnSpLocks': {},
  };
  if (connector.startConnection) {
    cNvCxnSpPr['a:stCxn'] = {
      '@_id': connector.startConnection.shapeId,
      '@_idx': POSITION_TO_INDEX_MAP[connector.startConnection.position] ?? 0,
    };
  }
  if (connector.endConnection) {
    cNvCxnSpPr['a:endCxn'] = {
      '@_id': connector.endConnection.shapeId,
      '@_idx': POSITION_TO_INDEX_MAP[connector.endConnection.position] ?? 0,
    };
  }

  const nvCxnSpPr = {
    'p:cNvPr': cNvPr,
    'p:cNvCxnSpPr': cNvCxnSpPr,
    'p:nvPr': {},
  };

  const spPr: Record<string, unknown> = {
    'a:xfrm': {
      'a:off': {
        '@_x': Math.round(Number(connector.position?.x ?? 0)),
        '@_y': Math.round(Number(connector.position?.y ?? 0)),
      },
      'a:ext': {
        '@_cx': Math.round(Number(connector.position?.cx ?? 100000)),
        '@_cy': Math.round(Number(connector.position?.cy ?? 0)),
      },
    },
    'a:prstGeom': {
      '@_prst': connector.shapeType || 'line',
      'a:avLst': {},
    },
  };

  if (connector.line) {
    const ln = serializeLine(connector.line);
    if (ln) spPr['a:ln'] = ln;
  }

  return {
    'p:nvCxnSpPr': nvCxnSpPr,
    'p:spPr': spPr,
  };
}
