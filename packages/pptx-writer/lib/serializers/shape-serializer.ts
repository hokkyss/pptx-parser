import type {
  PptxConnectionPosition,
  PptxConnectorElement,
  PptxGeometry,
  PptxLine,
  PptxLineEnd,
  PptxShadow,
  PptxShapeElement,
  PptxShapeLocks,
} from '@hokkyss/pptx-core';
import { el, type XmlElement } from '../xml/xml-element';
import { serializeFill, serializeHyperlink, serializeTextBody } from './text-serializer';

/**
 * Serializes shape locks `<a:spLocks>`.
 */
export function serializeShapeLocks(locks?: PptxShapeLocks): undefined | XmlElement {
  if (!locks) return undefined;

  const attrs: Record<string, number | string | undefined> = {};
  if (locks.noAdjustHandles) attrs.noAdjustHandles = '1';
  if (locks.noChangeAspect) attrs.noChangeAspect = '1';
  if (locks.noChangeShapeType) attrs.noChangeShapeType = '1';
  if (locks.noCrop) attrs.noCrop = '1';
  if (locks.noEditPoints) attrs.noEditPoints = '1';
  if (locks.noGrp) attrs.noGrp = '1';
  if (locks.noMove) attrs.noMove = '1';
  if (locks.noResize) attrs.noResize = '1';
  if (locks.noRot) attrs.noRot = '1';
  if (locks.noSelect) attrs.noSelect = '1';
  if (locks.noUngrp) attrs.noUngrp = '1';

  return Object.keys(attrs).length > 0 ? el('a:spLocks', attrs) : undefined;
}

/**
 * Serializes line end arrowhead properties `<a:headEnd>` or `<a:tailEnd>`.
 */
export function serializeLineEnd(lineEnd?: PptxLineEnd, tag: 'a:headEnd' | 'a:tailEnd' = 'a:headEnd'): undefined | XmlElement {
  if (!lineEnd) return undefined;
  const attrs: Record<string, string | undefined> = {};
  if (lineEnd.type !== undefined) attrs.type = lineEnd.type;
  if (lineEnd.width !== undefined) attrs.w = lineEnd.width;
  if (lineEnd.length !== undefined) attrs.len = lineEnd.length;
  return Object.keys(attrs).length > 0 ? el(tag, attrs) : undefined;
}

/**
 * Serializes line/outline properties `<a:ln>`.
 */
export function serializeLine(line?: PptxLine): undefined | XmlElement {
  if (!line) return undefined;

  const attrs: Record<string, number | undefined> = {};
  if (line.width !== undefined) {
    attrs.w = Math.round(Number(line.width));
  }

  const children: (undefined | XmlElement)[] = [];

  if (line.fill) {
    const fillNode = serializeFill(line.fill);
    if (fillNode) children.push(fillNode);
  }
  if (line.dashStyle) {
    children.push(el('a:prstDash', { val: line.dashStyle }));
  }
  if (line.headEnd) {
    const head = serializeLineEnd(line.headEnd, 'a:headEnd');
    if (head) children.push(head);
  }
  if (line.tailEnd) {
    const tail = serializeLineEnd(line.tailEnd, 'a:tailEnd');
    if (tail) children.push(tail);
  }

  return el('a:ln', attrs, children.filter(Boolean));
}

/**
 * Serializes geometry `<a:prstGeom>` or `<a:custGeom>`.
 */
export function serializeGeometry(geometry?: PptxGeometry): XmlElement {
  if (!geometry || !geometry.presetGeometry) {
    return el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]);
  }

  const avLstChildren: XmlElement[] = [];
  if (geometry.adjustments) {
    for (const [name, val] of Object.entries(geometry.adjustments)) {
      avLstChildren.push(el('a:gd', { fmla: `val ${val}`, name }));
    }
  }

  return el('a:prstGeom', { prst: geometry.presetGeometry }, [
    el('a:avLst', avLstChildren),
  ]);
}

/**
 * Serializes outer shadow effects `<a:effectLst><a:outerShdw>`.
 */
export function serializeShadow(shadow?: PptxShadow): undefined | XmlElement {
  if (!shadow) return undefined;

  const attrs: Record<string, number | string | undefined> = {};
  if (shadow.blurRadius !== undefined) attrs.blurRad = Math.round(Number(shadow.blurRadius));
  if (shadow.distance !== undefined) attrs.dist = Math.round(Number(shadow.distance));
  if (shadow.direction !== undefined) attrs.dir = Math.round(Number(shadow.direction));
  if (shadow.alignment) attrs.algn = shadow.alignment;
  if (shadow.rotateWithShape !== undefined) attrs.rotWithShape = shadow.rotateWithShape ? '1' : '0';

  const clrChildren: XmlElement[] = [];
  if (shadow.opacity !== undefined && shadow.opacity < 1) {
    clrChildren.push(el('a:alpha', { val: Math.round(shadow.opacity * 100000) }));
  }
  const srgbClrNode = el('a:srgbClr', { val: (shadow.color || '000000').replace(/^#/, '') }, clrChildren);
  const outerShdw = el('a:outerShdw', attrs, [srgbClrNode]);

  return el('a:effectLst', [outerShdw]);
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

/**
 * Serializes a shape element into OpenXML `<p:sp>` strictly following DrawingML schema sequence.
 */
export function serializeShape(shape: PptxShapeElement): XmlElement {
  const cNvPrAttrs: Record<string, string> = {
    id: shape.id || '2',
    name: shape.name || `Shape ${shape.id || '2'}`,
  };
  if (shape.isVisible === false) {
    cNvPrAttrs.hidden = '1';
  }
  const cNvPrChildren: XmlElement[] = [];
  if (shape.hyperlink) {
    const hlinkNode = serializeHyperlink(shape.hyperlink);
    if (hlinkNode) {
      cNvPrChildren.push(hlinkNode);
    }
  }
  const cNvPr = el('p:cNvPr', cNvPrAttrs, cNvPrChildren);

  const cNvSpPrAttrs: Record<string, string> = {};
  if (shape.isTextBox) {
    cNvSpPrAttrs.txBox = '1';
  }
  const cNvSpPrChildren: XmlElement[] = [];
  if (shape.locks) {
    const locks = serializeShapeLocks(shape.locks);
    if (locks) cNvSpPrChildren.push(locks);
  }
  const cNvSpPr = el('p:cNvSpPr', cNvSpPrAttrs, cNvSpPrChildren);

  const nvPrChildren: XmlElement[] = [];
  if (shape.placeholder) {
    const phAttrs: Record<string, number | string> = {
      type: shape.placeholder.type,
    };
    if (shape.placeholder.idx !== undefined) {
      phAttrs.idx = shape.placeholder.idx;
    }
    nvPrChildren.push(el('p:ph', phAttrs));
  }
  const nvPr = el('p:nvPr', nvPrChildren);

  const nvSpPr = el('p:nvSpPr', [cNvPr, cNvSpPr, nvPr]);

  const spPrChildren: XmlElement[] = [];

  const hasExplicitSize = shape.position && (Number(shape.position.cx) > 0 || Number(shape.position.cy) > 0);
  if (hasExplicitSize) {
    const xfrmAttrs: Record<string, number | undefined> = {};
    if (shape.rotation) {
      xfrmAttrs.rot = Math.round(Number(shape.rotation));
    }
    const xfrm = el('a:xfrm', xfrmAttrs, [
      el('a:off', {
        x: Math.round(Number(shape.position?.x ?? 0)),
        y: Math.round(Number(shape.position?.y ?? 0)),
      }),
      el('a:ext', {
        cx: Math.round(Number(shape.position?.cx ?? 1000000)),
        cy: Math.round(Number(shape.position?.cy ?? 1000000)),
      }),
    ]);
    spPrChildren.push(xfrm);
  } else if (!shape.placeholder) {
    const xfrm = el('a:xfrm', [
      el('a:off', { x: Math.round(Number(shape.position?.x ?? 0)), y: Math.round(Number(shape.position?.y ?? 0)) }),
      el('a:ext', { cx: 2000000, cy: 1000000 }),
    ]);
    spPrChildren.push(xfrm);
  }

  // Geometry
  if (shape.geometry) {
    spPrChildren.push(serializeGeometry(shape.geometry));
  } else if (!hasExplicitSize && shape.placeholder) {
    // Inherits geometry from layout
  } else {
    const rawType = shape.shapeType || 'rect';
    const mappedType = PRESET_GEOMETRY_MAP[rawType] || rawType;
    spPrChildren.push(el('a:prstGeom', { prst: mappedType }, [el('a:avLst')]));
  }

  // Fill
  if (shape.fill) {
    const fillNode = serializeFill(shape.fill);
    if (fillNode) spPrChildren.push(fillNode);
  }

  // Line
  if (shape.line) {
    const lnNode = serializeLine(shape.line);
    if (lnNode) spPrChildren.push(lnNode);
  }

  // Effects (Shadows)
  if (shape.shadow) {
    const effectNode = serializeShadow(shape.shadow);
    if (effectNode) spPrChildren.push(effectNode);
  }

  const spPr = el('p:spPr', spPrChildren);

  // Text Body (Strictly required for p:sp in PresentationML)
  const txBody = shape.textBody
    ? serializeTextBody(shape.textBody, 'p:txBody')
    : el('p:txBody', [
        el('a:bodyPr'),
        el('a:lstStyle'),
        el('a:p', [el('a:endParaRPr')]),
      ]);

  return el('p:sp', [nvSpPr, spPr, txBody]);
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
export function serializeConnector(connector: PptxConnectorElement): XmlElement {
  const cNvPrAttrs: Record<string, string> = {
    id: connector.id || '2',
    name: connector.name || `Connector ${connector.id || '2'}`,
  };
  if (connector.isVisible === false) {
    cNvPrAttrs.hidden = '1';
  }
  const cNvPrChildren: XmlElement[] = [];
  if (connector.hyperlink) {
    const hlinkNode = serializeHyperlink(connector.hyperlink);
    if (hlinkNode) {
      cNvPrChildren.push(hlinkNode);
    }
  }
  const cNvPr = el('p:cNvPr', cNvPrAttrs, cNvPrChildren);

  const cNvCxnSpPrChildren: XmlElement[] = [el('a:cxnSpLocks')];
  if (connector.startConnection) {
    cNvCxnSpPrChildren.push(
      el('a:stCxn', {
        id: connector.startConnection.shapeId,
        idx: POSITION_TO_INDEX_MAP[connector.startConnection.position] ?? 0,
      }),
    );
  }
  if (connector.endConnection) {
    cNvCxnSpPrChildren.push(
      el('a:endCxn', {
        id: connector.endConnection.shapeId,
        idx: POSITION_TO_INDEX_MAP[connector.endConnection.position] ?? 0,
      }),
    );
  }
  const cNvCxnSpPr = el('p:cNvCxnSpPr', cNvCxnSpPrChildren);
  const nvCxnSpPr = el('p:nvCxnSpPr', [cNvPr, cNvCxnSpPr, el('p:nvPr')]);

  const xfrm = el('a:xfrm', [
    el('a:off', {
      x: Math.round(Number(connector.position?.x ?? 0)),
      y: Math.round(Number(connector.position?.y ?? 0)),
    }),
    el('a:ext', {
      cx: Math.round(Number(connector.position?.cx ?? 100000)),
      cy: Math.round(Number(connector.position?.cy ?? 0)),
    }),
  ]);

  const geom = el('a:prstGeom', {
    prst: connector.shapeType || 'line',
  }, [el('a:avLst')]);

  const spPrChildren: XmlElement[] = [xfrm, geom];

  if (connector.line) {
    const ln = serializeLine(connector.line);
    if (ln) spPrChildren.push(ln);
  }

  const spPr = el('p:spPr', spPrChildren);

  return el('p:cxnSp', [nvCxnSpPr, spPr]);
}
