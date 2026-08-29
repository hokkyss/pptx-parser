import type { RelationshipResolver } from '@hokkyss/pptx-core';
import {
  Emu,
  EmuDegree,
  PptxConnectionPosition,
  PptxLine,
  PptxLineEnd,
  PptxLineEndLength,
  PptxLineEndType,
  PptxLineEndWidth,
  PptxPlaceholder,
  PptxShape,
  PptxShapeAttachment,
  PptxShapeElement,
} from '../types/ast';
import { defaultXmlParser, getXmlChild, getXmlChildren, XmlParser } from '../xml/xml-parser';
import { parseFill } from './fill-parser';
import { parseHyperlink } from './hyperlink-parser';
import { parseTextBody } from './text-parser';

export { getXmlChild, getXmlChildren };

/**
 * Parses all visual elements/shapes from a slide XML string or parsed shape tree node (`<p:spTree>`).
 *
 * Extracts shapes, text boxes, pictures, table graphicFrames, chart graphicFrames, groups, and connectors.
 * @param slideXml Raw slide XML string or parsed XML object node.
 * @param parser Optional custom `XmlParser` instance.
 * @returns Array of parsed `PptxShape` elements with zIndex assigned in rendering order.
 * @example
 * ```ts
 * const shapes = parseShapes(slideXmlString);
 * console.log(`Parsed ${shapes.length} shapes.`);
 * ```
 */
export function parseShapes(
  slideXml: Record<string, unknown> | string,
  parser?: XmlParser,
): PptxShape[];
export function parseShapes(
  slideXml: Record<string, unknown> | string,
  relationshipResolver?: RelationshipResolver,
  parser?: XmlParser,
): PptxShape[];
export function parseShapes(
  slideXml: Record<string, unknown> | string,
  resolverOrParser?: RelationshipResolver | XmlParser,
  parser?: XmlParser,
): PptxShape[] {
  let resolver: RelationshipResolver | undefined;
  let activeParser = defaultXmlParser;

  if (resolverOrParser) {
    if ('parse' in resolverOrParser && typeof resolverOrParser.parse === 'function') {
      activeParser = resolverOrParser;
    } else {
      resolver = resolverOrParser as RelationshipResolver;
      if (parser) {
        activeParser = parser;
      }
    }
  }

  let parsed: Record<string, unknown>;
  if (typeof slideXml === 'string') {
    parsed = activeParser.parse<Record<string, unknown>>(slideXml);
  } else {
    parsed = slideXml;
  }

  const rootNode = getXmlChild(parsed, 'sld') || getXmlChild(parsed, 'sldLayout') || getXmlChild(parsed, 'sldMaster') || parsed;
  const cSldNode = getXmlChild(rootNode, 'cSld') || rootNode;
  const spTree = getXmlChild(cSldNode, 'spTree') || cSldNode;

  if (!spTree) return [];

  return parseShapeTree(spTree, resolver);
}

/**
 * Traverses a shape tree XML node (`<p:spTree>`), parsing shapes, pictures, graphic frames, groups, and connectors.
 *
 * Automatically computes and assigns 0-based `zIndex` values representing back-to-front rendering order.
 * @param spTree Raw XML object node representing shape tree (`<p:spTree>`).
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Array of parsed `PptxShape` elements.
 */
export function parseShapeTree(
  spTree: Record<string, unknown>,
  relationshipResolver?: RelationshipResolver,
): PptxShape[] {
  const shapes: PptxShape[] = [];

  // Parse Auto Shapes (<p:sp>)
  for (const spNode of getXmlChildren(spTree, 'sp')) {
    shapes.push(parseSingleShape(spNode, 'shape', relationshipResolver));
  }

  // Parse Pictures (<p:pic>)
  for (const picNode of getXmlChildren(spTree, 'pic')) {
    shapes.push(parseSingleShape(picNode, 'picture', relationshipResolver));
  }

  // Parse Graphic Frames (<p:graphicFrame>)
  for (const gfNode of getXmlChildren(spTree, 'graphicFrame')) {
    shapes.push(parseSingleShape(gfNode, 'graphicFrame', relationshipResolver));
  }

  // Parse Group Shapes (<p:grpSp>)
  for (const grpNode of getXmlChildren(spTree, 'grpSp')) {
    shapes.push(parseSingleShape(grpNode, 'group', relationshipResolver));
  }

  // Parse Connection Shapes (<p:cxnSp>)
  for (const cxnNode of getXmlChildren(spTree, 'cxnSp')) {
    shapes.push(parseSingleShape(cxnNode, 'connector', relationshipResolver));
  }

  // Assign 0-based zIndex to reflect rendering layer order
  for (let i = 0; i < shapes.length; i++) {
    shapes[i].zIndex = i;
  }

  return shapes;
}

/** Helper to parse line end arrowhead properties from <a:headEnd> or <a:tailEnd> */
function parseLineEnd(lineEndNode?: Record<string, unknown>): PptxLineEnd | undefined {
  if (!lineEndNode) return undefined;
  const type = lineEndNode['@_type'] as PptxLineEndType | undefined;
  const width = lineEndNode['@_w'] as PptxLineEndWidth | undefined;
  const length = lineEndNode['@_len'] as PptxLineEndLength | undefined;
  if (!type && !width && !length) return undefined;
  return {
    ...(type && { type }),
    ...(width && { width }),
    ...(length && { length }),
  };
}

const INDEX_TO_POSITION_MAP: Record<number, PptxConnectionPosition> = {
  0: 'top',
  1: 'left',
  2: 'bottom',
  3: 'right',
};

/** Helper to parse connector attachment points from <a:stCxn> or <a:endCxn> */
function parseConnectionPoint(cxnNode?: Record<string, unknown>): PptxShapeAttachment | undefined {
  if (!cxnNode) return undefined;
  const rawId = cxnNode['@_id'];
  if (rawId === undefined || rawId === null) return undefined;
  const shapeId = String(rawId);
  const rawIdx = Number(cxnNode['@_idx'] ?? 0);
  const position = INDEX_TO_POSITION_MAP[rawIdx] || 'top';
  return { position, shapeId };
}

/**
 * Parses a single OpenXML shape node (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`, `<p:cxnSp>`).
 * @param node Raw XML object node.
 * @param shapeType OpenXML shape tag classification.
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Parsed `PptxShape` AST structure.
 */
export function parseSingleShape(
  node: Record<string, unknown>,
  shapeType: 'connector' | 'graphicFrame' | 'group' | 'picture' | 'shape',
  relationshipResolver?: RelationshipResolver,
): PptxShape {
  // Non-visual shape properties node (nvSpPr, nvPicPr, nvGraphicFramePr, nvGrpSpPr, nvCxnSpPr)
  const nvPrNode
    = getXmlChild(node, 'nvSpPr')
      || getXmlChild(node, 'nvPicPr')
      || getXmlChild(node, 'nvGraphicFramePr')
      || getXmlChild(node, 'nvGrpSpPr')
      || getXmlChild(node, 'nvCxnSpPr')
      || {};

  const cNvPrNode = getXmlChild(nvPrNode, 'cNvPr') || {};
  const id = (cNvPrNode['@_id'] as string) || '0';
  const name = (cNvPrNode['@_name'] as string) || '';
  const hiddenAttr = cNvPrNode['@_hidden'];
  const isHidden = hiddenAttr === '1' || hiddenAttr === 'true' || hiddenAttr === true ? true : undefined;
  const isVisible = isHidden ? false : true;

  const hlinkNode = getXmlChild(cNvPrNode, 'hlinkClick');
  const hyperlink = hlinkNode ? parseHyperlink(hlinkNode, relationshipResolver) : undefined;

  // Extract shape locks (<a:spLocks>, <a:picLocks>, <a:grpSpLocks>, etc.)
  const cNvChild
    = getXmlChild(nvPrNode, 'cNvSpPr')
      || getXmlChild(nvPrNode, 'cNvPicPr')
      || getXmlChild(nvPrNode, 'cNvGrpSpPr')
      || getXmlChild(nvPrNode, 'cNvCxnSpPr')
      || getXmlChild(nvPrNode, 'cNvGraphicFramePr')
      || {};

  const locksNode
    = getXmlChild(cNvChild, 'spLocks')
      || getXmlChild(cNvChild, 'picLocks')
      || getXmlChild(cNvChild, 'grpSpLocks')
      || getXmlChild(cNvChild, 'cxnSpLocks')
      || getXmlChild(cNvChild, 'graphicFrameLocks');

  let locks: import('../types/ast').PptxShapeLocks | undefined;
  let isLocked: boolean | undefined;

  if (locksNode) {
    const parseLockAttr = (key: string): boolean | undefined => {
      const val = locksNode[key];
      return val === '1' || val === 'true' || val === true ? true : undefined;
    };

    const parsedLocks: import('../types/ast').PptxShapeLocks = {
      noGrp: parseLockAttr('@_noGrp'),
      noSelect: parseLockAttr('@_noSelect'),
      noRot: parseLockAttr('@_noRot'),
      noChangeAspect: parseLockAttr('@_noChangeAspect'),
      noMove: parseLockAttr('@_noMove'),
      noResize: parseLockAttr('@_noResize'),
      noEditPoints: parseLockAttr('@_noEditPoints'),
      noAdjustHandles: parseLockAttr('@_noAdjustHandles'),
      noChangeShapeType: parseLockAttr('@_noChangeShapeType'),
      noCrop: parseLockAttr('@_noCrop'),
      noUngrp: parseLockAttr('@_noUngrp'),
    };

    const cleanLocks = Object.fromEntries(Object.entries(parsedLocks).filter(([, v]) => v !== undefined));
    if (Object.keys(cleanLocks).length > 0) {
      locks = cleanLocks;
      isLocked = true;
    }
  }

  // Extract placeholder info (<p:ph>)
  const nvPrChild = getXmlChild(nvPrNode, 'nvPr') || {};
  const phNode = getXmlChild(nvPrChild, 'ph') || getXmlChild(nvPrNode, 'ph');
  let placeholder: PptxPlaceholder | undefined;
  if (phNode) {
    placeholder = {
      idx: phNode['@_idx'] !== undefined ? String(phNode['@_idx']) : undefined,
      type: (phNode['@_type'] as string) || 'body',
    };
  }

  // Check if text box (<p:cNvSpPr txBox="1"/>)
  const cNvSpPrNode = getXmlChild(nvPrNode, 'cNvSpPr') || {};
  const isTextBox = cNvSpPrNode['@_txBox'] === '1' || cNvSpPrNode['@_txBox'] === true || cNvSpPrNode['@_txBox'] === 'true';

  // Extract transform (position & size)
  const spPrNode = getXmlChild(node, 'spPr') || getXmlChild(node, 'grpSpPr') || {};
  const xfrmNode = getXmlChild(spPrNode, 'xfrm') || getXmlChild(node, 'xfrm') || {};

  const offNode = getXmlChild(xfrmNode, 'off') || {};
  const extNode = getXmlChild(xfrmNode, 'ext') || {};

  const x = (offNode['@_x'] !== undefined ? Number(offNode['@_x']) : 0) as Emu;
  const y = (offNode['@_y'] !== undefined ? Number(offNode['@_y']) : 0) as Emu;
  const cx = (extNode['@_cx'] !== undefined ? Number(extNode['@_cx']) : 0) as Emu;
  const cy = (extNode['@_cy'] !== undefined ? Number(extNode['@_cy']) : 0) as Emu;

  const rotation = (xfrmNode['@_rot'] !== undefined ? Number(xfrmNode['@_rot']) : 0) as EmuDegree;

  // Extract text body
  const txBodyNode = getXmlChild(node, 'txBody');
  let textBody = txBodyNode ? parseTextBody(txBodyNode, relationshipResolver) : undefined;
  if (textBody && textBody.paragraphs.length === 0) {
    textBody = undefined;
  }

  // Extract picture blip embed rId
  let picture: { mediaId: string } | undefined;
  let blipEmbedId: string | undefined;
  if (shapeType === 'picture') {
    const blipFill = getXmlChild(node, 'blipFill') || {};
    const blip = getXmlChild(blipFill, 'blip') || {};
    blipEmbedId = (blip['@_r:embed'] as string) || (blip['@_embed'] as string) || undefined;
    picture = {
      mediaId: blipEmbedId || '',
    };
  }

  // Parse children if group shape
  let children: PptxShape[] | undefined;
  if (shapeType === 'group') {
    children = parseShapeTree(node, relationshipResolver);
  }

  // Extract fill properties
  const fill = parseFill(spPrNode);

  // Extract preset geometry
  const prstGeomNode = getXmlChild(spPrNode, 'prstGeom');
  const shapePreset = prstGeomNode ? (prstGeomNode['@_prst'] as string) : undefined;

  // Extract outline / line properties (<a:ln>)
  const lnNode = getXmlChild(spPrNode, 'ln');
  let line: PptxLine | undefined;
  if (lnNode) {
    const w = lnNode['@_w'] !== undefined ? (Number(lnNode['@_w']) as Emu) : undefined;
    const lineFill = parseFill(lnNode);
    const prstDash = getXmlChild(lnNode, 'prstDash');
    const dashStyle = prstDash ? (prstDash['@_val'] as string) : undefined;
    const headEnd = parseLineEnd(getXmlChild(lnNode, 'headEnd'));
    const tailEnd = parseLineEnd(getXmlChild(lnNode, 'tailEnd'));
    line = {
      dashStyle,
      fill: lineFill,
      width: w,
      ...(headEnd && { headEnd }),
      ...(tailEnd && { tailEnd }),
    };
  }

  const baseResult = {
    id,
    name,
    type: shapeType,
    position: { cx, cy, x, y },
    rotation,
    zIndex: 0,
    isVisible,
    ...(fill !== undefined && { fill }),
    ...(line !== undefined && { line }),
    ...(shapePreset !== undefined && { shapeType: shapePreset }),
    ...(hyperlink !== undefined && { hyperlink }),
    ...(isLocked !== undefined && { isLocked }),
    ...(locks !== undefined && { locks }),
    layerSource: 'slide' as const,
    placeholder,
  };

  if (shapeType === 'picture') {
    const innerNvPr = getXmlChild(nvPrNode, 'nvPr') || {};
    const audioFileNode = getXmlChild(innerNvPr, 'audioFile');
    const videoFileNode = getXmlChild(innerNvPr, 'videoFile');

    if (audioFileNode) {
      const audioRelId = (audioFileNode['@_r:link'] as string) || (audioFileNode['@_link'] as string);
      const resolvedTarget = audioRelId ? relationshipResolver?.getTarget(audioRelId) : undefined;
      const mediaId = resolvedTarget || audioRelId || blipEmbedId || '';
      const ext = mediaId.split('.').pop()?.toLowerCase() || '';
      const mimeMap: Record<string, string> = {
        m4a: 'audio/mp4',
        mp3: 'audio/mpeg',
        ogg: 'audio/ogg',
        wav: 'audio/x-wav',
        wma: 'audio/x-ms-wma',
      };
      const mimeType = mimeMap[ext] || 'audio/mpeg';
      const posterTarget = blipEmbedId ? relationshipResolver?.getTarget(blipEmbedId) : undefined;
      const posterImageId = posterTarget || blipEmbedId || undefined;

      return {
        ...baseResult,
        audio: {
          mediaId,
          mimeType,
          ...(posterImageId !== undefined ? { posterImageId } : {}),
        },
        blipEmbedId,
        elementType: 'audio',
        type: 'picture',
      };
    }

    if (videoFileNode) {
      const videoRelId = (videoFileNode['@_r:link'] as string) || (videoFileNode['@_link'] as string);
      const resolvedTarget = videoRelId ? relationshipResolver?.getTarget(videoRelId) : undefined;
      const mediaId = resolvedTarget || videoRelId || blipEmbedId || '';
      const ext = mediaId.split('.').pop()?.toLowerCase() || '';
      const mimeMap: Record<string, string> = {
        avi: 'video/x-msvideo',
        mov: 'video/quicktime',
        mp4: 'video/mp4',
        webm: 'video/webm',
        wmv: 'video/x-ms-wmv',
      };
      const mimeType = mimeMap[ext] || 'video/mp4';
      const posterTarget = blipEmbedId ? relationshipResolver?.getTarget(blipEmbedId) : undefined;
      const posterImageId = posterTarget || blipEmbedId || undefined;

      return {
        ...baseResult,
        blipEmbedId,
        elementType: 'video',
        type: 'picture',
        video: {
          mediaId,
          mimeType,
          ...(posterImageId !== undefined ? { posterImageId } : {}),
        },
      };
    }

    return {
      ...baseResult,
      blipEmbedId,
      elementType: 'picture',
      picture: picture || { mediaId: blipEmbedId || '' },
      type: 'picture',
    };
  }

  if (shapeType === 'group') {
    return {
      ...baseResult,
      type: 'group',
      elementType: 'group',
      children: children || [],
    };
  }

  if (shapeType === 'connector') {
    const cNvCxnSpPrNode = getXmlChild(nvPrNode, 'cNvCxnSpPr') || cNvChild;
    const stCxnNode = getXmlChild(cNvCxnSpPrNode, 'stCxn');
    const endCxnNode = getXmlChild(cNvCxnSpPrNode, 'endCxn');
    const startConnection = parseConnectionPoint(stCxnNode);
    const endConnection = parseConnectionPoint(endCxnNode);

    return {
      ...baseResult,
      elementType: 'connector',
      textBody,
      type: 'connector',
      ...(startConnection && { startConnection }),
      ...(endConnection && { endConnection }),
    };
  }

  if (shapeType === 'graphicFrame') {
    const graphicNode = getXmlChild(node, 'graphic') || (node['a:graphic'] as Record<string, unknown>);
    const graphicDataNode = graphicNode ? (getXmlChild(graphicNode, 'graphicData') || (graphicNode['a:graphicData'] as Record<string, unknown>)) : undefined;
    const graphicUri = (graphicDataNode?.['@_uri'] as string) || '';
    const chartNode = graphicDataNode ? (getXmlChild(graphicDataNode, 'chart') || (graphicDataNode['c:chart'] as Record<string, unknown>)) : undefined;
    const chartRelId = (chartNode?.['@_r:id'] || chartNode?.['@_id'] || chartNode?.['r:id']) as string | undefined;
    const tblNode = graphicDataNode ? (getXmlChild(graphicDataNode, 'tbl') || (graphicDataNode['a:tbl'] as Record<string, unknown>)) : undefined;

    return {
      ...baseResult,
      type: 'graphicFrame',
      elementType: 'shape', // Temporary fallback, refined by PptxParser to 'table' or 'chart'
      textBody,
      _chartRelId: chartRelId,
      _graphicUri: graphicUri,
      _tblNode: tblNode,
    } as PptxShapeElement;
  }

  return {
    ...baseResult,
    type: 'shape',
    elementType: 'shape',
    textBody,
    isTextBox: isTextBox || undefined,
  };
}
