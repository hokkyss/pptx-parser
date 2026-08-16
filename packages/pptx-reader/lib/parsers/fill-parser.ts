import type {
  PptxBackground,
  PptxColor,
  PptxFill,
  PptxGradientFill,
  PptxGradientStop,
  PptxGradientType,
} from '@hokkyss/pptx-core';
import { thousandthsPercent } from '@hokkyss/pptx-core';
import { getXmlChild, getXmlChildren } from '../xml/xml-parser';

/**
 * Extracts a typed `PptxColor` from an OpenXML color container node (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`).
 * @param clrParent Node containing color elements.
 * @returns Parsed `PptxColor` or `undefined`.
 */
export function parseColorNode(clrParent?: Record<string, unknown>): PptxColor | undefined {
  if (!clrParent || typeof clrParent !== 'object') return undefined;

  const srgbNode = getXmlChild(clrParent, 'srgbClr');
  if (srgbNode) {
    const val = (srgbNode['@_val'] as string) || '';
    const alphaNode = getXmlChild(srgbNode, 'alpha');
    const alpha = alphaNode?.['@_val'] !== undefined ? thousandthsPercent(Number(alphaNode['@_val'])) : undefined;
    return {
      alpha,
      type: 'srgb',
      value: val.toUpperCase(),
    };
  }

  const schemeNode = getXmlChild(clrParent, 'schemeClr');
  if (schemeNode) {
    const val = (schemeNode['@_val'] as string) || '';
    const alphaNode = getXmlChild(schemeNode, 'alpha');
    const alpha = alphaNode?.['@_val'] !== undefined ? thousandthsPercent(Number(alphaNode['@_val'])) : undefined;
    return {
      alpha,
      type: 'scheme',
      value: val,
    };
  }

  const sysNode = getXmlChild(clrParent, 'sysClr');
  if (sysNode) {
    const val = (sysNode['@_lastClr'] || sysNode['@_val'] || '') as string;
    const alphaNode = getXmlChild(sysNode, 'alpha');
    const alpha = alphaNode?.['@_val'] !== undefined ? thousandthsPercent(Number(alphaNode['@_val'])) : undefined;
    return {
      alpha,
      type: 'system',
      value: val.toUpperCase(),
    };
  }

  return undefined;
}

/**
 * Parses OpenXML DrawingML `<a:gradFill>` into a typed `PptxGradientFill`.
 * @param gradNode Raw XML object node for `<a:gradFill>`.
 * @returns Parsed `PptxGradientFill` or `undefined`.
 */
export function parseGradientFill(gradNode?: Record<string, unknown>): PptxGradientFill | undefined {
  if (!gradNode || typeof gradNode !== 'object') return undefined;

  const flip = gradNode['@_flip'] as 'none' | 'x' | 'xy' | 'y' | undefined;
  const rotWithShape = gradNode['@_rotWithShape'];
  const rotateWithShape = rotWithShape !== undefined ? (rotWithShape === '1' || rotWithShape === true || rotWithShape === 'true') : undefined;

  // 1. Parse Gradient Stops (<a:gsLst><a:gs>)
  const gsLstNode = getXmlChild(gradNode, 'gsLst');
  const gsNodes = gsLstNode ? getXmlChildren(gsLstNode, 'gs') : [];
  const stops: PptxGradientStop[] = [];

  for (const gs of gsNodes) {
    const posRaw = gs['@_pos'];
    const position = posRaw !== undefined ? Number(posRaw) : 0;
    const color = parseColorNode(gs);
    if (color) {
      const opacity = color.alpha !== undefined ? Number(color.alpha) / 100000 : undefined;
      stops.push({
        color,
        opacity,
        position,
      });
    }
  }

  // 2. Parse Direction / Type (<a:lin> or <a:path>)
  const linNode = getXmlChild(gradNode, 'lin');
  const pathNode = getXmlChild(gradNode, 'path');

  let type: PptxGradientType = 'linear';
  let angle: number | undefined;
  let pathBounds: PptxGradientFill['pathBounds'] | undefined;

  if (linNode) {
    type = 'linear';
    if (linNode['@_ang'] !== undefined) {
      angle = Number(linNode['@_ang']);
    }
  } else if (pathNode) {
    const rawPath = pathNode['@_path'] as string | undefined;
    type = rawPath === 'circle' ? 'radial' : 'path';

    const fillToRect = getXmlChild(pathNode, 'fillToRect');
    if (fillToRect) {
      pathBounds = {
        bottom: fillToRect['@_b'] !== undefined ? Number(fillToRect['@_b']) / 100000 : undefined,
        left: fillToRect['@_l'] !== undefined ? Number(fillToRect['@_l']) / 100000 : undefined,
        right: fillToRect['@_r'] !== undefined ? Number(fillToRect['@_r']) / 100000 : undefined,
        top: fillToRect['@_t'] !== undefined ? Number(fillToRect['@_t']) / 100000 : undefined,
      };
    }
  }

  return {
    angle,
    flip: flip && flip !== 'none' ? flip : undefined,
    pathBounds,
    rotateWithShape,
    stops,
    type,
  };
}

/**
 * Parses DrawingML fill properties from a container node (e.g. `<p:spPr>`, `<p:bgPr>`, `<a:ln>`).
 * @param container Raw XML object node containing fill tags.
 * @returns Typed `PptxFill` or `undefined`.
 */
export function parseFill(container?: Record<string, unknown>): PptxFill | undefined {
  if (!container || typeof container !== 'object') return undefined;

  // 1. Check for <a:noFill>
  if (getXmlChild(container, 'noFill')) {
    return { type: 'none' };
  }

  // 2. Check for <a:solidFill>
  const solidNode = getXmlChild(container, 'solidFill');
  if (solidNode) {
    const solidColor = parseColorNode(solidNode);
    if (solidColor) {
      return {
        solidColor,
        type: 'solid',
      };
    }
  }

  // 3. Check for <a:gradFill>
  const gradNode = getXmlChild(container, 'gradFill');
  if (gradNode) {
    const gradient = parseGradientFill(gradNode);
    if (gradient) {
      return {
        gradient,
        type: 'gradient',
      };
    }
  }

  return undefined;
}

/**
 * Parses slide background properties from `<p:bg>`.
 * @param bgNode Raw XML object node representing `<p:bg>`.
 * @returns Parsed `PptxBackground` or `undefined`.
 */
export function parseBackground(bgNode?: Record<string, unknown>): PptxBackground | undefined {
  if (!bgNode || typeof bgNode !== 'object') return undefined;

  const bgPr = getXmlChild(bgNode, 'bgPr') || bgNode;
  const fill = parseFill(bgPr);
  if (fill) {
    return { fill };
  }

  return undefined;
}
