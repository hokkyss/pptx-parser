import { PptxAnimation } from '../types/ast';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

/**
 * Parses slide animation timeline nodes (`<p:timing>`) into structured `PptxAnimation` AST elements.
 *
 * Recursively traverses OpenXML animation tree nodes (`<p:tnLst>`, `<p:par>`, `<p:childTnLst>`, `<p:animEffect>`) to extract target shape IDs, effect types, trigger conditions, and sequence indices.
 * @param slideXml Raw XML string or parsed object node representing a slide (`<p:sld>`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Array of parsed `PptxAnimation` timeline items.
 * @example
 * ```ts
 * const animations = parseAnimations(slideXmlString);
 * console.log(`Found ${animations.length} slide animations.`);
 * ```
 */
export function parseAnimations(slideXml: Record<string, unknown> | string, parser: XmlParser = defaultXmlParser): PptxAnimation[] {
  let parsed: Record<string, unknown>;
  if (typeof slideXml === 'string') {
    parsed = parser.parse<Record<string, unknown>>(slideXml);
  } else {
    parsed = slideXml;
  }

  const sldNode = (parsed['p:sld'] || parsed['sld'] || parsed) as Record<string, unknown>;
  const timingNode = (sldNode['p:timing'] || sldNode['timing']) as Record<string, unknown> | undefined;

  if (!timingNode) return [];

  const animations: PptxAnimation[] = [];
  traverseTimingTree(timingNode, animations, 0);
  return animations;
}

/**
 * Helper method performing recursive traversal over the OpenXML timing tree to discover target shape IDs (`<p:spTgt>`) and animation behaviors.
 * @param node Current timing XML object node.
 * @param results Accumulator array for parsed `PptxAnimation` items.
 * @param sequenceIndex Current zero-based animation sequence counter.
 */
function traverseTimingTree(node: Record<string, unknown>, results: PptxAnimation[], sequenceIndex: number): void {
  if (!node || typeof node !== 'object') return;

  // Check if current node defines a target shape target (<p:spTgt spid="..."/>)
  const tgtEl = (node['p:tgtEl'] || node['tgtEl']) as Record<string, unknown> | undefined;
  if (tgtEl) {
    const spTgt = (tgtEl['p:spTgt'] || tgtEl['spTgt']) as Record<string, unknown> | undefined;
    if (spTgt && spTgt['@_spid']) {
      const targetShapeId = String(spTgt['@_spid']);
      const effect = (node['@_presetClass'] as string) || 'appear';
      results.push({
        targetShapeId,
        effect,
        effectType: effect,
        trigger: (node['@_nodeType'] as string) || 'onClick',
        sequence: sequenceIndex++,
      });
    }
  }

  // Recurse into all object keys / array children
  for (const value of Object.values(node)) {
    if (Array.isArray(value)) {
      for (const item of value) {
        if (typeof item === 'object' && item !== null) {
          traverseTimingTree(item as Record<string, unknown>, results, sequenceIndex);
        }
      }
    } else if (typeof value === 'object' && value !== null) {
      traverseTimingTree(value as Record<string, unknown>, results, sequenceIndex);
    }
  }
}
