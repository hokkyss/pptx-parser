import { PptxTransition } from '../types/ast';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

/**
 * Parses slide transition nodes (`<p:transition>`) into structured `PptxTransition` AST elements.
 *
 * Identifies transition type (`fade`, `push`, `wipe`, `split`, `morph`, `wheel`), transition speed (`fast`, `medium`, `slow`), click advance settings (`advanceOnClick`), and auto-advance delay (`advanceAfterMs`).
 * @param slideXml Raw XML string or parsed object node representing a slide (`<p:sld>`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Parsed `PptxTransition` object or `undefined` if no transition node exists.
 * @example
 * ```ts
 * const transition = parseTransition(slideXmlString);
 * console.log(transition?.type, transition?.speed, transition?.advanceOnClick);
 * ```
 */
export function parseTransition(slideXml: Record<string, unknown> | string, parser: XmlParser = defaultXmlParser): PptxTransition | undefined {
  let parsed: Record<string, unknown>;
  if (typeof slideXml === 'string') {
    parsed = parser.parse<Record<string, unknown>>(slideXml);
  } else {
    parsed = slideXml;
  }

  const sldNode = (parsed['p:sld'] || parsed['sld'] || parsed) as Record<string, unknown>;
  const transitionNode = (sldNode['p:transition'] || sldNode['transition']) as Record<string, unknown> | undefined;

  if (!transitionNode) return undefined;

  const speed = (transitionNode['@_spd'] as string) || 'medium';
  const advanceOnClick = transitionNode['@_advClick'] !== '0' && transitionNode['@_advClick'] !== false;
  const advanceAfterMs = transitionNode['@_advTm'] !== undefined ? Number(transitionNode['@_advTm']) : undefined;

  // Identify transition type tag (e.g. p:fade, p:push, p:wipe, p:morph, p:split, p:wheel)
  const typeKey = Object.keys(transitionNode).find((k) => k.startsWith('p:') || (!k.startsWith('@_') && k !== '#text'));
  const type = typeKey ? typeKey.replace('p:', '') : 'none';

  return {
    type,
    speed,
    advanceOnClick,
    advanceAfterMs,
  };
}
