import type { PptxTransition, PptxTransitionDirection, PptxTransitionSpeed } from '@hokkyss/pptx-core';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

const REVERSE_DIRECTION_MAP: Record<string, PptxTransitionDirection> = {
  d: 'down',
  down: 'down',
  horz: 'horz',
  in: 'in',
  l: 'left',
  left: 'left',
  out: 'out',
  r: 'right',
  right: 'right',
  u: 'up',
  up: 'up',
  vert: 'vert',
};

const SPEED_DURATION_MAP: Record<string, number> = {
  fast: 500,
  med: 1000,
  medium: 1000,
  slow: 2000,
};

/**
 * Parses slide transition nodes (`<p:transition>`) into structured `PptxTransition` AST elements.
 *
 * Identifies transition type (`fade`, `push`, `wipe`, `split`, `wheel`, `dissolve`, etc.),
 * transition speed/duration, direction (`left`, `right`, `up`, `down`), click advance settings (`advanceOnClick`),
 * and auto-advance delay (`advanceAfterMs`).
 * @param slideXml Raw XML string or parsed object node representing a slide (`<p:sld>`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Parsed `PptxTransition` object or `undefined` if no transition node exists.
 * @example
 * ```ts
 * const transition = parseTransition(slideXmlString);
 * console.log(transition?.type, transition?.speed, transition?.direction);
 * ```
 */
export function parseTransition(
  slideXml: Record<string, unknown> | string,
  parser: XmlParser = defaultXmlParser,
): PptxTransition | undefined {
  let parsed: Record<string, unknown>;
  if (typeof slideXml === 'string') {
    parsed = parser.parse<Record<string, unknown>>(slideXml);
  } else {
    parsed = slideXml;
  }

  const sldNode = (parsed['p:sld'] || parsed['sld'] || parsed) as Record<string, unknown>;
  const transitionNode = (sldNode['p:transition'] || sldNode['transition']) as Record<string, unknown> | undefined;

  if (!transitionNode) return undefined;

  const rawSpeed = transitionNode['@_spd'] as string | undefined;
  const speed: PptxTransitionSpeed | undefined = rawSpeed ? (rawSpeed === 'medium' ? 'med' : rawSpeed) : undefined;
  const rawDur = transitionNode['@_dur'] !== undefined ? Number(transitionNode['@_dur']) : undefined;
  const durationMs = rawDur ?? (speed ? SPEED_DURATION_MAP[speed] : undefined);

  const advanceOnClick = transitionNode['@_advClick'] !== '0' && transitionNode['@_advClick'] !== false;
  const advanceAfterMs = transitionNode['@_advTm'] !== undefined ? Number(transitionNode['@_advTm']) : undefined;

  // Identify transition type tag (e.g. p:fade, p:push, p:wipe, p:split, p:wheel, p:cut, etc.)
  const typeKey = Object.keys(transitionNode).find((k) => k.startsWith('p:') || (!k.startsWith('@_') && k !== '#text'));
  const type = typeKey ? typeKey.replace('p:', '') : 'none';

  const childNode = (typeKey ? transitionNode[typeKey] : undefined) as Record<string, unknown> | undefined;

  let direction: PptxTransitionDirection | undefined;
  let throughBlack: boolean | undefined;
  let spokes: number | undefined;

  if (childNode && typeof childNode === 'object') {
    const rawDir = (childNode['@_dir'] as string | undefined) || (childNode['@_orient'] as string | undefined);
    if (rawDir) {
      direction = REVERSE_DIRECTION_MAP[rawDir] ?? (rawDir);
    }
    if (childNode['@_thruBlk'] === '1' || childNode['@_thruBlk'] === true) {
      throughBlack = true;
    }
    if (childNode['@_spokes'] !== undefined) {
      spokes = Number(childNode['@_spokes']);
    }
  }

  return {
    advanceAfterMs,
    advanceOnClick,
    direction,
    durationMs,
    speed,
    spokes,
    throughBlack,
    type,
  };
}
