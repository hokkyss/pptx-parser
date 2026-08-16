import type { PptxHyperlink, PptxHyperlinkAction, RelationshipResolver } from '@hokkyss/pptx-core';
import {
  sanitizeHyperlinkAction,
  sanitizeHyperlinkTooltip,
  sanitizeHyperlinkUrl,
  sanitizeSlideIndex,
} from '@hokkyss/pptx-core';

/**
 * Parses an OpenXML DrawingML `<a:hlinkClick>` node into a typed `PptxHyperlink` AST element.
 *
 * Resolves and sanitizes external URLs, target slide indexes, navigation actions, and tooltip screentips.
 * @param hlinkNode Raw XML object node for `<a:hlinkClick>`.
 * @param relationshipResolver Optional RelationshipResolver to map `r:id` to external/internal targets.
 * @returns Parsed `PptxHyperlink` or `undefined` if no valid hyperlink data.
 */
export function parseHyperlink(
  hlinkNode?: Record<string, unknown>,
  relationshipResolver?: RelationshipResolver,
): PptxHyperlink | undefined {
  if (!hlinkNode || typeof hlinkNode !== 'object') return undefined;

  const rId = (hlinkNode['@_r:id'] || hlinkNode['@_id']) as string | undefined;
  const rawAction = hlinkNode['@_action'] as string | undefined;
  const rawTooltip = hlinkNode['@_tooltip'] as string | undefined;
  const rawInvalidUrl = hlinkNode['@_invalidUrl'] as string | undefined;

  let url: string | undefined = sanitizeHyperlinkUrl(rawInvalidUrl);
  let slideIndex: number | undefined;
  let normalizedAction: PptxHyperlinkAction | undefined;

  if (rId && relationshipResolver) {
    const rel = relationshipResolver.getRelationship(rId);
    if (rel && rel.target) {
      if (
        rel.targetMode === 'External'
        || rel.target.startsWith('http://')
        || rel.target.startsWith('https://')
        || rel.target.startsWith('mailto:')
      ) {
        url = sanitizeHyperlinkUrl(rel.target);
      } else {
        const match = rel.target.match(/(?:^|\/)slide(\d+)\.xml$/);
        if (match) {
          slideIndex = sanitizeSlideIndex(match[1]);
        }
      }
    }
  }

  if (rawAction) {
    if (rawAction.includes('nextslide') || rawAction.includes('nextSlide')) {
      normalizedAction = 'nextSlide';
    } else if (
      rawAction.includes('prevslide')
      || rawAction.includes('prevSlide')
      || rawAction.includes('previousslide')
      || rawAction.includes('previousSlide')
    ) {
      normalizedAction = 'previousSlide';
    } else if (rawAction.includes('firstslide') || rawAction.includes('firstSlide')) {
      normalizedAction = 'firstSlide';
    } else if (rawAction.includes('lastslide') || rawAction.includes('lastSlide')) {
      normalizedAction = 'lastSlide';
    } else if (rawAction.includes('endshow') || rawAction.includes('endShow')) {
      normalizedAction = 'endShow';
    } else if (!rawAction.includes('hlinksldjump')) {
      normalizedAction = sanitizeHyperlinkAction(rawAction);
    }
  }

  const tooltip = sanitizeHyperlinkTooltip(rawTooltip);

  const result: PptxHyperlink = {
    action: normalizedAction,
    rId,
    slideIndex,
    tooltip,
    url,
  };

  const hasAnyField = Boolean(
    result.action
    || result.rId
    || result.slideIndex !== undefined
    || result.tooltip
    || result.url,
  );
  return hasAnyField ? result : undefined;
}
