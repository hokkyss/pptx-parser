import type { PptxHyperlink, PptxHyperlinkAction, RelationshipResolver } from '@hokkyss/pptx-core';

/**
 * Parses an OpenXML DrawingML `<a:hlinkClick>` node into a typed `PptxHyperlink` AST element.
 *
 * Resolves external URLs, target slide indexes, navigation actions, and tooltip screentips.
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
  const action = hlinkNode['@_action'] as string | undefined;
  const tooltip = hlinkNode['@_tooltip'] as string | undefined;
  const invalidUrl = hlinkNode['@_invalidUrl'] as string | undefined;

  let url: string | undefined = invalidUrl;
  let slideIndex: number | undefined;
  let normalizedAction: PptxHyperlinkAction | undefined;

  if (rId && relationshipResolver) {
    const rel = relationshipResolver.getRelationship(rId);
    if (rel) {
      if (
        rel.targetMode === 'External'
        || rel.target.startsWith('http://')
        || rel.target.startsWith('https://')
        || rel.target.startsWith('mailto:')
      ) {
        url = rel.target;
      } else if (rel.target.includes('slide')) {
        const match = rel.target.match(/slide(\d+)\.xml/);
        if (match) {
          slideIndex = parseInt(match[1], 10);
        }
      }
    }
  }

  if (action) {
    if (action.includes('nextslide') || action.includes('nextSlide')) {
      normalizedAction = 'nextSlide';
    } else if (
      action.includes('prevslide')
      || action.includes('prevSlide')
      || action.includes('previousslide')
      || action.includes('previousSlide')
    ) {
      normalizedAction = 'previousSlide';
    } else if (action.includes('firstslide') || action.includes('firstSlide')) {
      normalizedAction = 'firstSlide';
    } else if (action.includes('lastslide') || action.includes('lastSlide')) {
      normalizedAction = 'lastSlide';
    } else if (action.includes('endshow') || action.includes('endShow')) {
      normalizedAction = 'endShow';
    } else if (!action.includes('hlinksldjump')) {
      normalizedAction = action;
    }
  }

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
