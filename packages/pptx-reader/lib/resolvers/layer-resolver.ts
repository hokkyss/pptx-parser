import { PptxDocument, PptxElement, PptxShape } from '../types/ast';

/**
 * Resolved 3-tier layer structure containing elements from Master, Layout, and Slide rendering layers.
 */
export interface PptxResolvedSlideLayers {
  /** All elements composed in back-to-front rendering order (Master -> Layout -> Slide) */
  allElementsInRenderOrder: PptxElement[];
  /** All shapes composed in back-to-front rendering order (alias for allElementsInRenderOrder) */
  allShapesInRenderOrder: PptxShape[];
  /** Layout background elements (Middle layer) */
  layoutElements: PptxElement[];
  /** Layout background shapes (alias) */
  layoutShapes: PptxShape[];
  /** Master background elements (Bottom layer) */
  masterElements: PptxElement[];
  /** Master background shapes (alias) */
  masterShapes: PptxShape[];
  /** Slide content elements (Top layer) */
  slideElements: PptxElement[];
  /** Slide content shapes (alias) */
  slideShapes: PptxShape[];
}

/**
 * Resolves and composes the 3-tier OpenXML rendering layer stack (Master -> Layout -> Slide) for a target slide.
 *
 * PowerPoint renders slide visual elements in 3 stacked layers:
 * 1. **Bottom Layer (`masterElements`)**: Slide Master background graphics and theme branding.
 * 2. **Middle Layer (`layoutElements`)**: Slide Layout template frames and background placeholders.
 * 3. **Top Layer (`slideElements`)**: Slide-specific shapes, text boxes, tables, charts, and media.
 * @param doc The parsed `PptxDocument` instance.
 * @param slideNumberOrId 1-based slide number (e.g. `1`) or internal slide ID string (e.g. `'rId2'`).
 * @returns Resolved layer structure containing masterElements, layoutElements, slideElements, and allElementsInRenderOrder, or `undefined` if slide is not found.
 * @example
 * ```ts
 * import { parsePptx, resolveSlideLayers } from '@hokkyss/pptx-reader';
 *
 * const doc = await parsePptx(buffer);
 * const layers = resolveSlideLayers(doc, 1);
 *
 * if (layers) {
 *   // Render all elements in exact back-to-front depth order
 *   for (const element of layers.allElementsInRenderOrder) {
 *     console.log(element.zIndex, element.layerSource, element.name);
 *   }
 * }
 * ```
 */
export function resolveSlideLayers(
  doc: PptxDocument,
  slideNumberOrId: number | string,
): PptxResolvedSlideLayers | undefined {
  const slide
    = typeof slideNumberOrId === 'number'
      ? doc.slides[slideNumberOrId - 1]
      : doc.slides.find((s) => s.slideId === slideNumberOrId);

  if (!slide) return undefined;

  // 1. Resolve Slide Layout
  const layout = doc.slideLayouts.find((l) => l.id === slide.layoutId);
  // 2. Resolve Slide Master
  const master = layout
    ? doc.slideMasters.find((m) => m.id === layout.masterId)
    : doc.slideMasters[0];

  const masterElements = (master?.elements || master?.shapes || []).map((s) => ({
    ...s,
    layerSource: 'master' as const,
  }));

  const layoutElements = (layout?.elements || layout?.shapes || []).map((s) => ({
    ...s,
    layerSource: 'layout' as const,
  }));

  const slideElements = (slide.elements || slide.shapes || []).map((s) => ({
    ...s,
    layerSource: 'slide' as const,
  }));

  const allElementsInRenderOrder = [
    ...masterElements,
    ...layoutElements,
    ...slideElements,
  ];

  return {
    allElementsInRenderOrder,
    allShapesInRenderOrder: allElementsInRenderOrder,
    layoutElements,
    layoutShapes: layoutElements,
    masterElements,
    masterShapes: masterElements,
    slideElements,
    slideShapes: slideElements,
  };
}
