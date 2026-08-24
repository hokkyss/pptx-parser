import type {
  Emu,
  PptxChartElement,
  PptxCustomXmlPart,
  PptxDocument,
  PptxMetadata,
  PptxShape,
  PptxSlide,
  PptxTableElement,
  PptxTheme,
} from '@hokkyss/pptx-core';
import type { PptxMediaAsset } from '@hokkyss/pptx-core';
import type { XmlParser } from '@hokkyss/pptx-core';
import type { ZipReader } from '@hokkyss/pptx-core';
import { parseAnimations } from './parsers/animation-parser';
import { parseChart } from './parsers/chart-parser';
import { parseBackground } from './parsers/fill-parser';
import { parseShapes } from './parsers/shape-parser';
import { parseTable } from './parsers/table-parser';
import { parseTextBody } from './parsers/text-parser';
import { parseTransition } from './parsers/transition-parser';
import { createMasterLayoutResolver } from './resolvers/master-layout-resolver';
import { extractMedia } from './resolvers/media-resolver';
import { createRelationshipResolver } from './resolvers/relationship-resolver';
import { createThemeResolver } from './resolvers/theme-resolver';
import { PptxParseOptions } from './types/options';
import { defaultXmlParser } from './xml/xml-parser';
import { createZipReader } from './zip/zip-reader';

/**
 * Helper function to parse document core properties (`docProps/core.xml`) into a `PptxMetadata` structure.
 * @param zipReader Active ZipReader instance for reading package files.
 * @param xmlParser Active XmlParser instance.
 * @param slideWidth Presentation width in EMU.
 * @param slideHeight Presentation height in EMU.
 * @returns Parsed metadata containing title, creator, revision, and dimensions.
 */
function parseMetadata(
  zipReader: ZipReader,
  xmlParser: XmlParser,
  slideWidth: Emu,
  slideHeight: Emu,
): PptxMetadata {
  const coreXml = zipReader.getFileText('docProps/core.xml');
  const defaultMeta: PptxMetadata = {
    slideCount: 0,
    slideHeight,
    slideWidth,
  };

  if (!coreXml) return defaultMeta;

  const parsed = xmlParser.parse<Record<string, unknown>>(coreXml);
  const coreNode = (parsed['cp:coreProperties'] || parsed['coreProperties'] || {}) as Record<string, unknown>;

  const title = coreNode['dc:title'] ? String(coreNode['dc:title']) : undefined;
  const creator = coreNode['dc:creator'] ? String(coreNode['dc:creator']) : undefined;
  const lastModifiedBy = coreNode['cp:lastModifiedBy'] ? String(coreNode['cp:lastModifiedBy']) : undefined;
  const revision = coreNode['cp:revision'] !== undefined ? Number(coreNode['cp:revision']) : undefined;

  return {
    creator,
    lastModifiedBy,
    revision,
    slideCount: 0,
    slideHeight,
    slideWidth,
    title,
  };
}

/**
 * Parses speaker notes text from a notesSlide XML document.
 */
export function parseSpeakerNotes(notesXml: string, xmlParser: XmlParser): { notes?: string; notesBody?: import('@hokkyss/pptx-core').PptxTextBody } | undefined {
  try {
    const parsed = xmlParser.parse<Record<string, unknown>>(notesXml);
    const notesRoot = (parsed['p:notes'] || parsed.notes) as Record<string, unknown> | undefined;
    if (!notesRoot) return undefined;

    const cSld = (notesRoot['p:cSld'] || notesRoot.cSld) as Record<string, unknown> | undefined;
    if (!cSld) return undefined;

    const spTree = (cSld['p:spTree'] || cSld.spTree) as Record<string, unknown> | undefined;
    if (!spTree) return undefined;

    const rawSps = spTree['p:sp'] || spTree.sp;
    const sps = Array.isArray(rawSps) ? rawSps : (rawSps ? [rawSps] : []);

    for (const spNode of sps) {
      const sp = spNode as Record<string, unknown>;
      const nvSpPr = (sp['p:nvSpPr'] || sp.nvSpPr) as Record<string, unknown> | undefined;
      const nvPr = nvSpPr ? ((nvSpPr['p:nvPr'] || nvSpPr.nvPr) as Record<string, unknown> | undefined) : undefined;
      const ph = nvPr ? ((nvPr['p:ph'] || nvPr.ph) as Record<string, unknown> | undefined) : undefined;
      const phType = ph ? String(ph['@_type'] || ph.type || '') : undefined;

      // Skip slide image placeholder
      if (phType === 'sldImg') continue;

      const txBody = (sp['p:txBody'] || sp.txBody) as Record<string, unknown> | undefined;
      if (!txBody) continue;

      const notesBody = parseTextBody(txBody);
      const textLines: string[] = [];
      for (const p of notesBody.paragraphs) {
        const lineText = p.runs.map((r: { text?: string }) => r.text || '').join('');
        textLines.push(lineText);
      }
      const notes = textLines.join('\n').trim();

      return { notes: notes.length > 0 ? notes : undefined, notesBody };
    }

    return undefined;
  } catch {
    return undefined;
  }
}

/**
 * Helper function to refine generic shapes and graphicFrames into specific discriminated union variants (`table`, `chart`, `picture`, `group`).
 * @param shape Target shape element to inspect and mutate.
 * @param slideXml Raw XML string of the slide.
 * @param chartRels Array of resolved chart relationships for the slide.
 * @param zipReader Active ZipReader instance.
 * @param xmlParser Active XmlParser instance.
 * @param parseChartFn Optional chart parser function.
 * @param parseTableFn Optional table parser function.
 */
function refineElementType(
  shape: PptxShape,
  slideXml: string,
  chartRels: Array<{ id?: string; resolvedTarget: string }>,
  zipReader: ZipReader,
  xmlParser: XmlParser,
  parseChartFn?: (xml: string, parser?: XmlParser) => import('@hokkyss/pptx-core').PptxChart | undefined,
  parseTableFn?: (xml: string, parser?: XmlParser) => import('@hokkyss/pptx-core').PptxTable | undefined,
): void {
  if (shape.type === 'picture') {
    shape.elementType = 'picture';
    return;
  }

  if (shape.type === 'graphicFrame') {
    const internalShape = shape as unknown as {
      _chartRelId?: string;
      _graphicUri?: string;
      _tblNode?: Record<string, unknown>;
    };

    // 1. Table graphicFrame
    if (internalShape._tblNode || internalShape._graphicUri?.includes('table')) {
      if (parseTableFn) {
        const table = internalShape._tblNode
          ? parseTableFn(internalShape._tblNode as unknown as string, xmlParser)
          : parseTableFn(slideXml, xmlParser);
        if (table) {
          const tableElem = (shape as unknown) as PptxTableElement;
          tableElem.table = table;
          tableElem.elementType = 'table';
          delete (tableElem as unknown as Record<string, unknown>)._tblNode;
          delete (tableElem as unknown as Record<string, unknown>)._graphicUri;
          delete (tableElem as unknown as Record<string, unknown>)._chartRelId;
          return;
        }
      }
    }

    // 2. Chart graphicFrame
    if (internalShape._chartRelId || internalShape._graphicUri?.includes('chart') || chartRels.length > 0) {
      if (parseChartFn) {
        const targetRel = internalShape._chartRelId
          ? chartRels.find((r) => r.id === internalShape._chartRelId) || chartRels[0]
          : chartRels[0];

        if (targetRel) {
          const chartPath = targetRel.resolvedTarget;
          const chartXml = zipReader.getFileText(chartPath);
          if (chartXml) {
            const chart = parseChartFn(chartXml, xmlParser);
            if (chart) {
              const chartElem = (shape as unknown) as PptxChartElement;
              chartElem.chart = chart;
              chartElem.elementType = 'chart';
              delete (chartElem as unknown as Record<string, unknown>)._tblNode;
              delete (chartElem as unknown as Record<string, unknown>)._graphicUri;
              delete (chartElem as unknown as Record<string, unknown>)._chartRelId;
              return;
            }
          }
        }
      }
    }

    if (parseTableFn) {
      const table = parseTableFn(slideXml, xmlParser);
      if (table && table.rows.length > 0) {
        const tableElem = (shape as unknown) as PptxTableElement;
        tableElem.table = table;
        tableElem.elementType = 'table';
        return;
      }
    }
  }

  if (shape.type === 'group') {
    shape.elementType = 'group';
    for (const child of shape.children) {
      refineElementType(child, slideXml, chartRels, zipReader, xmlParser, parseChartFn, parseTableFn);
    }
    return;
  }

  if (shape.type === 'connector') {
    shape.elementType = 'connector';
    return;
  }

  shape.elementType = 'shape';
}

/**
 * Functional convenience entrypoint to parse a PPTX binary buffer into a complete `PptxDocument` AST.
 *
 * Employs dynamic imports for optional sub-parsers (media, animations, transitions, charts, tables)
 * to guarantee optimal bundle tree-shaking when specific capabilities are disabled.
 * @param input Binary PPTX data as a `Uint8Array` or `ArrayBuffer`.
 * @param options Optional parse settings (media extraction, lazy getters, animations, etc.).
 * @returns Promise resolving to the parsed `PptxDocument` AST object.
 * @throws {Error} If binary data is invalid or missing `ppt/presentation.xml`.
 * @example
 * ```ts
 * import { parsePptx } from '@hokkyss/pptx-reader';
 *
 * const doc = await parsePptx(fileBuffer, { includeMedia: true });
 * console.log(`Presentation title: ${doc.metadata.title}`);
 * console.log(`Found ${doc.slides.length} slides.`);
 * ```
 */
export async function parsePptx(
  input: ArrayBuffer | Uint8Array,
  options: PptxParseOptions = {},
): Promise<PptxDocument> {
  const opts: Required<PptxParseOptions> = {
    customXml: false,
    includeMedia: true,
    lazyMedia: false,
    parseAnimations: true,
    parseTransitions: true,
    ...options,
  };

  const xmlParser = defaultXmlParser;
  const zipReader = await createZipReader(input);

  // 1. Parse Presentation XML & Dimensions
  const presXml = zipReader.getFileText('ppt/presentation.xml');
  if (!presXml) {
    throw new Error('Invalid PPTX package: missing ppt/presentation.xml');
  }

  const parsedPres = xmlParser.parse<Record<string, unknown>>(presXml);
  const presNode = (parsedPres['p:presentation'] || parsedPres['presentation'] || {}) as Record<string, unknown>;

  const sldSz = (presNode['p:sldSz'] || presNode['sldSz'] || {}) as Record<string, unknown>;
  const slideWidth = (sldSz['@_cx'] !== undefined ? Number(sldSz['@_cx']) : 9144000) as Emu;
  const slideHeight = (sldSz['@_cy'] !== undefined ? Number(sldSz['@_cy']) : 6858000) as Emu;

  const firstSlideNumAttr = presNode['@_firstSlideNum'];
  const firstSlideNumber = firstSlideNumAttr !== undefined ? Number(firstSlideNumAttr) : undefined;

  // 2. Parse Core Properties (Metadata)
  const metadata = parseMetadata(zipReader, xmlParser, slideWidth, slideHeight);
  if (firstSlideNumber !== undefined) {
    metadata.firstSlideNumber = firstSlideNumber;
  }

  // 3. Resolve Presentation Relationships
  const presRelsXml = zipReader.getFileText('ppt/_rels/presentation.xml.rels');
  const presRels = createRelationshipResolver(presRelsXml, 'ppt/presentation.xml', xmlParser);

  // 4. Extract Themes (ensure primary theme from presentation rels is first)
  const themeRels = presRels.getRelationshipsByType('theme');
  const themeResolver = createThemeResolver(xmlParser);
  const themes: PptxTheme[] = [];
  const processedThemePaths = new Set<string>();

  if (themeRels.length > 0) {
    const primaryPath = themeRels[0].resolvedTarget;
    const primaryXml = zipReader.getFileText(primaryPath);
    if (primaryXml) {
      const parsedTheme = themeResolver.parseTheme(primaryXml);
      parsedTheme.id = primaryPath.split('/').pop()?.replace('.xml', '') || 'theme1';
      parsedTheme.rawXml = primaryXml;
      themes.push(parsedTheme);
      processedThemePaths.add(primaryPath);
    }
  }

  const allThemePaths = zipReader.getPathsStartingWith('ppt/theme/');
  for (const themePath of allThemePaths) {
    if (!processedThemePaths.has(themePath)) {
      const themeXml = zipReader.getFileText(themePath);
      if (themeXml) {
        const parsedTheme = themeResolver.parseTheme(themeXml);
        parsedTheme.id = themePath.split('/').pop()?.replace('.xml', '') || 'theme';
        parsedTheme.rawXml = themeXml;
        themes.push(parsedTheme);
        processedThemePaths.add(themePath);
      }
    }
  }

  // 5. Extract Media Assets
  let media: PptxMediaAsset[] = [];
  if (opts.includeMedia) {
    media = extractMedia(zipReader, Boolean(opts.lazyMedia));
  }

  // 6. Extract Slide Masters and Layouts
  const masterLayoutResolver = createMasterLayoutResolver(zipReader, xmlParser);
  const masterPaths = zipReader.getPathsStartingWith('ppt/slideMasters/').filter((p) => p.endsWith('.xml') && !p.includes('_rels'));
  const layoutPaths = zipReader.getPathsStartingWith('ppt/slideLayouts/').filter((p) => p.endsWith('.xml') && !p.includes('_rels'));

  const slideMasters = masterLayoutResolver.parseMasters(masterPaths).map((m) => {
    const masterPath = masterPaths.find((p) => p.endsWith(`${m.id}.xml`));
    const masterXml = masterPath ? zipReader.getFileText(masterPath) || '' : '';
    const masterRelsPath = masterPath ? masterPath.replace('slideMasters/', 'slideMasters/_rels/').concat('.rels') : '';
    const masterRelsXml = masterRelsPath ? zipReader.getFileText(masterRelsPath) || '' : '';
    for (const shape of m.shapes) {
      refineElementType(shape, masterXml, [], zipReader, xmlParser, parseChart, parseTable);
    }
    return {
      ...m,
      elements: m.shapes,
      rawXml: masterXml,
      relsXml: masterRelsXml,
    };
  });

  const slideLayouts = masterLayoutResolver.parseLayouts(layoutPaths).map((l) => {
    const layoutPath = layoutPaths.find((p) => p.endsWith(`${l.id}.xml`));
    const layoutXml = layoutPath ? zipReader.getFileText(layoutPath) || '' : '';
    const layoutRelsPath = layoutPath ? layoutPath.replace('slideLayouts/', 'slideLayouts/_rels/').concat('.rels') : '';
    const layoutRelsXml = layoutRelsPath ? zipReader.getFileText(layoutRelsPath) || '' : '';
    for (const shape of l.shapes) {
      refineElementType(shape, layoutXml, [], zipReader, xmlParser, parseChart, parseTable);
    }
    return {
      ...l,
      elements: l.shapes,
      rawXml: layoutXml,
      relsXml: layoutRelsXml,
    };
  });

  // 7. Animation & Transition parsers
  const parseAnimationsFn = opts.parseAnimations ? parseAnimations : undefined;
  const parseTransitionFn = opts.parseTransitions ? parseTransition : undefined;

  // 8. Parse Slide List in Presentation Order (<p:sldIdLst>)
  const sldIdLstNode = (presNode['p:sldIdLst'] || presNode['sldIdLst'] || {}) as Record<string, unknown>;
  const rawSldIds = sldIdLstNode['p:sldId'] || sldIdLstNode['sldId'];
  const sldIdArray = Array.isArray(rawSldIds) ? rawSldIds : (rawSldIds ? [rawSldIds] : []);
  const orderedSlideRelIds = sldIdArray.map((item: Record<string, unknown>) => item['@_r:id'] || item['r:id']).filter(Boolean);

  let slideRels: Array<{ id: string; resolvedTarget: string; type: string }> = [];
  if (orderedSlideRelIds.length > 0) {
    for (const rId of orderedSlideRelIds) {
      const rel = presRels.getRelationship(String(rId));
      if (rel) {
        slideRels.push(rel);
      }
    }
  } else {
    slideRels = presRels.getRelationshipsByType('slide');
  }

  const slides: PptxSlide[] = [];

  for (let i = 0; i < slideRels.length; i++) {
    const rel = slideRels[i];
    const slidePath = rel.resolvedTarget;
    const slideXml = zipReader.getFileText(slidePath);
    if (!slideXml) continue;

    const slideNumber = i + 1;
    const slideId = rel.id;

    // Parse Slide .rels for layout, chart, and media references
    const slideRelsPath = slidePath.replace('slides/', 'slides/_rels/').concat('.rels');
    const slideRelsXml = zipReader.getFileText(slideRelsPath);
    const slideRelResolver = createRelationshipResolver(slideRelsXml, slidePath, xmlParser);

    const layoutRel = slideRelResolver.getRelationshipsByType('slideLayout')[0];
    const layoutId = layoutRel ? layoutRel.resolvedTarget.split('/').pop()?.replace('.xml', '') : undefined;

    // Parse Shapes
    const shapes = parseShapes(slideXml, slideRelResolver, xmlParser);

    // Check for GraphicFrames containing tables or charts & assign high-level elementType
    const chartRels = slideRelResolver.getRelationshipsByType('chart');

    for (const shape of shapes) {
      refineElementType(shape, slideXml, chartRels, zipReader, xmlParser, parseChart, parseTable);
    }

    // Parse Animations & Transitions
    const animations = parseAnimationsFn ? parseAnimationsFn(slideXml, xmlParser) : [];
    const transition = parseTransitionFn ? parseTransitionFn(slideXml, xmlParser) : undefined;

    // Parse Speaker Notes
    let notes: string | undefined;
    let notesBody: import('@hokkyss/pptx-core').PptxTextBody | undefined;
    const notesRel = slideRelResolver.getRelationshipsByType('notesSlide')[0];
    const notesPath = notesRel ? notesRel.resolvedTarget : `ppt/notesSlides/notesSlide${slideNumber}.xml`;
    if (zipReader.hasFile(notesPath)) {
      const notesXml = zipReader.getFileText(notesPath);
      if (notesXml) {
        const parsedNotes = parseSpeakerNotes(notesXml, xmlParser);
        if (parsedNotes) {
          notes = parsedNotes.notes;
          notesBody = parsedNotes.notesBody;
        }
      }
    }

    // Parse Slide Background
    const parsedSlideXml = xmlParser.parse<Record<string, unknown>>(slideXml);
    const sldNode = (parsedSlideXml['p:sld'] || parsedSlideXml['sld'] || {}) as Record<string, unknown>;
    const cSldNode = (sldNode['p:cSld'] || sldNode['cSld'] || {}) as Record<string, unknown>;
    const bgNode = (cSldNode['p:bg'] || cSldNode['bg'] || sldNode['p:bg'] || sldNode['bg']) as Record<string, unknown> | undefined;
    const background = parseBackground(bgNode);

    slides.push({
      animations,
      background,
      elements: shapes,
      layoutId,
      notes,
      notesBody,
      rawXml: slideXml,
      relsXml: slideRelsXml,
      shapes,
      slideId,
      slideNumber,
      transition,
    });
  }

  metadata.slideCount = slides.length;

  // 9. Collect Auxiliary Package Parts (Charts, Embeddings, Notes, Comments, Handouts, CustomXml)
  const customXml: PptxCustomXmlPart[] = [];
  const auxiliaryPaths = [
    ...zipReader.getPathsStartingWith('ppt/charts/'),
    ...zipReader.getPathsStartingWith('ppt/embeddings/'),
    ...zipReader.getPathsStartingWith('ppt/notesSlides/'),
    ...zipReader.getPathsStartingWith('ppt/notesMasters/'),
    ...zipReader.getPathsStartingWith('ppt/handoutMasters/'),
    ...zipReader.getPathsStartingWith('ppt/comments/'),
    ...zipReader.getPathsStartingWith('customXml/'),
  ];
  if (zipReader.hasFile('ppt/commentAuthors.xml')) auxiliaryPaths.push('ppt/commentAuthors.xml');
  if (zipReader.hasFile('ppt/presProps.xml')) auxiliaryPaths.push('ppt/presProps.xml');
  if (zipReader.hasFile('ppt/viewProps.xml')) auxiliaryPaths.push('ppt/viewProps.xml');
  if (zipReader.hasFile('ppt/tableStyles.xml')) auxiliaryPaths.push('ppt/tableStyles.xml');

  const processedAux = new Set<string>();
  for (const aPath of auxiliaryPaths) {
    if (processedAux.has(aPath)) continue;
    processedAux.add(aPath);

    if (aPath.endsWith('.xml') || aPath.endsWith('.rels')) {
      const xmlString = zipReader.getFileText(aPath);
      if (xmlString) {
        customXml.push({ path: aPath, xmlString });
      }
    } else {
      const binaryData = zipReader.getFileData(aPath);
      if (binaryData) {
        customXml.push({ binaryData, path: aPath });
      }
    }
  }

  return {
    customXml,
    media,
    metadata,
    slideLayouts,
    slideMasters,
    slides,
    themes,
  };
}
