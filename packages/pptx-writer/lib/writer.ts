import type { PptxDocument } from '@hokkyss/pptx-core';
import { serializeChart } from './serializers/chart-serializer';
import { serializeContentTypes } from './serializers/content-types-serializer';
import { serializeNotesSlide, serializeNotesSlideRels } from './serializers/notes-serializer';
import { serializePresentation } from './serializers/presentation-serializer';
import { serializeAppProperties, serializeCoreProperties } from './serializers/properties-serializer';
import { type RelationshipEntry, serializeRelationships } from './serializers/relationship-serializer';
import { serializeSlide } from './serializers/slide-serializer';
import { serializeTheme } from './serializers/theme-serializer';
import {
  DEFAULT_NOTES_MASTER_XML,
  DEFAULT_PRES_PROPS_XML,
  DEFAULT_SLIDE_LAYOUT_XML,
  DEFAULT_SLIDE_MASTER_XML,
  DEFAULT_TABLE_STYLES_XML,
  DEFAULT_THEME_XML,
  DEFAULT_VIEW_PROPS_XML,
} from './templates/defaults';
import { createZipPackage } from './zip/zip-writer';

export interface WritePptxOptions {
  /**
   * Validation and fallback mode:
   * - 'lenient' (default): Automatically injects missing layouts, themes, IDs, and dimensions.
   * - 'strict': Fails fast with descriptive error messages on invalid AST structures.
   */
  mode?: 'lenient' | 'strict';
}

const REL_TYPES = {
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  commentAuthors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  handoutMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
};

/**
 * Serializes a `PptxDocument` AST into a valid `.pptx` presentation binary buffer.
 * Fully preserves themes, slide masters, slide layouts, media assets, and auxiliary parts.
 * @param document The parsed or constructed PptxDocument AST.
 * @param options Optional serializer options.
 * @returns Promise resolving to a Uint8Array representing the zipped `.pptx` archive.
 */
export async function writePptx(
  document: PptxDocument,
  options: WritePptxOptions = {},
): Promise<Uint8Array> {
  const mode = options.mode ?? 'lenient';

  if (mode === 'strict') {
    if (!document.slides || document.slides.length === 0) {
      throw new Error('Strict mode error: Document must contain at least one slide.');
    }
  }

  const slides = document.slides && document.slides.length > 0
    ? document.slides
    : [{
        animations: [],
        elements: [],
        shapes: [],
        slideId: 'rId1',
        slideNumber: 1,
      }];

  const files: Record<string, string | Uint8Array> = {};

  // 1. Root relationships (_rels/.rels)
  const rootRels: RelationshipEntry[] = [
    { id: 'rId1', target: 'ppt/presentation.xml', type: REL_TYPES.officeDocument },
    { id: 'rId2', target: 'docProps/core.xml', type: REL_TYPES.coreProperties },
    { id: 'rId3', target: 'docProps/app.xml', type: REL_TYPES.extendedProperties },
  ];
  files['_rels/.rels'] = serializeRelationships(rootRels);

  // 2. Metadata (docProps/core.xml & docProps/app.xml)
  const metadata = {
    ...document.metadata,
    slideCount: slides.length,
  };
  files['docProps/core.xml'] = serializeCoreProperties(metadata);
  files['docProps/app.xml'] = serializeAppProperties(metadata);

  // 3. Presentation auxiliary files
  files['ppt/presProps.xml'] = DEFAULT_PRES_PROPS_XML;
  files['ppt/viewProps.xml'] = DEFAULT_VIEW_PROPS_XML;
  files['ppt/tableStyles.xml'] = DEFAULT_TABLE_STYLES_XML;

  // 4. Themes
  const themeNames: string[] = [];
  if (document.themes && document.themes.length > 0) {
    for (let i = 0; i < document.themes.length; i++) {
      const th = document.themes[i];
      const thName = th.id ? (th.id.endsWith('.xml') ? th.id : `${th.id}.xml`) : `theme${i + 1}.xml`;
      themeNames.push(thName);
      files[`ppt/theme/${thName}`] = serializeTheme(th);
    }
  } else {
    themeNames.push('theme1.xml');
    files['ppt/theme/theme1.xml'] = DEFAULT_THEME_XML;
  }

  // 5. Slide Masters & Slide Master Relationships
  const masterNames: string[] = [];
  const masterRelIds: string[] = [];
  let presRelCounter = 1;

  if (document.slideMasters && document.slideMasters.length > 0) {
    for (let i = 0; i < document.slideMasters.length; i++) {
      const sm = document.slideMasters[i];
      const smName = sm.id ? (sm.id.endsWith('.xml') ? sm.id : `${sm.id}.xml`) : `slideMaster${i + 1}.xml`;
      masterNames.push(smName);

      const relId = `rId${presRelCounter++}`;
      masterRelIds.push(relId);

      files[`ppt/slideMasters/${smName}`] = sm.rawXml || DEFAULT_SLIDE_MASTER_XML;
      if (sm.relsXml) {
        files[`ppt/slideMasters/_rels/${smName}.rels`] = sm.relsXml;
      } else {
        files[`ppt/slideMasters/_rels/${smName}.rels`] = serializeRelationships([
          { id: 'rId1', target: '../slideLayouts/slideLayout1.xml', type: REL_TYPES.slideLayout },
          { id: 'rId2', target: `../theme/${themeNames[0]}`, type: REL_TYPES.theme },
        ]);
      }
    }
  } else {
    masterNames.push('slideMaster1.xml');
    masterRelIds.push('rId1');
    presRelCounter = 2;
    files['ppt/slideMasters/slideMaster1.xml'] = DEFAULT_SLIDE_MASTER_XML;
    files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = serializeRelationships([
      { id: 'rId1', target: '../slideLayouts/slideLayout1.xml', type: REL_TYPES.slideLayout },
      { id: 'rId2', target: `../theme/${themeNames[0]}`, type: REL_TYPES.theme },
    ]);
  }

  // 6. Slide Layouts & Slide Layout Relationships
  const layoutNames: string[] = [];
  if (document.slideLayouts && document.slideLayouts.length > 0) {
    for (let i = 0; i < document.slideLayouts.length; i++) {
      const sl = document.slideLayouts[i];
      const slName = sl.id ? (sl.id.endsWith('.xml') ? sl.id : `${sl.id}.xml`) : `slideLayout${i + 1}.xml`;
      layoutNames.push(slName);

      files[`ppt/slideLayouts/${slName}`] = sl.rawXml || DEFAULT_SLIDE_LAYOUT_XML;
      if (sl.relsXml) {
        files[`ppt/slideLayouts/_rels/${slName}.rels`] = sl.relsXml;
      } else {
        const targetMaster = sl.masterId ? (sl.masterId.endsWith('.xml') ? sl.masterId : `${sl.masterId}.xml`) : masterNames[0];
        files[`ppt/slideLayouts/_rels/${slName}.rels`] = serializeRelationships([
          { id: 'rId1', target: `../slideMasters/${targetMaster}`, type: REL_TYPES.slideMaster },
        ]);
      }
    }
  } else {
    layoutNames.push('slideLayout1.xml');
    files['ppt/slideLayouts/slideLayout1.xml'] = DEFAULT_SLIDE_LAYOUT_XML;
    files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = serializeRelationships([
      { id: 'rId1', target: `../slideMasters/${masterNames[0]}`, type: REL_TYPES.slideMaster },
    ]);
  }

  // 7. Media assets
  const mediaMap = new Map<string, string>(); // mediaId/path -> media file name
  const mediaExtensions = new Set<string>();

  for (const asset of document.media || []) {
    const fileName = asset.fileName || asset.filename || `image_${asset.id || '1'}.png`;
    const targetPath = `ppt/media/${fileName}`;

    let assetData: null | Uint8Array = asset.data;
    if (!assetData && asset.lazyGetter) {
      assetData = await asset.lazyGetter();
    }

    if (assetData) {
      files[targetPath] = assetData;
    }

    const ext = fileName.split('.').pop() || 'png';
    mediaExtensions.add(ext);

    mediaMap.set(asset.id, fileName);
    if (asset.path) mediaMap.set(asset.path, fileName);
    mediaMap.set(fileName, fileName);
  }

  // 8. Custom XML & Auxiliary parts (Charts, Embeddings, Notes, Comments, etc.)
  const customPartOverrides: Array<{ contentType: string; partName: string }> = [];

  for (const custom of document.customXml || []) {
    files[custom.path] = custom.binaryData ?? custom.xmlString ?? '';

    const normalized = custom.path.startsWith('/') ? custom.path : `/${custom.path}`;
    if (normalized.startsWith('/ppt/charts/chart') && normalized.endsWith('.xml')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        partName: normalized,
      });
    } else if (normalized.startsWith('/ppt/embeddings/') && normalized.endsWith('.xlsx')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        partName: normalized,
      });
    } else if (normalized.startsWith('/ppt/notesSlides/') && normalized.endsWith('.xml')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
        partName: normalized,
      });
    } else if (normalized.startsWith('/ppt/notesMasters/') && normalized.endsWith('.xml')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
        partName: normalized,
      });
    } else if (normalized.startsWith('/ppt/handoutMasters/') && normalized.endsWith('.xml')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml',
        partName: normalized,
      });
    } else if (normalized.startsWith('/ppt/comments/') && normalized.endsWith('.xml')) {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
        partName: normalized,
      });
    } else if (normalized === '/ppt/commentAuthors.xml') {
      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
        partName: normalized,
      });
    }
  }

  // 9. Slides and Slide Relationships
  const presentationSlideRelIds: string[] = [];
  let notesSlideCounter = 1;
  let chartCounter = 1;

  for (let idx = 0; idx < slides.length; idx++) {
    const slide = slides[idx];
    const slideNum = idx + 1;
    const slideRelId = `rId${presRelCounter++}`;
    presentationSlideRelIds.push(slideRelId);

    const rawLayoutId = slide.layoutId || (document.slideLayouts && document.slideLayouts.length > 0 ? document.slideLayouts[0].id : 'slideLayout1');
    const layoutTarget = rawLayoutId.endsWith('.xml') ? rawLayoutId : `${rawLayoutId}.xml`;

    const slideRels: RelationshipEntry[] = [
      { id: 'rId1', target: `../slideLayouts/${layoutTarget}`, type: REL_TYPES.slideLayout },
    ];

    const slidePictureEmbedMap = new Map<string, string>();
    const slideChartRelIds: string[] = [];
    let slideRelCounter = 2;

    const elements = (slide.elements && slide.elements.length > 0) ? slide.elements : (slide.shapes || []);
    for (const elem of elements) {
      if (elem.elementType === 'picture') {
        const mediaId = elem.picture.mediaId || elem.id;
        const matchedFileName = mediaMap.get(mediaId) || `${mediaId}.png`;
        const picRelId = `rId${slideRelCounter++}`;

        slideRels.push({
          id: picRelId,
          target: `../media/${matchedFileName}`,
          type: REL_TYPES.image,
        });

        slidePictureEmbedMap.set(mediaId, picRelId);
        slidePictureEmbedMap.set(elem.id, picRelId);
      } else if (elem.elementType === 'chart') {
        if (elem.chart) {
          const chartIndex = chartCounter++;
          const chartRelId = `rId${slideRelCounter++}`;

          slideRels.push({
            id: chartRelId,
            target: `../charts/chart${chartIndex}.xml`,
            type: REL_TYPES.chart,
          });

          slideChartRelIds.push(chartRelId);

          const chartXml = serializeChart(elem.chart);
          files[`ppt/charts/chart${chartIndex}.xml`] = chartXml;

          customPartOverrides.push({
            contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
            partName: `/ppt/charts/chart${chartIndex}.xml`,
          });
        }
      }
    }

    // Handle speaker notes (strictly sequential notesSlide indexing)
    const hasNotes = (slide.notes && slide.notes.trim().length > 0) || slide.notesBody;
    if (hasNotes) {
      const notesIndex = notesSlideCounter++;
      const notesRelId = `rId${slideRelCounter++}`;
      slideRels.push({
        id: notesRelId,
        target: `../notesSlides/notesSlide${notesIndex}.xml`,
        type: REL_TYPES.notesSlide,
      });

      files[`ppt/notesSlides/notesSlide${notesIndex}.xml`] = serializeNotesSlide(slide.notesBody || slide.notes || '');
      files[`ppt/notesSlides/_rels/notesSlide${notesIndex}.xml.rels`] = serializeNotesSlideRels(slideNum);

      customPartOverrides.push({
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
        partName: `/ppt/notesSlides/notesSlide${notesIndex}.xml`,
      });
    }

    if (slide.relsXml && !hasNotes) {
      files[`ppt/slides/_rels/slide${slideNum}.xml.rels`] = slide.relsXml;
    } else {
      files[`ppt/slides/_rels/slide${slideNum}.xml.rels`] = serializeRelationships(slideRels);
    }

    if (slide.rawXml) {
      files[`ppt/slides/slide${slideNum}.xml`] = slide.rawXml;
    } else {
      files[`ppt/slides/slide${slideNum}.xml`] = serializeSlide(slide, slidePictureEmbedMap, slideChartRelIds);
    }
  }

  // 10. Presentation Relationships (ppt/_rels/presentation.xml.rels)
  const presRels: RelationshipEntry[] = [];

  for (let i = 0; i < masterNames.length; i++) {
    presRels.push({
      id: masterRelIds[i],
      target: `slideMasters/${masterNames[i]}`,
      type: REL_TYPES.slideMaster,
    });
  }

  for (let i = 0; i < slides.length; i++) {
    presRels.push({
      id: presentationSlideRelIds[i],
      target: `slides/slide${i + 1}.xml`,
      type: REL_TYPES.slide,
    });
  }

  let notesMasterRelId: string | undefined;
  let handoutMasterRelId: string | undefined;

  for (const custom of document.customXml || []) {
    if (custom.path === 'ppt/notesMasters/notesMaster1.xml') {
      notesMasterRelId = `rId${presRelCounter++}`;
      presRels.push({
        id: notesMasterRelId,
        target: 'notesMasters/notesMaster1.xml',
        type: REL_TYPES.notesMaster,
      });
    } else if (custom.path === 'ppt/handoutMasters/handoutMaster1.xml') {
      handoutMasterRelId = `rId${presRelCounter++}`;
      presRels.push({
        id: handoutMasterRelId,
        target: 'handoutMasters/handoutMaster1.xml',
        type: REL_TYPES.handoutMaster,
      });
    } else if (custom.path === 'ppt/commentAuthors.xml') {
      presRels.push({
        id: `rId${presRelCounter++}`,
        target: 'commentAuthors.xml',
        type: REL_TYPES.commentAuthors,
      });
    }
  }

  // If any slide has notes and notesMaster is not present, generate default notesMaster
  const hasAnySlideNotes = slides.some((s) => (s.notes && s.notes.trim().length > 0) || s.notesBody);
  if (hasAnySlideNotes && !notesMasterRelId) {
    notesMasterRelId = `rId${presRelCounter++}`;
    presRels.push({
      id: notesMasterRelId,
      target: 'notesMasters/notesMaster1.xml',
      type: REL_TYPES.notesMaster,
    });

    files['ppt/notesMasters/notesMaster1.xml'] = DEFAULT_NOTES_MASTER_XML;
    files['ppt/theme/theme2.xml'] = DEFAULT_THEME_XML;
    files['ppt/notesMasters/_rels/notesMaster1.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../theme/theme2.xml" Type="${REL_TYPES.theme}"/></Relationships>`;

    customPartOverrides.push(
      {
        contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        partName: '/ppt/theme/theme2.xml',
      },
      {
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
        partName: '/ppt/notesMasters/notesMaster1.xml',
      },
    );
  }

  presRels.push(
    { id: `rId${presRelCounter++}`, target: 'presProps.xml', type: REL_TYPES.presProps },
    { id: `rId${presRelCounter++}`, target: 'viewProps.xml', type: REL_TYPES.viewProps },
    { id: `rId${presRelCounter++}`, target: `theme/${themeNames[0]}`, type: REL_TYPES.theme },
    { id: `rId${presRelCounter++}`, target: 'tableStyles.xml', type: REL_TYPES.tableStyles },
  );

  files['ppt/_rels/presentation.xml.rels'] = serializeRelationships(presRels);

  // 11. Presentation XML (ppt/presentation.xml)
  files['ppt/presentation.xml'] = serializePresentation(document, {
    handoutMasterRelId,
    masterRelIds,
    notesMasterRelId,
    slideRelIds: presentationSlideRelIds,
  });

  // 12. Content Types ([Content_Types].xml)
  files['[Content_Types].xml'] = serializeContentTypes({
    customPartOverrides,
    layoutNames,
    masterNames,
    mediaExtensions: Array.from(mediaExtensions),
    slideCount: slides.length,
    themeNames,
  });

  // 13. Create compressed ZIP archive
  return createZipPackage(files);
}
