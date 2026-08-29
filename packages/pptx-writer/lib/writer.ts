import type { PptxDocument, PptxElement, PptxHyperlink } from '@hokkyss/pptx-core';
import { sanitizeHyperlinkUrl, sanitizeSlideIndex } from '@hokkyss/pptx-core';
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

/**
 * Decodes a base64 string to a Uint8Array byte buffer.
 * @param base64 - Base64 encoded string.
 * @returns Uint8Array byte array.
 */
function decodeBase64ToBytes(base64: string): Uint8Array {
  if (typeof Buffer !== 'undefined') {
    return new Uint8Array(Buffer.from(base64, 'base64'));
  }
  const binaryString = atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

const DEFAULT_AUDIO_POSTER_BASE64 = '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';
const DEFAULT_VIDEO_POSTER_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAUAAAAC0CAYAAADl5PURAAAE5UlEQVR4nO3X2U6cRxSF0VJewBM2NJPnITaeZ+f18m55rIqIRAtCg8H5/z7V2aukdX+0Lz6p2o3tlx0gUas+AKCKAAKxBBCIJYBArHZj52UHSNRu7rzqAIkEEIglgEAsAQRiCSAQSwCBWAIIxBJAIJYAArHazcWrDpCo3VwcdYBEAgjEarcWRx0gkQACsQQQiCWAQCwBBGIJIBBLAIFY7dbuUQdIJIBArHZr93UHSCSAQCwBBGK127uvO0AiAQRiCSAQSwCBWAIIxGq39153gEQCCMQSQCBWu733pgMkEkAglgACsQQQiNXu7L3pAIkEEIglgECsdmf/TQdIJIBALAEEYrU7+287QCIBBGIJIBBLAIFYAgjEEkAgVtvaf9sBErWtg7cdIJEAArEEEIglgECstnXwrgMkEkBK/fbnX+U3kEsAmd1x5H5V9e38vwkgs/gv0RND1kUAmcwc0RND5iSATGKd8RNBptK2Dt91+FUV4TsXwgF2YDO1u4fvOlxXdfRWqd6EzSOAXFt16ESQqQgg11IdOBFkSu3u4fsOV1EdtutFsH4vxieAXEl10ESQOQggP1UdMhFkLgLIpaoDJoLMSQC5VHW8BJA5CSAXqg6XCDK3dvf++w6rVEdr0gAOsCfjEUBWmio8x686fiLIRQSQc6aMzsmrjp8Iskq7d/99h9PmCOAoEazelrEIIOfMFcARQli9LWNp9+5/6HBi6uBc9GojWL8zYxBAzlhXACtDWL0x4xBAzlh3ACsiWL0x4xBAluaIzXWeCLJuAshSdQDXGcHqrRlDu/fgQ4djIwRwXRGs3poxCCBLowRwHSGs3poxCCBLowVwzghWb80YBJClEQM4Vwirt2YMAsjSyAGcOoLVWzOGtv3gY4e5vppzvKluq96cegLI0ugBnPKu6q0ZgwCyNHIAp76remvGIIAsjRjAOW4SQE4IIEujBXCu+AkgJ9r2w48djo0SwDnDtwzgAHtTTwBZGiGA64ifAHJCAFmqDuC64ieAnBBAzqgI4DrDJ36cJoCcse4Arjt+AshpAsgZ6wpgRfgEkH9r2w8/dTht7gDWxq9+X8bRdh5+6nDaXAGsDN+J6m0ZiwByzhwBrA6fALKKALLSlAGsjp74cZG28+hTh1WqgzVp/AbYk/EIIBeqjpYAMjcB5FLV4RI/5iSAXKo6XgLInASQn6oOmPgxFwHkSqpDJn7MQQC5suqgiR9TazuPPne4quqwXS1+9TuxGQSQa6sOnPgxlbZ49LnDdVWHbpXqTdg8bfH4c4dfVR29f8I3wA5sJgFkEsLHJhJAJiN+bBoBZBaixyYQQGYneoxKACklclQSQCBWWzz+0gESCSAQSwCBWG3x5EsHSNR2n3zpAIkEEIglgEAsAQRiCSAQSwCBWAIIxGq7T752gEQCCMRqu0+/doBEAgjEEkAgVtt7+rUDJBJAIJYAArEEEIglgECstvf0WwdIJIBArLb37FsHSCSAQCwBBGIJIBBLAIFYbf/Ztw6QSACBWAIIxBJAIFbbf/a9AyRq+8+/d4BEAgjEEkAglgACsQQQiCWAQCwBBGK1g+ffO0AiAQRiCSAQqx28+NEBEgkgEEsAgVgCCMQSQCCWAAKxBBCIJYBALAEEYrXDFz86QKJ2+PsfHSCRAAKxBBCIJYBALAEEYv0NFG41taG9cqwAAAAASUVORK5CYII=';

const DEFAULT_AUDIO_POSTER_PNG = decodeBase64ToBytes(DEFAULT_AUDIO_POSTER_BASE64);
const DEFAULT_VIDEO_POSTER_PNG = decodeBase64ToBytes(DEFAULT_VIDEO_POSTER_BASE64);

const REL_TYPES = {
  audio: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  commentAuthors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  handoutMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  media: 'http://schemas.microsoft.com/office/2007/relationships/media',
  notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  video: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
  viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
};

/**
 * Validates, sanitizes, and registers a hyperlink target in the slide relationship table.
 */
function processHyperlinkTarget(
  hlink: PptxHyperlink | string,
  slideRels: RelationshipEntry[],
  getSlideRelId: () => string,
): PptxHyperlink | undefined {
  if (typeof hlink === 'string') {
    const safeUrl = sanitizeHyperlinkUrl(hlink);
    if (!safeUrl) return undefined;
    const rId = getSlideRelId();
    slideRels.push({
      id: rId,
      target: safeUrl,
      targetMode: 'External',
      type: REL_TYPES.hyperlink,
    });
    return { rId, url: safeUrl };
  }

  if (hlink.rId) return hlink;

  if (hlink.url) {
    const safeUrl = sanitizeHyperlinkUrl(hlink.url);
    if (safeUrl) {
      const rId = getSlideRelId();
      hlink.rId = rId;
      hlink.url = safeUrl;
      slideRels.push({
        id: rId,
        target: safeUrl,
        targetMode: 'External',
        type: REL_TYPES.hyperlink,
      });
      return hlink;
    }
    return undefined;
  }

  if (hlink.slideIndex !== undefined) {
    const safeIndex = sanitizeSlideIndex(hlink.slideIndex);
    if (safeIndex !== undefined) {
      const rId = getSlideRelId();
      hlink.rId = rId;
      hlink.slideIndex = safeIndex;
      slideRels.push({
        id: rId,
        target: `slide${safeIndex}.xml`,
        type: REL_TYPES.slide,
      });
      return hlink;
    }
    return undefined;
  }

  return hlink;
}

/**
 * Traverses an element, child elements, and text runs to register and allocate relationship IDs for hyperlinks.
 */
function registerElementHyperlinks(
  elem: PptxElement,
  slideRels: RelationshipEntry[],
  getSlideRelId: () => string,
): void {
  if (elem.hyperlink) {
    const processed = processHyperlinkTarget(elem.hyperlink, slideRels, getSlideRelId);
    if (processed) {
      elem.hyperlink = processed;
    } else {
      delete elem.hyperlink;
    }
  }

  if (elem.textBody) {
    for (const p of elem.textBody.paragraphs || []) {
      for (const r of p.runs || []) {
        if (r.properties?.hyperlink) {
          const processed = processHyperlinkTarget(r.properties.hyperlink, slideRels, getSlideRelId);
          if (processed) {
            r.properties.hyperlink = processed;
          } else {
            delete r.properties.hyperlink;
          }
        }
      }
    }
  }

  if (elem.elementType === 'table' && elem.table) {
    for (const row of elem.table.rows || []) {
      for (const cell of row.cells || []) {
        if (cell.textBody) {
          for (const p of cell.textBody.paragraphs || []) {
            for (const r of p.runs || []) {
              if (r.properties?.hyperlink) {
                const processed = processHyperlinkTarget(r.properties.hyperlink, slideRels, getSlideRelId);
                if (processed) {
                  r.properties.hyperlink = processed;
                } else {
                  delete r.properties.hyperlink;
                }
              }
            }
          }
        }
      }
    }
  }

  if (elem.elementType === 'group' && elem.children) {
    for (const child of elem.children) {
      registerElementHyperlinks(child, slideRels, getSlideRelId);
    }
  }
}

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
    const slideAudioVideoRelMap = new Map<string, { embedRelId: string; imageRelId?: string; linkRelId: string }>();
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
      } else if (elem.elementType === 'audio') {
        const mediaId = elem.audio.mediaId;
        const matchedFileName = mediaMap.get(mediaId) || `${mediaId}.mp3`;

        // Ensure default audio poster image exists in package
        files['ppt/media/audio_poster.png'] = DEFAULT_AUDIO_POSTER_PNG;
        mediaExtensions.add('png');

        // r:embed — image relationship (for <p:pic> <a:blip r:embed>)
        const imageRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: imageRelId,
          target: '../media/audio_poster.png',
          type: REL_TYPES.image,
        });

        // r:link — audio relationship (for <p:audioFile r:link>)
        const linkRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: linkRelId,
          target: `../media/${matchedFileName}`,
          type: REL_TYPES.audio,
        });
        // r:embed — media relationship (for p14:media r:embed extension)
        const embedRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: embedRelId,
          target: `../media/${matchedFileName}`,
          type: REL_TYPES.media,
        });
        slideAudioVideoRelMap.set(mediaId, { embedRelId, imageRelId, linkRelId });
      } else if (elem.elementType === 'video') {
        const mediaId = elem.video.mediaId;
        const matchedFileName = mediaMap.get(mediaId) || `${mediaId}.mp4`;

        // Ensure default video poster image exists in package
        files['ppt/media/video_poster.png'] = DEFAULT_VIDEO_POSTER_PNG;
        mediaExtensions.add('png');

        // r:embed — image relationship (for <p:pic> <a:blip r:embed>)
        const imageRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: imageRelId,
          target: '../media/video_poster.png',
          type: REL_TYPES.image,
        });

        // r:link — video relationship (for <p:videoFile r:link>)
        const linkRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: linkRelId,
          target: `../media/${matchedFileName}`,
          type: REL_TYPES.video,
        });
        // r:embed — media relationship (for p14:media r:embed extension)
        const embedRelId = `rId${slideRelCounter++}`;
        slideRels.push({
          id: embedRelId,
          target: `../media/${matchedFileName}`,
          type: REL_TYPES.media,
        });
        slideAudioVideoRelMap.set(mediaId, { embedRelId, imageRelId, linkRelId });
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

      // Register hyperlinks across element, text runs, and tables/groups
      registerElementHyperlinks(elem, slideRels, () => `rId${slideRelCounter++}`);
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
      files[`ppt/slides/slide${slideNum}.xml`] = serializeSlide(slide, slidePictureEmbedMap, slideChartRelIds, slideAudioVideoRelMap);
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
