import type { PptxMetadata } from '@hokkyss/pptx-core';
import { serializeXml } from '../xml/xml-builder';

/**
 * Serializes extended document metadata to `docProps/app.xml`.
 * @param metadata
 */
export function serializeAppProperties(metadata: PptxMetadata): string {
  const appObj: Record<string, unknown> = {
    Properties: {
      '@_xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
      '@_xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
      Application: 'Microsoft Office PowerPoint',
      AppVersion: '16.0000',
      Company: '',
      HeadingPairs: {
        'vt:vector': {
          '@_baseType': 'variant',
          '@_size': 2,
          'vt:variant': [
            { 'vt:lpstr': 'Theme' },
            { 'vt:i4': 1 },
          ],
        },
      },
      HiddenSlides: 0,
      HyperlinksChanged: false,
      LinksUpToDate: false,
      MMClips: 0,
      Notes: 0,
      Paragraphs: 0,
      PresentationFormat: 'Widescreen',
      ScaleCrop: false,
      SharedDoc: false,
      Slides: metadata.slideCount ?? 1,
      TitlesOfParts: {
        'vt:vector': {
          '@_baseType': 'lpstr',
          '@_size': 1,
          'vt:lpstr': 'Office Theme',
        },
      },
      TotalTime: 0,
      Words: 0,
    },
  };

  return serializeXml(appObj);
}

/**
 * Serializes document metadata to `docProps/core.xml` (Dublin Core & OpenXML metadata).
 * @param metadata
 */
export function serializeCoreProperties(metadata: PptxMetadata): string {
  const createdDate = metadata.created ? metadata.created.toISOString() : new Date().toISOString();
  const modifiedDate = metadata.modified ? metadata.modified.toISOString() : createdDate;

  const coreObj: Record<string, unknown> = {
    'cp:coreProperties': {
      '@_xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
      '@_xmlns:dc': 'http://purl.org/dc/elements/1.1/',
      '@_xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
      '@_xmlns:dcterms': 'http://purl.org/dc/terms/',
      '@_xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
      'cp:lastModifiedBy': metadata.lastModifiedBy ?? metadata.creator ?? 'PPTX Writer',
      'cp:revision': metadata.revision ?? 1,
      'dc:creator': metadata.creator ?? 'PPTX Writer',
      'dc:title': metadata.title ?? '',
      'dcterms:created': {
        '@_xsi:type': 'dcterms:W3CDTF',
        '#text': createdDate,
      },
      'dcterms:modified': {
        '@_xsi:type': 'dcterms:W3CDTF',
        '#text': modifiedDate,
      },
    },
  };

  return serializeXml(coreObj);
}
