import type { PptxMetadata } from '@hokkyss/pptx-core';
import { el, serializeXml } from '../xml/xml-element';

/**
 * Serializes extended document metadata to `docProps/app.xml`.
 * @param metadata
 */
export function serializeAppProperties(metadata: Partial<PptxMetadata> = {}): string {
  const root = el('Properties', {
    xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  }, [
    el('Application', 'Microsoft Office PowerPoint'),
    el('AppVersion', '16.0000'),
    el('Company', ''),
    el('HeadingPairs', [
      el('vt:vector', {
        baseType: 'variant',
        size: 2,
      }, [
        el('vt:variant', [el('vt:lpstr', 'Theme')]),
        el('vt:variant', [el('vt:i4', 1)]),
      ]),
    ]),
    el('HiddenSlides', 0),
    el('HyperlinksChanged', 'false'),
    el('LinksUpToDate', 'false'),
    el('MMClips', 0),
    el('Notes', 0),
    el('Paragraphs', 0),
    el('PresentationFormat', 'Widescreen'),
    el('ScaleCrop', 'false'),
    el('SharedDoc', 'false'),
    el('Slides', metadata.slideCount ?? 1),
    el('TitlesOfParts', [
      el('vt:vector', {
        baseType: 'lpstr',
        size: 1,
      }, [
        el('vt:lpstr', 'Office Theme'),
      ]),
    ]),
    el('TotalTime', 0),
    el('Words', 0),
  ]);

  return serializeXml(root);
}

/**
 * Serializes document metadata to `docProps/core.xml` (Dublin Core & OpenXML metadata).
 * @param metadata
 */
export function serializeCoreProperties(metadata: Partial<PptxMetadata> = {}): string {
  const createdDate = metadata.created ? metadata.created.toISOString() : new Date().toISOString();
  const modifiedDate = metadata.modified ? metadata.modified.toISOString() : createdDate;

  const root = el('cp:coreProperties', {
    'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
    'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
    'xmlns:dcterms': 'http://purl.org/dc/terms/',
    'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
  }, [
    el('cp:lastModifiedBy', metadata.lastModifiedBy ?? metadata.creator ?? 'PPTX Writer'),
    el('cp:revision', metadata.revision ?? 1),
    el('dc:creator', metadata.creator ?? 'PPTX Writer'),
    el('dc:title', metadata.title ?? ''),
    el('dcterms:created', { 'xsi:type': 'dcterms:W3CDTF' }, createdDate),
    el('dcterms:modified', { 'xsi:type': 'dcterms:W3CDTF' }, modifiedDate),
  ]);

  return serializeXml(root);
}
