import type { PptxTextBody } from '@hokkyss/pptx-core';
import { el, serializeXml, type XmlElement } from '../xml/xml-element';
import { serializeTextBody } from './text-serializer';

/**
 * Serializes speaker notes (plain string or structured PptxTextBody) into an OpenXML `<p:notes>` document.
 */
export function serializeNotesSlide(notesInput: PptxTextBody | string): string {
  let txBodyNode: XmlElement;

  if (typeof notesInput === 'string') {
    const paragraphs = notesInput.split('\n').map((line) => ({
      properties: {},
      runs: [
        {
          properties: {},
          text: line,
        },
      ],
    }));
    txBodyNode = serializeTextBody({
      bodyProperties: {},
      paragraphs,
    }, 'p:txBody');
  } else {
    txBodyNode = serializeTextBody(notesInput, 'p:txBody');
  }

  const notesRoot = el('p:notes', {
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  }, [
    el('p:cSld', [
      el('p:spTree', [
        el('p:nvGrpSpPr', [
          el('p:cNvPr', { id: '1', name: '' }),
          el('p:cNvGrpSpPr'),
          el('p:nvPr'),
        ]),
        el('p:grpSpPr', [
          el('a:xfrm', [
            el('a:off', { x: 0, y: 0 }),
            el('a:ext', { cx: 0, cy: 0 }),
            el('a:chOff', { x: 0, y: 0 }),
            el('a:chExt', { cx: 0, cy: 0 }),
          ]),
        ]),
        el('p:sp', [
          el('p:nvSpPr', [
            el('p:cNvPr', { id: '2', name: 'Slide Image Placeholder 1' }),
            el('p:cNvSpPr', [
              el('a:spLocks', { noChangeAspect: '1', noGrp: '1', noRot: '1' }),
            ]),
            el('p:nvPr', [
              el('p:ph', { type: 'sldImg' }),
            ]),
          ]),
          el('p:spPr'),
        ]),
        el('p:sp', [
          el('p:nvSpPr', [
            el('p:cNvPr', { id: '3', name: 'Notes Placeholder 2' }),
            el('p:cNvSpPr', [
              el('a:spLocks', { noGrp: '1' }),
            ]),
            el('p:nvPr', [
              el('p:ph', { idx: 1, type: 'body' }),
            ]),
          ]),
          el('p:spPr'),
          txBodyNode,
        ]),
      ]),
    ]),
    el('p:clrMapOvr', [
      el('a:masterClrMapping'),
    ]),
  ]);

  return serializeXml(notesRoot);
}

/**
 * Serializes the relationships for a notesSlide pointing back to its parent slide and notesMaster.
 */
export function serializeNotesSlideRels(slideNum: number): string {
  const root = el('Relationships', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
  }, [
    el('Relationship', {
      Id: 'rId1',
      Target: '../notesMasters/notesMaster1.xml',
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
    }),
    el('Relationship', {
      Id: 'rId2',
      Target: `../slides/slide${slideNum}.xml`,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    }),
  ]);

  return serializeXml(root);
}
