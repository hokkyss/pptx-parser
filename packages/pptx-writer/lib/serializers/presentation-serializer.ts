import type { PptxDocument } from '@hokkyss/pptx-core';
import { el, serializeXml, type XmlElement } from '../xml/xml-element';

export interface PresentationSerializerOptions {
  handoutMasterRelId?: string;
  masterRelIds?: string[];
  notesMasterRelId?: string;
  slideRelIds: string[];
}

function buildDefaultTextStyle(): XmlElement {
  const children: XmlElement[] = [
    el('a:defPPr', [
      el('a:defRPr', { lang: 'en-US' }),
    ]),
  ];

  for (let i = 1; i <= 9; i++) {
    const marL = (i - 1) * 457200;
    children.push(
      el(`a:lvl${i}pPr`, {
        algn: 'l',
        defTabSz: 914400,
        eaLnBrk: 1,
        hangingPunct: 1,
        latinLnBrk: 0,
        marL,
        rtl: 0,
      }, [
        el('a:defRPr', {
          kern: 1200,
          sz: 1800,
        }, [
          el('a:solidFill', [
            el('a:schemeClr', { val: 'tx1' }),
          ]),
          el('a:latin', { typeface: '+mn-lt' }),
          el('a:ea', { typeface: '+mn-ea' }),
          el('a:cs', { typeface: '+mn-cs' }),
        ]),
      ]),
    );
  }

  return el('p:defaultTextStyle', children);
}

/**
 * Serializes `ppt/presentation.xml` adhering strictly to ECMA-376 sequence.
 * Sequence: p:sldMasterIdLst -> p:notesMasterIdLst -> p:handoutMasterIdLst -> p:sldIdLst -> p:sldSz -> p:notesSz -> p:defaultTextStyle
 */
export function serializePresentation(
  document: PptxDocument,
  options: PresentationSerializerOptions,
): string {
  const masterIds = options.masterRelIds || ['rId1'];
  const sldMasterIdNodes = masterIds.map((relId, idx) =>
    el('p:sldMasterId', {
      id: 2147483648 + idx,
      'r:id': relId,
    }),
  );

  const sldIdNodes = options.slideRelIds.map((relId, idx) =>
    el('p:sldId', {
      id: 256 + idx,
      'r:id': relId,
    }),
  );

  const slideWidth = document.metadata?.slideWidth ?? 12192000;
  const slideHeight = document.metadata?.slideHeight ?? 6858000;

  const presAttrs: Record<string, number | string | undefined> = {
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  };
  if (document.metadata?.firstSlideNumber !== undefined) {
    presAttrs.firstSlideNum = document.metadata.firstSlideNumber;
  }

  const presChildren: XmlElement[] = [
    el('p:sldMasterIdLst', sldMasterIdNodes),
  ];

  if (options.notesMasterRelId) {
    presChildren.push(
      el('p:notesMasterIdLst', [
        el('p:notesMasterId', { 'r:id': options.notesMasterRelId }),
      ]),
    );
  }

  if (options.handoutMasterRelId) {
    presChildren.push(
      el('p:handoutMasterIdLst', [
        el('p:handoutMasterId', { 'r:id': options.handoutMasterRelId }),
      ]),
    );
  }

  presChildren.push(
    el('p:sldIdLst', sldIdNodes),
    el('p:sldSz', { cx: slideWidth, cy: slideHeight }),
    el('p:notesSz', { cx: 6858000, cy: 9144000 }),
    buildDefaultTextStyle(),
  );

  const root = el('p:presentation', presAttrs, presChildren);

  return serializeXml(root);
}
