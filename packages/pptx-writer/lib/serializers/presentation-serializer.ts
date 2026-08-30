import type { PptxDocument } from '@hokkyss/pptx-core';
import { serializeXml } from '../xml/xml-builder';

export interface PresentationSerializerOptions {
  handoutMasterRelId?: string;
  masterRelIds?: string[];
  notesMasterRelId?: string;
  slideRelIds: string[];
}

function buildDefaultTextStyle(): Record<string, unknown> {
  const levels: Record<string, unknown> = {
    'a:defPPr': {
      'a:defRPr': { '@_lang': 'en-US' },
    },
  };

  for (let i = 1; i <= 9; i++) {
    const marL = (i - 1) * 457200;
    levels[`a:lvl${i}pPr`] = {
      '@_algn': 'l',
      '@_defTabSz': 914400,
      '@_eaLnBrk': '1',
      '@_hangingPunct': '1',
      '@_latinLnBrk': '0',
      '@_marL': marL,
      '@_rtl': '0',
      'a:defRPr': {
        '@_kern': 1200,
        '@_sz': 1800,
        'a:cs': { '@_typeface': '+mn-cs' },
        'a:ea': { '@_typeface': '+mn-ea' },
        'a:latin': { '@_typeface': '+mn-lt' },
        'a:solidFill': {
          'a:schemeClr': { '@_val': 'tx1' },
        },
      },
    };
  }

  return levels;
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
  const sldMasterIdList = masterIds.map((relId, idx) => ({
    '@_id': 2147483648 + idx,
    '@_r:id': relId,
  }));

  const sldIdList = options.slideRelIds.map((relId, idx) => ({
    '@_id': 256 + idx,
    '@_r:id': relId,
  }));

  const slideWidth = document.metadata?.slideWidth ?? 12192000;
  const slideHeight = document.metadata?.slideHeight ?? 6858000;

  const presObj: Record<string, unknown> = {
    '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    ...(document.metadata?.firstSlideNumber !== undefined
      ? { '@_firstSlideNum': document.metadata.firstSlideNumber }
      : {}),
    'p:sldMasterIdLst': {
      'p:sldMasterId': sldMasterIdList,
    },
  };

  if (options.notesMasterRelId) {
    presObj['p:notesMasterIdLst'] = {
      'p:notesMasterId': { '@_r:id': options.notesMasterRelId },
    };
  }

  if (options.handoutMasterRelId) {
    presObj['p:handoutMasterIdLst'] = {
      'p:handoutMasterId': { '@_r:id': options.handoutMasterRelId },
    };
  }

  presObj['p:sldIdLst'] = {
    'p:sldId': sldIdList,
  };
  presObj['p:sldSz'] = {
    '@_cx': slideWidth,
    '@_cy': slideHeight,
  };
  presObj['p:notesSz'] = {
    '@_cx': 6858000,
    '@_cy': 9144000,
  };
  presObj['p:defaultTextStyle'] = buildDefaultTextStyle();

  const root = {
    'p:presentation': presObj,
  };

  return serializeXml(root);
}
