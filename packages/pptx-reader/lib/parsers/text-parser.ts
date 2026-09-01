import { emu, hundredthsPoint, type RelationshipResolver } from '@hokkyss/pptx-core';
import {
  HundredthsPoint,
  PptxBullet,
  PptxParagraph,
  PptxParagraphProperties,
  PptxRun,
  PptxTextBody,
  PptxTextBodyProperties,
} from '../types/ast';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';
import { parseHyperlink } from './hyperlink-parser';

/**
 * Parses a single OpenXML paragraph node (`<a:p>`).
 *
 * Extracts paragraph properties (`alignment`, `leftMargin`, `rightMargin`, `firstLineIndent`, `spaceBefore`, `spaceAfter`, `lineSpacing`, `bullet`) and text runs.
 * @param pNode Raw XML paragraph object node (`<a:p>`).
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Parsed `PptxParagraph` structure.
 */
export function parseParagraph(
  pNode: Record<string, unknown>,
  relationshipResolver?: RelationshipResolver,
): PptxParagraph {
  const pPr = (pNode['a:pPr'] || pNode['pPr'] || {}) as Record<string, unknown>;

  const alignMap: Record<string, 'center' | 'justify' | 'left' | 'right'> = {
    center: 'center',
    ctr: 'center',
    just: 'justify',
    justify: 'justify',
    l: 'left',
    left: 'left',
    r: 'right',
    right: 'right',
  };

  const algnRaw = (pPr['@_algn'] as string) || '';
  const alignment = alignMap[algnRaw];
  const level = pPr['@_lvl'] !== undefined ? Number(pPr['@_lvl']) : 0;

  const leftMargin = pPr['@_marL'] !== undefined ? emu(Number(pPr['@_marL'])) : undefined;
  const rightMargin = pPr['@_marR'] !== undefined ? emu(Number(pPr['@_marR'])) : undefined;
  const firstLineIndent = pPr['@_indent'] !== undefined ? emu(Number(pPr['@_indent'])) : undefined;

  // Space before paragraph (<a:spcBef>)
  let spaceBefore: HundredthsPoint | undefined;
  const spcBefNode = (pPr['a:spcBef'] || pPr['spcBef']) as Record<string, unknown> | undefined;
  if (spcBefNode) {
    const befPts = (spcBefNode['a:spcPts'] || spcBefNode['spcPts']) as Record<string, unknown> | undefined;
    if (befPts && befPts['@_val'] !== undefined) {
      spaceBefore = hundredthsPoint(Number(befPts['@_val']));
    }
  }

  // Space after paragraph (<a:spcAft>)
  let spaceAfter: HundredthsPoint | undefined;
  const spcAftNode = (pPr['a:spcAft'] || pPr['spcAft']) as Record<string, unknown> | undefined;
  if (spcAftNode) {
    const aftPts = (spcAftNode['a:spcPts'] || spcAftNode['spcPts']) as Record<string, unknown> | undefined;
    if (aftPts && aftPts['@_val'] !== undefined) {
      spaceAfter = hundredthsPoint(Number(aftPts['@_val']));
    }
  }

  // Line spacing (<a:lnSpc>)
  let lineSpacing: HundredthsPoint | undefined;
  const lnSpcNode = (pPr['a:lnSpc'] || pPr['lnSpc']) as Record<string, unknown> | undefined;
  if (lnSpcNode) {
    const spcPct = (lnSpcNode['a:spcPct'] || lnSpcNode['spcPct'] || lnSpcNode['a:spcPts'] || lnSpcNode['spcPts']) as Record<string, unknown> | undefined;
    if (spcPct && spcPct['@_val'] !== undefined) {
      lineSpacing = hundredthsPoint(Number(spcPct['@_val']));
    }
  }

  // Extract bullet point properties
  let bullet: PptxBullet | undefined;
  const buChar = (pPr['a:buChar'] || pPr['buChar']) as Record<string, unknown> | undefined;
  const buAutoNum = (pPr['a:buAutoNum'] || pPr['buAutoNum']) as Record<string, unknown> | undefined;
  const buNone = pPr['a:buNone'] !== undefined ? pPr['a:buNone'] : pPr['buNone'];

  if (buNone !== undefined) {
    bullet = { type: 'none' };
  } else if (buChar) {
    bullet = {
      char: (buChar['@_char'] as string) || '•',
      type: 'char',
    };
  } else if (buAutoNum) {
    bullet = {
      autoNumType: (buAutoNum['@_type'] as string) || 'arabicPeriod',
      startAt: buAutoNum['@_startAt'] !== undefined ? Number(buAutoNum['@_startAt']) : 1,
      type: 'autoNum',
    };
  }

  // Check default run properties (<a:defRPr>) inside paragraph properties
  const defRPr = (pPr['a:defRPr'] || pPr['defRPr']) as Record<string, unknown> | undefined;
  const fallbackProps = defRPr ? parseRunProperties(defRPr) : undefined;

  // Extract text runs
  let rNodes = pNode['a:r'] || pNode['r'];
  const runs: PptxRun[] = [];

  if (rNodes) {
    if (!Array.isArray(rNodes)) {
      rNodes = [rNodes];
    }
    for (const rNode of rNodes as Record<string, unknown>[]) {
      const run = parseRun(rNode, fallbackProps, relationshipResolver);
      if (run) runs.push(run);
    }
  }

  const fldNodes = pNode['a:fld'] || pNode['fld'];
  if (fldNodes) {
    const flds = Array.isArray(fldNodes) ? fldNodes : [fldNodes];
    for (const fld of flds as Record<string, unknown>[]) {
      const run = parseRun(fld, fallbackProps, relationshipResolver);
      if (run) runs.push(run);
    }
  }

  const brNodes = pNode['a:br'] || pNode['br'];
  if (brNodes) {
    const brs = Array.isArray(brNodes) ? brNodes : [brNodes];
    for (const brNode of brs as Record<string, unknown>[]) {
      const rPr = (brNode['a:rPr'] || brNode['rPr'] || {}) as Record<string, unknown>;
      const explicitProps = parseRunProperties(rPr, relationshipResolver);
      const cleanExplicit = Object.fromEntries(Object.entries(explicitProps).filter(([, v]) => v !== undefined));
      const properties = {
        ...fallbackProps,
        ...cleanExplicit,
      };
      runs.push({
        break: true,
        properties,
        text: '',
      });
    }
  }

  const properties: PptxParagraphProperties = {
    alignment,
    bullet,
    firstLineIndent,
    leftMargin,
    level,
    lineSpacing,
    rightMargin,
    spaceAfter,
    spaceBefore,
  };

  return {
    properties,
    runs,
  };
}

/**
 * Parses a single OpenXML text run node (`<a:r>` or `<a:fld>`).
 *
 * Extracts text string and formatting properties (`fontSize`, `bold`, `italic`, `underline`, `strikethrough`, `color`, `fontFamily`, `hyperlink`).
 * @param rNode Raw XML text run node (`<a:r>`).
 * @param fallbackProps Optional default run properties from `<a:defRPr>`.
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Parsed `PptxRun` structure or `null` if run is empty.
 */
export function parseRun(
  rNode: Record<string, unknown>,
  fallbackProps?: PptxRun['properties'],
  relationshipResolver?: RelationshipResolver,
): null | PptxRun {
  const tText = rNode['a:t'] !== undefined ? rNode['a:t'] : rNode['t'];
  let text = '';

  if (typeof tText === 'string') {
    text = tText;
  } else if (typeof tText === 'number') {
    text = String(tText);
  } else if (tText && typeof tText === 'object' && '#text' in tText) {
    text = String((tText as Record<string, unknown>)['#text']);
  }

  const rPr = (rNode['a:rPr'] || rNode['rPr'] || {}) as Record<string, unknown>;
  const explicitProps = parseRunProperties(rPr, relationshipResolver);

  const cleanExplicit = Object.fromEntries(Object.entries(explicitProps).filter(([, v]) => v !== undefined));
  const properties = {
    ...fallbackProps,
    ...cleanExplicit,
  };

  return {
    properties,
    text,
  };
}

/**
 * Extracts run formatting properties from an XML run properties node (`<a:rPr>` or `<a:defRPr>`).
 * @param rPr Raw XML run properties object.
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Parsed `PptxRun['properties']` structure.
 */
export function parseRunProperties(
  rPr: Record<string, unknown>,
  relationshipResolver?: RelationshipResolver,
): PptxRun['properties'] {
  const fontSize = rPr['@_sz'] !== undefined ? hundredthsPoint(Number(rPr['@_sz'])) : undefined;
  const bold = rPr['@_b'] === '1' || rPr['@_b'] === true;
  const italic = rPr['@_i'] === '1' || rPr['@_i'] === true;
  const underline = rPr['@_u'] !== undefined && rPr['@_u'] !== 'none';
  const strikethrough = rPr['@_strike'] !== undefined && rPr['@_strike'] !== 'noStrike';
  const baseline = rPr['@_baseline'] !== undefined ? Number(rPr['@_baseline']) : undefined;
  const superscript = baseline !== undefined ? baseline > 0 : undefined;
  const subscript = baseline !== undefined ? baseline < 0 : undefined;

  const latin = (rPr['a:latin'] || rPr['latin']) as Record<string, unknown> | undefined;
  const fontFamily = latin && latin['@_typeface'] ? (latin['@_typeface'] as string) : undefined;

  let color: string | undefined;
  const solidFill = (rPr['a:solidFill'] || rPr['solidFill']) as Record<string, unknown> | undefined;
  if (solidFill) {
    const srgbClr = (solidFill['a:srgbClr'] || solidFill['srgbClr']) as Record<string, unknown> | undefined;
    if (srgbClr && srgbClr['@_val']) {
      color = srgbClr['@_val'] as string;
    }
  }

  const hlinkNode = (rPr['a:hlinkClick'] || rPr['hlinkClick']) as Record<string, unknown> | undefined;
  const hyperlink = hlinkNode ? parseHyperlink(hlinkNode, relationshipResolver) : undefined;

  return {
    baseline,
    bold: bold || undefined,
    color,
    fontFamily,
    fontSize,
    hyperlink,
    italic: italic || undefined,
    strikethrough: strikethrough || undefined,
    subscript: subscript || undefined,
    superscript: superscript || undefined,
    underline: underline || undefined,
  };
}

/**
 * Parses an OpenXML `<a:txBody>` or `<p:txBody>` node into a `PptxTextBody` AST structure.
 *
 * Extracts text body container properties (`verticalAlignment`, `wrap`, `insets`, `columns`, `columnSpacing`) and all child paragraphs.
 * @param txBodyNode Raw XML object node for `<p:txBody>`.
 * @param relationshipResolver Optional resolver for mapping hyperlink `r:id` references.
 * @returns Parsed `PptxTextBody` AST.
 * @example
 * ```ts
 * const textBody = parseTextBody(txBodyNode);
 * console.log(textBody.paragraphs[0].runs[0].text);
 * ```
 */
export function parseTextBody(
  txBodyNode: Record<string, unknown>,
  relationshipResolver?: RelationshipResolver,
): PptxTextBody {
  const bodyPrNode = (txBodyNode['a:bodyPr'] || txBodyNode['bodyPr'] || {}) as Record<string, unknown>;

  const vertAlignMap: Record<string, 'bottom' | 'middle' | 'top'> = {
    b: 'bottom',
    bottom: 'bottom',
    ctr: 'middle',
    middle: 'middle',
    t: 'top',
    top: 'top',
  };

  const anchorRaw = (bodyPrNode['@_anchor'] as string) || '';
  const verticalAlignment = vertAlignMap[anchorRaw];
  const wrap = (bodyPrNode['@_wrap'] as 'none' | 'square') || undefined;

  const leftInset = bodyPrNode['@_lIns'] !== undefined ? emu(Number(bodyPrNode['@_lIns'])) : undefined;
  const topInset = bodyPrNode['@_tIns'] !== undefined ? emu(Number(bodyPrNode['@_tIns'])) : undefined;
  const rightInset = bodyPrNode['@_rIns'] !== undefined ? emu(Number(bodyPrNode['@_rIns'])) : undefined;
  const bottomInset = bodyPrNode['@_bIns'] !== undefined ? emu(Number(bodyPrNode['@_bIns'])) : undefined;

  const columns = bodyPrNode['@_numCol'] !== undefined ? Number(bodyPrNode['@_numCol']) : undefined;
  const columnSpacing = bodyPrNode['@_spcCol'] !== undefined ? emu(Number(bodyPrNode['@_spcCol'])) : undefined;

  const bodyProperties: PptxTextBodyProperties = {
    bottomInset,
    columns,
    columnSpacing,
    leftInset,
    rightInset,
    topInset,
    verticalAlignment,
    wrap,
  };

  let pNodes = txBodyNode['a:p'] || txBodyNode['p'];
  const paragraphs: PptxParagraph[] = [];

  if (pNodes) {
    if (!Array.isArray(pNodes)) {
      pNodes = [pNodes];
    }
    for (const pNode of pNodes as Record<string, unknown>[]) {
      paragraphs.push(parseParagraph(pNode, relationshipResolver));
    }
  }

  return {
    bodyProperties,
    paragraphs,
  };
}

/**
 * Convenience XML entrypoint for parsing text body from XML string.
 * @param xml Raw text body XML string.
 * @param parser Optional custom `XmlParser` instance.
 * @returns Parsed `PptxTextBody` AST structure.
 */
export function parseTextBodyXml(xml: string, parser: XmlParser = defaultXmlParser): PptxTextBody {
  const parsed = parser.parse<Record<string, unknown>>(xml);
  const txBodyNode = (parsed['p:txBody'] || parsed['a:txBody'] || parsed['txBody'] || parsed) as Record<string, unknown>;
  return parseTextBody(txBodyNode);
}
