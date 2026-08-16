import {
  Emu,
  HundredthsPoint,
  PptxBullet,
  PptxParagraph,
  PptxParagraphProperties,
  PptxRun,
  PptxTextBody,
  PptxTextBodyProperties,
} from '../types/ast';
import { defaultXmlParser, XmlParser } from '../xml/xml-parser';

/**
 * Parses a single OpenXML paragraph node (`<a:p>`).
 *
 * Extracts paragraph properties (`alignment`, `leftMargin`, `rightMargin`, `firstLineIndent`, `spaceBefore`, `spaceAfter`, `lineSpacing`, `bullet`) and text runs.
 * @param pNode Raw XML paragraph object node (`<a:p>`).
 * @returns Parsed `PptxParagraph` structure.
 */
export function parseParagraph(pNode: Record<string, unknown>): PptxParagraph {
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
  const level = pPr['@_lvl'] !== undefined ? Number(pPr['@_lvl']) : undefined;

  const leftMargin = pPr['@_marL'] !== undefined ? ((Number(pPr['@_marL']) as unknown) as Emu) : undefined;
  const rightMargin = pPr['@_marR'] !== undefined ? ((Number(pPr['@_marR']) as unknown) as Emu) : undefined;
  const firstLineIndent = pPr['@_indent'] !== undefined ? ((Number(pPr['@_indent']) as unknown) as Emu) : undefined;

  // Space before paragraph (<a:spcBef>)
  let spaceBefore: HundredthsPoint | undefined;
  const spcBefNode = (pPr['a:spcBef'] || pPr['spcBef']) as Record<string, unknown> | undefined;
  if (spcBefNode) {
    const befPts = (spcBefNode['a:spcPts'] || spcBefNode['spcPts']) as Record<string, unknown> | undefined;
    if (befPts && befPts['@_val'] !== undefined) {
      spaceBefore = (Number(befPts['@_val']) as unknown) as HundredthsPoint;
    }
  }

  // Space after paragraph (<a:spcAft>)
  let spaceAfter: HundredthsPoint | undefined;
  const spcAftNode = (pPr['a:spcAft'] || pPr['spcAft']) as Record<string, unknown> | undefined;
  if (spcAftNode) {
    const aftPts = (spcAftNode['a:spcPts'] || spcAftNode['spcPts']) as Record<string, unknown> | undefined;
    if (aftPts && aftPts['@_val'] !== undefined) {
      spaceAfter = (Number(aftPts['@_val']) as unknown) as HundredthsPoint;
    }
  }

  // Line spacing (<a:lnSpc>)
  let lineSpacing: HundredthsPoint | undefined;
  const lnSpcNode = (pPr['a:lnSpc'] || pPr['lnSpc']) as Record<string, unknown> | undefined;
  if (lnSpcNode) {
    const spcPct = (lnSpcNode['a:spcPct'] || lnSpcNode['spcPct'] || lnSpcNode['a:spcPts'] || lnSpcNode['spcPts']) as Record<string, unknown> | undefined;
    if (spcPct && spcPct['@_val'] !== undefined) {
      lineSpacing = (Number(spcPct['@_val']) as unknown) as HundredthsPoint;
    }
  }

  // Extract bullet point properties
  let bullet: PptxBullet | undefined;
  const buChar = (pPr['a:buChar'] || pPr['buChar']) as Record<string, unknown> | undefined;
  const buAutoNum = (pPr['a:buAutoNum'] || pPr['buAutoNum']) as Record<string, unknown> | undefined;
  const buNone = pPr['a:buNone'] || pPr['buNone'];

  if (buNone) {
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
      const run = parseRun(rNode, fallbackProps);
      if (run) runs.push(run);
    }
  }

  const fldNodes = pNode['a:fld'] || pNode['fld'];
  if (fldNodes) {
    const flds = Array.isArray(fldNodes) ? fldNodes : [fldNodes];
    for (const fld of flds as Record<string, unknown>[]) {
      const run = parseRun(fld, fallbackProps);
      if (run) runs.push(run);
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
 * Extracts text string and formatting properties (`fontSize`, `bold`, `italic`, `underline`, `strikethrough`, `color`, `fontFamily`).
 * @param rNode Raw XML text run node (`<a:r>`).
 * @param fallbackProps Optional default run properties from `<a:defRPr>`.
 * @returns Parsed `PptxRun` structure or `null` if run is empty.
 */
export function parseRun(rNode: Record<string, unknown>, fallbackProps?: PptxRun['properties']): null | PptxRun {
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
  const explicitProps = parseRunProperties(rPr);

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
 * @returns Parsed `PptxRun['properties']` structure.
 */
export function parseRunProperties(rPr: Record<string, unknown>): PptxRun['properties'] {
  const fontSize = rPr['@_sz'] !== undefined ? ((Number(rPr['@_sz']) as unknown) as HundredthsPoint) : undefined;
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

  return {
    baseline,
    bold: bold || undefined,
    color,
    fontFamily,
    fontSize,
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
 * @returns Parsed `PptxTextBody` AST.
 * @example
 * ```ts
 * const textBody = parseTextBody(txBodyNode);
 * console.log(textBody.paragraphs[0].runs[0].text);
 * ```
 */
export function parseTextBody(txBodyNode: Record<string, unknown>): PptxTextBody {
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

  const leftInset = bodyPrNode['@_lIns'] !== undefined ? ((Number(bodyPrNode['@_lIns']) as unknown) as Emu) : undefined;
  const topInset = bodyPrNode['@_tIns'] !== undefined ? ((Number(bodyPrNode['@_tIns']) as unknown) as Emu) : undefined;
  const rightInset = bodyPrNode['@_rIns'] !== undefined ? ((Number(bodyPrNode['@_rIns']) as unknown) as Emu) : undefined;
  const bottomInset = bodyPrNode['@_bIns'] !== undefined ? ((Number(bodyPrNode['@_bIns']) as unknown) as Emu) : undefined;

  const columns = bodyPrNode['@_numCol'] !== undefined ? Number(bodyPrNode['@_numCol']) : undefined;
  const columnSpacing = bodyPrNode['@_spcCol'] !== undefined ? ((Number(bodyPrNode['@_spcCol']) as unknown) as Emu) : undefined;

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
      paragraphs.push(parseParagraph(pNode));
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
