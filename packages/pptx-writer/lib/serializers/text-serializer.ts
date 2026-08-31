import type {
  PptxBullet,
  PptxColor,
  PptxFill,
  PptxHyperlink,
  PptxParagraph,
  PptxParagraphProperties,
  PptxRun,
  PptxTextBody,
  PptxTextBodyProperties,
} from '@hokkyss/pptx-core';
import {
  degreesToGradientAngle,
  sanitizeHyperlinkAction,
  sanitizeHyperlinkTooltip,
  sanitizeSlideIndex,
} from '@hokkyss/pptx-core';
import { XMLBuilder } from 'fast-xml-parser';
import { sanitizeXmlText } from '../xml/xml-builder';

/**
 * Standard Step increment per indentation level in EMU (0.5" / 36 pt).
 * In Microsoft Office PowerPoint, each tab/indent level increases the left margin by 0.5 inches.
 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties.leftmargin
 * @see https://ecma-international.org/publications-and-standards/standards/ecma-376/ ECMA-376 Part 1, Section 21.1.2.2.7 (pPr)
 */
export const DEFAULT_LEVEL_INDENT = 457200; // 0.5 inches in EMU (0.5 * 914400)

/**
 * Standard hanging indent distance for character bullets (`•`, `-`, etc.) in EMU (0.3125" / 5/16" / 22.5 pt).
 * Default hanging indent geometry used in Microsoft PowerPoint default blank templates (`Office Theme.potx`).
 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties.indent
 * @see https://ecma-international.org/publications-and-standards/standards/ecma-376/ ECMA-376 Part 1, Section 21.1.2.2.7 (pPr)
 */
export const DEFAULT_CHAR_BULLET_GAP = 285750; // 5/16 (0.3125) inches in EMU (0.3125 * 914400)

/**
 * Standard hanging indent distance for auto-numbered lists (`1.`, `10.`, etc.) in EMU (0.375" / 3/8" / 27 pt).
 * Canonical DrawingML implied schema default when `@_indent` is omitted: "-342900".
 * "If this attribute is omitted, then a value of -342900 is implied." (ECMA-376 Part 1, Section 21.1.2.2.7, p. 3219).
 * Expanded hanging indent clearance prevents multi-digit number collisions.
 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties.indent
 * @see https://ecma-international.org/publications-and-standards/standards/ecma-376/ ECMA-376 Part 1, Section 21.1.2.2.7 (pPr, page 3219)
 */
export const DEFAULT_AUTONUM_BULLET_GAP = 342900; // 3/8 (0.375) inches in EMU (0.375 * 914400)

/**
 * Standard default tab stop interval in EMU (1.0" / 72 pt).
 * Default tab interval specified by the ECMA-376 DrawingML standard.
 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties.defaulttabsize
 * @see https://ecma-international.org/publications-and-standards/standards/ecma-376/ ECMA-376 Part 1, Section 21.1.2.2.7 (pPr)
 */
export const DEFAULT_TAB_SIZE = 914400; // 1.0 inches in EMU (1.0 * 914400)


const ALIGNMENT_MAP: Record<string, string> = {
  center: 'ctr',
  distributed: 'dist',
  justified: 'just',
  left: 'l',
  right: 'r',
};

const SCHEME_COLOR_NAMES = new Set([
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'bg1',
  'bg2',
  'dk1',
  'dk2',
  'folHlink',
  'hlink',
  'lt1',
  'lt2',
  'tx1',
  'tx2',
]);

/**
 * Serializes a color into DrawingML `<a:srgbClr>` or `<a:schemeClr>`.
 */
export function serializeColorNode(
  colorInput: PptxColor | string,
  opacityOverride?: number,
): Record<string, unknown> {
  let colorVal: string;
  let colorType: 'scheme' | 'srgb' = 'srgb';
  let alphaVal: number | undefined;

  if (typeof colorInput === 'string') {
    const clean = colorInput.trim();
    if (clean.startsWith('#') || /^[0-9A-Fa-f]{6}$/.test(clean)) {
      colorVal = clean.replace(/^#/, '').toUpperCase();
      colorType = 'srgb';
    } else if (SCHEME_COLOR_NAMES.has(clean) || clean.startsWith('accent')) {
      colorVal = clean;
      colorType = 'scheme';
    } else {
      colorVal = clean.replace(/^#/, '').toUpperCase();
      colorType = 'srgb';
    }
  } else {
    colorVal = colorInput.value.replace(/^#/, '');
    colorType = colorInput.type === 'scheme' ? 'scheme' : 'srgb';
    if (colorType === 'srgb') colorVal = colorVal.toUpperCase();
    if (colorInput.alpha !== undefined) {
      alphaVal = Math.round(Number(colorInput.alpha));
    }
  }

  if (opacityOverride !== undefined) {
    alphaVal = Math.round(opacityOverride <= 1 ? opacityOverride * 100000 : opacityOverride);
  }

  const clrKey = colorType === 'scheme' ? 'a:schemeClr' : 'a:srgbClr';
  const clrNode: Record<string, unknown> = {
    '@_val': colorVal,
  };
  if (alphaVal !== undefined) {
    clrNode['a:alpha'] = { '@_val': alphaVal };
  }

  return { [clrKey]: clrNode };
}

/**
 * Serializes fill properties `<a:solidFill>`, `<a:gradFill>`, `<a:noFill>`, etc.
 */
export function serializeFill(fill?: PptxFill): Record<string, unknown> | undefined {
  if (!fill) return undefined;

  if (fill.type === 'none') {
    return { 'a:noFill': {} };
  }

  if (fill.type === 'solid' && fill.solidColor) {
    const colorNode = serializeColorNode(fill.solidColor);
    return { 'a:solidFill': colorNode };
  }

  if (fill.type === 'gradient' && fill.gradient) {
    const grad = fill.gradient;
    const gradFill: Record<string, unknown> = {};

    if (grad.flip && grad.flip !== 'none') {
      gradFill['@_flip'] = grad.flip;
    }
    if (grad.rotateWithShape !== undefined) {
      gradFill['@_rotWithShape'] = grad.rotateWithShape ? '1' : '0';
    }

    // 1. Gradient Stop List <a:gsLst>
    const stops = grad.stops || [];
    const gsList = stops.map((stop, idx) => {
      let posVal: number;
      if (typeof stop.position === 'number') {
        posVal = stop.position <= 1 ? Math.round(stop.position * 100000) : Math.round(stop.position);
      } else {
        posVal = stops.length > 1 ? Math.round((idx / (stops.length - 1)) * 100000) : 0;
      }
      posVal = Math.max(0, Math.min(100000, posVal));

      const colorNode = serializeColorNode(stop.color, stop.opacity);
      return {
        '@_pos': posVal,
        ...colorNode,
      };
    });

    gradFill['a:gsLst'] = {
      'a:gs': gsList,
    };

    // 2. Gradient Type / Direction
    if (grad.type === 'radial' || grad.type === 'path') {
      const pathNode: Record<string, unknown> = {
        '@_path': grad.type === 'radial' ? 'circle' : 'rect',
      };
      if (grad.pathBounds) {
        const fillToRect: Record<string, unknown> = {};
        if (grad.pathBounds.left !== undefined) fillToRect['@_l'] = Math.round(grad.pathBounds.left <= 1 ? grad.pathBounds.left * 100000 : grad.pathBounds.left);
        if (grad.pathBounds.top !== undefined) fillToRect['@_t'] = Math.round(grad.pathBounds.top <= 1 ? grad.pathBounds.top * 100000 : grad.pathBounds.top);
        if (grad.pathBounds.right !== undefined) fillToRect['@_r'] = Math.round(grad.pathBounds.right <= 1 ? grad.pathBounds.right * 100000 : grad.pathBounds.right);
        if (grad.pathBounds.bottom !== undefined) fillToRect['@_b'] = Math.round(grad.pathBounds.bottom <= 1 ? grad.pathBounds.bottom * 100000 : grad.pathBounds.bottom);
        pathNode['a:fillToRect'] = fillToRect;
      } else {
        pathNode['a:fillToRect'] = {
          '@_b': 50000,
          '@_l': 50000,
          '@_r': 50000,
          '@_t': 50000,
        };
      }
      gradFill['a:path'] = pathNode;
    } else {
      // Linear gradient (default)
      let angVal = 5400000; // Default 90 degrees (top to bottom)
      if (grad.angle !== undefined) {
        const numAngle = Number(grad.angle);
        if (Math.abs(numAngle) <= 360) {
          angVal = degreesToGradientAngle(numAngle);
        } else {
          angVal = Math.round(numAngle);
        }
      }
      gradFill['a:lin'] = {
        '@_ang': angVal,
        '@_scaled': '1',
      };
    }

    return { 'a:gradFill': gradFill };
  }

  return undefined;
}

const VERTICAL_ALIGNMENT_MAP: Record<string, string> = {
  bottom: 'b',
  middle: 'ctr',
  top: 't',
};

/**
 * Serializes text body properties `<a:bodyPr>`.
 */
export function serializeBodyProperties(props?: PptxTextBodyProperties): Record<string, unknown> {
  if (!props) return {};

  const bodyPr: Record<string, unknown> = {};

  if (props.verticalAlignment && props.verticalAlignment !== 'top') {
    bodyPr['@_anchor'] = VERTICAL_ALIGNMENT_MAP[props.verticalAlignment] ?? props.verticalAlignment;
  }
  if (props.wrap) {
    bodyPr['@_wrap'] = props.wrap;
  }
  if (props.leftInset !== undefined) bodyPr['@_lIns'] = Math.round(Number(props.leftInset));
  if (props.topInset !== undefined) bodyPr['@_tIns'] = Math.round(Number(props.topInset));
  if (props.rightInset !== undefined) bodyPr['@_rIns'] = Math.round(Number(props.rightInset));
  if (props.bottomInset !== undefined) bodyPr['@_bIns'] = Math.round(Number(props.bottomInset));

  return bodyPr;
}

/**
 * Serializes bullet properties inside `<a:pPr>`.
 */
export function serializeBulletProperties(bullet?: PptxBullet): Record<string, unknown> | undefined {
  if (!bullet) return undefined;

  if (bullet.type === 'none') {
    return { 'a:buNone': {} };
  }

  if (bullet.type === 'char' && bullet.char) {
    const isStandardBullet = bullet.char === '•' || bullet.char === '&#8226;' || bullet.char === '\u2022';
    const font = bullet.fontFamily || (isStandardBullet ? 'Arial' : '+mn-lt');
    const buFont: Record<string, unknown> = { '@_typeface': font };
    if (font === 'Arial') {
      buFont['@_panose'] = '020B0604020202020204';
      buFont['@_pitchFamily'] = '34';
      buFont['@_charset'] = '0';
    }
    return {
      'a:buFont': buFont,
      'a:buChar': { '@_char': bullet.char },
    };
  }

  if (bullet.type === 'autoNum') {
    return {
      'a:buAutoNum': {
        '@_type': bullet.autoNumType ?? 'arabicPeriod',
        ...(bullet.startAt ? { '@_startAt': bullet.startAt } : {}),
      },
    };
  }

  return undefined;
}

/**
 * Serializes an OpenXML DrawingML Hyperlink `<a:hlinkClick>`.
 * @param hyperlink Hyperlink configuration or URL string.
 * @param relIdOverride Optional relationship ID override.
 * @returns Serialized XML node object or `undefined`.
 */
export function serializeHyperlink(
  hyperlink?: PptxHyperlink | string,
  relIdOverride?: string,
): Record<string, unknown> | undefined {
  if (!hyperlink) return undefined;

  const hlinkNode: Record<string, unknown> = {};

  if (typeof hyperlink === 'string') {
    const rId = relIdOverride;
    if (rId) {
      hlinkNode['@_r:id'] = rId;
    }
    return Object.keys(hlinkNode).length > 0 ? hlinkNode : undefined;
  }

  const rId = relIdOverride || hyperlink.rId;
  if (rId) {
    hlinkNode['@_r:id'] = rId;
  }

  const cleanTooltip = sanitizeHyperlinkTooltip(hyperlink.tooltip);
  if (cleanTooltip) {
    hlinkNode['@_tooltip'] = cleanTooltip;
  }

  const cleanAction = sanitizeHyperlinkAction(hyperlink.action);
  if (cleanAction) {
    switch (cleanAction) {
      case 'endShow':
        hlinkNode['@_action'] = 'ppaction://hlinkshowjump?jump=endshow';
        break;
      case 'firstSlide':
        hlinkNode['@_action'] = 'ppaction://hlinkshowjump?jump=firstslide';
        break;
      case 'lastSlide':
        hlinkNode['@_action'] = 'ppaction://hlinkshowjump?jump=lastslide';
        break;
      case 'nextSlide':
        hlinkNode['@_action'] = 'ppaction://hlinkshowjump?jump=nextslide';
        break;
      case 'previousSlide':
        hlinkNode['@_action'] = 'ppaction://hlinkshowjump?jump=previousslide';
        break;
      default:
        hlinkNode['@_action'] = cleanAction;
    }
  } else {
    const cleanSlideIndex = sanitizeSlideIndex(hyperlink.slideIndex);
    if (cleanSlideIndex && !hlinkNode['@_action']) {
      hlinkNode['@_action'] = 'ppaction://hlinksldjump';
    }
  }

  return Object.keys(hlinkNode).length > 0 ? hlinkNode : undefined;
}

/**
 * Serializes run properties `<a:rPr>`.
 */
export function serializeRunProperties(props?: PptxRun['properties']): Record<string, unknown> {
  const rPr: Record<string, unknown> = {};

  if (!props) return rPr;

  if (props.bold) rPr['@_b'] = '1';
  if (props.italic) rPr['@_i'] = '1';
  if (props.underline) rPr['@_u'] = typeof props.underline === 'string' ? props.underline : 'sng';
  if (props.strikethrough) rPr['@_strike'] = typeof props.strikethrough === 'string' ? props.strikethrough : 'sngStrike';
  if (props.baseline !== undefined) {
    rPr['@_baseline'] = Math.round(Number(props.baseline));
  } else if (props.superscript) {
    rPr['@_baseline'] = '30000';
  } else if (props.subscript) {
    rPr['@_baseline'] = '-25000';
  }
  if (props.fontSize !== undefined) rPr['@_sz'] = Math.round(Number(props.fontSize));

  if (props.color) {
    if (typeof props.color === 'string') {
      rPr['a:solidFill'] = {
        'a:srgbClr': { '@_val': props.color.replace(/^#/, '').toUpperCase() },
      };
    } else {
      const fillNode = serializeFill({ solidColor: props.color, type: 'solid' });
      if (fillNode) {
        Object.assign(rPr, fillNode);
      }
    }
  }

  if (props.hyperlink) {
    const hlinkNode = serializeHyperlink(props.hyperlink);
    if (hlinkNode) {
      rPr['a:hlinkClick'] = hlinkNode;
    }
  }

  const typeface = props.fontFamily || '+mn-lt';
  rPr['a:latin'] = { '@_typeface': typeface };
  rPr['a:cs'] = { '@_typeface': typeface };

  return rPr;
}


/**
 * Content builder: entity-aware XML builder used for building individual element nodes
 * (`<a:r>`, `<a:br>`, `<a:pPr>`, etc.) where text content and attribute values must have
 * `&`, `<`, `>` properly escaped. Used for the raw-XML paragraph path.
 */
const _contentBuilder = new XMLBuilder({
  attributeNamePrefix: '@_',
  format: false,
  ignoreAttributes: false,
  suppressBooleanAttributes: false,
  suppressEmptyNode: true,
  textNodeName: '#text',
});

/**
 * Wrap builder: used ONLY to inject a pre-built XML string as raw content into a parent
 * tag (e.g. injecting the pre-built txBody content into `<p:txBody>...</p:txBody>`).
 * processEntities: false prevents double-encoding of already-escaped XML strings.
 */
const _wrapBuilder = new XMLBuilder({
  attributeNamePrefix: '@_',
  format: false,
  ignoreAttributes: false,
  processEntities: false,
  suppressBooleanAttributes: false,
  suppressEmptyNode: true,
  textNodeName: '#text',
});

/**
 * Builds the `<a:pPr>` XML string for a paragraph's properties object.
 * Used internally by the raw-XML paragraph path.
 */
function buildPPrXml(pPr: Record<string, unknown>): string {
  if (Object.keys(pPr).length === 0) return '';
  return _contentBuilder.build({ 'a:pPr': pPr }) as string;
}

/**
 * Builds the XML for a single run node — either `<a:r>` or `<a:br>`.
 * Uses `_contentBuilder` (entity-aware) so text content like `&` and `<` is correctly escaped.
 * Used internally by the raw-XML paragraph path.
 */
function buildRunXml(run: PptxRun): string {
  if (run.break === true) {
    const brNode: Record<string, unknown> = {};
    if (run.properties && Object.keys(run.properties).length > 0) {
      const rPr = serializeRunProperties(run.properties);
      if (Object.keys(rPr).length > 0) {
        brNode['a:rPr'] = rPr;
      }
    }
    return _contentBuilder.build({ 'a:br': brNode }) as string;
  }

  const rNode: Record<string, unknown> = {};
  if (run.properties) {
    const rPr = serializeRunProperties(run.properties);
    if (Object.keys(rPr).length > 0) {
      rNode['a:rPr'] = rPr;
    }
  }
  rNode['a:t'] = sanitizeXmlText((run.text ?? '').replace(/[\r\n]+/g, ' '));
  return _contentBuilder.build({ 'a:r': rNode }) as string;
}


/**
 * Computes the shared `<a:pPr>` attributes and bullet properties for a paragraph,
 * returning a plain pPr object. Used by both the object-path and raw-XML-path.
 */
function buildPPrObject(paragraph: PptxParagraph): Record<string, unknown> {
  const pPr: Record<string, unknown> = {};
  const props = (paragraph.properties || paragraph) as { margin?: number; indent?: number } & PptxParagraphProperties;

  if (props.alignment && props.alignment !== 'left') {
    pPr['@_algn'] = ALIGNMENT_MAP[props.alignment] ?? props.alignment;
  }
  if (props.leftMargin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.leftMargin));
  } else if (props.margin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.margin));
  }
  if (props.level !== undefined && props.level > 0) {
    pPr['@_lvl'] = props.level;
  }
  if (props.firstLineIndent !== undefined) {
    pPr['@_indent'] = Math.round(Number(props.firstLineIndent));
  } else if (props.indent !== undefined) {
    pPr['@_indent'] = Math.round(Number(props.indent));
  }

  if (props.bullet) {
    const bulletNode = serializeBulletProperties(props.bullet);
    if (bulletNode) {
      Object.assign(pPr, bulletNode);
    }
  }

  return pPr;
}

/**
 * Serializes paragraph `<a:p>`.
 * Follows schema order: a:pPr -> a:r / a:br (with a:rPr then a:t) -> a:endParaRPr
 *
 * Returns a plain object for pure-text-run paragraphs (fast path), or a raw XML string
 * for paragraphs containing line breaks (`{ break: true }` runs) to preserve exact element order.
 */
export function serializeParagraph(paragraph: PptxParagraph): Record<string, unknown> | string {
  const runs = paragraph.runs || [];
  const hasBreaks = runs.some((r) => r.break === true);

  const pPr = buildPPrObject(paragraph);

  if (!hasBreaks) {
    // Fast path: only text runs — return a plain JS object for fast-xml-parser to handle
    const textRuns = runs.map((run: PptxRun) => {
      const rNode: Record<string, unknown> = {};
      if (run.properties) {
        const rPr = serializeRunProperties(run.properties);
        if (Object.keys(rPr).length > 0) {
          rNode['a:rPr'] = rPr;
        }
      }
      // Strict OpenXML compliance: <a:t> must not have raw newlines or invalid XML 1.0 control characters
      rNode['a:t'] = sanitizeXmlText((run.text ?? '').replace(/[\r\n]+/g, ' '));
      return rNode;
    });

    const pNode: Record<string, unknown> = {};
    if (Object.keys(pPr).length > 0) {
      pNode['a:pPr'] = pPr;
    }
    if (textRuns.length > 0) {
      pNode['a:r'] = textRuns;
    } else {
      pNode['a:endParaRPr'] = {};
    }
    return pNode;
  }

  // Raw-XML path: paragraph has line breaks — build raw XML to preserve interleaved element order
  const pPrXml = buildPPrXml(pPr);
  const runsXml = runs.map(buildRunXml).join('');
  return `<a:p>${pPrXml}${runsXml}</a:p>`;
}

/**
 * Serializes complete `<p:txBody>` or `<a:txBody>`.
 *
 * When all paragraphs are plain objects, returns a structured JS object (default fast path).
 * When any paragraph contains line breaks, falls back to a raw XML string approach to
 * correctly preserve the interleaved `<a:r>` / `<a:br>` element order in the output.
 * In this case, the returned string contains only the INNER content (without `<p:txBody>` wrapper).
 * Callers that assign this to an object key must use `_wrapBuilder` (processEntities: false)
 * when building the final shape XML, which `shape-serializer.ts` does automatically via the
 * `serializeShapeWithTextBody` helper.
 */
export function serializeTextBody(textBody: PptxTextBody): Record<string, unknown> | string {
  const bodyPr = serializeBodyProperties(textBody.bodyProperties);
  const paragraphs = (textBody.paragraphs || []).map(serializeParagraph);

  const hasRawParagraphs = paragraphs.some((p) => typeof p === 'string');

  if (!hasRawParagraphs) {
    // Fast path: return a plain JS object
    return {
      'a:bodyPr': Object.keys(bodyPr).length > 0 ? bodyPr : {},
      'a:lstStyle': {},
      'a:p': (paragraphs as Record<string, unknown>[]).length > 0
        ? paragraphs as Record<string, unknown>[]
        : [{ 'a:pPr': {}, 'a:endParaRPr': {} }],
    };
  }

  // Raw-XML path: at least one paragraph has a line break — build the entire txBody inner content as raw XML.
  // Use _contentBuilder (entity-aware) for bodyPr and plain paragraphs, since their content may have
  // text values that need entity encoding. The pre-built paragraph strings (returned by serializeParagraph)
  // are already correctly encoded, so they are concatenated as-is.
  const bodyPrXml = _contentBuilder.build({ 'a:bodyPr': Object.keys(bodyPr).length > 0 ? bodyPr : {} }) as string;
  const lstStyleXml = '<a:lstStyle/>';

  const paragraphsXml = paragraphs.length > 0
    ? paragraphs.map((p) => (typeof p === 'string' ? p : (_contentBuilder.build({ 'a:p': p }) as string))).join('')
    : '<a:p><a:pPr/><a:endParaRPr/></a:p>';

  return `${bodyPrXml}${lstStyleXml}${paragraphsXml}`;
}

/**
 * Re-exported `_wrapBuilder` for use by callers that need to inject `serializeTextBody`'s
 * raw XML string result into a parent object (e.g. `shape-serializer.ts`).
 * This builder has `processEntities: false` so it does not double-encode pre-built XML strings.
 * @internal
 */
export { _wrapBuilder as _rawXmlWrapBuilder };
