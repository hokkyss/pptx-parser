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
import { sanitizeXmlText } from '../xml/xml-builder';

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

  if (props.verticalAlignment) {
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
    return { 'a:buChar': { '@_char': bullet.char } };
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
 * Serializes paragraph `<a:p>`.
 * Follows schema order: a:pPr -> a:r (with a:rPr then a:t) -> a:endParaRPr
 */
export function serializeParagraph(paragraph: PptxParagraph): Record<string, unknown> {
  const pPr: Record<string, unknown> = {};
  const props = (paragraph.properties || paragraph) as { margin?: number; indent?: number } & PptxParagraphProperties;

  if (props.alignment) {
    pPr['@_algn'] = ALIGNMENT_MAP[props.alignment] ?? props.alignment;
  }
  if (props.level !== undefined) {
    pPr['@_lvl'] = props.level;
  }
  if (props.leftMargin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.leftMargin));
  } else if (props.margin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.margin));
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
    if (props.bullet.type !== 'none' && pPr['@_marL'] === undefined && pPr['@_indent'] === undefined) {
      const lvl = props.level ?? 0;
      const isNumbering = props.bullet.type === 'autoNum';
      // Numbered lists ("1.") use ~16pt (203200 EMU); single char bullets ("•") use ~12pt (152400 EMU)
      const bulletGap = isNumbering ? 203200 : 152400;
      const levelIndent = 228600; // 0.25 in per nested indentation level
      pPr['@_marL'] = (lvl * levelIndent) + bulletGap;
      pPr['@_indent'] = -bulletGap;
    }
  }

  const runs = (paragraph.runs || []).map((run: PptxRun) => {
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

  if (runs.length > 0) {
    pNode['a:r'] = runs;
  } else {
    pNode['a:endParaRPr'] = {};
  }

  return pNode;
}

/**
 * Serializes complete `<p:txBody>` or `<a:txBody>`.
 */
export function serializeTextBody(textBody: PptxTextBody): Record<string, unknown> {
  const bodyPr = serializeBodyProperties(textBody.bodyProperties);
  const paragraphs = (textBody.paragraphs || []).map(serializeParagraph);

  return {
    'a:bodyPr': Object.keys(bodyPr).length > 0 ? bodyPr : {},
    'a:lstStyle': {},
    'a:p': paragraphs.length > 0 ? paragraphs : [{ 'a:pPr': {}, 'a:endParaRPr': {} }],
  };
}
