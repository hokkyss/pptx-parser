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
import { el, sanitizeXmlText, type XmlElement } from '../xml/xml-element';

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
): XmlElement {
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

  const clrTag = colorType === 'scheme' ? 'a:schemeClr' : 'a:srgbClr';
  const children = alphaVal !== undefined ? [el('a:alpha', { val: alphaVal })] : undefined;
  return el(clrTag, { val: colorVal }, children);
}

/**
 * Serializes fill properties `<a:solidFill>`, `<a:gradFill>`, `<a:noFill>`, etc.
 */
export function serializeFill(fill?: PptxFill): undefined | XmlElement {
  if (!fill) return undefined;

  if (fill.type === 'none') {
    return el('a:noFill');
  }

  if (fill.type === 'solid' && fill.solidColor) {
    const colorNode = serializeColorNode(fill.solidColor);
    return el('a:solidFill', [colorNode]);
  }

  if (fill.type === 'gradient' && fill.gradient) {
    const grad = fill.gradient;
    const gradAttrs: Record<string, number | string | undefined> = {};

    if (grad.flip && grad.flip !== 'none') {
      gradAttrs.flip = grad.flip;
    }
    if (grad.rotateWithShape !== undefined) {
      gradAttrs.rotWithShape = grad.rotateWithShape ? '1' : '0';
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
      return el('a:gs', { pos: posVal }, [colorNode]);
    });

    const gsLstNode = el('a:gsLst', gsList);

    // 2. Gradient Type / Direction
    let directionNode: XmlElement;
    if (grad.type === 'radial' || grad.type === 'path') {
      const fillToRectAttrs: Record<string, number> = {};
      if (grad.pathBounds) {
        if (grad.pathBounds.left !== undefined) fillToRectAttrs.l = Math.round(grad.pathBounds.left <= 1 ? grad.pathBounds.left * 100000 : grad.pathBounds.left);
        if (grad.pathBounds.top !== undefined) fillToRectAttrs.t = Math.round(grad.pathBounds.top <= 1 ? grad.pathBounds.top * 100000 : grad.pathBounds.top);
        if (grad.pathBounds.right !== undefined) fillToRectAttrs.r = Math.round(grad.pathBounds.right <= 1 ? grad.pathBounds.right * 100000 : grad.pathBounds.right);
        if (grad.pathBounds.bottom !== undefined) fillToRectAttrs.b = Math.round(grad.pathBounds.bottom <= 1 ? grad.pathBounds.bottom * 100000 : grad.pathBounds.bottom);
      } else {
        fillToRectAttrs.b = 50000;
        fillToRectAttrs.l = 50000;
        fillToRectAttrs.r = 50000;
        fillToRectAttrs.t = 50000;
      }
      directionNode = el('a:path', { path: grad.type === 'radial' ? 'circle' : 'rect' }, [
        el('a:fillToRect', fillToRectAttrs),
      ]);
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
      directionNode = el('a:lin', { ang: angVal, scaled: '1' });
    }

    return el('a:gradFill', gradAttrs, [gsLstNode, directionNode]);
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
export function serializeBodyProperties(props?: PptxTextBodyProperties): XmlElement {
  const attrs: Record<string, number | string | undefined> = {};

  if (props) {
    if (props.verticalAlignment) {
      attrs.anchor = VERTICAL_ALIGNMENT_MAP[props.verticalAlignment] ?? props.verticalAlignment;
    }
    if (props.wrap) {
      attrs.wrap = props.wrap;
    }
    if (props.leftInset !== undefined) attrs.lIns = Math.round(Number(props.leftInset));
    if (props.topInset !== undefined) attrs.tIns = Math.round(Number(props.topInset));
    if (props.rightInset !== undefined) attrs.rIns = Math.round(Number(props.rightInset));
    if (props.bottomInset !== undefined) attrs.bIns = Math.round(Number(props.bottomInset));
  }

  return el('a:bodyPr', attrs);
}

/**
 * Serializes bullet properties inside `<a:pPr>`.
 */
export function serializeBulletProperties(bullet?: PptxBullet): undefined | XmlElement {
  if (!bullet) return undefined;

  if (bullet.type === 'none') {
    return el('a:buNone');
  }

  if (bullet.type === 'char' && bullet.char) {
    return el('a:buChar', { char: bullet.char });
  }

  if (bullet.type === 'autoNum') {
    const attrs: Record<string, number | string | undefined> = {
      type: bullet.autoNumType ?? 'arabicPeriod',
    };
    if (bullet.startAt) {
      attrs.startAt = bullet.startAt;
    }
    return el('a:buAutoNum', attrs);
  }

  return undefined;
}

/**
 * Serializes an OpenXML DrawingML Hyperlink `<a:hlinkClick>`.
 * @param hyperlink Hyperlink configuration or URL string.
 * @param relIdOverride Optional relationship ID override.
 * @returns Serialized XmlElement or `undefined`.
 */
export function serializeHyperlink(
  hyperlink?: PptxHyperlink | string,
  relIdOverride?: string,
): undefined | XmlElement {
  if (!hyperlink) return undefined;

  const attrs: Record<string, string | undefined> = {};

  if (typeof hyperlink === 'string') {
    const rId = relIdOverride;
    if (rId) {
      attrs['r:id'] = rId;
    }
    return Object.keys(attrs).length > 0 ? el('a:hlinkClick', attrs) : undefined;
  }

  const rId = relIdOverride || hyperlink.rId;
  if (rId) {
    attrs['r:id'] = rId;
  }

  const cleanTooltip = sanitizeHyperlinkTooltip(hyperlink.tooltip);
  if (cleanTooltip) {
    attrs.tooltip = cleanTooltip;
  }

  const cleanAction = sanitizeHyperlinkAction(hyperlink.action);
  if (cleanAction) {
    switch (cleanAction) {
      case 'endShow':
        attrs.action = 'ppaction://hlinkshowjump?jump=endshow';
        break;
      case 'firstSlide':
        attrs.action = 'ppaction://hlinkshowjump?jump=firstslide';
        break;
      case 'lastSlide':
        attrs.action = 'ppaction://hlinkshowjump?jump=lastslide';
        break;
      case 'nextSlide':
        attrs.action = 'ppaction://hlinkshowjump?jump=nextslide';
        break;
      case 'previousSlide':
        attrs.action = 'ppaction://hlinkshowjump?jump=previousslide';
        break;
      default:
        attrs.action = cleanAction;
    }
  } else {
    const cleanSlideIndex = sanitizeSlideIndex(hyperlink.slideIndex);
    if (cleanSlideIndex && !attrs.action) {
      attrs.action = 'ppaction://hlinksldjump';
    }
  }

  return Object.keys(attrs).length > 0 ? el('a:hlinkClick', attrs) : undefined;
}

/**
 * Serializes run properties `<a:rPr>`.
 */
export function serializeRunProperties(props?: PptxRun['properties']): undefined | XmlElement {
  if (!props) return undefined;

  const attrs: Record<string, number | string | undefined> = {};

  if (props.bold) attrs.b = '1';
  if (props.italic) attrs.i = '1';
  if (props.underline) attrs.u = typeof props.underline === 'string' ? props.underline : 'sng';
  if (props.strikethrough) attrs.strike = typeof props.strikethrough === 'string' ? props.strikethrough : 'sngStrike';
  if (props.baseline !== undefined) {
    attrs.baseline = Math.round(Number(props.baseline));
  } else if (props.superscript) {
    attrs.baseline = '30000';
  } else if (props.subscript) {
    attrs.baseline = '-25000';
  }
  if (props.fontSize !== undefined) attrs.sz = Math.round(Number(props.fontSize));

  const children: (undefined | XmlElement)[] = [];

  if (props.color) {
    if (typeof props.color === 'string') {
      children.push(el('a:solidFill', [
        serializeColorNode(props.color),
      ]));
    } else {
      const fillNode = serializeFill({ solidColor: props.color, type: 'solid' });
      if (fillNode) {
        children.push(fillNode);
      }
    }
  }

  if (props.hyperlink) {
    const hlinkNode = serializeHyperlink(props.hyperlink);
    if (hlinkNode) {
      children.push(hlinkNode);
    }
  }

  const typeface = props.fontFamily || '+mn-lt';
  children.push(el('a:latin', { typeface }));
  children.push(el('a:cs', { typeface }));

  return el('a:rPr', attrs, children.filter(Boolean));
}

/**
 * Serializes paragraph `<a:p>`.
 * Follows strict schema sequence: a:pPr -> interleaved (a:r | a:br) -> a:endParaRPr (if empty).
 */
export function serializeParagraph(paragraph: PptxParagraph): XmlElement {
  const pPrAttrs: Record<string, number | string | undefined> = {};
  const props = (paragraph.properties || paragraph) as { margin?: number; indent?: number } & PptxParagraphProperties;

  if (props.alignment) {
    pPrAttrs.algn = ALIGNMENT_MAP[props.alignment] ?? props.alignment;
  }
  if (props.level !== undefined) {
    pPrAttrs.lvl = props.level;
  }
  if (props.leftMargin !== undefined) {
    pPrAttrs.marL = Math.round(Number(props.leftMargin));
  } else if (props.margin !== undefined) {
    pPrAttrs.marL = Math.round(Number(props.margin));
  }
  if (props.firstLineIndent !== undefined) {
    pPrAttrs.indent = Math.round(Number(props.firstLineIndent));
  } else if (props.indent !== undefined) {
    pPrAttrs.indent = Math.round(Number(props.indent));
  }

  const pPrChildren: XmlElement[] = [];

  if (props.bullet) {
    const bulletNode = serializeBulletProperties(props.bullet);
    if (bulletNode) {
      pPrChildren.push(bulletNode);
    }
    if (props.bullet.type !== 'none' && pPrAttrs.marL === undefined && pPrAttrs.indent === undefined) {
      const lvl = props.level ?? 0;
      const isNumbering = props.bullet.type === 'autoNum';
      // Numbered lists ("1.") use ~16pt (203200 EMU); single char bullets ("•") use ~12pt (152400 EMU)
      const bulletGap = isNumbering ? 203200 : 152400;
      const levelIndent = 228600; // 0.25 in per nested indentation level
      pPrAttrs.marL = (lvl * levelIndent) + bulletGap;
      pPrAttrs.indent = -bulletGap;
    }
  }

  const pPrNode = (Object.keys(pPrAttrs).length > 0 || pPrChildren.length > 0)
    ? el('a:pPr', pPrAttrs, pPrChildren)
    : undefined;

  const runAndBreakNodes: XmlElement[] = [];

  for (const run of (paragraph.runs || [])) {
    const rPr = serializeRunProperties(run.properties);

    if (run.break) {
      runAndBreakNodes.push(el('a:br', rPr ? [rPr] : []));
      continue;
    }

    const text = run.text ?? '';
    // Handle potential newlines in text by generating interleaved <a:r> and <a:br>
    if (text.includes('\n')) {
      const lines = text.split(/\r?\n/);
      for (let i = 0; i < lines.length; i++) {
        if (i > 0) {
          runAndBreakNodes.push(el('a:br', rPr ? [rPr] : []));
        }
        if (lines[i].length > 0 || lines.length === 1) {
          runAndBreakNodes.push(
            el('a:r', [
              rPr,
              el('a:t', sanitizeXmlText(lines[i])),
            ].filter(Boolean)),
          );
        }
      }
    } else {
      runAndBreakNodes.push(
        el('a:r', [
          rPr,
          el('a:t', sanitizeXmlText(text)),
        ].filter(Boolean)),
      );
    }
  }

  const pChildren: (undefined | XmlElement)[] = [];
  if (pPrNode) {
    pChildren.push(pPrNode);
  }
  pChildren.push(...runAndBreakNodes);
  if (runAndBreakNodes.length === 0) {
    pChildren.push(el('a:endParaRPr'));
  }

  return el('a:p', pChildren.filter(Boolean));
}

/**
 * Serializes complete `<p:txBody>` or `<a:txBody>`.
 */
export function serializeTextBody(textBody: PptxTextBody, tag: string = 'p:txBody'): XmlElement {
  const bodyPrNode = serializeBodyProperties(textBody.bodyProperties);
  const paragraphs = (textBody.paragraphs && textBody.paragraphs.length > 0)
    ? textBody.paragraphs.map(serializeParagraph)
    : [el('a:p', [el('a:pPr'), el('a:endParaRPr')])];

  return el(tag, [
    bodyPrNode,
    el('a:lstStyle'),
    ...paragraphs,
  ]);
}
