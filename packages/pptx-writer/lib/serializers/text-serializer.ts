import type { PptxBullet, PptxParagraph, PptxRun, PptxTextBody, PptxTextBodyProperties } from '@hokkyss/pptx-core';
import type { PptxColor, PptxFill } from '@hokkyss/pptx-core';

const ALIGNMENT_MAP: Record<string, string> = {
  center: 'ctr',
  distributed: 'dist',
  justified: 'just',
  left: 'l',
  right: 'r',
};

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
 * Serializes fill properties `<a:solidFill>`, `<a:gradFill>`, `<a:noFill>`, etc.
 */
export function serializeFill(fill?: PptxFill): Record<string, unknown> | undefined {
  if (!fill) return undefined;

  if (fill.type === 'none') {
    return { 'a:noFill': {} };
  }

  if (fill.type === 'solid' && fill.solidColor) {
    const color: PptxColor = fill.solidColor;
    if (color.type === 'srgb') {
      const srgbClr: Record<string, unknown> = {
        '@_val': color.value.replace(/^#/, '').toUpperCase(),
      };
      if (color.alpha !== undefined) {
        srgbClr['a:alpha'] = { '@_val': Math.round(Number(color.alpha)) };
      }
      return { 'a:solidFill': { 'a:srgbClr': srgbClr } };
    }

    if (color.type === 'scheme') {
      const schemeClr: Record<string, unknown> = {
        '@_val': color.value,
      };
      if (color.alpha !== undefined) {
        schemeClr['a:alpha'] = { '@_val': Math.round(Number(color.alpha)) };
      }
      return { 'a:solidFill': { 'a:schemeClr': schemeClr } };
    }
  }

  return undefined;
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
  const props = paragraph.properties || (paragraph as unknown as { alignment?: string; bullet?: PptxBullet; indent?: number; level?: number; margin?: number });

  if (props.alignment) {
    pPr['@_algn'] = ALIGNMENT_MAP[props.alignment] ?? props.alignment;
  }
  if (props.level !== undefined) {
    pPr['@_lvl'] = props.level;
  }
  if ('leftMargin' in props && props.leftMargin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.leftMargin));
  } else if ('margin' in props && props.margin !== undefined) {
    pPr['@_marL'] = Math.round(Number(props.margin));
  }
  if ('firstLineIndent' in props && props.firstLineIndent !== undefined) {
    pPr['@_indent'] = Math.round(Number(props.firstLineIndent));
  } else if ('indent' in props && props.indent !== undefined) {
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
    // Strict OpenXML compliance: <a:t> must not have raw newlines
    rNode['a:t'] = (run.text ?? '').replace(/[\r\n]+/g, ' ');
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
