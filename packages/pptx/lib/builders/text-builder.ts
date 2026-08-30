import type { PptxBullet, PptxHyperlink, PptxParagraph, PptxRun, PptxTextBody } from '@hokkyss/pptx-core';
import { hundredthsPoint, type Points } from '@hokkyss/pptx-core';

export interface TextRunConfig {
  baseline?: number;
  bold?: boolean;
  /**
   * When `true`, this entry represents a soft line break (Shift+Enter in PowerPoint).
   * Serialized as `<a:br>`. The `text` field is not required and is ignored.
   * Optional run properties (bold, italic, etc.) are forwarded to `<a:rPr>` inside `<a:br>`.
   */
  break?: true;
  color?: string; // Hex string e.g. '38BDF8'
  font?: string;
  fontSize?: Points;
  hyperlink?: PptxHyperlink | string;
  italic?: boolean;
  strikethrough?: 'dblStrike' | 'sngStrike' | boolean;
  subscript?: boolean;
  superscript?: boolean;
  text?: string;
  underline?: 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave' | boolean;
}

// NOTE: Intended
/* eslint-disable @typescript-eslint/no-redundant-type-constituents */
export type BulletInput = 'bullet'
  | 'none'
  | 'number'
  | boolean
  | PptxBullet
  | string;
/* eslint-enable @typescript-eslint/no-redundant-type-constituents */

export interface ParagraphConfig {
  align?: 'center' | 'justify' | 'left' | 'right';
  baseline?: number;
  bold?: boolean;
  bullet?: BulletInput;
  color?: string;
  font?: string;
  fontSize?: Points;
  hyperlink?: PptxHyperlink | string;
  italic?: boolean;
  level?: number;
  lineSpacing?: Points;
  runs?: (string | TextRunConfig)[];
  spaceAfter?: Points;
  spaceBefore?: Points;
  strikethrough?: 'dblStrike' | 'sngStrike' | boolean;
  subscript?: boolean;
  superscript?: boolean;
  text?: (string | TextRunConfig)[] | string | TextRunConfig;
  underline?: 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave' | boolean;
}

export interface TextOptions {
  align?: 'center' | 'justify' | 'left' | 'right';
  baseline?: number;
  bold?: boolean;
  bullet?: BulletInput;
  color?: string;
  font?: string;
  fontSize?: Points;
  hyperlink?: PptxHyperlink | string;
  italic?: boolean;
  level?: number;
  spaceAfter?: Points;
  spaceBefore?: Points;
  strikethrough?: 'dblStrike' | 'sngStrike' | boolean;
  subscript?: boolean;
  superscript?: boolean;
  underline?: 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave' | boolean;
  verticalAlignment?: 'bottom' | 'middle' | 'top';
}

/**
 * Normalizes user bullet configuration into a typed OpenXML PptxBullet.
 */
export function normalizeBullet(bullet?: BulletInput): PptxBullet | undefined {
  if (bullet === undefined) return undefined;
  if (bullet === false || bullet === 'none') {
    return { type: 'none' };
  }
  if (bullet === true || bullet === 'bullet') {
    return { char: '•', type: 'char' };
  }
  if (bullet === 'number') {
    return { autoNumType: 'arabicPeriod', type: 'autoNum' };
  }
  if (typeof bullet === 'string') {
    return { char: bullet, type: 'char' };
  }
  if (typeof bullet === 'object') {
    return bullet;
  }
  return undefined;
}

/**
 * Normalizes string or TextRunConfig into a typed PptxRun.
 * When `input` is a `TextRunConfig` with `break: true`, returns a line-break sentinel
 * (`PptxRun` with `break: true`) that serializes as `<a:br>` instead of `<a:r>`.
 */
export function buildTextRun(input: string | TextRunConfig, defaultOptions?: TextOptions): PptxRun {
  if (typeof input === 'string') {
    return {
      properties: {
        baseline: defaultOptions?.baseline,
        bold: defaultOptions?.bold,
        color: defaultOptions?.color,
        fontFamily: defaultOptions?.font,
        fontSize: defaultOptions?.fontSize ? hundredthsPoint(Math.round(defaultOptions.fontSize * 100)) : undefined,
        hyperlink: defaultOptions?.hyperlink,
        italic: defaultOptions?.italic,
        strikethrough: defaultOptions?.strikethrough,
        subscript: defaultOptions?.subscript,
        superscript: defaultOptions?.superscript,
        underline: defaultOptions?.underline,
      },
      text: input,
    };
  }

  const properties = {
    baseline: input.baseline ?? defaultOptions?.baseline,
    bold: input.bold ?? defaultOptions?.bold,
    color: input.color ?? defaultOptions?.color,
    fontFamily: input.font ?? defaultOptions?.font,
    fontSize: input.fontSize ? hundredthsPoint(Math.round(input.fontSize * 100)) : (defaultOptions?.fontSize ? hundredthsPoint(Math.round(defaultOptions.fontSize * 100)) : undefined),
    hyperlink: input.hyperlink ?? defaultOptions?.hyperlink,
    italic: input.italic ?? defaultOptions?.italic,
    strikethrough: input.strikethrough ?? defaultOptions?.strikethrough,
    subscript: input.subscript ?? defaultOptions?.subscript,
    superscript: input.superscript ?? defaultOptions?.superscript,
    underline: input.underline ?? defaultOptions?.underline,
  };

  if (input.break === true) {
    return {
      break: true,
      properties,
      text: '',
    };
  }

  return {
    properties,
    text: input.text ?? '',
  };
}

/**
 * Builds a PptxTextBody AST node from text content and styling options.
 * Accurately handles multilevel indentation, custom bullet hierarchies, newlines, and inline run modifiers.
 */
export function buildTextBody(
  content: (ParagraphConfig | string | TextRunConfig)[] | string,
  options?: TextOptions,
): PptxTextBody {
  const paragraphs: PptxParagraph[] = [];

  if (typeof content === 'string') {
    const lines = content.split(/\r?\n/);
    const hasTabs = lines.some((l) => l.startsWith('\t'));

    for (let rawLine of lines) {
      let level = options?.level ?? (hasTabs ? 0 : undefined);
      if (hasTabs) {
        let tabCount = 0;
        while (rawLine.startsWith('\t')) {
          tabCount++;
          rawLine = rawLine.slice(1);
        }
        level = tabCount;
      }

      paragraphs.push({
        properties: {
          alignment: options?.align || 'left',
          bullet: normalizeBullet(options?.bullet),
          level,
          spaceAfter: options?.spaceAfter ? hundredthsPoint(Math.round(options.spaceAfter * 100)) : undefined,
          spaceBefore: options?.spaceBefore ? hundredthsPoint(Math.round(options.spaceBefore * 100)) : undefined,
        },
        runs: [buildTextRun(rawLine, options)],
      });
    }
  } else if (Array.isArray(content) && content.length > 0) {
    const isParagraphConfigList = content.some(
      (item) =>
        typeof item === 'object'
        && item !== null
        && ('level' in item
          || 'runs' in item
          || 'bullet' in item
          || 'align' in item
          || 'spaceAfter' in item
          || 'spaceBefore' in item
          || 'lineSpacing' in item
          || (Array.isArray((item as Record<string, unknown>).text))),
    );

    if (isParagraphConfigList) {
      for (const item of content) {
        if (typeof item === 'string') {
          paragraphs.push({
            properties: {
              alignment: options?.align || 'left',
              bullet: normalizeBullet(options?.bullet),
              level: options?.level ?? 0,
            },
            runs: [buildTextRun(item, options)],
          });
        } else {
          const pConfig = item as ParagraphConfig;
          const level = pConfig.level ?? options?.level ?? 0;
          const bullet = normalizeBullet(pConfig.bullet ?? options?.bullet);
          const alignment = pConfig.align || options?.align || 'left';

          const mergedOptions: TextOptions = {
            ...options,
            baseline: pConfig.baseline ?? options?.baseline,
            bold: pConfig.bold ?? options?.bold,
            color: pConfig.color ?? options?.color,
            font: pConfig.font ?? options?.font,
            fontSize: pConfig.fontSize ?? options?.fontSize,
            hyperlink: pConfig.hyperlink ?? options?.hyperlink,
            italic: pConfig.italic ?? options?.italic,
            strikethrough: pConfig.strikethrough ?? options?.strikethrough,
            subscript: pConfig.subscript ?? options?.subscript,
            superscript: pConfig.superscript ?? options?.superscript,
            underline: pConfig.underline ?? options?.underline,
          };

          let runs: PptxRun[] = [];
          if (pConfig.runs && Array.isArray(pConfig.runs)) {
            runs = pConfig.runs.map((r) => buildTextRun(r, mergedOptions));
          } else if (typeof pConfig.text === 'string') {
            runs = [buildTextRun(pConfig.text, mergedOptions)];
          } else if (Array.isArray(pConfig.text)) {
            runs = pConfig.text.map((r) => buildTextRun(r, mergedOptions));
          } else if (typeof pConfig.text === 'object') {
            runs = [buildTextRun(pConfig.text, mergedOptions)];
          }

          paragraphs.push({
            properties: {
              alignment,
              bullet,
              level,
              spaceAfter: pConfig.spaceAfter ? hundredthsPoint(Math.round(pConfig.spaceAfter * 100)) : undefined,
              spaceBefore: pConfig.spaceBefore ? hundredthsPoint(Math.round(pConfig.spaceBefore * 100)) : undefined,
            },
            runs,
          });
        }
      }
    } else {
      let currentRuns: PptxRun[] = [];
      for (const item of content as (string | TextRunConfig)[]) {
        if (typeof item === 'object' && item !== null && 'break' in item && item.break === true) {
          currentRuns.push(buildTextRun(item, options));
          continue;
        }
        const text = typeof item === 'string' ? item : (item.text ?? '');
        const lines = text.split(/\r?\n/);
        for (let i = 0; i < lines.length; i++) {
          if (i > 0) {
            paragraphs.push({
              properties: {
                alignment: options?.align || 'left',
                bullet: normalizeBullet(options?.bullet),
                level: options?.level,
              },
              runs: currentRuns,
            });
            currentRuns = [];
          }
          if (lines[i].length > 0 || lines.length === 1) {
            const runConfig = typeof item === 'string' ? lines[i] : { ...item, text: lines[i] };
            currentRuns.push(buildTextRun(runConfig, options));
          }
        }
      }
      if (currentRuns.length > 0 || paragraphs.length === 0) {
        paragraphs.push({
          properties: {
            alignment: options?.align || 'left',
            bullet: normalizeBullet(options?.bullet),
            level: options?.level,
          },
          runs: currentRuns,
        });
      }
    }
  }

  return {
    bodyProperties: {
      verticalAlignment: options?.verticalAlignment || 'top',
    },
    paragraphs,
  };
}
