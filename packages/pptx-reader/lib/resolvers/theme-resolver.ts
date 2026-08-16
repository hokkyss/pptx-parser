import type { PptxColorScheme, PptxFontScheme, PptxTheme } from '@hokkyss/pptx-core';
import type { XmlParser } from '@hokkyss/pptx-core';
import { defaultXmlParser } from '../xml/xml-parser';

export type { PptxColorScheme, PptxFontScheme, PptxTheme };

/**
 * Resolved theme helper interface.
 */
export interface ThemeResolver {
  /**
   * Parses theme XML string or object into a structured `PptxTheme` AST node.
   * @param themeXml Raw XML string of a theme file (e.g. `ppt/theme/theme1.xml`).
   * @returns Structured `PptxTheme` containing colorScheme, fontScheme, and customColors.
   */
  parseTheme(themeXml: string): PptxTheme;
}

/**
 * Helper extracting hex color string from color XML node (`srgbClr` or `sysClr`).
 * @param colorNode Target color XML node.
 * @returns Hex color string without `#` prefix (e.g. `'000F78'`).
 */
function extractHexColor(colorNode: Record<string, unknown>): string {
  // Check srgbClr
  const srgbClr = (colorNode['a:srgbClr'] || colorNode['srgbClr']) as Record<string, unknown> | undefined;
  if (srgbClr && srgbClr['@_val']) {
    return srgbClr['@_val'] as string;
  }
  // Check sysClr
  const sysClr = (colorNode['a:sysClr'] || colorNode['sysClr']) as Record<string, unknown> | undefined;
  if (sysClr && sysClr['@_lastClr']) {
    return sysClr['@_lastClr'] as string;
  }
  return '000000';
}

/**
 * Helper method parsing color scheme palette elements (`dk1`, `lt1`, `accent1-6`, `hlink`, `folHlink`).
 * @param themeElements Raw theme elements XML node.
 * @returns Resolved `PptxColorScheme` object.
 */
function parseColorScheme(themeElements?: Record<string, unknown>): PptxColorScheme {
  const defaultScheme: PptxColorScheme = {
    accent1: '4F81BD',
    accent2: 'C0504D',
    accent3: '9BBB59',
    accent4: '8064A2',
    accent5: '4BACC6',
    accent6: 'F79646',
    dk1: '000000',
    dk2: '1F497D',
    folHlink: '800080',
    hlink: '0000FF',
    lt1: 'FFFFFF',
    lt2: 'EEECE1',
  };

  if (!themeElements) return defaultScheme;

  const clrScheme = (themeElements['a:clrScheme'] || themeElements['clrScheme']) as Record<string, unknown> | undefined;
  if (!clrScheme) return defaultScheme;

  const keys: (keyof PptxColorScheme)[] = [
    'dk1', 'lt1', 'dk2', 'lt2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
  ];

  const result: Partial<PptxColorScheme> = {};

  for (const key of keys) {
    const colorNode = (clrScheme[`a:${key}`] || clrScheme[key]) as Record<string, unknown> | undefined;
    if (colorNode) {
      result[key] = extractHexColor(colorNode);
    }
  }

  return { ...defaultScheme, ...result };
}

/**
 * Helper extracting latin font typeface string from font XML node.
 * @param fontNode Font XML node.
 * @returns Font family name (e.g. `'Calibri'`) or `undefined`.
 */
function extractFontTypeface(fontNode?: Record<string, unknown>): string | undefined {
  if (!fontNode) return undefined;
  const latin = (fontNode['a:latin'] || fontNode['latin']) as Record<string, unknown> | undefined;
  if (latin && latin['@_typeface']) {
    return latin['@_typeface'] as string;
  }
  return undefined;
}

/**
 * Helper parsing font scheme elements (`majorFont` and `minorFont`).
 * @param themeElements Raw theme elements XML node.
 * @returns Resolved `PptxFontScheme` object.
 */
function parseFontScheme(themeElements?: Record<string, unknown>): PptxFontScheme {
  const defaultScheme: PptxFontScheme = {
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    name: 'Default Font Scheme',
  };

  if (!themeElements) return defaultScheme;

  const fontScheme = (themeElements['a:fontScheme'] || themeElements['fontScheme']) as Record<string, unknown> | undefined;
  if (!fontScheme) return defaultScheme;

  const schemeName = (fontScheme['@_name'] as string) || defaultScheme.name;

  const majorNode = (fontScheme['a:majorFont'] || fontScheme['majorFont']) as Record<string, unknown> | undefined;
  const minorNode = (fontScheme['a:minorFont'] || fontScheme['minorFont']) as Record<string, unknown> | undefined;

  const majorFont = extractFontTypeface(majorNode) || defaultScheme.majorFont;
  const minorFont = extractFontTypeface(minorNode) || defaultScheme.minorFont;

  return {
    majorFont,
    minorFont,
    name: schemeName,
  };
}

/**
 * Helper parsing custom color list (`a:custClrLst`).
 * @param themeTag Theme root XML node.
 * @returns Record mapping custom color names to hex values.
 */
function parseCustomColors(themeTag: Record<string, unknown>): Record<string, string> {
  const custClrLst = (themeTag['a:custClrLst'] || themeTag['custClrLst']) as Record<string, unknown> | undefined;
  if (!custClrLst) return {};

  let custClrs = custClrLst['a:custClr'] || custClrLst['custClr'];
  if (!custClrs) return {};
  if (!Array.isArray(custClrs)) {
    custClrs = [custClrs];
  }

  const results: Record<string, string> = {};
  for (const item of custClrs as Record<string, unknown>[]) {
    const name = (item['@_name'] as string) || 'custom';
    const hex = extractHexColor(item);
    results[name] = hex;
  }
  return results;
}

/**
 * Fallback default theme structure.
 * @returns Default fallback PptxTheme object.
 */
function createDefaultTheme(): PptxTheme {
  return {
    colorScheme: {
      accent1: '4F81BD',
      accent2: 'C0504D',
      accent3: '9BBB59',
      accent4: '8064A2',
      accent5: '4BACC6',
      accent6: 'F79646',
      dk1: '000000',
      dk2: '1F497D',
      folHlink: '800080',
      hlink: '0000FF',
      lt1: 'FFFFFF',
      lt2: 'EEECE1',
    },
    customColors: {},
    fontScheme: { majorFont: 'Calibri', minorFont: 'Calibri', name: 'Default' },
    formatScheme: {},
    name: 'Default Theme',
  };
}

/**
 * Creates a ThemeResolver instance for parsing Presentation Theme XML (`ppt/theme/theme1.xml`).
 * @param parser Optional custom `XmlParser` instance.
 * @returns Frozen `ThemeResolver` instance.
 * @example
 * ```ts
 * const themeResolver = createThemeResolver();
 * const theme = themeResolver.parseTheme(themeXmlString);
 * console.log(theme.name, theme.colorScheme.accent1, theme.fontScheme.majorFont);
 * ```
 */
export function createThemeResolver(parser: XmlParser = defaultXmlParser): ThemeResolver {
  return Object.freeze({
    parseTheme(themeXml: string): PptxTheme {
      const parsed = parser.parse<Record<string, unknown>>(themeXml);
      const themeTag = (parsed['a:theme'] || parsed['theme']) as Record<string, unknown> | undefined;

      if (!themeTag) {
        return createDefaultTheme();
      }

      const themeName = (themeTag['@_name'] as string) || 'Default Theme';
      const themeElements = (themeTag['a:themeElements'] || themeTag['themeElements']) as Record<string, unknown> | undefined;

      // Parse Color Scheme
      const colorScheme = parseColorScheme(themeElements);
      // Parse Font Scheme
      const fontScheme = parseFontScheme(themeElements);
      // Parse Custom Colors
      const customColors = parseCustomColors(themeTag);

      return {
        colorScheme,
        customColors,
        fontScheme,
        formatScheme: {
          backgroundFillStyles: [],
          fillStyles: [],
          lineStyles: [],
        },
        name: themeName,
      };
    },
  });
}
