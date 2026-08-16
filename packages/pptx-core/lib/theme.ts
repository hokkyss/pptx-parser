import type { PptxFill } from './color';
import type { Emu } from './units';

/** Represents line (stroke) properties */
export interface PptxLine {
  /** Line dash style. OpenXML: `<a:ln><a:prstDash @_val>` */
  dashStyle?: string;
  /** Line fill. OpenXML: `<a:ln><a:solidFill>` */
  fill?: PptxFill;
  /** Line width in EMU. OpenXML: `<a:ln @_w>` */
  width?: Emu;
}

/** Represents a color scheme within a theme */
export interface PptxColorScheme {
  /** Accent 1 color (hex). OpenXML: `<a:accent1>` */
  accent1: string;
  /** Accent 2 color (hex). OpenXML: `<a:accent2>` */
  accent2: string;
  /** Accent 3 color (hex). OpenXML: `<a:accent3>` */
  accent3: string;
  /** Accent 4 color (hex). OpenXML: `<a:accent4>` */
  accent4: string;
  /** Accent 5 color (hex). OpenXML: `<a:accent5>` */
  accent5: string;
  /** Accent 6 color (hex). OpenXML: `<a:accent6>` */
  accent6: string;
  /** Dark 1 color (hex). OpenXML: `<a:dk1>` */
  dk1: string;
  /** Dark 2 color (hex). OpenXML: `<a:dk2>` */
  dk2: string;
  /** Followed hyperlink color (hex). OpenXML: `<a:folHlink>` */
  folHlink: string;
  /** Hyperlink color (hex). OpenXML: `<a:hlink>` */
  hlink: string;
  /** Light 1 color (hex). OpenXML: `<a:lt1>` */
  lt1: string;
  /** Light 2 color (hex). OpenXML: `<a:lt2>` */
  lt2: string;
}

/** Represents a font scheme within a theme */
export interface PptxFontScheme {
  /** Major font family. OpenXML: `<a:majorFont><a:latin @_typeface>` */
  majorFont: string;
  /** Minor font family. OpenXML: `<a:minorFont><a:latin @_typeface>` */
  minorFont: string;
  /** Font scheme name. OpenXML: `<a:fontScheme @_name>` */
  name: string;
}

/** Strongly-typed theme format scheme */
export interface PptxFormatScheme {
  /** Background fill styles. OpenXML: `<a:fmtScheme><a:bgFillStyleLst>` */
  backgroundFillStyles?: PptxFill[];
  /** Fill styles. OpenXML: `<a:fmtScheme><a:fillStyleLst>` */
  fillStyles?: PptxFill[];
  /** Line styles. OpenXML: `<a:fmtScheme><a:lnStyleLst>` */
  lineStyles?: PptxLine[];
}

/** Represents a theme */
export interface PptxTheme {
  /** Color scheme. OpenXML: `<a:clrScheme>` */
  colorScheme: PptxColorScheme;
  /** Custom colors map (e.g. { navy: '000F78' }). OpenXML: `<a:custClrLst>` */
  customColors: Record<string, string>;
  /** Font scheme. OpenXML: `<a:fontScheme>` */
  fontScheme: PptxFontScheme;
  /** Format scheme. OpenXML: `<a:fmtScheme>` */
  formatScheme: PptxFormatScheme;
  /** Optional theme part ID (e.g. 'theme1') */
  id?: string;
  /** Theme name. OpenXML: `<a:theme @_name>` */
  name: string;
  /** Optional raw OpenXML content for exact round-trip fidelity */
  rawXml?: string;
}

/** Color scheme overrides input (partial, accepts hex strings with or without #) */
export type ThemeColorInput = Partial<Record<keyof PptxColorScheme, string>>;

/** Font scheme overrides input */
export interface ThemeFontInput {
  /** Major font (used for headings) */
  major?: string;
  /** Minor font (used for body text) */
  minor?: string;
  /** Font scheme name */
  name?: string;
}
