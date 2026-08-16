import type { Emu, HundredthsPoint } from './units';

/** Represents bullet point formatting */
export interface PptxBullet {
  /** Auto-numbering type for 'autoNum' type (e.g. 'arabicPeriod'). OpenXML: `<a:buAutoNum @_type>` */
  autoNumType?: string;
  /** Character for 'char' type. OpenXML: `<a:buChar @_char>` */
  char?: string;
  /** Starting number for 'autoNum' type. OpenXML: `<a:buAutoNum @_startAt>` */
  startAt?: number;
  /** Bullet type. OpenXML: `<a:buChar>`, `<a:buAutoNum>`, `<a:buBlip>`, `<a:buNone>` */
  type: 'autoNum' | 'char' | 'none' | 'picture';
}

/** Strongly-typed paragraph properties (margins, padding, spacing, indentation, bullets) */
export interface PptxParagraphProperties {
  /** Text alignment ('left' | 'center' | 'right' | 'justify'). OpenXML: `<a:pPr @_algn>` */
  alignment?: 'center' | 'justify' | 'left' | 'right';
  /** Bullet point configuration. OpenXML: `<a:buChar>`, `<a:buAutoNum>`, `<a:buNone>` */
  bullet?: PptxBullet;
  /** First line indent or hanging indent in EMU. OpenXML: `<a:pPr @_indent>` */
  firstLineIndent?: Emu;
  /** Left margin / Indentation in EMU. OpenXML: `<a:pPr @_marL>` */
  leftMargin?: Emu;
  /** Indentation level (0-based level). OpenXML: `<a:pPr @_lvl>` */
  level?: number;
  /** Line spacing value in Hundredths of a Point or Percent. OpenXML: `<a:pPr><a:lnSpc>` */
  lineSpacing?: HundredthsPoint;
  /** Right margin in EMU. OpenXML: `<a:pPr @_marR>` */
  rightMargin?: Emu;
  /** Space after paragraph in Hundredths of a Point. OpenXML: `<a:pPr><a:spcAft><a:spcPts @_val>` */
  spaceAfter?: HundredthsPoint;
  /** Space before paragraph in Hundredths of a Point. OpenXML: `<a:pPr><a:spcBef><a:spcPts @_val>` */
  spaceBefore?: HundredthsPoint;
}

/** Represents a run of text with uniform formatting */
export interface PptxRun {
  /** Text properties. OpenXML: `<a:rPr>` */
  properties: {
    /** Baseline shift in thousandths of a percent (e.g. 30000 for +30% superscript, -25000 for -25% subscript). OpenXML: `<a:rPr @_baseline>` */
    baseline?: number;
    /** Bold. OpenXML: `<a:rPr @_b>` */
    bold?: boolean;
    /** Text color. OpenXML: `<a:rPr><a:solidFill><a:srgbClr @_val>` */
    color?: string;
    /** Font family name. OpenXML: `<a:rPr><a:latin @_typeface>` */
    fontFamily?: string;
    /** Font size in Hundredths of a Point. OpenXML: `<a:rPr @_sz>` */
    fontSize?: HundredthsPoint;
    /** Hyperlink URL or reference. OpenXML: `<a:rPr><a:hlinkClick>` */
    hyperlink?: string;
    /** Italic. OpenXML: `<a:rPr @_i>` */
    italic?: boolean;
    /** Language. OpenXML: `<a:rPr @_lang>` */
    language?: string;
    /** Strikethrough. OpenXML: `<a:rPr @_strike>` */
    strikethrough?: 'dblStrike' | 'sngStrike' | boolean;
    /** Subscript convenience helper. OpenXML: `<a:rPr @_baseline="-25000">` */
    subscript?: boolean;
    /** Superscript convenience helper. OpenXML: `<a:rPr @_baseline="30000">` */
    superscript?: boolean;
    /** Underline. OpenXML: `<a:rPr @_u>` */
    underline?: 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave' | boolean;
  };
  /** Text content. OpenXML: `<a:t>` */
  text: string;
}

/** Represents a single paragraph of text */
export interface PptxParagraph {
  /** Strongly-typed paragraph properties. OpenXML: `<a:pPr>` */
  properties: PptxParagraphProperties;
  /** Text runs in the paragraph. OpenXML: `<a:r>` & `<a:fld>` */
  runs: PptxRun[];
}

/** Strongly-typed text body properties (all camelCase, no @_ attributes) */
export interface PptxTextBodyProperties {
  /** Bottom padding inset in EMU. OpenXML: `<a:bodyPr @_bIns>` */
  bottomInset?: Emu;
  /** Number of text columns. OpenXML: `<a:bodyPr @_numCol>` */
  columns?: number;
  /** Spacing between columns in EMU. OpenXML: `<a:bodyPr @_spcCol>` */
  columnSpacing?: Emu;
  /** Left padding inset in EMU. OpenXML: `<a:bodyPr @_lIns>` */
  leftInset?: Emu;
  /** Right padding inset in EMU. OpenXML: `<a:bodyPr @_rIns>` */
  rightInset?: Emu;
  /** Top padding inset in EMU. OpenXML: `<a:bodyPr @_tIns>` */
  topInset?: Emu;
  /** Upright vertical text setting. OpenXML: `<a:bodyPr @_upright>` */
  upright?: boolean;
  /** Vertical alignment of text within text frame ('top' | 'middle' | 'bottom'). OpenXML: `<a:bodyPr @_anchor>` */
  verticalAlignment?: 'bottom' | 'middle' | 'top';
  /** Vertical text orientation. OpenXML: `<a:bodyPr @_vert>` */
  verticalText?: boolean;
  /** Text wrapping style ('square' | 'none'). OpenXML: `<a:bodyPr @_wrap>` */
  wrap?: 'none' | 'square';
}

/** Represents text content within a shape */
export interface PptxTextBody {
  /** Strongly-typed body properties. OpenXML: `<a:bodyPr>` */
  bodyProperties: PptxTextBodyProperties;
  /** Paragraphs in the text body. OpenXML: `<a:p>` */
  paragraphs: PptxParagraph[];
}
