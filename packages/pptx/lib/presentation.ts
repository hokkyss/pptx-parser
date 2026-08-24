import type { PptxColorScheme, PptxDocument, PptxMetadata, PptxSlide } from '@hokkyss/pptx-core';
import type { ThemeColorInput, ThemeFontInput } from '@hokkyss/pptx-core';
import {
  emu,
  type Inches,
  inchesToEmu,
} from '@hokkyss/pptx-core';
import { Slide } from './slide';
import { SlideMaster } from './slide-master';
import type { WritePptxOptions } from '@hokkyss/pptx-writer';

export interface CreatePresentationOptions {
  author?: string;
  company?: string;
  /**
   * Slide height in inches.
   * Standard Landscape references:
   * - 16:9 Widescreen (Default): `inches(7.5)`
   * - 16:10 Widescreen: `inches(6.25)`
   * - 4:3 Standard: `inches(7.5)`
   * - A4 Landscape: `inches(8.27)`
   */
  height?: Inches;
  title?: string;
  /**
   * First slide number displayed on the presentation (starting slide number).
   * OpenXML: `<p:presentation firstSlideNum="...">` (Defaults to 1 in PowerPoint, set to 0 for decks with Title covers).
   */
  firstSlideNumber?: number;
  /**
   * Slide width in inches.
   * Standard Landscape references:
   * - 16:9 Widescreen (Default): `inches(13.333)`
   * - 16:10 Widescreen: `inches(10)`
   * - 4:3 Standard: `inches(10)`
   * - A4 Landscape: `inches(11.69)`
   */
  width?: Inches;
}

export interface AddSlideOptions {
  /**
   * Slide Layout to use for this slide.
   * Accepts a layout name (e.g. 'Title and Content', 'Blank', 'Title Slide') or layout ID ('slideLayout1').
   */
  layout?: string;
  /** Explicit layout ID (e.g. 'slideLayout1') */
  layoutId?: string;
  /**
   * Slide Master to use for this slide.
   * Accepts a SlideMaster instance, master name (e.g. 'Office Theme'), or master ID ('slideMaster1').
   */
  master?: SlideMaster | string;
  /** Speaker notes for the slide */
  notes?: string;
}

/**
 * Top-level Presentation class for programmatically creating and mutating `.pptx` presentations.
 */
export class Presentation {
  private _ast: PptxDocument;
  private _slideWrappers: Slide[] = [];

  constructor(ast: PptxDocument) {
    this._ast = ast;
    this.refreshSlideWrappers();
  }

  private refreshSlideWrappers(): void {
    this._slideWrappers = (this._ast.slides || []).map((s) => new Slide(s, this));
  }

  /**
   * Creates a new, blank PowerPoint presentation with default master, theme, and layout.
   */
  static create(options: CreatePresentationOptions = {}): Presentation {
    const widthEmu = options.width ? inchesToEmu(options.width) : emu(12192000); // 13.333 in EMU
    const heightEmu = options.height ? inchesToEmu(options.height) : emu(6858000); // 7.5 in EMU

    const metadata: PptxMetadata = {
      created: new Date(),
      creator: options.author || 'Pptx SDK',
      firstSlideNumber: options.firstSlideNumber,
      lastModifiedBy: options.author || 'Pptx SDK',
      modified: new Date(),
      revision: 1,
      slideCount: 0,
      slideHeight: heightEmu,
      slideWidth: widthEmu,
      title: options.title,
    };

    const doc: PptxDocument = {
      customXml: [],
      media: [],
      metadata,
      slideLayouts: [
        {
          elements: [],
          id: 'slideLayout1',
          masterId: 'slideMaster1',
          name: 'Blank',
          shapes: [],
          type: 'layout',
        },
      ],
      slideMasters: [
        {
          elements: [],
          id: 'slideMaster1',
          layoutIds: ['slideLayout1'],
          name: 'Office Theme',
          shapes: [],
          theme: {
            colorScheme: {
              accent1: '2563EB',
              accent2: '10B981',
              accent3: 'F59E0B',
              accent4: 'EF4444',
              accent5: '8B5CF6',
              accent6: 'EC4899',
              dk1: '000000',
              dk2: '1F2937',
              folHlink: '6D28D9',
              hlink: '2563EB',
              lt1: 'FFFFFF',
              lt2: 'F3F4F6',
            },
            customColors: {},
            fontScheme: {
              majorFont: 'Calibri',
              minorFont: 'Calibri',
              name: 'Office',
            },
            formatScheme: {},
            id: 'theme1',
            name: 'Office Theme',
          },
        },
      ],
      slides: [],
      themes: [
        {
          colorScheme: {
            accent1: '2563EB',
            accent2: '10B981',
            accent3: 'F59E0B',
            accent4: 'EF4444',
            accent5: '8B5CF6',
            accent6: 'EC4899',
            dk1: '000000',
            dk2: '1F2937',
            folHlink: '6D28D9',
            hlink: '2563EB',
            lt1: 'FFFFFF',
            lt2: 'F3F4F6',
          },
          customColors: {},
          fontScheme: {
            majorFont: 'Calibri',
            minorFont: 'Calibri',
            name: 'Office',
          },
          formatScheme: {},
          id: 'theme1',
          name: 'Office Theme',
        },
      ],
    };

    return new Presentation(doc);
  }

  /**
   * Loads an existing `.pptx` presentation buffer into a fluent `Presentation` model.
   */
  static async load(input: ArrayBuffer | Uint8Array): Promise<Presentation> {
    const { parsePptx } = await import('@hokkyss/pptx-reader');
    const doc = await parsePptx(input, {
      customXml: true,
      includeMedia: true,
      parseAnimations: true,
      parseTransitions: true,
    });
    return new Presentation(doc);
  }

  /** Underlying `PptxDocument` AST node */
  get ast(): PptxDocument {
    return this._ast;
  }

  /** Presentation metadata (title, author, slide dimensions) */
  get metadata(): PptxMetadata {
    return this._ast.metadata;
  }

  /** Array of all `Slide` instances in the presentation */
  get slides(): Slide[] {
    return this._slideWrappers;
  }

  /**
   * Returns all Slide Masters in this presentation.
   */
  getMasters(): SlideMaster[] {
    return (this._ast.slideMasters || []).map((m) => new SlideMaster(m, this));
  }

  /**
   * Finds a Slide Master by name, ID, or 1-based index.
   */
  getMaster(nameOrIdOrIndex: number | string): SlideMaster | undefined {
    const masters = this.getMasters();

    if (typeof nameOrIdOrIndex === 'number') {
      return masters[nameOrIdOrIndex - 1] || masters[nameOrIdOrIndex];
    }

    const query = nameOrIdOrIndex.toLowerCase();
    return (
      masters.find((m) => m.id.toLowerCase() === query)
      || masters.find((m) => m.name && m.name.toLowerCase() === query)
      || masters.find((m) => m.name && m.name.toLowerCase().includes(query))
    );
  }

  /**
   * Resolves the appropriate `layoutId` based on master and layout options.
   */
  private resolveLayoutId(options: AddSlideOptions): string {
    const layouts = this._ast.slideLayouts || [];
    const defaultLayoutId = layouts[0]?.id || 'slideLayout1';

    // 1. If master is specified
    if (options.master) {
      const masterInstance = options.master instanceof SlideMaster
        ? options.master
        : this.getMaster(options.master);

      if (masterInstance) {
        const layoutTarget = options.layout || options.layoutId;
        if (layoutTarget) {
          const match = masterInstance.getLayout(layoutTarget);
          if (match) return match.id;
        }
        // If layout is not specified or not found on master, use first layout of master
        const masterLayouts = masterInstance.getLayouts();
        if (masterLayouts.length > 0) {
          return masterLayouts[0].id;
        }
      }
    }

    // 2. If layout / layoutId is specified without master
    const layoutTarget = options.layout || options.layoutId;
    if (layoutTarget) {
      const query = layoutTarget.toLowerCase();
      const match
        = layouts.find((l) => l.id.toLowerCase() === query)
          || layouts.find((l) => l.name.toLowerCase() === query)
          || layouts.find((l) => l.name.toLowerCase().includes(query))
          || layouts.find((l) => l.type?.toLowerCase() === query)
          || layouts.find((l) => l.matchingName?.toLowerCase().includes(query));

      if (match) return match.id;
    }

    return defaultLayoutId;
  }

  /**
   * Appends a new slide to the presentation.
   * @param options Optional slide configuration (master, layout, notes).
   * @returns The newly created `Slide` instance.
   */
  addSlide(options: AddSlideOptions = {}): Slide {
    const slideNumber = this._ast.slides.length + 1;
    const slideId = `rId${slideNumber + 1}`;
    const layoutId = this.resolveLayoutId(options);

    const rawSlide: PptxSlide = {
      animations: [],
      elements: [],
      layoutId,
      notes: options.notes,
      shapes: [],
      slideId,
      slideNumber,
    };

    this._ast.slides.push(rawSlide);
    this._ast.metadata.slideCount = this._ast.slides.length;

    const slide = new Slide(rawSlide, this);
    this._slideWrappers.push(slide);
    return slide;
  }

  /**
   * Retrieves a slide by its 1-based slide number or its internal slide ID (`rId*`).
   */
  getSlide(indexOrId: number | string): Slide | undefined {
    if (typeof indexOrId === 'number') {
      return this._slideWrappers[indexOrId - 1];
    }
    return this._slideWrappers.find((s) => s.slideId === indexOrId);
  }

  /**
   * Removes a slide by its 1-based slide number or its internal slide ID (`rId*`).
   * Automatically re-numbers subsequent slides.
   */
  removeSlide(indexOrId: number | string): boolean {
    const slideIdx = typeof indexOrId === 'number'
      ? indexOrId - 1
      : this._slideWrappers.findIndex((s) => s.slideId === indexOrId);

    if (slideIdx < 0 || slideIdx >= this._ast.slides.length) {
      return false;
    }

    this._ast.slides.splice(slideIdx, 1);
    this._slideWrappers.splice(slideIdx, 1);

    // Re-index remaining slides
    for (let i = 0; i < this._ast.slides.length; i++) {
      this._ast.slides[i].slideNumber = i + 1;
    }

    this._ast.metadata.slideCount = this._ast.slides.length;
    return true;
  }

  /**
   * Duplicates an existing slide with deep clone of all elements, notes, and background.
   */
  duplicateSlide(indexOrId: number | string): Slide {
    const original = this.getSlide(indexOrId);
    if (!original) {
      throw new Error(`Cannot duplicate slide: slide ${indexOrId} not found.`);
    }

    const slideNumber = this._ast.slides.length + 1;
    const slideId = `rId${slideNumber + 1}`;

    const clonedSlide: PptxSlide = {
      animations: structuredClone(original.ast.animations || []),
      background: original.ast.background ? structuredClone(original.ast.background) : undefined,
      elements: structuredClone(original.ast.elements || []),
      layoutId: original.ast.layoutId,
      notes: original.ast.notes,
      shapes: structuredClone(original.ast.shapes || []),
      slideId,
      slideNumber,
    };

    this._ast.slides.push(clonedSlide);
    this._ast.metadata.slideCount = this._ast.slides.length;

    const slide = new Slide(clonedSlide, this);
    this._slideWrappers.push(slide);
    return slide;
  }

  /**
   * Moves a slide from one 1-based index to another.
   */
  moveSlide(fromIndex: number, toIndex: number): void {
    if (
      fromIndex < 1
      || fromIndex > this._slideWrappers.length
      || toIndex < 1
      || toIndex > this._slideWrappers.length
    ) {
      throw new Error(`Invalid slide index for move: from ${fromIndex} to ${toIndex}.`);
    }

    const [movedSlide] = this._ast.slides.splice(fromIndex - 1, 1);
    this._ast.slides.splice(toIndex - 1, 0, movedSlide);

    const [movedWrapper] = this._slideWrappers.splice(fromIndex - 1, 1);
    this._slideWrappers.splice(toIndex - 1, 0, movedWrapper);

    // Re-index
    for (let i = 0; i < this._ast.slides.length; i++) {
      this._ast.slides[i].slideNumber = i + 1;
    }
  }

  /**
   * Partially merges color scheme overrides into the primary slide master theme.
   * Accepts '#hex' or raw hex strings. Only provided color fields are updated.
   * @param colors Partial color scheme object (e.g. `{ accent1: '#EE957F', accent2: 'FD8628' }`)
   * @param name Optional name for the color palette
   */
  setThemeColors(colors: ThemeColorInput, name?: string): this {
    const theme = this._ast.themes[0];
    if (!theme) return this;
    if (name) {
      theme.colorScheme.name = name;
    }
    for (const [key, value] of Object.entries(colors) as [keyof PptxColorScheme, string | undefined][]) {
      if (value !== undefined) {
        theme.colorScheme[key] = value.replace(/^#/, '').toUpperCase();
      }
    }
    if (this._ast.slideMasters[0]?.theme) {
      this._ast.slideMasters[0].theme.colorScheme = structuredClone(theme.colorScheme);
    }
    return this;
  }

  /**
   * Partially merges font scheme overrides into the primary slide master theme.
   * @param fonts Font configuration containing `major` (headings), `minor` (body text), and/or font scheme `name`.
   */
  setThemeFonts(fonts: ThemeFontInput): this {
    const theme = this._ast.themes[0];
    if (!theme) return this;
    if (fonts.major) {
      theme.fontScheme.majorFont = fonts.major;
    }
    if (fonts.minor) {
      theme.fontScheme.minorFont = fonts.minor;
    }
    if (fonts.name) {
      theme.fontScheme.name = fonts.name;
    }
    if (this._ast.slideMasters[0]?.theme) {
      this._ast.slideMasters[0].theme.fontScheme = structuredClone(theme.fontScheme);
    }
    return this;
  }

  /**
   * Sets the display name of the primary theme (shown in PowerPoint's Theme/Color schemes dropdown).
   * @param name Theme name (e.g. 'Ajinomoto Corporate Theme')
   */
  setThemeName(name: string): this {
    const theme = this._ast.themes[0];
    if (!theme) return this;
    theme.name = name;
    theme.colorScheme.name = name;
    if (this._ast.slideMasters[0]?.theme) {
      this._ast.slideMasters[0].theme.name = name;
      this._ast.slideMasters[0].theme.colorScheme.name = name;
    }
    return this;
  }

  /**
   * First slide number displayed on the presentation (starting slide number).
   * OpenXML: `<p:presentation firstSlideNum="...">` (Defaults to 1, set to 0 for decks with Title covers).
   */
  get firstSlideNumber(): number | undefined {
    return this._ast.metadata?.firstSlideNumber;
  }

  /**
   * Sets the starting slide number for the presentation.
   * @param num Starting slide number (e.g. 0 or 1).
   * @example
   * ```ts
   * const pres = Presentation.create();
   * pres.setFirstSlideNumber(0); // Cover slide is numbered 0, content starts at 1
   * ```
   */
  setFirstSlideNumber(num: number): this {
    if (this._ast.metadata) {
      this._ast.metadata.firstSlideNumber = num;
    }
    return this;
  }

  /**
   * Serializes the presentation to a binary `.pptx` Uint8Array.
   */
  async toBuffer(options?: WritePptxOptions): Promise<Uint8Array> {
    const { writePptx } = await import('@hokkyss/pptx-writer');
    return writePptx(this._ast, options);
  }

  /**
   * Serializes the presentation to an ArrayBuffer.
   */
  async toArrayBuffer(options?: WritePptxOptions): Promise<ArrayBuffer> {
    const uint8 = await this.toBuffer(options);
    const copy = new Uint8Array(uint8.byteLength);
    copy.set(uint8);
    return copy.buffer;
  }
}
