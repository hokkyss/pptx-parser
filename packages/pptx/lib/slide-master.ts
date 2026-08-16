import type { PptxSlideLayout, PptxSlideMaster } from '@hokkyss/pptx-core';
import type { AddSlideOptions, Presentation } from './presentation';
import type { Slide } from './slide';

/**
 * Wrapper class representing a PowerPoint Slide Master (`ppt/slideMasters/slideMaster*.xml`).
 * Allows discovering child layouts and creating slides bound to this master.
 */
export class SlideMaster {
  private _ast: PptxSlideMaster;
  private _presentation: Presentation;

  constructor(ast: PptxSlideMaster, presentation: Presentation) {
    this._ast = ast;
    this._presentation = presentation;
  }

  /** Slide Master ID (e.g. 'slideMaster1') */
  get id(): string {
    return this._ast.id;
  }

  /** Slide Master name (e.g. 'Office Theme') */
  get name(): string | undefined {
    return this._ast.name || this._ast.theme?.name;
  }

  /** Direct reference to underlying `PptxSlideMaster` AST node */
  get ast(): PptxSlideMaster {
    return this._ast;
  }

  /** Presentation instance owning this master */
  get presentation(): Presentation {
    return this._presentation;
  }

  /**
   * Returns all `PptxSlideLayout`s associated with this Slide Master.
   */
  getLayouts(): PptxSlideLayout[] {
    const layoutIdSet = new Set(this._ast.layoutIds || []);
    return (this._presentation.ast.slideLayouts || []).filter(
      (l) => l.masterId === this._ast.id || layoutIdSet.has(l.id),
    );
  }

  /**
   * Finds a specific layout belonging to this master by name, type, ID, or index.
   * @param nameOrIdOrIndex Name (e.g. 'Title and Content'), ID ('slideLayout1'), or 1-based index.
   */
  getLayout(nameOrIdOrIndex: number | string): PptxSlideLayout | undefined {
    const layouts = this.getLayouts();

    if (typeof nameOrIdOrIndex === 'number') {
      return layouts[nameOrIdOrIndex - 1] || layouts[nameOrIdOrIndex];
    }

    const query = nameOrIdOrIndex.toLowerCase();
    return (
      layouts.find((l) => l.id.toLowerCase() === query)
      || layouts.find((l) => l.name.toLowerCase() === query)
      || layouts.find((l) => l.name.toLowerCase().includes(query))
      || layouts.find((l) => l.type?.toLowerCase() === query)
      || layouts.find((l) => l.matchingName?.toLowerCase().includes(query))
    );
  }

  /**
   * Adds a new slide bound to this Slide Master.
   * @param options Slide configuration options (layout name/id, speaker notes).
   * @returns The newly created `Slide` instance.
   */
  addSlide(options: Omit<AddSlideOptions, 'master'> = {}): Slide {
    return this._presentation.addSlide({
      ...options,
      master: this,
    });
  }
}
