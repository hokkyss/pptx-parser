/**
 * Single relationship entry mapping an OpenXML ID (`rId*`) to a target file path and type URI.
 */
export interface Relationship {
  /** Relationship ID (e.g. `'rId1'`) */
  id: string;
  /** Resolved absolute path within the PPTX ZIP archive (e.g. `'ppt/slides/slide1.xml'`) */
  resolvedTarget: string;
  /** Raw target file path as specified in the `.rels` file (e.g. `'../slides/slide1.xml'`) */
  target: string;
  /** Target mode (`'External'` for remote web URLs) */
  targetMode?: string;
  /** Relationship type URI (e.g. `'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'`) */
  type: string;
}

/**
 * Resolves OpenXML relationship files (`.rels`) mapping relationship IDs (`rId1`, `rId2`) to resource target paths.
 */
export interface RelationshipResolver {
  /**
   * Manually adds relationships scoped to a source file path.
   * @param sourcePath Source file path.
   * @param rels List of relationship objects with id, type, and target.
   */
  addRelationships(sourcePath: string, rels: { id: string; target: string; type: string }[]): void;

  /**
   * Returns all parsed relationships in this resolver instance.
   * @returns Array of all `Relationship` objects.
   */
  getAll(): Relationship[];

  /**
   * Retrieves relationship by ID or by `(sourcePath, id)`.
   * @param sourcePathOrId Relationship ID string (e.g. `'rId1'`) or source path string.
   * @param id Optional relationship ID if first argument is a source path.
   * @returns Matching `Relationship` object or `undefined`.
   */
  getRelationship(sourcePathOrId: string, id?: string): Relationship | undefined;

  /**
   * Retrieves all relationships matching a specific type keyword or full URI.
   * @param typeKeyword Keyword to match against type URI (e.g. `'slide'`, `'slideLayout'`, `'image'`, `'chart'`).
   * @returns Array of matching `Relationship` objects.
   */
  getRelationshipsByType(typeKeyword: string): Relationship[];

  /**
   * Retrieves resolved target file path for a relationship ID.
   * @param id Relationship ID string (e.g. `'rId1'`).
   * @returns Resolved path string or `undefined`.
   */
  getTarget(id: string): string | undefined;

  /**
   * Parses `.rels` XML string content into `Relationship` objects.
   * @param relsXml Raw XML string.
   * @param sourcePath Base file path for resolving relative targets.
   * @returns Array of parsed `Relationship` entries.
   */
  parseRels(relsXml: string, sourcePath?: string): Relationship[];
}
