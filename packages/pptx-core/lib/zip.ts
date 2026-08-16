/**
 * Single decompressed file entry in a PPTX ZIP container.
 */
export interface ZipEntry {
  /** Raw decompressed binary content */
  data: Uint8Array;
  /** Relative file path within the ZIP archive (e.g. `'ppt/presentation.xml'`) */
  path: string;
}

/**
 * Security and resource limit options for decompressing PPTX ZIP archives.
 */
export interface ZipReaderOptions {
  /** Maximum allowed number of file entries in the ZIP archive (default: 10,000) */
  maxEntries?: number;
  /** Maximum allowed uncompressed byte size for any individual file entry (default: 100MB) */
  maxSingleFileBytes?: number;
  /** Maximum allowed total uncompressed byte size for the entire archive (default: 500MB) */
  maxTotalBytes?: number;
}

/**
 * Handles non-blocking asynchronous decompression, file entry lookup, and text decoding for OpenXML ZIP packages (`.pptx`).
 */
export interface ZipReader {
  /**
   * Asynchronous getter for text content.
   * @param path Relative file path in the ZIP.
   * @returns Promise resolving to text content or empty string if missing.
   */
  getFileAsString(path: string): Promise<string>;

  /**
   * Retrieves raw binary data for a file entry in the ZIP archive.
   * @param path Relative file path in the ZIP.
   * @returns Raw `Uint8Array` data or `undefined` if file is not found.
   */
  getFileData(path: string): Uint8Array | undefined;

  /**
   * Reads and decodes a text file entry in the ZIP archive as a UTF-8 string.
   * @param path Relative file path in the ZIP (e.g. `'ppt/presentation.xml'`).
   * @returns Decoded UTF-8 string content or `undefined` if file is missing.
   */
  getFileText(path: string): string | undefined;

  /**
   * Returns an array of all decompressed file paths present in the ZIP container.
   * @returns Array of path strings (e.g. `['[Content_Types].xml', 'ppt/presentation.xml']`).
   */
  getPaths(): string[];

  /**
   * Finds all file paths matching a given prefix (e.g., `'ppt/media/'` or `'ppt/slides/'`).
   * @param prefix Path prefix filter string.
   * @returns Array of matching file path strings.
   */
  getPathsStartingWith(prefix: string): string[];

  /**
   * Checks whether a file path exists in the decompressed ZIP archive.
   * @param path Relative file path in the ZIP (e.g. `'ppt/presentation.xml'`).
   * @returns `true` if file exists, `false` otherwise.
   */
  hasFile(path: string): boolean;

  /**
   * Alias for `getPaths()`.
   * @returns Array of file paths in package.
   */
  listFiles(): string[];
}
