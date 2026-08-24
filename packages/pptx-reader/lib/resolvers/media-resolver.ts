import type { PptxMediaAsset } from '@hokkyss/pptx-core';
import type { ZipReader } from '@hokkyss/pptx-core';

/**
 * Extended `PptxMediaAsset` interface supporting an optional `getData()` getter function.
 */
export interface PptxMediaAssetExtended extends PptxMediaAsset {
  /** Asynchronous getter resolving the raw binary buffer of the media asset */
  getData?: () => Promise<null | Uint8Array>;
}

/**
 * Maps file extensions to standard MIME content types for PPTX media assets.
 */
const EXTENSION_TO_MIME: Record<string, string> = {
  bmp: 'image/bmp',
  emf: 'image/x-emf',
  gif: 'image/gif',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  m4a: 'audio/mp4',
  mp3: 'audio/mpeg',
  mp4: 'video/mp4',
  png: 'image/png',
  svg: 'image/svg+xml',
  wav: 'audio/wav',
  webm: 'video/webm',
  wmf: 'image/x-wmf',
};

/**
 * Discovers and extracts all media files stored in `ppt/media/` inside the ZIP archive.
 * @param zipReader Active `ZipReader` instance.
 * @param lazy Whether to defer binary buffer allocation using lazy getters (defaults to `false`).
 * @returns Array of parsed `PptxMediaAssetExtended` assets with detected MIME types and binary data/getters.
 * @example
 * ```ts
 * const assets = extractMedia(zipReader, false);
 * console.log(`Extracted ${assets.length} media assets.`);
 * ```
 */
export function extractMedia(zipReader: ZipReader, lazy: boolean = false): PptxMediaAssetExtended[] {
  const mediaPaths = zipReader.getPathsStartingWith('ppt/media/');
  const assets: PptxMediaAssetExtended[] = [];

  for (const path of mediaPaths) {
    const filename = path.split('/').pop() || path;
    const ext = filename.split('.').pop()?.toLowerCase() || '';
    const mimeType = EXTENSION_TO_MIME[ext] || 'application/octet-stream';
    const id = path;

    const getData = async () => Promise.resolve(zipReader.getFileData(path) || null);

    if (lazy) {
      assets.push({
        data: null,
        filename,
        getData,
        id,
        lazyGetter: getData,
        mimeType,
        path,
      });
    } else {
      const data = zipReader.getFileData(path) || new Uint8Array(0);
      assets.push({
        data,
        filename,
        getData,
        id,
        lazyGetter: getData,
        mimeType,
        path,
      });
    }
  }

  return assets;
}

export interface MediaResolver {
  addRelationship(rId: string, mediaPath: string): void;
  getAllMedia(): PptxMediaAssetExtended[];
  getMediaByRelId(rId: string): string | undefined;
  loadFromFiles(paths: string[], getBinary: (path: string) => Promise<Uint8Array>): Promise<PptxMediaAssetExtended[]>;
}

/**
 * Creates a MediaResolver for custom media loading and relationship mapping.
 * @param options Optional settings including `lazy` flag.
 * @returns Frozen `MediaResolver` instance.
 */
export function createMediaResolver(options?: { lazy?: boolean }): MediaResolver {
  const isLazy = Boolean(options?.lazy);
  let mediaList: PptxMediaAssetExtended[] = [];
  const relMap = new Map<string, string>();

  return Object.freeze({
    addRelationship(rId: string, mediaPath: string): void {
      relMap.set(rId, mediaPath);
    },
    getAllMedia(): PptxMediaAssetExtended[] {
      return mediaList;
    },
    getMediaByRelId(rId: string): string | undefined {
      return relMap.get(rId);
    },
    async loadFromFiles(
      paths: string[],
      getBinary: (path: string) => Promise<Uint8Array>,
    ): Promise<PptxMediaAssetExtended[]> {
      mediaList = [];
      for (const path of paths) {
        const filename = path.split('/').pop() || path;
        const ext = filename.split('.').pop()?.toLowerCase() || '';
        const mimeType = EXTENSION_TO_MIME[ext] || 'application/octet-stream';
        const id = path;

        const getData = async () => getBinary(path);

        if (isLazy) {
          mediaList.push({
            data: undefined as unknown as null,
            filename,
            getData,
            id,
            lazyGetter: getData,
            mimeType,
            path,
          });
        } else {
          const data = await getBinary(path);
          mediaList.push({
            data,
            filename,
            getData,
            id,
            lazyGetter: getData,
            mimeType,
            path,
          });
        }
      }
      return mediaList;
    },
  });
}
