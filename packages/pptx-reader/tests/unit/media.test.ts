import type { ZipReader } from '@hokkyss/pptx-core';
import { describe, it, expect } from 'vitest';
import { extractMedia } from '../../lib/resolvers/media-resolver';

describe('extractMedia', () => {
  it('should return empty list when ppt/media is missing', () => {
    const mockZip = {
      getPathsStartingWith: () => [],
      getFileData: () => undefined,
      getFileAsBinary: () => undefined,
    } as unknown as ZipReader;

    const media = extractMedia(mockZip);
    expect(media).toEqual([]);
  });

  it('should discover and extract binary media files eagerly', () => {
    const mockZip = {
      getPathsStartingWith: () => ['ppt/media/image1.png', 'ppt/media/sample.mp4'],
      getFileData: () => new Uint8Array([1, 2, 3, 4]),
      getFileAsBinary: () => new Uint8Array([1, 2, 3, 4]),
    } as unknown as ZipReader;

    const media = extractMedia(mockZip, false);
    expect(media.length).toBe(2);

    expect(media[0].filename).toBe('image1.png');
    expect(media[0].mimeType).toBe('image/png');
    expect(media[0].data).toEqual(new Uint8Array([1, 2, 3, 4]));

    expect(media[1].filename).toBe('sample.mp4');
    expect(media[1].mimeType).toBe('video/mp4');
  });

  it('should support lazy loading getters when lazy option is true', async () => {
    const mockZip = {
      getPathsStartingWith: () => ['ppt/media/image2.jpg'],
      getFileData: () => new Uint8Array([9, 9, 9]),
      getFileAsBinary: () => new Uint8Array([9, 9, 9]),
    } as unknown as ZipReader;

    const media = extractMedia(mockZip, true);
    expect(media[0].data).toBeNull();
    expect(media[0].lazyGetter).toBeDefined();

    const loadedData = await media[0].lazyGetter!();
    expect(loadedData).toEqual(new Uint8Array([9, 9, 9]));
  });
});
