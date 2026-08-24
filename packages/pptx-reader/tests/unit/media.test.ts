import type { ZipReader } from '@hokkyss/pptx-core';
import { describe, it, expect } from 'vitest';
import { createMediaResolver, extractMedia } from '../../lib/resolvers/media-resolver';

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

describe('createMediaResolver', () => {
  it('getAllMedia returns empty array initially', () => {
    const resolver = createMediaResolver();
    expect(resolver.getAllMedia()).toEqual([]);
  });

  it('addRelationship and getMediaByRelId map rId to path', () => {
    const resolver = createMediaResolver();
    resolver.addRelationship('rId1', 'ppt/media/image1.png');
    expect(resolver.getMediaByRelId('rId1')).toBe('ppt/media/image1.png');
  });

  it('getMediaByRelId returns undefined for unknown rId', () => {
    const resolver = createMediaResolver();
    expect(resolver.getMediaByRelId('rId99')).toBeUndefined();
  });

  it('loadFromFiles eagerly fetches binary data when lazy=false (default)', async () => {
    const data = new Uint8Array([1, 2, 3]);
    const getBinary = async (_path: string) => data;

    const resolver = createMediaResolver();
    const assets = await resolver.loadFromFiles(['ppt/media/image1.png'], getBinary);

    expect(assets).toHaveLength(1);
    expect(assets[0].filename).toBe('image1.png');
    expect(assets[0].mimeType).toBe('image/png');
    expect(assets[0].data).toEqual(data);
    expect(assets[0].getData).toBeDefined();
  });

  it('loadFromFiles defers binary fetch when lazy=true', async () => {
    const data = new Uint8Array([9, 8, 7]);
    const getBinary = async (_path: string) => data;

    const resolver = createMediaResolver({ lazy: true });
    const assets = await resolver.loadFromFiles(['ppt/media/image2.jpg'], getBinary);

    expect(assets).toHaveLength(1);
    expect(assets[0].data).toBeUndefined();
    expect(assets[0].getData).toBeDefined();
    const fetched = await assets[0].getData!();
    expect(fetched).toEqual(data);
  });

  it('getAllMedia returns the assets loaded by loadFromFiles', async () => {
    const getBinary = async (_path: string) => new Uint8Array([0]);

    const resolver = createMediaResolver();
    await resolver.loadFromFiles(['ppt/media/a.gif', 'ppt/media/b.mp4'], getBinary);

    const all = resolver.getAllMedia();
    expect(all).toHaveLength(2);
    expect(all[0].mimeType).toBe('image/gif');
    expect(all[1].mimeType).toBe('video/mp4');
  });

  it('uses application/octet-stream for unknown extension', async () => {
    const getBinary = async (_path: string) => new Uint8Array([0]);
    const resolver = createMediaResolver();
    const assets = await resolver.loadFromFiles(['ppt/media/file.xyz'], getBinary);
    expect(assets[0].mimeType).toBe('application/octet-stream');
  });

  it('correctly maps all known MIME extensions', async () => {
    const getBinary = async (_path: string) => new Uint8Array([0]);
    const cases: Record<string, string> = {
      'file.bmp': 'image/bmp',
      'file.emf': 'image/x-emf',
      'file.gif': 'image/gif',
      'file.jpeg': 'image/jpeg',
      'file.jpg': 'image/jpeg',
      'file.m4a': 'audio/mp4',
      'file.mp3': 'audio/mpeg',
      'file.mp4': 'video/mp4',
      'file.png': 'image/png',
      'file.svg': 'image/svg+xml',
      'file.wav': 'audio/wav',
      'file.webm': 'video/webm',
      'file.wmf': 'image/x-wmf',
    };
    for (const [filename, expectedMime] of Object.entries(cases)) {
      const resolver = createMediaResolver();
      const assets = await resolver.loadFromFiles([`ppt/media/${filename}`], getBinary);
      expect(assets[0].mimeType, `MIME for ${filename}`).toBe(expectedMime);
    }
  });

  it('loadFromFiles resets the media list on each call', async () => {
    const getBinary = async (_path: string) => new Uint8Array([0]);
    const resolver = createMediaResolver();
    await resolver.loadFromFiles(['ppt/media/a.png'], getBinary);
    await resolver.loadFromFiles(['ppt/media/b.png', 'ppt/media/c.png'], getBinary);
    expect(resolver.getAllMedia()).toHaveLength(2);
  });
});

describe('extractMedia extended coverage', () => {
  it('covers getData callback and null fallback in eager extractMedia', async () => {
    const mockZip = {
      getPathsStartingWith: () => ['ppt/media/test.png'],
      getFileData: () => null,
      getFileAsBinary: () => null,
    } as unknown as ZipReader;

    const media = extractMedia(mockZip, false);
    expect(media[0].data).toEqual(new Uint8Array(0));
    expect(await media[0].getData!()).toBeNull();
  });
});
