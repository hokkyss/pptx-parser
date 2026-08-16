import { describe, it, expect } from 'vitest';
import { createZipReader } from '../../lib/zip/zip-reader';

describe('createZipReader', () => {
  it('should instantiate and return empty path list for empty ZIP buffer', async () => {
    const emptyZipBuffer = new Uint8Array([80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    const reader = await createZipReader(emptyZipBuffer);

    expect(reader.listFiles()).toEqual([]);
  });

  it('should return undefined when reading non-existent path', async () => {
    const emptyZipBuffer = new Uint8Array([80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    const reader = await createZipReader(emptyZipBuffer);

    expect(reader.getFileText('ppt/presentation.xml')).toBeUndefined();
    expect(reader.getFileData('ppt/media/image1.png')).toBeUndefined();
  });
});
