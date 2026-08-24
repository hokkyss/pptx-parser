import { zipSync, strToU8 } from 'fflate';
import { describe, it, expect } from 'vitest';
import { createZipReader, sanitizeZipPath } from '../../lib/zip/zip-reader';

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

  describe('Security & Hardening Guardrails', () => {
    it('should sanitize path traversal attempts (Zip Slip defense)', () => {
      expect(sanitizeZipPath('../../etc/passwd')).toBe('etc/passwd');
      expect(sanitizeZipPath('ppt/../../secret.xml')).toBe('secret.xml');
      expect(sanitizeZipPath('/ppt/slides/slide1.xml')).toBe('ppt/slides/slide1.xml');
      expect(sanitizeZipPath('ppt\\slides\\slide1.xml')).toBe('ppt/slides/slide1.xml');
      expect(sanitizeZipPath('')).toBe('');
    });

    it('should enforce maxSingleFileBytes security limit', async () => {
      const mockZip = zipSync({
        'large.xml': strToU8('x'.repeat(2000)),
      });

      await expect(
        createZipReader(mockZip, { maxSingleFileBytes: 1000 }),
      ).rejects.toThrow(/Security limit exceeded: File entry "large.xml" is 2000 bytes/);
    });

    it('should enforce maxTotalBytes security limit against Zip bombs', async () => {
      const mockZip = zipSync({
        'f1.xml': strToU8('a'.repeat(600)),
        'f2.xml': strToU8('b'.repeat(600)),
      });

      await expect(
        createZipReader(mockZip, { maxTotalBytes: 1000 }),
      ).rejects.toThrow(/Security limit exceeded: Total uncompressed archive size exceeded/);
    });

    it('should enforce maxEntries security limit', async () => {
      const mockZip = zipSync({
        '1.xml': strToU8('a'),
        '2.xml': strToU8('b'),
        '3.xml': strToU8('c'),
      });

      await expect(
        createZipReader(mockZip, { maxEntries: 2 }),
      ).rejects.toThrow(/Security limit exceeded: ZIP archive contains 3 entries/);
    });
  });
});

describe('createZipReader extended API', () => {
  it('handles getFileAsString, hasFile, getPathsStartingWith, and error handling', async () => {
    const mockZip = zipSync({
      'ppt/slides/slide1.xml': strToU8('<slide1/>'),
      'ppt/media/img.png': strToU8('PNG'),
    });

    const reader = await createZipReader(mockZip);
    expect(reader.hasFile('ppt/slides/slide1.xml')).toBe(true);
    expect(reader.hasFile('/ppt/slides/slide1.xml')).toBe(true);
    expect(reader.hasFile('nonexistent')).toBe(false);

    expect(await reader.getFileAsString('ppt/slides/slide1.xml')).toBe('<slide1/>');
    expect(await reader.getFileAsString('nonexistent')).toBe('');

    expect(reader.getPathsStartingWith('ppt/media/')).toEqual(['ppt/media/img.png']);
    expect(reader.getPathsStartingWith('/ppt/media/')).toEqual(['ppt/media/img.png']);

    // Invalid non-zip buffer
    await expect(createZipReader(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow('Failed to unzip PPTX file container');

    // Incomplete mock buffer starting with PK\x03\x04
    const pkMock = new Uint8Array([80, 75, 3, 4, 0, 0]);
    const mockReader = await createZipReader(pkMock);
    expect(mockReader.listFiles()).toEqual([]);
  });
});

describe('createZipReader all-zero buffer rejection', () => {
  it('throws error for all-zero corrupt buffer', async () => {
    await expect(createZipReader(new Uint8Array([0, 0, 0, 0]))).rejects.toThrow('Invalid or corrupt PPTX container');
  });
});
