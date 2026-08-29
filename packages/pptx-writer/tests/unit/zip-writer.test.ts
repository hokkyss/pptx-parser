import { describe, expect, it } from 'vitest';
import { createZipPackage } from '../../lib/zip/zip-writer';

describe('ZIP Writer', () => {
  it('packages files into a valid ZIP archive buffer', async () => {
    const files: Record<string, string | Uint8Array> = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types/>',
      'docProps/core.xml': '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties/>',
      'ppt/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
    };

    const zipBuffer = await createZipPackage(files);
    expect(zipBuffer).toBeInstanceOf(Uint8Array);
    expect(zipBuffer.length).toBeGreaterThan(0);
    // ZIP magic bytes: PK\x03\x04
    expect(zipBuffer[0]).toBe(0x50);
    expect(zipBuffer[1]).toBe(0x4b);
    expect(zipBuffer[2]).toBe(0x03);
    expect(zipBuffer[3]).toBe(0x04);
  });

  it('decodes base64 string to bytes in both Buffer and atob environments', async () => {
    const { decodeBase64ToBytes } = await import('../../lib/writer');
    const originalBuffer = globalThis.Buffer;
    const base64 = 'AAAA';
    const bytesWithBuffer = decodeBase64ToBytes(base64);
    expect(bytesWithBuffer).toEqual(new Uint8Array([0, 0, 0]));

    // Test atob fallback
    (globalThis as any).Buffer = undefined;
    try {
      const bytesWithAtob = decodeBase64ToBytes(base64);
      expect(bytesWithAtob).toEqual(new Uint8Array([0, 0, 0]));
    } finally {
      globalThis.Buffer = originalBuffer;
    }
  });
});
