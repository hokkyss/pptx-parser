import type { AsyncZipOptions } from 'fflate';
import { describe, expect, it, vi } from 'vitest';
import { createZipPackage } from '../../lib/zip/zip-writer';

vi.mock('fflate', () => ({
  strToU8: (str: string) => new TextEncoder().encode(str),
  zip: (
    _zippable: Record<string, Uint8Array>,
    _opts: AsyncZipOptions,
    cb: (err: Error | null, data: Uint8Array) => void,
  ) => {
    cb(new Error('Forced ZIP failure'), new Uint8Array());
  },
}));

describe('ZIP Writer error callback', () => {
  it('rejects when fflate zip yields an error', async () => {
    await expect(createZipPackage({ 'test.txt': 'hello' })).rejects.toThrow('Forced ZIP failure');
  });
});
