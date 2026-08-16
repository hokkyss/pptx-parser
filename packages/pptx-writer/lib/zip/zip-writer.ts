import { type AsyncZippable, strToU8, zip } from 'fflate';

/**
 * Packages an in-memory dictionary of file paths to strings or Uint8Array buffers into a valid ZIP archive buffer.
 * @param files Key-value map where key is the internal archive path (e.g. `[Content_Types].xml`, `ppt/slides/slide1.xml`).
 * @returns Promise resolving to the compressed ZIP `Uint8Array`.
 */
export function createZipPackage(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const zippable: AsyncZippable = {};

    for (const [path, content] of Object.entries(files)) {
      if (typeof content === 'string') {
        zippable[path] = strToU8(content);
      } else {
        zippable[path] = content;
      }
    }

    zip(zippable, { level: 6 }, (err, data) => {
      if (err) {
        reject(err);
      } else {
        resolve(data);
      }
    });
  });
}
