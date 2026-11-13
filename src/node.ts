// Node-only helpers. The main entry point stays free of Node APIs so it can run in a browser.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';
import type { BlobStore } from './store.ts';

/** A {@link BlobStore} backed by a folder, e.g. a mounted volume. Keys become file paths under it. */
export function directoryStore(directory: string): BlobStore {
  const root = resolve(directory);
  const locate = (key: string): string => {
    const file = resolve(join(root, key));
    if (file !== root && !file.startsWith(root + sep)) throw new Error(`Store key leaves the store folder: ${key}`);
    return file;
  };
  return {
    async get(key) {
      try {
        return new Uint8Array(await readFile(locate(key)));
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
        throw error;
      }
    },
    async set(key, data) {
      const file = locate(key);
      await mkdir(dirname(file), { recursive: true });
      await writeFile(file, data);
    },
  };
}
