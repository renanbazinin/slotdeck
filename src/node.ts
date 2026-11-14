// Node-only helpers. The main entry point stays free of Node APIs so it can run in a browser.
import { readdirSync, readFileSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';
import type { BlobStore } from './store.ts';
import type { FontFiles, FontkitModule } from './text/fontkit.ts';

/**
 * Finds the font files of the given typefaces in a folder, such as the system's font folder, by
 * the family name each file states. The result goes to `fontkitMetrics`. A typeface that has no
 * file in the folder is left out; every weight and slant found of the others is included.
 */
export function fontFiles(fontkit: FontkitModule, directory: string, typefaces: Iterable<string>): FontFiles {
  const wanted = new Map([...typefaces].map((name) => [name.toLowerCase(), name]));
  const found: Record<string, Uint8Array[]> = {};
  for (const name of readdirSync(directory).sort()) {
    if (!/\.(ttf|otf|ttc|woff2?)$/i.test(name)) continue;
    const bytes = new Uint8Array(readFileSync(join(directory, name)));
    let families: string[];
    try {
      const opened = (fontkit.create as (buffer: Uint8Array) => unknown)(bytes) as { familyName?: string | null; fonts?: Array<{ familyName?: string | null }> };
      families = (opened.fonts ?? [opened]).map((font) => (font.familyName ?? '').toLowerCase());
    } catch {
      continue; // not a font this library reads
    }
    for (const family of new Set(families)) {
      const typeface = wanted.get(family);
      if (typeface) (found[typeface] ??= []).push(bytes);
    }
  }
  return found;
}

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
