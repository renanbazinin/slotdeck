// Node-only helpers. The main entry point stays free of Node APIs so it can run in a browser.
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve, sep } from 'node:path';
import type { AssetDef, IconDef } from './plan/catalog.ts';
import type { BlobStore } from './store.ts';
import type { FontFiles, FontkitModule } from './text/fontkit.ts';
import { readImageInfo } from './util/image.ts';

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

/** How a folder of pictures is read (`iconFiles`, `assetFiles`). From 0.8.0. */
export interface PictureFolderOptions {
  /**
   * What each picture shows, by its id: the model chooses a picture by this line and nothing else.
   * An object such as `{ "clock": "a clock: time, schedule" }`, or the name of a JSON file in the
   * folder that holds one. Such a file may also be a list of `{ id, description }`, each with a
   * `file` where a picture's file is not named after its id. Left out, the folder's `icons.json`
   * (`iconFiles`) or `assets.json` (`assetFiles`) is read where it has one. A picture that no line
   * describes is described by the words of its file name.
   */
  descriptions?: Record<string, string> | string;
  /**
   * Also reads the folders inside the folder. An id is still the picture's own file name, so two
   * pictures of one name in different folders are an error.
   */
  recursive?: boolean;
  /**
   * Told of every file that is passed over, with the reason: a picture in a format that is not
   * taken (SVG, WebP), a file named as a picture that is none, a line of the descriptions that
   * names no picture of the folder. Left out, such files are passed over in silence.
   */
  onSkipped?: (file: string, reason: string) => void;
}

/** A picture's file name ends so: the three formats a deck's pictures are read in. */
const TAKEN = /\.(png|jpe?g|gif)$/i;
/** Pictures in a format that is not taken: said, where someone asks (`onSkipped`). */
const NOT_TAKEN = /\.(svg|webp|bmp|ico|tiff?|avif|heic|emf|wmf|eps|ai|psd)$/i;

/**
 * A line of words from a file name, for a picture nobody described: `money-bag_2` is "money bag
 * 2", and `ClockFace` "Clock Face". A start, not a description: say what the picture stands for.
 */
const wordsOf = (name: string): string =>
  name
    .replace(/([\p{Ll}\p{Nd}])(\p{Lu})/gu, '$1 $2')
    .replace(/[-_.\s]+/g, ' ')
    .trim();

interface FolderPicture {
  id: string;
  description: string;
  data: Uint8Array;
}

/** Reads the pictures of a folder (see `iconFiles`). `list` is the file of descriptions read where none is named. */
function pictureFiles(directory: string, list: string, options: PictureFolderOptions): FolderPicture[] {
  const root = resolve(directory);
  const skipped = options.onSkipped ?? (() => {});

  // The lines that describe pictures: by id, and by file name for a line that names its file.
  const byId = new Map<string, string>();
  const byFile = new Map<string, { id: string; description: string }>();
  const given = options.descriptions;
  const listed = typeof given === 'string' ? given : given === undefined ? list : undefined;
  if (given !== undefined && typeof given !== 'string') {
    for (const [id, description] of Object.entries(given)) if (typeof description === 'string') byId.set(id, description);
  } else if (listed !== undefined) {
    const file = join(root, listed);
    // A list that was named has to be there; the usual one is read where the folder has it.
    if (existsSync(file) || typeof given === 'string') {
      let read: unknown;
      try {
        read = JSON.parse(readFileSync(file, 'utf8'));
      } catch (error) {
        throw new Error(`${file} cannot be read as the descriptions of the pictures in its folder: ${error instanceof Error ? error.message : String(error)}`);
      }
      const entries: unknown[] = Array.isArray(read) ? read : typeof read === 'object' && read !== null ? Object.entries(read).map(([id, description]) => ({ id, description })) : [];
      for (const entry of entries) {
        const line = (typeof entry === 'object' && entry !== null ? entry : {}) as { id?: unknown; description?: unknown; file?: unknown };
        if (typeof line.id !== 'string' || line.id === '' || typeof line.description !== 'string') continue;
        byId.set(line.id, line.description);
        if (typeof line.file === 'string' && line.file !== '') byFile.set(line.file.replace(/\\/g, '/').toLowerCase(), { id: line.id, description: line.description });
      }
    }
  }

  const found = new Map<string, FolderPicture & { file: string }>();
  const read = (folder: string, within: string): void => {
    for (const entry of readdirSync(folder, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0))) {
      const file = within ? `${within}/${entry.name}` : entry.name;
      if (entry.isDirectory()) {
        if (options.recursive) read(join(folder, entry.name), file);
        continue;
      }
      if (!entry.isFile()) continue;
      if (!TAKEN.test(entry.name)) {
        if (NOT_TAKEN.test(entry.name)) skipped(file, 'pictures are read as PNG, JPEG or GIF: save it as a PNG');
        continue;
      }
      const data = new Uint8Array(readFileSync(join(folder, entry.name)));
      try {
        readImageInfo(data);
      } catch {
        skipped(file, 'it is named as a picture, and its content is no PNG, JPEG or GIF');
        continue;
      }
      const name = entry.name.replace(/\.[^.]+$/, '');
      const named = byFile.get(file.toLowerCase()) ?? byFile.get(entry.name.toLowerCase());
      const id = named?.id ?? name;
      const before = found.get(id);
      if (before) throw new Error(`${root} has two pictures with the id "${id}": ${before.file} and ${file}. A picture is known by its id, so each needs one of its own`);
      found.set(id, { id, description: (named?.description ?? byId.get(id) ?? '').trim() || wordsOf(name), data, file });
    }
  };
  read(root, '');
  for (const id of byId.keys()) if (!found.has(id)) skipped(listed ?? id, `it describes "${id}", and the folder has no such picture`);
  return [...found.values()].sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0)).map(({ id, description, data }) => ({ id, description, data }));
}

/**
 * An app's own icons, read from a folder: every PNG, JPEG and GIF in it is an icon, known by its
 * file name without the ending (`clock.png` is `clock`). The result goes to `icons` of
 * `presentationTools`, `generateDeck`, `planDeck`, `resolveAssets` and the other calls that take
 * an icon set. A model chooses an icon by its description alone, so describe each one: what it
 * shows and what it stands for ("a clock: time, schedule, duration"), in the folder's
 * `icons.json` or in `descriptions` (see {@link PictureFolderOptions}). An icon without a line
 * is described by the words of its file name, which is a start. The icons come back in the order
 * of their ids, the same on every machine. From 0.8.0.
 */
export function iconFiles(directory: string, options: PictureFolderOptions = {}): IconDef[] {
  return pictureFiles(directory, 'icons.json', options);
}

/**
 * An app's own pictures of one kind, read from a folder as `iconFiles` reads icons: `photo` for
 * photographs, which fill their frame and may be cropped, `graphic` for pictures shown whole,
 * such as diagrams, screenshots and logos. The result goes to `assets`. Their descriptions are
 * read from the folder's `assets.json` where nothing else is named. From 0.8.0.
 */
export function assetFiles(directory: string, kind: AssetDef['kind'], options: PictureFolderOptions = {}): AssetDef[] {
  return pictureFiles(directory, 'assets.json', options).map((picture) => ({ id: picture.id, kind, description: picture.description, data: picture.data }));
}
