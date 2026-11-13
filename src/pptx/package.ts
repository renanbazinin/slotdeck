import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { Zippable } from 'fflate';
import { NS, XML_HEADER, childElements, escapeXmlAttr, parseXml, rootElement } from './xml.ts';

const CONTENT_TYPES = '[Content_Types].xml';

/** Parts that are already compressed; deflating them again only costs time. */
const STORE_ONLY = /\.(png|jpe?g|gif|webp|fntdata|mp4|m4a|mp3|wmv|xlsx|docx|zip)$/i;

/** An Office Open XML package: a flat map of part path to bytes. */
export class PptxPackage {
  private readonly files: Map<string, Uint8Array>;

  private constructor(files: Map<string, Uint8Array>) {
    this.files = files;
  }

  static load(bytes: Uint8Array): PptxPackage {
    const entries = unzipSync(bytes);
    const files = new Map<string, Uint8Array>();
    for (const [name, data] of Object.entries(entries)) {
      if (name.endsWith('/')) continue;
      files.set(name, data);
    }
    if (!files.has(CONTENT_TYPES)) {
      throw new Error('Not an Office Open XML package: [Content_Types].xml is missing');
    }
    return new PptxPackage(files);
  }

  /** Cheap copy: the byte arrays are shared and must be treated as immutable. */
  clone(): PptxPackage {
    return new PptxPackage(new Map(this.files));
  }

  has(partPath: string): boolean {
    return this.files.has(partPath);
  }

  paths(): string[] {
    return [...this.files.keys()];
  }

  read(partPath: string): Uint8Array {
    const data = this.files.get(partPath);
    if (!data) throw new Error(`Part not found in package: ${partPath}`);
    return data;
  }

  readText(partPath: string): string {
    const text = strFromU8(this.read(partPath));
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  }

  write(partPath: string, data: Uint8Array | string): void {
    this.files.set(partPath, typeof data === 'string' ? strToU8(data) : data);
  }

  delete(partPath: string): void {
    this.files.delete(partPath);
  }

  save(): Uint8Array {
    const zippable: Zippable = {};
    const names = [CONTENT_TYPES, ...this.paths().filter((name) => name !== CONTENT_TYPES)];
    for (const name of names) {
      zippable[name] = [this.read(name), { level: STORE_ONLY.test(name) ? 0 : 6 }];
    }
    // Entries carry a fixed timestamp, so the same content always gives the same bytes.
    return zipSync(zippable, { mtime: new Date(1980, 0, 1) });
  }
}

// ---------------------------------------------------------------------------
// Relationships

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/** `ppt/slides/slide1.xml` -> `ppt/slides/_rels/slide1.xml.rels`; the package root is `''`. */
export function relsPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Inverse of {@link relsPath}, or `undefined` when the path is not a relationships part. */
export function relsOwner(path: string): string | undefined {
  const match = /^(.*\/)?_rels\/([^/]*)\.rels$/.exec(path);
  return match ? `${match[1] ?? ''}${match[2]}` : undefined;
}

export function readRels(pkg: PptxPackage, partPath: string): Relationship[] {
  const file = relsPath(partPath);
  if (!pkg.has(file)) return [];
  const root = rootElement(parseXml(pkg.readText(file)));
  return childElements(root, NS.rels, 'Relationship').map((el) => ({
    id: el.getAttribute('Id') ?? '',
    type: el.getAttribute('Type') ?? '',
    target: el.getAttribute('Target') ?? '',
    external: el.getAttribute('TargetMode') === 'External',
  }));
}

export function writeRels(pkg: PptxPackage, partPath: string, rels: Relationship[]): void {
  const body = rels
    .map(
      (rel) =>
        `<Relationship Id="${escapeXmlAttr(rel.id)}" Type="${escapeXmlAttr(rel.type)}" Target="${escapeXmlAttr(rel.target)}"` +
        (rel.external ? ' TargetMode="External"' : '') +
        '/>',
    )
    .join('');
  pkg.write(relsPath(partPath), `${XML_HEADER}<Relationships xmlns="${NS.rels}">${body}</Relationships>`);
}

/** True when the relationship type ends with the given name, e.g. `isRelType(rel, 'slide')`. */
export function isRelType(rel: Relationship, name: string): boolean {
  return rel.type.endsWith(`/${name}`);
}

export function nextRelId(rels: Relationship[]): string {
  let max = 0;
  for (const rel of rels) {
    const match = /^rId(\d+)$/.exec(rel.id);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `rId${max + 1}`;
}

/** Resolves a relationship target against the part that owns it. */
export function resolveTarget(partPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segments = partPath.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') segments.pop();
    else if (segment !== '.' && segment !== '') segments.push(segment);
  }
  return segments.join('/');
}

/** The relative target that points from one part to another. */
export function relativeTarget(fromPart: string, toPart: string): string {
  const from = fromPart.split('/').slice(0, -1);
  const to = toPart.split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
  return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

// ---------------------------------------------------------------------------
// Content types

export class ContentTypes {
  private readonly defaults = new Map<string, string>();
  private readonly overrides = new Map<string, string>();

  static read(pkg: PptxPackage): ContentTypes {
    const types = new ContentTypes();
    const root = rootElement(parseXml(pkg.readText(CONTENT_TYPES)));
    for (const el of childElements(root, NS.contentTypes, 'Default')) {
      types.defaults.set((el.getAttribute('Extension') ?? '').toLowerCase(), el.getAttribute('ContentType') ?? '');
    }
    for (const el of childElements(root, NS.contentTypes, 'Override')) {
      types.overrides.set(el.getAttribute('PartName') ?? '', el.getAttribute('ContentType') ?? '');
    }
    return types;
  }

  ensureDefault(extension: string, contentType: string): void {
    const key = extension.toLowerCase();
    if (!this.defaults.has(key)) this.defaults.set(key, contentType);
  }

  /** `partPath` is the package path without a leading slash. */
  setOverride(partPath: string, contentType: string): void {
    this.overrides.set(`/${partPath}`, contentType);
  }

  removeOverride(partPath: string): void {
    this.overrides.delete(`/${partPath}`);
  }

  save(pkg: PptxPackage): void {
    const defaults = [...this.defaults]
      .map(([ext, type]) => `<Default Extension="${escapeXmlAttr(ext)}" ContentType="${escapeXmlAttr(type)}"/>`)
      .join('');
    const overrides = [...this.overrides]
      .map(([part, type]) => `<Override PartName="${escapeXmlAttr(part)}" ContentType="${escapeXmlAttr(type)}"/>`)
      .join('');
    pkg.write(CONTENT_TYPES, `${XML_HEADER}<Types xmlns="${NS.contentTypes}">${defaults}${overrides}</Types>`);
  }
}

// ---------------------------------------------------------------------------
// Reachability

/** Every part that can be reached from the package root by following relationships. */
export function reachableParts(pkg: PptxPackage): Set<string> {
  const seen = new Set<string>();
  const queue = [''];
  while (queue.length > 0) {
    const part = queue.pop() as string;
    for (const rel of readRels(pkg, part)) {
      if (rel.external) continue;
      const target = resolveTarget(part, rel.target);
      if (seen.has(target) || !pkg.has(target)) continue;
      seen.add(target);
      queue.push(target);
    }
  }
  return seen;
}

/** Deletes parts nothing points to any more (old slides, their notes, unused media). */
export function pruneUnreachable(pkg: PptxPackage, types: ContentTypes): void {
  const reachable = reachableParts(pkg);
  for (const file of pkg.paths()) {
    if (file === CONTENT_TYPES) continue;
    const owner = relsOwner(file);
    if (owner !== undefined) {
      if (owner !== '' && !reachable.has(owner)) pkg.delete(file);
      continue;
    }
    if (!reachable.has(file)) {
      pkg.delete(file);
      types.removeOverride(file);
    }
  }
}
