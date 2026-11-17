import { unzlibSync, zlibSync } from 'fflate';

/** A flat-colour PNG: a stand-in picture that needs no image model and works in any runtime. */
export function solidPng(width: number, height: number, [red, green, blue]: [number, number, number]): Uint8Array {
  const stride = width * 3 + 1;
  const raw = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const offset = y * stride + 1 + x * 3;
      raw[offset] = red;
      raw[offset + 1] = green;
      raw[offset + 2] = blue;
    }
  }
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = 8; // bit depth
  header[9] = 2; // colour type: RGB

  const chunks = [chunk('IHDR', header), chunk('IDAT', zlibSync(raw)), chunk('IEND', new Uint8Array(0))];
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  const png = new Uint8Array(signature.length + chunks.reduce((sum, part) => sum + part.length, 0));
  png.set(signature);
  let offset = signature.length;
  for (const part of chunks) {
    png.set(part, offset);
    offset += part.length;
  }
  return png;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

interface PngChunk {
  type: string;
  data: Uint8Array;
}

function readChunks(png: Uint8Array): PngChunk[] | undefined {
  if (png.length < 8 || SIGNATURE.some((byte, index) => png[index] !== byte)) return undefined;
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const chunks: PngChunk[] = [];
  for (let offset = 8; offset + 12 <= png.length; ) {
    const length = view.getUint32(offset);
    if (offset + 12 + length > png.length) return undefined;
    const type = String.fromCharCode(png[offset + 4] as number, png[offset + 5] as number, png[offset + 6] as number, png[offset + 7] as number);
    chunks.push({ type, data: png.subarray(offset + 8, offset + 8 + length) });
    offset += 12 + length;
    if (type === 'IEND') return chunks;
  }
  return undefined;
}

function writeChunks(chunks: PngChunk[]): Uint8Array {
  const parts = chunks.map((part) => chunk(part.type, part.data));
  const png = new Uint8Array(SIGNATURE.length + parts.reduce((sum, part) => sum + part.length, 0));
  png.set(SIGNATURE);
  let offset = SIGNATURE.length;
  for (const part of parts) {
    png.set(part, offset);
    offset += part.length;
  }
  return png;
}

/** The lines of an 8-bit RGB or RGBA picture as plain bytes, with the difference coding of each line undone. */
function plainLines(chunks: PngChunk[], width: number, height: number, bytes: number): Uint8Array | undefined {
  const stride = width * bytes;
  const packed = chunks.filter((part) => part.type === 'IDAT');
  const joined = new Uint8Array(packed.reduce((sum, part) => sum + part.data.length, 0));
  let at = 0;
  for (const part of packed) {
    joined.set(part.data, at);
    at += part.data.length;
  }
  let raw: Uint8Array;
  try {
    raw = unzlibSync(joined);
  } catch {
    return undefined;
  }
  if (raw.length < (stride + 1) * height) return undefined;

  const out = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)] as number;
    if (filter > 4) return undefined;
    const from = y * (stride + 1) + 1;
    const to = y * stride;
    for (let x = 0; x < stride; x++) {
      const left = x >= bytes ? (out[to + x - bytes] as number) : 0;
      const up = y > 0 ? (out[to + x - stride] as number) : 0;
      const upLeft = y > 0 && x >= bytes ? (out[to + x - stride - bytes] as number) : 0;
      let predicted = 0;
      if (filter === 1) predicted = left;
      else if (filter === 2) predicted = up;
      else if (filter === 3) predicted = (left + up) >> 1;
      else if (filter === 4) {
        const estimate = left + up - upLeft;
        const byLeft = Math.abs(estimate - left);
        const byUp = Math.abs(estimate - up);
        const byUpLeft = Math.abs(estimate - upLeft);
        predicted = byLeft <= byUp && byLeft <= byUpLeft ? left : byUp <= byUpLeft ? up : upLeft;
      }
      out[to + x] = ((raw[from + x] as number) + predicted) & 0xff;
    }
  }
  return out;
}

interface PngShape {
  chunks: PngChunk[];
  width: number;
  height: number;
  depth: number;
  colorType: number;
  interlaced: boolean;
}

function readShape(png: Uint8Array): PngShape | undefined {
  const chunks = readChunks(png);
  const header = chunks?.[0];
  if (!chunks || !header || header.type !== 'IHDR' || header.data.length < 13) return undefined;
  const view = new DataView(header.data.buffer, header.data.byteOffset, header.data.byteLength);
  return { chunks, width: view.getUint32(0), height: view.getUint32(4), depth: header.data[8] as number, colorType: header.data[9] as number, interlaced: header.data[12] !== 0 };
}

/**
 * A PNG with every colour passed through `map` (red, green and blue 0..255 in, the same out);
 * transparency is kept as it is. `undefined` when no colour changes, or when the file is of a
 * kind this does not read: sixteen bits a channel, interlaced, or not a PNG. A greyscale file
 * can only hold greys, so it is left alone.
 *
 * A picture with a colour table has only its table rewritten. Any other is unpacked, mapped
 * pixel by pixel and packed again, so the same picture and map always give the same bytes.
 */
export function mapPngColors(png: Uint8Array, map: (red: number, green: number, blue: number) => readonly [number, number, number]): Uint8Array | undefined {
  const shape = readShape(png);
  if (!shape) return undefined;
  const { chunks, width, height, depth, colorType, interlaced } = shape;
  // An animated picture, and one written for a phone's own reader, are stored in ways this does not read.
  if (chunks.some((part) => part.type === 'acTL' || part.type === 'CgBI')) return undefined;

  // The same colour is asked for once: a pictogram has a handful of them.
  const known = new Map<number, number>();
  let changed = false;
  const mapped = (red: number, green: number, blue: number): number => {
    const key = (red << 16) | (green << 8) | blue;
    let value = known.get(key);
    if (value === undefined) {
      const [r, g, b] = map(red, green, blue);
      value = ((r & 0xff) << 16) | ((g & 0xff) << 8) | (b & 0xff);
      known.set(key, value);
    }
    if (value !== key) changed = true;
    return value;
  };

  if (colorType === 3) {
    const table = chunks.find((part) => part.type === 'PLTE');
    if (!table || table.data.length % 3 !== 0) return undefined;
    const next = new Uint8Array(table.data.length);
    for (let offset = 0; offset < next.length; offset += 3) {
      const value = mapped(table.data[offset] as number, table.data[offset + 1] as number, table.data[offset + 2] as number);
      next[offset] = value >> 16;
      next[offset + 1] = (value >> 8) & 0xff;
      next[offset + 2] = value & 0xff;
    }
    return changed ? writeChunks(chunks.map((part) => (part === table ? { type: 'PLTE', data: next } : part))) : undefined;
  }

  if ((colorType !== 2 && colorType !== 6) || depth !== 8 || interlaced || width === 0 || height === 0 || width * height > 4_000_000) return undefined;
  // A picture without an alpha channel may name one colour as see-through: changing colours would lose that.
  if (colorType === 2 && chunks.some((part) => part.type === 'tRNS')) return undefined;
  const bytes = colorType === 6 ? 4 : 3;
  const stride = width * bytes;
  const plain = plainLines(chunks, width, height, bytes);
  if (!plain) return undefined;

  // The mapped lines are stored plain: a zero before each says so.
  const out = new Uint8Array((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    const from = y * stride;
    const to = y * (stride + 1) + 1;
    for (let x = 0; x < stride; x += bytes) {
      const value = mapped(plain[from + x] as number, plain[from + x + 1] as number, plain[from + x + 2] as number);
      out[to + x] = value >> 16;
      out[to + x + 1] = (value >> 8) & 0xff;
      out[to + x + 2] = value & 0xff;
      if (bytes === 4) out[to + x + 3] = plain[from + x + 3] as number;
    }
  }
  if (!changed) return undefined;

  const first = chunks.findIndex((part) => part.type === 'IDAT');
  const rest = (parts: PngChunk[]) => parts.filter((part) => part.type !== 'IDAT');
  return writeChunks([...rest(chunks.slice(0, first)), { type: 'IDAT', data: zlibSync(out) }, ...rest(chunks.slice(first))]);
}

/** The pixels of an 8-bit RGB or RGBA PNG that is not interlaced: red, green, blue and alpha for each. `undefined` for any other file. */
export function readPngPixels(png: Uint8Array): { width: number; height: number; pixels: Uint8Array } | undefined {
  const shape = readShape(png);
  if (!shape || (shape.colorType !== 2 && shape.colorType !== 6) || shape.depth !== 8 || shape.interlaced) return undefined;
  const { width, height } = shape;
  const bytes = shape.colorType === 6 ? 4 : 3;
  const plain = plainLines(shape.chunks, width, height, bytes);
  if (!plain) return undefined;
  const pixels = new Uint8Array(width * height * 4);
  for (let pixel = 0; pixel < width * height; pixel++) {
    pixels[pixel * 4] = plain[pixel * bytes] as number;
    pixels[pixel * 4 + 1] = plain[pixel * bytes + 1] as number;
    pixels[pixel * 4 + 2] = plain[pixel * bytes + 2] as number;
    pixels[pixel * 4 + 3] = bytes === 4 ? (plain[pixel * bytes + 3] as number) : 255;
  }
  return { width, height, pixels };
}
