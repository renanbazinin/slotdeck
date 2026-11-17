import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { deflateSync, zlibSync } from 'fflate';
import { readImageInfo } from '../src/util/image.ts';
import { mapPngColors, readPngPixels, solidPng } from '../src/util/png.ts';

const swap = (red: number, green: number, blue: number) => [blue, green, red] as const;

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** A PNG put together by hand, so that the tests do not lean on the code they test. */
function png(header: number[], parts: Array<[string, Uint8Array]>): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  };
  const all = [chunk('IHDR', new Uint8Array(header)), ...parts.map(([type, data]) => chunk(type, data)), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(8 + all.reduce((sum, part) => sum + part.length, 0));
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  let offset = 8;
  for (const part of all) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

test('a flat picture changes colour and stays a picture of the same size', () => {
  const teal = solidPng(5, 3, [14, 124, 123]);
  const next = mapPngColors(teal, swap) as Uint8Array;
  assert.deepEqual(readImageInfo(next), readImageInfo(teal));
  const { pixels, width, height } = readPngPixels(next) as { pixels: Uint8Array; width: number; height: number };
  assert.deepEqual([width, height], [5, 3]);
  for (let i = 0; i < pixels.length; i += 4) assert.deepEqual([...pixels.subarray(i, i + 4)], [123, 124, 14, 255]);
  // The same picture and map give the same bytes; a map that changes nothing gives nothing.
  assert.deepEqual(mapPngColors(teal, swap), next);
  assert.equal(mapPngColors(teal, (red, green, blue) => [red, green, blue]), undefined);
});

test('transparency is kept, and every way a line can be stored is read', () => {
  // Four pixels wide, five lines, one for each way of storing a line; the pixels are a ramp.
  const width = 4;
  const rows: number[][] = [];
  for (let y = 0; y < 5; y++) rows.push(Array.from({ length: width * 4 }, (_, x) => (x % 4 === 3 ? 40 * (y + 1) : (y * 53 + x * 29) & 0xff)));
  const raw: number[] = [];
  rows.forEach((row, y) => {
    const above = y > 0 ? (rows[y - 1] as number[]) : row.map(() => 0);
    raw.push(y);
    row.forEach((value, x) => {
      const left = x >= 4 ? (row[x - 4] as number) : 0;
      const up = above[x] as number;
      const upLeft = x >= 4 ? (above[x - 4] as number) : 0;
      let predicted = 0;
      if (y === 1) predicted = left;
      else if (y === 2) predicted = up;
      else if (y === 3) predicted = (left + up) >> 1;
      else if (y === 4) {
        const estimate = left + up - upLeft;
        const a = Math.abs(estimate - left);
        const b = Math.abs(estimate - up);
        const c = Math.abs(estimate - upLeft);
        predicted = a <= b && a <= c ? left : b <= c ? up : upLeft;
      }
      raw.push((value - predicted) & 0xff);
    });
  });
  const packed = zlibSync(new Uint8Array(raw));
  // The data may come in several pieces.
  const picture = png([0, 0, 0, width, 0, 0, 0, 5, 8, 6, 0, 0, 0], [['IDAT', packed.subarray(0, 7)], ['IDAT', packed.subarray(7)]]);
  const before = readPngPixels(picture) as { pixels: Uint8Array };
  const { pixels } = readPngPixels(mapPngColors(picture, swap) as Uint8Array) as { pixels: Uint8Array };
  rows.forEach((row, y) => {
    for (let x = 0; x < width; x++) {
      const [red, green, blue, alpha] = row.slice(x * 4, x * 4 + 4) as [number, number, number, number];
      const at = (y * width + x) * 4;
      assert.deepEqual([...before.pixels.subarray(at, at + 4)], [red, green, blue, alpha], `line ${y}, pixel ${x} as stored`);
      assert.deepEqual([...pixels.subarray(at, at + 4)], [blue, green, red, alpha], `line ${y}, pixel ${x} mapped`);
    }
  });
});

test('a picture with a colour table has only its table rewritten', () => {
  const table = new Uint8Array([14, 124, 123, 255, 255, 255]);
  const data = zlibSync(new Uint8Array([0, 0b01000000, 0, 0b10000000]));
  const picture = png([0, 0, 0, 2, 0, 0, 0, 2, 1, 3, 0, 0, 0], [['PLTE', table], ['tRNS', new Uint8Array([255, 0])], ['IDAT', data]]);
  const next = mapPngColors(picture, (red, green, blue) => (red === 14 ? [200, 16, 46] : [red, green, blue])) as Uint8Array;
  assert.equal(next.length, picture.length);
  const at = 8 + 25 + 8; // after the signature and the header chunk: the table's bytes
  assert.deepEqual([...next.subarray(at, at + 6)], [200, 16, 46, 255, 255, 255]);
  // Everything after the table and its checksum is as it was: transparency and pixels.
  assert.deepEqual(next.subarray(at + 10), picture.subarray(at + 10));
  assert.deepEqual(next.subarray(0, at), picture.subarray(0, at));
});

test('what cannot be read is left alone', () => {
  assert.equal(mapPngColors(new Uint8Array([1, 2, 3]), swap), undefined);
  assert.equal(mapPngColors(new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 16, 74, 70, 73, 70]), swap), undefined);
  const grey = png([0, 0, 0, 1, 0, 0, 0, 1, 8, 0, 0, 0, 0], [['IDAT', zlibSync(new Uint8Array([0, 90]))]]);
  assert.equal(mapPngColors(grey, swap), undefined);
  const deep = png([0, 0, 0, 1, 0, 0, 0, 1, 16, 2, 0, 0, 0], [['IDAT', zlibSync(new Uint8Array([0, 1, 2, 3, 4, 5, 6]))]]);
  assert.equal(mapPngColors(deep, swap), undefined);
  const interlaced = png([0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 1], [['IDAT', zlibSync(new Uint8Array([0, 1, 2, 3]))]]);
  assert.equal(mapPngColors(interlaced, swap), undefined);
  const broken = png([0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0], [['IDAT', deflateSync(new Uint8Array([0, 1, 2, 3])).subarray(0, 2)]]);
  assert.equal(mapPngColors(broken, swap), undefined);
  const short = png([0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0, 0, 0], [['IDAT', zlibSync(new Uint8Array([0, 1, 2, 3]))]]);
  assert.equal(mapPngColors(short, swap), undefined);
});

test('the icons of the examples are pictures this reads', () => {
  const icon = new Uint8Array(readFileSync(new URL('../examples/icons/clock.png', import.meta.url)));
  let seen = 0;
  const next = mapPngColors(icon, (red, green, blue) => {
    seen++;
    return red === green && green === blue ? [red, green, blue] : [200, 16, 46];
  });
  assert.ok(seen > 0, 'the colours of the icon were read');
  assert.ok(next, 'the icon has a colour that is not a grey, and it was changed');
  assert.deepEqual(readImageInfo(next as Uint8Array), readImageInfo(icon));
});
