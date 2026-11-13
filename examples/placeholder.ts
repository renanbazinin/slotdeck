// Stand-in pictures drawn in code, so demos and tests need no image files and no image model.
import { crc32, deflateSync } from 'node:zlib';

type Rgb = [number, number, number];

/**
 * A stand-in photo: a diagonal gradient with a centred ring.
 * The ring makes crop mistakes obvious: it must stay round and centred in any frame.
 */
export function placeholderPhoto(width: number, height: number, from: Rgb, to: Rgb): Uint8Array {
  const cx = width / 2;
  const cy = height / 2;
  const radius = Math.min(width, height) * 0.3;
  const thickness = Math.min(width, height) * 0.035;
  return encodePng(width, height, (x, y) => {
    if (Math.abs(Math.hypot(x - cx, y - cy) - radius) < thickness) return [255, 255, 255];
    const t = (x / width + y / height) / 2;
    return [from[0] + (to[0] - from[0]) * t, from[1] + (to[1] - from[1]) * t, from[2] + (to[2] - from[2]) * t];
  });
}

/** A stand-in for a chart the app supplies: bars on a white ground, with a baseline. */
export function placeholderChart(width: number, height: number, values: number[] = [0.35, 0.5, 0.42, 0.7, 0.62, 0.9]): Uint8Array {
  const margin = Math.round(Math.min(width, height) * 0.08);
  const baseline = height - margin;
  const slot = (width - margin * 2) / values.length;
  const colours: Rgb[] = [
    [14, 124, 123],
    [244, 162, 89],
  ];
  return encodePng(width, height, (x, y) => {
    if (y >= baseline && y < baseline + 3 && x >= margin && x < width - margin) return [136, 136, 136];
    const index = Math.floor((x - margin) / slot);
    const value = values[index];
    if (x < margin || value === undefined || y >= baseline) return [255, 255, 255];
    const inBar = (x - margin) % slot > slot * 0.2 && (x - margin) % slot < slot * 0.8;
    const top = baseline - value * (baseline - margin);
    return inBar && y >= top ? (colours[index % colours.length] as Rgb) : [255, 255, 255];
  });
}

function encodePng(width: number, height: number, pixel: (x: number, y: number) => number[]): Uint8Array {
  const stride = width * 3 + 1;
  const raw = Buffer.alloc(stride * height);
  for (let y = 0; y < height; y++) {
    raw[y * stride] = 0; // filter: none
    for (let x = 0; x < width; x++) {
      const colour = pixel(x, y);
      const offset = y * stride + 1 + x * 3;
      for (let channel = 0; channel < 3; channel++) raw[offset + channel] = Math.round(colour[channel] as number);
    }
  }

  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header.writeUInt8(8, 8); // bit depth
  header.writeUInt8(2, 9); // colour type: RGB
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', header),
      chunk('IDAT', deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ]),
  );
}

function chunk(type: string, data: Buffer): Buffer {
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([length, body, checksum]);
}
