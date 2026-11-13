export type ImageFormat = 'png' | 'jpeg' | 'gif';

export interface ImageInfo {
  format: ImageFormat;
  /** File extension to store the image under. */
  extension: string;
  contentType: string;
  width: number;
  height: number;
}

/** Reads format and pixel size from the image header. Throws for formats PowerPoint cannot embed as a picture. */
export function readImageInfo(data: Uint8Array): ImageInfo {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);

  if (data.length >= 24 && view.getUint32(0) === 0x89504e47 && view.getUint32(4) === 0x0d0a1a0a) {
    return {
      format: 'png',
      extension: 'png',
      contentType: 'image/png',
      width: view.getUint32(16),
      height: view.getUint32(20),
    };
  }

  if (data.length >= 10 && data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) {
    return {
      format: 'gif',
      extension: 'gif',
      contentType: 'image/gif',
      width: view.getUint16(6, true),
      height: view.getUint16(8, true),
    };
  }

  if (data.length >= 4 && data[0] === 0xff && data[1] === 0xd8) {
    const size = readJpegSize(view);
    if (size) return { format: 'jpeg', extension: 'jpg', contentType: 'image/jpeg', ...size };
    throw new Error('JPEG image has no frame header');
  }

  throw new Error('Unsupported image format: expected PNG, JPEG or GIF');
}

function readJpegSize(view: DataView): { width: number; height: number } | undefined {
  let offset = 2;
  while (offset + 4 <= view.byteLength) {
    if (view.getUint8(offset) !== 0xff) {
      offset++;
      continue;
    }
    const marker = view.getUint8(offset + 1);
    if (marker === 0xff) {
      offset++; // fill byte
      continue;
    }
    // Markers without a length field.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      offset += 2;
      continue;
    }
    const length = view.getUint16(offset + 2);
    const isFrameHeader = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isFrameHeader) {
      if (offset + 9 > view.byteLength) return undefined;
      return { height: view.getUint16(offset + 5), width: view.getUint16(offset + 7) };
    }
    offset += 2 + length;
  }
  return undefined;
}
