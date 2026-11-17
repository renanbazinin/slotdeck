// Colours as people see them. A template's colours are changed in OKLCH (lightness, chroma, hue),
// where a tint of a colour keeps its hue and equal steps of lightness look equal, so that a light
// teal becomes a light blue and not a grey one. https://bottosson.github.io/posts/oklab/

/** A colour in OKLCH: lightness 0..1, chroma from 0 (grey), hue in degrees 0..360. */
export interface Oklch {
  l: number;
  c: number;
  h: number;
}

const HEX = /^#?([0-9a-fA-F]{6})$/;

/**
 * A colour as slotdeck stores it: six hex digits in upper case, without `#`. `undefined` for
 * anything else, such as a colour name: a value that reaches a file unchecked would break it.
 */
export function parseHex(value: unknown): string | undefined {
  const match = typeof value === 'string' ? HEX.exec(value.trim()) : null;
  return match ? (match[1] as string).toUpperCase() : undefined;
}

const toLinear = (channel: number): number => (channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4);
const fromLinear = (channel: number): number => (channel <= 0.0031308 ? channel * 12.92 : 1.055 * channel ** (1 / 2.4) - 0.055);

/** Red, green and blue of a hex colour, each 0..255. */
export function hexToRgb(hex: string): [number, number, number] {
  const value = Number.parseInt(hex.replace(/^#/, ''), 16);
  return [(value >> 16) & 0xff, (value >> 8) & 0xff, value & 0xff];
}

export function rgbToHex(red: number, green: number, blue: number): string {
  const byte = (channel: number) => Math.min(255, Math.max(0, Math.round(channel)));
  return ((byte(red) << 16) | (byte(green) << 8) | byte(blue)).toString(16).padStart(6, '0').toUpperCase();
}

export function rgbToOklch(red: number, green: number, blue: number): Oklch {
  const r = toLinear(red / 255);
  const g = toLinear(green / 255);
  const b = toLinear(blue / 255);
  const l = Math.cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b);
  const m = Math.cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b);
  const s = Math.cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b);
  const lightness = 0.2104542553 * l + 0.793617785 * m - 0.0040720468 * s;
  const a = 1.9779984951 * l - 2.428592205 * m + 0.4505937099 * s;
  const bb = 0.0259040371 * l + 0.7827717662 * m - 0.808675766 * s;
  const chroma = Math.hypot(a, bb);
  // A grey has no hue; 0 keeps the result the same on every machine.
  const hue = chroma < 1e-7 ? 0 : (Math.atan2(bb, a) * 180) / Math.PI;
  return { l: lightness, c: chroma, h: hue < 0 ? hue + 360 : hue };
}

export function hexToOklch(hex: string): Oklch {
  return rgbToOklch(...hexToRgb(hex));
}

/** Linear red, green and blue of an OKLCH colour; outside 0..1 when a screen cannot show it. */
function toLinearRgb({ l, c, h }: Oklch): [number, number, number] {
  const angle = (h * Math.PI) / 180;
  const a = c * Math.cos(angle);
  const b = c * Math.sin(angle);
  const l3 = (l + 0.3963377774 * a + 0.2158037573 * b) ** 3;
  const m3 = (l - 0.1055613458 * a - 0.0638541728 * b) ** 3;
  const s3 = (l - 0.0894841775 * a - 1.291485548 * b) ** 3;
  return [
    4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3,
    -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3,
    -0.0041960863 * l3 - 0.7034186147 * m3 + 1.707614701 * s3,
  ];
}

const shows = (rgb: [number, number, number]): boolean => rgb.every((channel) => channel >= -0.0005 && channel <= 1.0005);

/**
 * The red, green and blue (0..255) of an OKLCH colour. A colour too vivid for a screen keeps its
 * lightness and hue and gives up chroma until it can be shown.
 */
export function oklchToRgb(color: Oklch): [number, number, number] {
  const lightness = Math.min(1, Math.max(0, color.l));
  let wanted: Oklch = { l: lightness, c: Math.max(0, color.c), h: color.h };
  let rgb = toLinearRgb(wanted);
  if (!shows(rgb)) {
    let low = 0;
    let high = wanted.c;
    for (let step = 0; step < 24; step++) {
      const middle = (low + high) / 2;
      if (shows(toLinearRgb({ ...wanted, c: middle }))) low = middle;
      else high = middle;
    }
    wanted = { ...wanted, c: low };
    rgb = toLinearRgb(wanted);
  }
  return rgb.map((channel) => Math.round(Math.min(1, Math.max(0, fromLinear(Math.min(1, Math.max(0, channel))))) * 255)) as [number, number, number];
}

export function oklchToHex(color: Oklch): string {
  return rgbToHex(...oklchToRgb(color));
}

/** How far apart two hues are, in degrees: 0 to 180. */
export function hueDistance(a: number, b: number): number {
  const distance = Math.abs(a - b) % 360;
  return distance > 180 ? 360 - distance : distance;
}

/** The signed way from hue `from` to hue `to`, in degrees: -180 to 180. */
export function hueOffset(from: number, to: number): number {
  const offset = (((to - from) % 360) + 540) % 360;
  return offset - 180;
}

/** How bright a colour is to the eye, 0 (black) to 1 (white), as the contrast rule of WCAG counts it. */
export function luminance(hex: string): number {
  const [red, green, blue] = hexToRgb(hex).map((channel) => toLinear(channel / 255)) as [number, number, number];
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
}

/** The contrast of two colours, from 1 (the same) to 21 (black on white). Text reads well from about 4.5, large text from 3. */
export function contrast(a: string, b: string): number {
  const first = luminance(a);
  const second = luminance(b);
  return (Math.max(first, second) + 0.05) / (Math.min(first, second) + 0.05);
}

/** A colour between two others: `share` of the first (0 to 1) and the rest of the second. */
export function mix(first: string, second: string, share: number): string {
  const a = hexToRgb(first);
  const b = hexToRgb(second);
  const part = Math.min(1, Math.max(0, share));
  return rgbToHex(a[0] * part + b[0] * (1 - part), a[1] * part + b[1] * (1 - part), a[2] * part + b[2] * (1 - part));
}

/**
 * A colour of the same hue that reads on a background: `color` made darker or lighter, whichever
 * is the shorter way, until its contrast to `background` is at least `wanted`. Black or white
 * when no shade of the hue gets there.
 */
export function readable(color: string, background: string, wanted = 4.5): string {
  if (contrast(color, background) >= wanted) return color;
  const given = hexToOklch(color);
  let best: { hex: string; away: number } | undefined;
  for (let step = 1; step <= 50; step++) {
    for (const lightness of [given.l - step * 0.02, given.l + step * 0.02]) {
      if (lightness < 0 || lightness > 1) continue;
      const hex = oklchToHex({ ...given, l: lightness });
      if (contrast(hex, background) >= wanted && (!best || step < best.away)) best = { hex, away: step };
    }
    if (best) return best.hex;
  }
  return contrast('000000', background) >= contrast('FFFFFF', background) ? '000000' : 'FFFFFF';
}
