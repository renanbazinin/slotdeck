// Font metrics from the `fontkit` library. slotdeck does not depend on it: the developer installs
// it and passes the module in, so the core keeps to its two dependencies.
import type { FontFace, FontMetrics, FontStyle } from './fonts.ts';

/** The part of the `fontkit` module this adapter uses: `import * as fontkit from 'fontkit'`. */
export interface FontkitModule {
  create(buffer: never, postscriptName?: string): unknown;
}

/** Font files by the typeface name the template uses. Several files give a family its weights and slants. */
export type FontFiles = Record<string, Uint8Array | Uint8Array[]>;

interface KitFont {
  familyName?: string | null;
  unitsPerEm: number;
  italicAngle?: number;
  head?: { macStyle?: { bold?: boolean; italic?: boolean } };
  'OS/2'?: { usWeightClass?: number; fsSelection?: { bold?: boolean; italic?: boolean } };
  variationAxes?: Record<string, { min: number; default: number; max: number }>;
  getVariation?(settings: Record<string, number>): KitFont;
  hasGlyphForCodePoint(codePoint: number): boolean;
  layout(text: string, features?: Record<string, boolean>): { positions: Array<{ xAdvance: number }> };
  /** Present on a collection (.ttc): the fonts it holds. */
  fonts?: KitFont[];
}

interface Loaded {
  font: KitFont;
  bold: boolean;
  italic: boolean;
}

/** A synthesized bold is drawn a little wider than the regular face it is made from. */
const FAKE_BOLD = 0.02;

/**
 * Builds {@link FontMetrics} from font files, read with `fontkit` (TrueType, OpenType, WOFF,
 * WOFF2 and .ttc collections). Each key is a typeface name as the template states it; its
 * value is one file or several (regular, bold, italic), told apart by what each file says of
 * itself. A file may stand in for a typeface of another name, such as an open metric-compatible
 * font for a licensed one.
 *
 * A weight or slant that was not supplied is measured with the nearest face that was. A
 * character that a typeface lacks is measured in the first supplied typeface that has it.
 */
export function fontkitMetrics(fontkit: FontkitModule, files: FontFiles): FontMetrics {
  const families = new Map<string, Loaded[]>();
  for (const [name, value] of Object.entries(files)) {
    const loaded: Loaded[] = [];
    for (const bytes of Array.isArray(value) ? value : [value]) {
      const opened = (fontkit.create as (buffer: Uint8Array) => unknown)(bytes) as KitFont;
      // A collection holds several families; the one asked for is taken, or the first.
      const members = opened.fonts ?? [opened];
      const named = members.filter((font) => (font.familyName ?? '').toLowerCase() === name.toLowerCase());
      for (const font of named.length > 0 ? named : members.slice(0, 1)) {
        loaded.push({
          font,
          bold: !!(font['OS/2']?.fsSelection?.bold ?? font.head?.macStyle?.bold) || (font['OS/2']?.usWeightClass ?? 400) >= 600,
          italic: !!(font['OS/2']?.fsSelection?.italic ?? font.head?.macStyle?.italic) || (font.italicAngle ?? 0) !== 0,
        });
      }
    }
    if (loaded.length > 0) families.set(name.toLowerCase(), loaded);
  }

  const faces = new Map<string, FontFace | undefined>();
  const build = (family: Loaded[], style: FontStyle): FontFace | undefined => {
    const exact = family.find((entry) => entry.bold === style.bold && entry.italic === style.italic);
    let chosen = exact ?? family.find((entry) => entry.bold === style.bold) ?? family.find((entry) => !entry.bold && !entry.italic) ?? family[0];
    if (!chosen) return undefined;
    let font = chosen.font;
    let substitute = !exact;
    // A variable font carries its bold in the same file.
    if (!exact && style.bold && !chosen.bold && font.variationAxes?.wght && font.getVariation) {
      try {
        font = font.getVariation({ wght: Math.min(700, font.variationAxes.wght.max) });
        chosen = { font, bold: true, italic: chosen.italic };
        substitute = chosen.italic !== style.italic;
      } catch {
        // Kept as the regular face.
      }
    }
    const widen = style.bold && !chosen.bold ? FAKE_BOLD : 0;
    const known = new Map<string, number[]>();
    const face: FontFace = {
      has: (codePoint) => font.hasGlyphForCodePoint(codePoint),
      advances(text, kerning) {
        const key = kerning ? `k${text}` : `n${text}`;
        let advances = known.get(key);
        if (!advances) {
          advances = font.layout(text, kerning ? undefined : { kern: false }).positions.map((position) => position.xAdvance / font.unitsPerEm + widen);
          if (known.size < 50000) known.set(key, advances);
        }
        return advances;
      },
    };
    if (substitute) face.substitute = true;
    return face;
  };
  const lookup = (name: string, style: FontStyle): FontFace | undefined => {
    const key = `${name.toLowerCase()}|${style.bold ? 'b' : ''}${style.italic ? 'i' : ''}`;
    if (!faces.has(key)) {
      const family = families.get(name.toLowerCase());
      faces.set(key, family ? build(family, style) : undefined);
    }
    return faces.get(key);
  };

  return {
    face: lookup,
    fallback(codePoint, style) {
      for (const name of families.keys()) {
        const face = lookup(name, style);
        if (face?.has(codePoint)) return face;
      }
      return undefined;
    },
  };
}
