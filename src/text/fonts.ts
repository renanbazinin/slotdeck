// What measuring text needs from a font library, and which typeface a character is set in.

/** Weight and slant of a face. */
export interface FontStyle {
  bold: boolean;
  italic: boolean;
}

/** A typeface in one weight and slant, as far as measuring needs it. */
export interface FontFace {
  /** Whether the face has a glyph for a character, given as a code point. */
  has(codePoint: number): boolean;
  /**
   * The advance of each glyph of a text set on one line, in ems, after the font's own shaping
   * (ligatures, Arabic joining) and with kerning when asked for. Their sum times the font size is
   * the width of the text. They are wanted one by one because PowerPoint rounds each of them.
   */
  advances(text: string, kerning: boolean): number[];
  /** Set when the face stands in for a weight or slant that was not supplied, so widths are a little off. */
  substitute?: boolean;
}

/**
 * The one thing slotdeck needs from a font library: the widths of glyphs in a typeface.
 * Implement this to plug in any library; {@link fontkitMetrics} does it for `fontkit`.
 */
export interface FontMetrics {
  /** The face for a typeface name as the template states it, or `undefined` when it was not supplied. */
  face(typeface: string, style: FontStyle): FontFace | undefined;
  /** A face for a character that the chosen face has no glyph for. */
  fallback?(codePoint: number, style: FontStyle): FontFace | undefined;
}

/**
 * The width of a piece of text at a font size, in points, as PowerPoint lays it out: on a grid of
 * an eighth of a point, to which the advance of every glyph is rounded. At small sizes that
 * rounding moves a line's width by a percent or two, enough to move a word to the next line.
 */
export function textWidth(face: FontFace, text: string, size: number, kerning: boolean): number {
  let units = 0;
  for (const advance of face.advances(text, kerning)) units += Math.round(advance * size * 8);
  return units / 8;
}

/**
 * The three typefaces a run of text carries. PowerPoint sets Hebrew, Arabic and the like in the
 * complex-script face, Chinese, Japanese and Korean in the East Asian one, and the rest in the Latin one.
 */
export type ScriptClass = 'latin' | 'cs' | 'ea';

const COMPLEX =
  /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Tibetan}\p{Script=Myanmar}]/u;
const EAST_ASIAN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}　-〿＀-￯]/u;
const STRONG = /[\p{L}\p{M}]/u;

const classes = new Map<string, ScriptClass | undefined>();

/** The typeface class of a character, or `undefined` for one that takes its neighbour's (a space, a digit, punctuation). */
export function scriptClass(char: string): ScriptClass | undefined {
  if (classes.has(char)) return classes.get(char);
  const found: ScriptClass | undefined = COMPLEX.test(char) ? 'cs' : EAST_ASIAN.test(char) ? 'ea' : STRONG.test(char) ? 'latin' : undefined;
  if (classes.size < 20000) classes.set(char, found);
  return found;
}

/** Whether a line may break before and after the character without a space, as between Chinese or Japanese characters. */
export function breaksAnywhere(char: string): boolean {
  return /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}]/u.test(char);
}

/** The theme's name for the script of a complex-script or East Asian character, as in `<a:font script="Hebr">`. */
export function themeScript(char: string): string | undefined {
  for (const [tag, test] of THEME_SCRIPTS) if (test.test(char)) return tag;
  return undefined;
}

const THEME_SCRIPTS: Array<[string, RegExp]> = [
  ['Hebr', /\p{Script=Hebrew}/u],
  ['Arab', /\p{Script=Arabic}/u],
  ['Syrc', /\p{Script=Syriac}/u],
  ['Thaa', /\p{Script=Thaana}/u],
  ['Thai', /\p{Script=Thai}/u],
  ['Laoo', /\p{Script=Lao}/u],
  ['Khmr', /\p{Script=Khmer}/u],
  ['Deva', /\p{Script=Devanagari}/u],
  ['Beng', /\p{Script=Bengali}/u],
  ['Guru', /\p{Script=Gurmukhi}/u],
  ['Gujr', /\p{Script=Gujarati}/u],
  ['Taml', /\p{Script=Tamil}/u],
  ['Telu', /\p{Script=Telugu}/u],
  ['Knda', /\p{Script=Kannada}/u],
  ['Mlym', /\p{Script=Malayalam}/u],
  ['Sinh', /\p{Script=Sinhala}/u],
  ['Tibt', /\p{Script=Tibetan}/u],
  ['Hang', /\p{Script=Hangul}/u],
  ['Jpan', /[\p{Script=Hiragana}\p{Script=Katakana}]/u],
  ['Hans', /\p{Script=Han}/u],
];
