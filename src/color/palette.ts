// A template's colours as a designer would name them: the main colour with its tints and shades,
// a second colour, the dark colour of the text, and the greys. Read from how often each colour is
// used; no model. A deck that asks for another colour has the whole family follow.
import type { DeckColors } from '../deck/spec.ts';
import { hexToOklch, hueDistance, hueOffset, oklchToHex, parseHex } from './space.ts';
import type { Oklch } from './space.ts';

/** A colour and how much the template uses it: fills count more than text, text more than lines. */
export interface ColorUse {
  hex: string;
  weight: number;
}

/** A colour of the template with every tint and shade of it that the template uses. */
export interface ColorFamily {
  /** The colour the family is named by: its most used strong member. */
  base: string;
  /** Every member, the base included, most used first. */
  members: string[];
  /** The family's share of all the colour the template uses, 0 to 1. */
  share: number;
}

export interface TemplatePalette {
  /** The template's main colour: what `accent` in a deck's colours takes the place of. */
  accent?: string;
  /** Its second colour. */
  accent2?: string;
  /** The dark colour its text is set in, when that is not black or a grey. */
  ink?: string;
  /** Families of colours, the most used first. Greys, black and white belong to none. */
  families: ColorFamily[];
  /** Every colour found, the most used first, with its share of all use. */
  colors: Array<{ hex: string; share: number }>;
}

/** Chroma under which a colour is a grey. Very light and very dark colours have little chroma even when they are tints or shades. */
const greyUnder = (lightness: number): number => (lightness > 0.85 ? 0.006 : lightness < 0.25 ? 0.02 : 0.03);
/** A colour strong enough to name a family: neither a tint, nor a shade, nor dull. */
const isVivid = (color: Oklch): boolean => color.c >= 0.07 && color.l >= 0.3 && color.l <= 0.9;
/** How far in hue a colour may be from a family's base and still belong to it. A pale colour's hue is less sure. */
const hueReach = (color: Oklch): number => (color.c < 0.04 ? 45 : 30);

export function isGrey(hex: string): boolean {
  const color = hexToOklch(hex);
  return color.c < greyUnder(color.l);
}

/**
 * Groups the colours a template uses into families and names the main ones. `pinned` states the
 * main and second colour outright, for a template where the count picks the wrong one.
 */
export function buildPalette(uses: ColorUse[], pinned: { accent?: string; accent2?: string } = {}): TemplatePalette {
  const weights = new Map<string, number>();
  for (const use of uses) {
    const hex = parseHex(use.hex);
    if (hex && use.weight > 0) weights.set(hex, (weights.get(hex) ?? 0) + use.weight);
  }
  const total = [...weights.values()].reduce((sum, weight) => sum + weight, 0) || 1;
  // The same colours always come in the same order, whatever order they were found in.
  const all = [...weights].map(([hex, weight]) => ({ hex, weight, color: hexToOklch(hex) })).sort((a, b) => b.weight - a.weight || a.hex.localeCompare(b.hex));
  const coloured = all.filter((entry) => entry.color.c >= greyUnder(entry.color.l));

  interface Group {
    base: (typeof all)[number];
    members: Array<(typeof all)[number]>;
    vivid: boolean;
  }
  const groups: Group[] = [];
  const join = (entry: (typeof all)[number]): boolean => {
    let best: Group | undefined;
    let nearest = Infinity;
    for (const group of groups) {
      const distance = hueDistance(entry.color.h, group.base.color.h);
      if (distance <= hueReach(entry.color) && distance < nearest) {
        best = group;
        nearest = distance;
      }
    }
    if (!best) return false;
    best.members.push(entry);
    return true;
  };
  // Pinned colours name their families first; then the strong colours, the most used and strongest first.
  for (const hex of [parseHex(pinned.accent), parseHex(pinned.accent2)]) {
    if (!hex || groups.some((group) => group.base.hex === hex)) continue;
    const entry = all.find((candidate) => candidate.hex === hex) ?? { hex, weight: 0, color: hexToOklch(hex) };
    groups.push({ base: entry, members: [entry], vivid: true });
  }
  const strength = (entry: (typeof all)[number]) => entry.weight * entry.color.c;
  const strong = coloured.filter((entry) => isVivid(entry.color)).sort((a, b) => strength(b) - strength(a) || a.hex.localeCompare(b.hex));
  for (const entry of strong) {
    if (groups.some((group) => group.members.includes(entry))) continue;
    if (!join(entry)) groups.push({ base: entry, members: [entry], vivid: true });
  }
  // Tints, shades and dull colours go with the family they lean to, or make one of their own.
  for (const entry of coloured) {
    if (groups.some((group) => group.members.includes(entry))) continue;
    if (!join(entry)) groups.push({ base: entry, members: [entry], vivid: false });
  }

  const weightOf = (group: Group) => group.members.reduce((sum, member) => sum + member.weight, 0);
  const pinnedBases = [parseHex(pinned.accent), parseHex(pinned.accent2)];
  // A family is named by the strong member in the middle of it: half of the family's strong colour
  // is as light or lighter, half as dark or darker. The pale end of a gradient or the darkest
  // shade would be a poor name: a deck's colour takes the base's place, and the rest follow it.
  for (const group of groups) {
    if (!group.vivid || pinnedBases.includes(group.base.hex)) continue;
    const strongMembers = group.members.filter((member) => isVivid(member.color)).sort((a, b) => a.color.l - b.color.l || a.hex.localeCompare(b.hex));
    const half = strongMembers.reduce((sum, member) => sum + member.weight, 0) / 2;
    let passed = 0;
    for (const member of strongMembers) {
      passed += member.weight;
      if (passed >= half) {
        group.base = member;
        break;
      }
    }
  }
  const ordered = groups
    .map((group) => ({ group, weight: weightOf(group) }))
    .sort((a, b) => b.weight - a.weight || a.group.base.hex.localeCompare(b.group.base.hex));
  const families: ColorFamily[] = ordered.map(({ group, weight }) => ({
    base: group.base.hex,
    members: group.members.slice().sort((a, b) => b.weight - a.weight || a.hex.localeCompare(b.hex)).map((member) => member.hex),
    share: Math.round((weight / total) * 10000) / 10000,
  }));

  const palette: TemplatePalette = { families, colors: all.map((entry) => ({ hex: entry.hex, share: Math.round((entry.weight / total) * 10000) / 10000 })) };
  const vivid = ordered.filter(({ group }) => group.vivid && !pinnedBases.includes(group.base.hex)).map(({ group }) => group.base.hex);
  const accent = pinnedBases[0] ?? vivid.shift();
  const accent2 = pinnedBases[1] ?? vivid.shift();
  if (accent) palette.accent = accent;
  if (accent2 && accent2 !== accent) palette.accent2 = accent2;
  const ink = ordered.find(({ group }) => !group.vivid && group.base.color.l < 0.4);
  if (ink) palette.ink = ink.group.base.hex;
  return palette;
}

/**
 * What is wrong with a deck's colours, in words, and the colours with everything wrong left out.
 * `undefined` colours when nothing usable is left.
 */
export function readColors(value: unknown): { colors?: DeckColors; problems: string[] } {
  const problems: string[] = [];
  if (value === undefined || value === null) return { problems };
  if (typeof value !== 'object' || Array.isArray(value)) return { problems: ['"colors" must be an object such as { "accent": "1565C0" }; it is ignored'] };
  const given = value as Record<string, unknown>;
  const colors: DeckColors = {};
  for (const name of ['accent', 'accent2'] as const) {
    if (given[name] === undefined || given[name] === null || given[name] === '') continue;
    const hex = parseHex(given[name]);
    if (hex) colors[name] = hex;
    else problems.push(`the colour "${name}" is ${JSON.stringify(given[name])}, which is not six hex digits such as "1565C0"; it is ignored`);
  }
  if (given.replace !== undefined && given.replace !== null) {
    if (typeof given.replace !== 'object' || Array.isArray(given.replace)) {
      problems.push('"replace" must map colours to colours, such as { "0E7C7B": "1565C0" }; it is ignored');
    } else {
      const replace: Record<string, string> = {};
      for (const [from, to] of Object.entries(given.replace as Record<string, unknown>)) {
        const source = parseHex(from);
        const target = parseHex(to);
        if (source && target) replace[source] = target;
        else problems.push(`the replacement ${JSON.stringify(from)}: ${JSON.stringify(to)} is not a pair of colours of six hex digits; it is ignored`);
      }
      if (Object.keys(replace).length > 0) colors.replace = replace;
    }
  }
  for (const name of Object.keys(given)) {
    if (name !== 'accent' && name !== 'accent2' && name !== 'replace') problems.push(`"colors" has no setting called "${name}"; it is ignored`);
  }
  return Object.keys(colors).length > 0 ? { colors, problems } : { problems };
}

/** A function from a template colour to the colour a deck shows in its place. */
export type ColorMap = (hex: string) => string;

/**
 * The colour a member of a family becomes when the family's base becomes `target`: it keeps its
 * place in the family. Its hue stays as far from the base's as it was, its chroma is scaled as
 * the base's is, and its lightness moves along the line that keeps black black and white white
 * and takes the base's lightness to the target's. So the base becomes the target exactly, a
 * tint becomes a tint of the target, and a shade a shade.
 */
function follow(color: Oklch, base: Oklch, target: Oklch): Oklch {
  const scale = Math.min(4, Math.max(0, target.c / Math.max(base.c, 0.02)));
  const lightness =
    color.l <= base.l
      ? base.l > 0 ? (color.l * target.l) / base.l : target.l
      : base.l < 1 ? 1 - ((1 - color.l) * (1 - target.l)) / (1 - base.l) : target.l;
  return { l: lightness, c: color.c * scale, h: target.h + hueOffset(base.h, color.h) };
}

/**
 * The colours of a deck as a function over the template's colours, or `undefined` when the deck
 * asks for nothing that changes them. It answers for any colour, also one the palette never saw,
 * such as the soft edge of a pictogram: a colour that leans to the main colour follows it.
 */
export function colorMap(palette: TemplatePalette, colors: DeckColors | undefined): ColorMap | undefined {
  const wanted = readColors(colors).colors;
  if (!wanted) return undefined;
  const moves: Array<{ base: Oklch; target: Oklch; from: string; to: string }> = [];
  for (const [from, to] of [[palette.accent, wanted.accent], [palette.accent2, wanted.accent2]] as const) {
    if (from && to && from !== to) moves.push({ base: hexToOklch(from), target: hexToOklch(to), from, to });
  }
  const replace = new Map(Object.entries(wanted.replace ?? {}).filter(([from, to]) => from !== to));
  if (moves.length === 0 && replace.size === 0) return undefined;
  // Bases of the families that stay as they are: a colour nearer to one of them is not taken along.
  const staying = palette.families.map((family) => family.base).filter((base) => !moves.some((move) => move.from === base)).map(hexToOklch);

  const known = new Map<string, string>();
  return (given) => {
    const hex = parseHex(given);
    if (!hex) return given;
    const cached = known.get(hex);
    if (cached !== undefined) return cached;
    let result = replace.get(hex);
    if (result === undefined) {
      result = hex;
      const exact = moves.find((move) => move.from === hex);
      if (exact) result = exact.to;
      else {
        const color = hexToOklch(hex);
        if (color.c >= greyUnder(color.l)) {
          const reach = hueReach(color);
          let nearest: (typeof moves)[number] | undefined;
          let distance = Infinity;
          for (const move of moves) {
            const away = hueDistance(color.h, move.base.h);
            if (away <= reach && away < distance) {
              nearest = move;
              distance = away;
            }
          }
          const kept = Math.min(Infinity, ...staying.map((base) => hueDistance(color.h, base.h)));
          if (nearest && distance <= kept) result = oklchToHex(follow(color, nearest.base, nearest.target));
        }
      }
    }
    known.set(hex, result);
    return result;
  };
}
