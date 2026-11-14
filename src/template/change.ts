// Changes to a manifest, one small operation at a time. A studio's own controls send them
// directly; `editManifest` has a model turn an instruction into the same operations. Either way
// each change is checked against the pptx before it is kept, so a manifest cannot be broken here.
import type { Template } from '../render/render.ts';
import type { PptxPackage } from '../pptx/package.ts';
import { checkManifest } from './check.ts';
import type { ManifestIssue } from './check.ts';
import { VARIANTS, deriveVariant } from './derive.ts';
import type { PatternVariant } from './derive.ts';
import { detectRepeat, rowFit, rowOnSlide } from './rows.ts';
import { inspectTemplate } from './inspect.ts';
import type { ShapeInfo, SlideInfo, TemplateInfo } from './inspect.ts';
import type { AlignPolicy, Direction, EmptyPolicy, ImageKind, PatternDef, SlideBox, SlotDef, TableTextStyle, TemplateManifest } from './manifest.ts';
import { ROLES, boxCapacity, rolesFor, rowRoom, shapeText, slotFromShape, slotName, tableLimits } from './measure.ts';
import type { Role } from './measure.ts';

/** New values for a slot. `null` clears an optional setting. Each applies to slots of the kind noted. */
export interface SlotPatch {
  /** A new name for the slot. */
  name?: string;
  description?: string | null;
  /** `error` required, `remove` optional, `keep` leave the template's content, `drop` always remove the shape. */
  empty?: EmptyPolicy;
  /** Text: the character budget. */
  maxChars?: number;
  /** Text: the most paragraphs. */
  maxParagraphs?: number | null;
  /** Text: take the value from this deck-level field (a header that is the same on every slide). */
  deckField?: string | null;
  /** Text: whether alignment flips for a deck that reads the other way. */
  alignOnDirectionChange?: AlignPolicy | null;
  /** Text: which run of a mixed-style paragraph supplies the formatting. */
  styleRun?: 'first' | 'last' | null;
  /** Text: a font size in points other than the template's. The budget follows unless one is given. */
  fontSize?: number | null;
  /** Text: space after each paragraph, in points. */
  paragraphGap?: number | null;
  /** Text: whether a deck's text size applies. */
  scales?: boolean | null;
  /** Text, image, chart: another area for the shape, as fractions of the slide. A text budget follows unless one is given. */
  box?: SlideBox | null;
  /** Image: what kind of picture fills it. */
  imageKind?: ImageKind;
  /** Image: crop to the frame, or shrink the frame to the picture. */
  fit?: 'cover' | 'contain' | null;
  /** Table: the most columns. The cell budget follows unless one is given. */
  maxColumns?: number;
  /** Table: the most body rows. The cell budget follows unless one is given. */
  maxRows?: number;
  /** Table: the character budget of one cell. */
  maxCellChars?: number;
  /** Table: text formatting, merged into what is already set. A new size changes the cell budget unless one is given. */
  headerText?: TableTextStyle | null;
  bodyText?: TableTextStyle | null;
  /** Chart: limits and styling. */
  maxCategories?: number;
  maxSeries?: number;
  colors?: string[] | null;
  text?: TableTextStyle | null;
  lineColor?: string | null;
}

export type ManifestChange =
  | { op: 'setManifest'; name?: string; language?: string; direction?: Direction; alignOnDirectionChange?: AlignPolicy; textScale?: number | null }
  /** Makes a slide a pattern, with no content slots yet. Header texts known from other patterns are filled in. */
  | { op: 'addPattern'; slide: number; id: string; description?: string }
  /** A second pattern from the same slide, to be changed into a variant. */
  | { op: 'copyPattern'; pattern: string; id: string; description?: string }
  /**
   * A second pattern from the same slide, laid out by the software: the `mirror` image, the
   * `text-only` version, or the pattern `without` the optional text named in `slot`.
   */
  | { op: 'derivePattern'; pattern: string; variant: PatternVariant; slot?: string; id?: string; description?: string }
  | { op: 'removePattern'; pattern: string }
  | { op: 'setPattern'; pattern: string; id?: string; description?: string; figures?: boolean }
  /** Makes a shape a slot, with limits measured from the shape. Replaces the shape's slot when it has one. */
  | { op: 'addSlot'; pattern: string; shape: string; role: Role; name?: string; optional?: boolean }
  /** The shape goes back to being part of the design. */
  | { op: 'removeSlot'; pattern: string; slot: string }
  | ({ op: 'setSlot'; pattern: string; slot: string } & SlotPatch)
  /**
   * Sets how many items a row or grid takes. Finds the row from numbered slot names when the pattern has none.
   * `widen` lets items grow wider, up to that factor, when a line holds fewer than the template draws.
   */
  | { op: 'setRepeat'; pattern: string; min?: number; max?: number; item?: string; span?: { left: number; right: number } | null; widen?: number | null }
  | { op: 'removeRepeat'; pattern: string };

export interface ChangeResult {
  manifest: TemplateManifest;
  /** The changes that were made, in order, each with a sentence saying what is now true. */
  applied: Array<{ index: number; change: ManifestChange; message: string }>;
  /** The changes that were not made, and why. The manifest is as if they had not been asked for. */
  rejected: Array<{ index: number; change: ManifestChange; reason: string }>;
  /** What {@link checkManifest} finds in the result. */
  issues: ManifestIssue[];
}

class Rejected extends Error {}
const reject = (reason: string): never => {
  throw new Rejected(reason);
};

const TEXT_KEYS = ['maxChars', 'maxParagraphs', 'deckField', 'alignOnDirectionChange', 'styleRun', 'fontSize', 'paragraphGap', 'scales'] as const;
const IMAGE_KEYS = ['imageKind', 'fit'] as const;
const TABLE_KEYS = ['maxColumns', 'maxRows', 'maxCellChars', 'headerText', 'bodyText'] as const;
const CHART_KEYS = ['maxCategories', 'maxSeries', 'colors', 'text', 'lineColor'] as const;
const KEYS_OF = { text: TEXT_KEYS, image: IMAGE_KEYS, table: TABLE_KEYS, chart: CHART_KEYS } as const;
const STYLE_KEYS = new Set(['headerText', 'bodyText', 'text']);
const COUNT_KEYS = new Set(['maxChars', 'maxParagraphs', 'maxColumns', 'maxRows', 'maxCellChars', 'maxCategories', 'maxSeries']);

const kebab = (raw: string) => raw.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const policyWord = (policy: EmptyPolicy) => ({ error: 'required', remove: 'optional', keep: 'kept as designed', drop: 'always removed' })[policy];

/**
 * Applies changes to a template's manifest and returns the new manifest; the template itself is
 * not touched. Changes are applied in order. One that cannot be made, or that would leave the
 * manifest pointing at something the pptx does not have, is rejected with a reason and skipped.
 */
export function changeManifest(template: Template, changes: ManifestChange[]): ChangeResult {
  return changeWith(template, changes, inspectTemplate(template.pkg));
}

/** {@link changeManifest} for a caller that has already read the pptx's shapes. */
export function changeWith(template: Template, changes: ManifestChange[], info: TemplateInfo): ChangeResult {
  const errors = (manifest: TemplateManifest) =>
    checkManifest({ pkg: template.pkg, manifest }, info)
      .filter((issue) => issue.severity === 'error')
      .map((issue) => `${issue.slot ?? issue.pattern ?? 'manifest'} ${issue.message}`);

  let manifest = structuredClone(template.manifest);
  const applied: ChangeResult['applied'] = [];
  const rejected: ChangeResult['rejected'] = [];
  changes.forEach((change, index) => {
    const next = structuredClone(manifest);
    try {
      const message = apply(next, change, info, template.pkg);
      const before = new Set(errors(manifest));
      const fresh = errors(next).filter((error) => !before.has(error));
      if (fresh.length > 0) reject(fresh.join('; '));
      manifest = next;
      applied.push({ index, change, message });
    } catch (error) {
      if (!(error instanceof Rejected)) throw error;
      rejected.push({ index, change, reason: error.message });
    }
  });
  return { manifest, applied, rejected, issues: checkManifest({ pkg: template.pkg, manifest }, info) };
}

function apply(manifest: TemplateManifest, change: ManifestChange, info: TemplateInfo, pkg: PptxPackage): string {
  const patternOf = (id: string): PatternDef => manifest.patterns.find((pattern) => pattern.id === id) ?? reject(`there is no pattern "${id}"`);
  const slideOf = (pattern: PatternDef): SlideInfo => info.slides[pattern.sourceSlide - 1] ?? reject(`pattern "${pattern.id}" points at slide ${pattern.sourceSlide}, which the template does not have`);
  const freeId = (raw: string): string => {
    const id = kebab(raw) || reject('a pattern needs an id made of letters and digits');
    if (manifest.patterns.some((pattern) => pattern.id === id)) reject(`another pattern is already called "${id}"`);
    return id;
  };

  switch (change.op) {
    case 'setManifest': {
      const done: string[] = [];
      if (change.name !== undefined) {
        manifest.name = change.name.trim() || reject('the name is empty');
        done.push(`is named "${manifest.name}"`);
      }
      if (change.language !== undefined) {
        manifest.language = change.language.trim() || reject('the language is empty');
        done.push(`is written in "${manifest.language}"`);
      }
      if (change.direction !== undefined) {
        if (change.direction !== 'rtl' && change.direction !== 'ltr') reject('the direction is "rtl" or "ltr"');
        manifest.direction = change.direction;
        done.push(`reads ${change.direction === 'rtl' ? 'right to left' : 'left to right'}`);
      }
      if (change.alignOnDirectionChange !== undefined) {
        if (change.alignOnDirectionChange !== 'flip' && change.alignOnDirectionChange !== 'keep') reject('alignment on a direction change is "flip" or "keep"');
        manifest.alignOnDirectionChange = change.alignOnDirectionChange;
        done.push(`${change.alignOnDirectionChange === 'flip' ? 'mirrors' : 'keeps'} its alignment in decks that read the other way`);
      }
      if (change.textScale !== undefined) {
        if (change.textScale === null || change.textScale === 1) {
          delete manifest.textScale;
          done.push('shows body text at its designed size');
        } else {
          if (!(typeof change.textScale === 'number' && change.textScale >= 0.5 && change.textScale <= 3)) reject('a text scale is a factor between 0.5 and 3, such as 1.3');
          manifest.textScale = change.textScale;
          done.push(`shows body text at ${change.textScale} times its designed size in every deck, with budgets to match`);
        }
      }
      if (done.length === 0) reject('nothing to set');
      return `The template ${done.join(', ')}.`;
    }

    case 'addPattern': {
      const slide = info.slides[change.slide - 1] ?? reject(`the template has ${info.slides.length} slides; there is no slide ${change.slide}`);
      const id = freeId(change.id);
      // A header text that other patterns already treat as one value for the deck is the same value here.
      const fields = new Map<string, string>();
      for (const pattern of manifest.patterns) {
        const shapes = info.slides[pattern.sourceSlide - 1]?.shapes ?? [];
        for (const slot of pattern.slots) {
          const shape = slot.kind === 'text' && slot.deckField ? shapes.find((candidate) => candidate.id === slot.shapeId) : undefined;
          if (shape && slot.kind === 'text' && slot.deckField) fields.set(shapeText(shape), slot.deckField);
        }
      }
      const slots: SlotDef[] = [];
      for (const shape of slide.shapes) {
        const field = shape.kind === 'text' ? fields.get(shapeText(shape)) : undefined;
        if (!field || slots.some((slot) => slot.name === field)) continue;
        const slot = slotFromShape(shape, info, 'text', field);
        if (slot?.kind === 'text') slots.push({ ...slot, deckField: field });
      }
      manifest.patterns.push({ id, sourceSlide: change.slide, description: change.description?.trim() || `Layout of slide ${change.slide}.`, slots });
      return `Slide ${change.slide} is now the pattern "${id}"${slots.length > 0 ? `, with the deck's ${slots.map((slot) => slot.name).join(', ')}` : ''}. It has no content slots yet.`;
    }

    case 'copyPattern': {
      const source = patternOf(change.pattern);
      const copy = structuredClone(source);
      copy.id = freeId(change.id);
      if (change.description?.trim()) copy.description = change.description.trim();
      manifest.patterns.splice(manifest.patterns.indexOf(source) + 1, 0, copy);
      return `"${copy.id}" is a new pattern, a copy of "${source.id}".`;
    }

    case 'derivePattern': {
      const source = patternOf(change.pattern);
      if (!VARIANTS.includes(change.variant)) reject(`"${String(change.variant)}" is not a variant; the variants are ${VARIANTS.join(', ')}`);
      const request = change.slot ? { variant: change.variant, slot: change.slot } : { variant: change.variant };
      const made = deriveVariant(source, request, { pkg, info, manifest });
      if ('refused' in made) return reject(made.refused);
      const copy = made.pattern;
      copy.id = freeId(change.id?.trim() || `${source.id}-${made.suffix}`);
      copy.description = change.description?.trim() || made.description;
      manifest.patterns.splice(manifest.patterns.indexOf(source) + 1, 0, copy);
      return `"${copy.id}" is a new pattern from slide ${copy.sourceSlide}: ${made.summary}`;
    }

    case 'removePattern': {
      const pattern = patternOf(change.pattern);
      if (manifest.patterns.length === 1) reject('a template needs at least one pattern');
      manifest.patterns.splice(manifest.patterns.indexOf(pattern), 1);
      return `The pattern "${pattern.id}" is removed. Slide ${pattern.sourceSlide} of the template is no longer used${manifest.patterns.some((other) => other.sourceSlide === pattern.sourceSlide) ? ' by it' : ''}.`;
    }

    case 'setPattern': {
      const pattern = patternOf(change.pattern);
      const done: string[] = [];
      if (change.id !== undefined && kebab(change.id) !== pattern.id) {
        const was = pattern.id;
        pattern.id = freeId(change.id);
        done.push(`"${was}" is now called "${pattern.id}"`);
      }
      if (change.description !== undefined) {
        pattern.description = change.description.trim() || reject('the description is empty');
        done.push(`"${pattern.id}" has a new description`);
      }
      if (change.figures !== undefined) {
        if (change.figures) pattern.figures = true;
        else delete pattern.figures;
        done.push(`"${pattern.id}" is ${change.figures ? 'offered only for briefs with figures' : 'offered for any brief'}`);
      }
      if (done.length === 0) reject('nothing to set');
      return `${done.join('; ')}.`.replace(/^./, (first) => first.toUpperCase());
    }

    case 'addSlot': {
      const pattern = patternOf(change.pattern);
      const slide = slideOf(pattern);
      const shape = slide.shapes.find((candidate) => candidate.id === String(change.shape)) ?? reject(`slide ${pattern.sourceSlide} has no shape #${change.shape}`);
      if (!ROLES.includes(change.role)) reject(`"${change.role}" is not a role; the roles are ${ROLES.join(', ')}`);
      const existing = pattern.slots.find((slot) => slot.shapeId === shape.id);
      if (change.role === 'fixed') {
        if (!existing) reject(`shape #${shape.id} is not a slot, so it already stays as designed`);
        return apply(manifest, { op: 'removeSlot', pattern: pattern.id, slot: (existing as SlotDef).name }, info, pkg);
      }
      if (existing && inRow(pattern, existing.name)) reject(`"${existing.name}" is part of the row of ${pattern.repeat?.item ?? 'item'}s; remove the row first`);
      const others = new Set(pattern.slots.filter((slot) => slot !== existing).map((slot) => slot.name));
      let name = slotName(change.name ?? existing?.name ?? '', change.role);
      for (let n = 2; others.has(name); n++) name = `${slotName(change.name ?? existing?.name ?? '', change.role)}${n}`;
      const slot =
        slotFromShape(shape, info, change.role, name, change.optional === true, slide) ??
        reject(`shape #${shape.id} is ${describeKind(shape)}; it can be ${rolesFor(shape).join(', ') || 'nothing but part of the design'}${change.role === 'chart' && shape.groupId ? ' (a chart cannot replace a picture inside a group)' : ''}`);
      if (existing) pattern.slots.splice(pattern.slots.indexOf(existing), 1, slot);
      else pattern.slots.push(slot);
      return `"${pattern.id}": shape #${shape.id} is ${existing ? 'now' : 'a new slot,'} "${name}", ${describeSlot(slot)}.`;
    }

    case 'removeSlot': {
      const pattern = patternOf(change.pattern);
      const slot = pattern.slots.find((candidate) => candidate.name === change.slot) ?? reject(`"${pattern.id}" has no slot "${change.slot}"`);
      if (inRow(pattern, slot.name)) reject(`"${slot.name}" is part of the row of ${pattern.repeat?.item ?? 'item'}s; remove the row first`);
      pattern.slots.splice(pattern.slots.indexOf(slot), 1);
      return `"${pattern.id}": "${slot.name}" is no longer a slot. Shape #${slot.shapeId} stays as the template shows it.`;
    }

    case 'setSlot': {
      const pattern = patternOf(change.pattern);
      const slide = slideOf(pattern);
      const slot = pattern.slots.find((candidate) => candidate.name === change.slot) ?? reject(`"${pattern.id}" has no slot "${change.slot}"`);
      const { op: _op, pattern: _pattern, slot: _slot, ...patch } = change;
      const keys = Object.keys(patch).filter((key) => (patch as Record<string, unknown>)[key] !== undefined);
      if (keys.length === 0) reject('nothing to set');

      // A setting that belongs to another kind of slot is a mistake worth saying, not ignoring.
      const allowed = new Set<string>(['name', 'description', 'empty', ...KEYS_OF[slot.kind], ...(slot.kind === 'table' ? [] : ['box'])]);
      const foreign = keys.filter((key) => !allowed.has(key));
      if (foreign.length > 0) reject(`"${slot.name}" is ${slot.kind === 'image' ? 'an image' : `a ${slot.kind}`} slot; ${foreign.join(', ')} ${foreign.length === 1 ? 'does' : 'do'} not apply to it`);
      for (const key of keys) {
        const value = (patch as Record<string, unknown>)[key];
        if (COUNT_KEYS.has(key) && value !== null && !(typeof value === 'number' && Number.isInteger(value) && value >= 1)) reject(`${key} must be a whole number of 1 or more`);
      }
      if (patch.empty !== undefined && !['error', 'remove', 'keep', 'drop'].includes(patch.empty)) reject('empty is "error", "remove", "keep" or "drop"');
      if (patch.box) checkBox(patch.box);
      if (typeof patch.fontSize === 'number' && !(patch.fontSize >= 4 && patch.fontSize <= 200)) reject('a font size is between 4 and 200 points');

      const row = inRow(pattern, slot.name);
      if (row && (patch.name !== undefined || patch.deckField !== undefined || patch.box !== undefined)) {
        reject(`"${slot.name}" is part of the row of ${pattern.repeat?.item ?? 'item'}s, so it cannot be renamed, moved or tied to a deck field; remove the row first`);
      }
      // Items of a row must match, so a change to one item's slot is made to that slot of every item.
      const targets = row ? row.names.map((name) => pattern.slots.find((candidate) => candidate.name === name)).filter((candidate): candidate is SlotDef => !!candidate) : [slot];

      const notes: string[] = [];
      for (const target of targets) {
        const shape = slide.shapes.find((candidate) => candidate.id === target.shapeId);
        const derived = derive(target, patch, shape, slide, info);
        const record = target as unknown as Record<string, unknown>;
        for (const [key, value] of Object.entries({ ...patch, ...derived })) {
          if (value === undefined || key === 'name') continue;
          if (value === null || (key === 'empty' && value === 'error')) delete record[key];
          // A text style is changed setting by setting: a new size keeps the colour already chosen.
          else if (STYLE_KEYS.has(key) && record[key] && typeof record[key] === 'object') record[key] = { ...(record[key] as object), ...(value as object) };
          else record[key] = value;
        }
        if (target === slot) for (const [key, value] of Object.entries(derived)) notes.push(`${key} follows at ${value}`);
      }
      if (patch.name !== undefined) {
        const name = slotName(patch.name, '') || reject('a slot name is made of letters and digits and starts with a letter');
        if (pattern.slots.some((candidate) => candidate !== slot && candidate.name === name)) reject(`"${pattern.id}" already has a slot "${name}"`);
        for (const sequence of pattern.sequences ?? []) for (const item of sequence) item.forEach((entry, position) => entry === slot.name && (item[position] = name));
        notes.unshift(`renamed from "${slot.name}"`);
        slot.name = name;
      }
      const who = targets.length > 1 ? `${targets.map((target) => `"${target.name}"`).join(', ')} are` : `"${slot.name}" is`;
      return `"${pattern.id}": ${who} now ${describeSlot(slot)}${notes.length > 0 ? ` (${notes.join('; ')})` : ''}.`;
    }

    case 'setRepeat': {
      const pattern = patternOf(change.pattern);
      const slide = slideOf(pattern);
      let found = false;
      if (!pattern.repeat) {
        const detected =
          detectRepeat(pattern.slots.filter((slot) => !(slot.kind === 'text' && slot.deckField)), slide) ??
          reject(`"${pattern.id}" has no row of like items. Name the slots of each item alike with a number (card1Heading, card1Text, card2Heading, card2Text) and they become a row`);
        pattern.repeat = detected;
        found = true;
      }
      const repeat = pattern.repeat as NonNullable<PatternDef['repeat']>;
      const drawn = repeat.shapes.length;
      let spanned = '';
      if (change.item !== undefined) repeat.item = change.item.trim().toLowerCase() || reject('the item needs a name, such as "card"');
      if (change.span === null) delete repeat.span;
      else if (change.span !== undefined) {
        if (!(change.span.left >= 0 && change.span.right <= 1 && change.span.left < change.span.right)) reject('a span runs from left to right, both between 0 and 1');
        repeat.span = { left: change.span.left, right: change.span.right };
        spanned = ` The row may now spread from ${Math.round(change.span.left * 100)}% to ${Math.round(change.span.right * 100)}% of the slide's width.`;
      }
      const min = change.min ?? repeat.min ?? drawn;
      const max = change.max ?? repeat.max ?? drawn;
      for (const value of [min, max]) if (!Number.isInteger(value) || value < 1) reject('the fewest and most items are whole numbers of 1 or more');
      if (min > max) reject(`the fewest (${min}) cannot be more than the most (${max})`);

      let widened = spanned;
      if (repeat.columns !== undefined) {
        if (max > drawn) reject(`a grid holds at most the ${drawn} ${repeat.item}s it draws; ${max} would need a new line of them`);
      } else {
        // A row may hold more than it draws, as long as the items fit side by side in the room there is.
        const room = rowRoom(slide, repeat.shapes);
        const onSlide = rowOnSlide(slide, repeat);
        const fits = (left: number, right: number) => rowFit(room.itemWidth, left, right);
        const span = repeat.span ?? { left: room.area.x, right: room.area.x + room.area.w };
        // What the template draws always fits, however close its items sit.
        if (max > drawn && max > fits(span.left, span.right)) {
          const most = change.span === undefined && onSlide ? fits(room.left, room.right) : fits(span.left, span.right);
          if (max > most) {
            reject(`${max} ${repeat.item}s do not fit side by side: each is ${Math.round(room.itemWidth * 100)}% of the slide wide and the row has ${Math.round(((change.span === undefined && onSlide ? room.right - room.left : span.right - span.left)) * 100)}%. The most is ${most}`);
          }
          repeat.span = { left: round(room.left), right: round(room.right) };
          widened = ` The row may now spread from ${Math.round(room.left * 100)}% to ${Math.round(room.right * 100)}% of the slide's width.`;
        }
      }
      repeat.min = min;
      repeat.max = max;
      if (change.widen === null || change.widen === 1) {
        if (repeat.widen !== undefined) widened += ` The ${repeat.item}s keep their drawn width whatever their number.`;
        delete repeat.widen;
      } else if (change.widen !== undefined) {
        if (!(typeof change.widen === 'number' && change.widen > 1 && change.widen <= 4)) reject('widen is a factor between 1 and 4, such as 1.5 to let an item become half as wide again');
        if ((repeat.columns ?? drawn) < 2) reject(`a line of the ${repeat.item}s holds one, so there is no shorter line to widen`);
        repeat.widen = change.widen;
        widened += ` With fewer ${repeat.item}s on a line than drawn, each grows up to ${change.widen} times as wide to fill the row, and holds more text.`;
      }
      return `"${pattern.id}": ${found ? `found a ${repeat.columns ? 'grid' : 'row'} of ${drawn} ${repeat.item}s; it` : `the ${repeat.columns ? 'grid' : 'row'} of ${repeat.item}s`} takes ${min === max ? `exactly ${min}` : `${min} to ${max}`}.${widened}`;
    }

    case 'removeRepeat': {
      const pattern = patternOf(change.pattern);
      if (!pattern.repeat) reject(`"${pattern.id}" has no row of items`);
      const repeat = pattern.repeat as NonNullable<PatternDef['repeat']>;
      delete pattern.repeat;
      return `"${pattern.id}": the ${repeat.item}s are no longer a row whose number can change. A deck fills all ${repeat.shapes.length}.`;
    }

    default:
      return reject(`"${String((change as { op?: unknown }).op)}" is not an operation`);
  }
}

const round = (value: number) => Math.round(value * 1000) / 1000;

function checkBox(box: SlideBox): void {
  const numbers = [box.x, box.y, box.w, box.h];
  if (numbers.some((value) => typeof value !== 'number' || Number.isNaN(value)) || box.x < 0 || box.y < 0 || box.w <= 0 || box.h <= 0 || box.x + box.w > 1.001 || box.y + box.h > 1.001) {
    reject('a box is x, y, w and h as fractions of the slide, and must lie on the slide');
  }
}

/** The row a slot belongs to, with the names of that slot across all the row's items. */
function inRow(pattern: PatternDef, name: string): { names: string[] } | undefined {
  const repeat = pattern.repeat;
  if (!repeat) return undefined;
  for (const template of repeat.slots) {
    const names = repeat.shapes.map((_, index) => template.replace('{n}', String(index + 1)));
    if (names.includes(name)) return { names };
  }
  return undefined;
}

/** Limits that follow from a change when the change does not state them: a budget after a new size or area. */
function derive(slot: SlotDef, patch: SlotPatch, shape: ShapeInfo | undefined, slide: SlideInfo, info: TemplateInfo): Record<string, number> {
  const derived: Record<string, number> = {};
  if (slot.kind === 'text' && patch.maxChars === undefined && slot.maxChars !== undefined && shape) {
    let factor = 1;
    if (patch.fontSize !== undefined) {
      const own = slot.size ?? shape.fontSizes?.[0] ?? 14;
      factor *= ((slot.fontSize ?? own) / (patch.fontSize ?? own)) ** 2;
    }
    if (patch.box !== undefined) {
      const before = slot.box ?? shape.box;
      const after = patch.box ?? shape.box;
      if (before.w * before.h > 0) factor *= (after.w * after.h) / (before.w * before.h);
    }
    let budget = Math.max(1, Math.round(slot.maxChars * factor));
    if (patch.box) {
      // A budget scaled up from a small box can promise more than the new area holds, so the area is measured too.
      const size = (patch.fontSize === undefined ? slot.fontSize : patch.fontSize) ?? slot.size ?? shape.fontSizes?.[0] ?? 14;
      const paragraphs = (patch.maxParagraphs === undefined ? slot.maxParagraphs : patch.maxParagraphs) ?? 1;
      const gap = (patch.paragraphGap === undefined ? slot.paragraphGap : patch.paragraphGap) ?? 0;
      budget = Math.min(budget, boxCapacity(patch.box, info, size, paragraphs, gap));
    }
    if (budget !== slot.maxChars) derived.maxChars = budget;
  }
  if (slot.kind === 'table' && patch.maxCellChars === undefined && shape && (patch.bodyText !== undefined || patch.maxColumns !== undefined || patch.maxRows !== undefined)) {
    const size = patch.bodyText === null ? undefined : (patch.bodyText?.size ?? slot.bodyText?.size);
    const measured = tableLimits(shape, info, {
      slide,
      columns: patch.maxColumns ?? slot.maxColumns ?? shape.table?.columns ?? 4,
      rows: patch.maxRows ?? slot.maxRows ?? Math.max(1, (shape.table?.rows ?? 4) - 1),
      ...(size !== undefined ? { fontSize: size } : {}),
    });
    if (measured.maxCellChars !== slot.maxCellChars) derived.maxCellChars = measured.maxCellChars;
  }
  return derived;
}

function describeKind(shape: ShapeInfo): string {
  return shape.kind === 'picture' ? 'a picture' : shape.kind === 'table' ? 'a table' : shape.kind === 'text' ? 'a text box' : `a ${shape.kind}`;
}

/** A slot in a few words: what it holds, how much, and what happens when a deck leaves it empty. */
export function describeSlot(slot: SlotDef): string {
  const policy = slot.empty ?? 'error';
  if (policy === 'drop') return 'a decoration that is always removed';
  if (policy === 'keep' && 'box' in slot && slot.box && slot.kind === 'text' && slot.maxChars === undefined) return 'part of the design, kept as it is in another place';
  const parts: string[] = [];
  if (slot.kind === 'text') {
    parts.push(slot.deckField ? `text from the deck's "${slot.deckField}"` : 'text');
    if (slot.maxChars !== undefined) parts.push(`up to ${slot.maxChars} characters`);
    if (slot.maxParagraphs !== undefined) parts.push(`${slot.maxParagraphs} paragraphs`);
    if (slot.fontSize !== undefined) parts.push(`set at ${slot.fontSize} pt`);
  } else if (slot.kind === 'image') {
    parts.push(slot.imageKind === 'icon' ? 'an icon' : slot.imageKind === 'photo' ? 'a photo' : 'a supplied graphic');
  } else if (slot.kind === 'table') {
    parts.push(`a table of up to ${slot.maxColumns ?? '?'} columns and ${slot.maxRows ?? '?'} rows`);
    if (slot.maxCellChars !== undefined) parts.push(`${slot.maxCellChars} characters a cell`);
    if (slot.bodyText?.size !== undefined) parts.push(`text at ${slot.bodyText.size} pt`);
  } else {
    parts.push(`a native chart of up to ${slot.maxCategories ?? '?'} categories and ${slot.maxSeries ?? '?'} series`);
  }
  if ('box' in slot && slot.box) parts.push('in its own area');
  parts.push(policyWord(policy));
  return parts.join(', ');
}
