// Prepares a manifest for use without a person going through it item by item: takes what the
// slides allow, solves the issues that have one answer, and says in a few sentences what is left.
// No model is asked. It is the step a studio runs after a draft, behind one control.
import { renderDeck } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { measuredBudgets } from './budgets.ts';
import { changeWith } from './change.ts';
import type { ManifestChange } from './change.ts';
import { checkManifest } from './check.ts';
import type { ManifestIssue } from './check.ts';
import { VARIANTS } from './derive.ts';
import type { PatternVariant } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import type { PatternDef, TemplateManifest } from './manifest.ts';
import { shapeText } from './measure.ts';
import { detectRepeat, strandedSlots } from './rows.ts';
import { sampleDeck } from './sample.ts';
import { suggestChanges } from './suggest.ts';
import type { SuggestionKind } from './suggest.ts';

export interface PrepareOptions {
  /**
   * Whether the app has an icon set for this template, which it gives the planner as `icons`. An
   * icon slot is filled from it. Without one, a layout with a required icon slot is not offered at
   * all, so required icon slots are then set to keep the template's own pictogram.
   */
  icons: boolean;
  /**
   * The variants to add wherever a slide allows them. Defaults to `mirror` and `text-only`. An empty
   * list adds none. `without` is left out by default: there is one for every optional text, and an
   * optional text can be left empty anyway.
   */
  variants?: readonly PatternVariant[];
  /**
   * How much the items of a row or grid may grow when a line holds fewer than the template draws,
   * as a factor above 1 and at most 4. Defaults to 2. `false` leaves the items at their drawn width.
   */
  widen?: number | false;
  /**
   * Whether a pattern that cannot be used may be taken out: one for which `checkManifest` reports
   * an error, or whose sample cannot be rendered. Defaults to `true`, which is right after a draft.
   * Pass `false` for a manifest that stored decks use: the pattern then stays, `sound` is `false`,
   * and `left` says what is wrong with it.
   */
  remove?: boolean;
  /** The values the app already has for deck fields. A field that has one is not asked about in `left`. */
  fields?: Record<string, string>;
}

export interface PreparedManifest {
  manifest: TemplateManifest;
  /** Fit to use: `checkManifest` reports no error, and the sample deck renders at the fewest and at the most items. */
  sound: boolean;
  /** What was done, in sentences, in the order it was done. */
  done: string[];
  /** What only a person can settle, in sentences. A short list by design: it is what a studio shows first. */
  left: string[];
  /** The ids the manifest has now and did not have before, all of them variants. */
  added: string[];
  /** The ids the manifest had and no longer has: patterns that were taken out because they cannot be used. Decks made with them can no longer be changed. */
  removed: string[];
  /** The icon slots that were set to keep the template's pictogram because the app has no icon set. Fill them again when it gets one. */
  kept: Array<{ pattern: string; slot: string }>;
  /** A value for each deck field: the one the app gave, or the text the template shows in its place. */
  fields: Record<string, string>;
}

const clip = (text: string, most = 48) => (text.length > most ? `${text.slice(0, most - 1).trimEnd()}…` : text);
const quote = (names: string[]) => names.map((name) => `"${name}"`).join(', ');

/**
 * Prepares a manifest for use in one step. It takes every suggestion of `suggestChanges` that is a
 * fact of the slide (`icon`, `row`, `row-max`), lets rows widen, adds the mirrored and text-only
 * variants, and then settles the rows of those variants. It keeps the template's own pictograms
 * where the app has no icon set, sets the measured budget where a budget lets text run over its
 * box, and takes out a pattern that cannot be used. No pattern and no slot is renamed, and a
 * second run changes nothing.
 *
 * The template passed in is not touched. Use it after `draftManifest`. For a manifest that is in
 * use, pass `remove: false`, so that every pattern stored decks were made with stays. What it
 * cannot decide, such as whether a description says the right thing, comes back in `left`.
 *
 * It takes again what a person undid: a variant that was removed by hand is added again, and a row
 * that was told not to widen widens again. Pass `variants: []` and `widen: false` to settle the
 * facts only.
 */
export function prepareManifest(template: Template, options: PrepareOptions): PreparedManifest {
  if (typeof options?.icons !== 'boolean') throw new Error('say whether the app has an icon set for this template: icons is true or false');
  const widen = options.widen ?? 2;
  if (widen !== false && !(typeof widen === 'number' && widen > 1 && widen <= 4)) throw new Error('widen is a factor above 1 and at most 4, such as 2, or false to leave rows as they are');
  const variants = new Set<PatternVariant>(options.variants ?? ['mirror', 'text-only']);
  for (const variant of variants) if (!VARIANTS.includes(variant)) throw new Error(`"${String(variant)}" is not a variant; the variants are ${VARIANTS.join(', ')}`);
  const remove = options.remove ?? true;

  const info = inspectTemplate(template.pkg);
  const before = new Set(template.manifest.patterns.map((pattern) => pattern.id));
  let current: Template = template;
  const done: string[] = [];
  const reasons = new Map<string, string>();
  const apply = (changes: ManifestChange[]): number => {
    if (changes.length === 0) return 0;
    const result = changeWith(current, changes, info);
    for (const entry of result.applied) if (!done.includes(entry.message)) done.push(entry.message);
    current = { ...current, manifest: result.manifest };
    return result.applied.length;
  };
  const takeOut = (ids: string[], why: (id: string) => string) => {
    const result = changeWith(current, ids.map((pattern): ManifestChange => ({ op: 'removePattern', pattern })), info);
    for (const entry of result.applied) {
      const id = (entry.change as { pattern: string }).pattern;
      reasons.set(id, why(id));
      done.push(`"${id}" was taken out: ${reasons.get(id)}.`);
    }
    current = { ...current, manifest: result.manifest };
  };
  const take = (kind: SuggestionKind) => {
    for (let turn = 0; turn < 3; turn++) {
      const offered = suggestChanges(current, { kinds: [kind], ...(widen === false ? {} : { widen }) }).filter((suggestion) => suggestion.variant === undefined || variants.has(suggestion.variant));
      if (apply(offered.flatMap((suggestion) => suggestion.changes)) === 0) break;
    }
  };

  // 1. A row whose items hold a slot that is not the row's, as an earlier draft made one where the
  // pictogram of each card was labelled a decoration. On a slide with fewer items that slot's
  // shape is gone with its item, and the slide cannot be rendered. Where every item has such a
  // slot under its own number and no deck fills it, it becomes part of the row. Otherwise the row
  // takes no fewer items than the one the slot is in. Every deck that rendered still does.
  const settled = current.manifest.patterns.map((pattern): PatternDef => {
    const slide = info.slides[pattern.sourceSlide - 1];
    if (!pattern.repeat || !slide) return pattern;
    let repeat = pattern.repeat;
    const fewest = repeat.min ?? repeat.shapes.length;
    const outside = () => strandedSlots({ ...pattern, repeat }, slide).filter((entry) => entry.item > fewest);
    if (outside().length === 0) return pattern;
    const at = new Map(strandedSlots(pattern, slide).map((entry) => [entry.slot, entry.item]));
    const unfilled = (name: string, item: number) => {
      const slot = pattern.slots.find((entry) => entry.name === name.replace('{n}', String(item)));
      return slot !== undefined && (slot.empty === 'keep' || slot.empty === 'drop') && at.get(slot.name) === item;
    };
    const found = detectRepeat(pattern.slots.filter((slot) => !(slot.kind === 'text' && slot.deckField)), slide);
    const more = (found?.slots ?? []).filter((name) => !repeat.slots.includes(name) && repeat.shapes.every((_, index) => unfilled(name, index + 1)));
    if (more.length > 0) {
      repeat = { ...repeat, slots: [...repeat.slots, ...more] };
      done.push(`"${pattern.id}": ${quote(more)} ${more.length === 1 ? 'is' : 'are'} now part of the row of ${repeat.item}s, because the shape lies in each ${repeat.item}. A slide with fewer ${repeat.item}s could not be rendered.`);
    }
    const rest = outside();
    if (rest.length > 0) {
      const item = Math.max(...rest.map((entry) => entry.item));
      const slots = rest.filter((entry) => entry.item === item).map((entry) => entry.slot);
      repeat = { ...repeat, min: item, ...(item > (repeat.max ?? repeat.shapes.length) ? { max: item } : {}) };
      done.push(`"${pattern.id}": the row of ${repeat.item}s takes no fewer than ${item}, because ${quote(slots)} ${slots.length === 1 ? 'lies' : 'lie'} in ${repeat.item} ${item} and ${slots.length === 1 ? 'is' : 'are'} not part of the row. A slide with fewer could not be rendered.`);
    }
    return { ...pattern, repeat };
  });
  if (settled.some((pattern, index) => pattern !== current.manifest.patterns[index])) current = { ...current, manifest: { ...current.manifest, patterns: settled } };

  // 2. A pattern whose slots do not match the pptx cannot be used. Taken out before anything is
  // made of it. Two patterns with one id are left for a person: nothing says which one is meant.
  const faulty = (issue: ManifestIssue) => issue.severity === 'error' && issue.pattern !== undefined && issue.message !== 'another pattern has the same id';
  if (remove) {
    for (let turn = 0; turn < 3; turn++) {
      const errors = checkManifest(current, info).filter(faulty);
      const ids = [...new Set(errors.map((issue) => issue.pattern as string))].filter((id) => current.manifest.patterns.filter((pattern) => pattern.id === id).length === 1);
      if (ids.length === 0) break;
      takeOut(ids, (id) => {
        const first = errors.find((issue) => issue.pattern === id) as ManifestIssue;
        return `${first.slot ? `"${first.slot}" ` : /^(points|has|is) /.test(first.message) ? 'it ' : ''}${first.message}`;
      });
    }
  }
  // A grid that is said to take more than it draws: what it draws.
  apply(
    current.manifest.patterns
      .filter((pattern) => pattern.repeat?.columns !== undefined && (pattern.repeat.max ?? 0) > pattern.repeat.shapes.length)
      .map((pattern): ManifestChange => ({ op: 'setRepeat', pattern: pattern.id, max: pattern.repeat?.shapes.length ?? 1 })),
  );

  // 3. Pictures that are pictograms become icon slots. Then, without an icon set, a required icon
  // slot would keep its layout from being offered: the template's own pictogram stays instead. An
  // optional one is left alone: a deck may leave it out as before.
  take('icon');
  const kept: PreparedManifest['kept'] = [];
  if (!options.icons) {
    const required = current.manifest.patterns.flatMap((pattern) =>
      pattern.slots.filter((slot) => slot.kind === 'image' && slot.imageKind === 'icon' && (slot.empty ?? 'error') === 'error').map((slot) => ({ pattern: pattern.id, slot: slot.name })),
    );
    apply(required.map((entry): ManifestChange => ({ op: 'setSlot', pattern: entry.pattern, slot: entry.slot, empty: 'keep' })));
    for (const entry of required) {
      const now = current.manifest.patterns.find((pattern) => pattern.id === entry.pattern)?.slots.find((slot) => slot.name === entry.slot);
      if (now?.empty === 'keep') kept.push(entry);
    }
  }

  // 4. What the slides allow: rows, then the variants, then the rows once more, because a variant
  // without its picture has room its pattern did not have.
  const rows: SuggestionKind[] = widen === false ? ['row-max'] : ['row-max', 'row-widen'];
  for (const kind of ['row', ...rows, ...(variants.size > 0 ? (['variant'] as const) : []), ...rows] as SuggestionKind[]) take(kind);

  // 5. A budget that lets text run over its box, which is reported when the template has its
  // fonts: the measured budget. The items of a row share one budget, the smallest any of them
  // measures, because a change to a slot of a row is made to that slot of every item.
  if (current.fonts) {
    const familyOf = (patternId: string | undefined, slot: string | undefined) => {
      const repeat = current.manifest.patterns.find((pattern) => pattern.id === patternId)?.repeat;
      const name = repeat?.slots.find((entry) => repeat.shapes.some((_, index) => entry.replace('{n}', String(index + 1)) === slot));
      return `${patternId}\n${name ?? slot}`;
    };
    const over = new Set(
      checkManifest(current, info)
        .filter((issue) => issue.message.startsWith('has a budget of '))
        .map((issue) => familyOf(issue.pattern, issue.slot)),
    );
    if (over.size > 0) {
      const smallest = new Map<string, { pattern: string; slot: string; suggested: number }>();
      const has = new Map<string, number>();
      for (const budget of measuredBudgets(current)) {
        const key = familyOf(budget.pattern, budget.slot);
        if (!over.has(key)) continue;
        if (budget.suggested < (smallest.get(key)?.suggested ?? Infinity)) smallest.set(key, budget);
        const slot = current.manifest.patterns.find((pattern) => pattern.id === budget.pattern)?.slots.find((entry) => entry.name === budget.slot);
        has.set(key, Math.max(has.get(key) ?? 0, slot?.kind === 'text' ? (slot.maxChars ?? Infinity) : 0));
      }
      // Only a budget that gets smaller: the one a slot has may be the measured one already, and
      // still be too much for the wider box of a row with fewer items.
      const lower = [...smallest].filter(([key, budget]) => budget.suggested < (has.get(key) ?? 0)).map(([, budget]) => budget);
      apply(lower.map((budget): ManifestChange => ({ op: 'setSlot', pattern: budget.pattern, slot: budget.slot, maxChars: budget.suggested })));
    }
  }

  // 6. Does it render. A pattern whose own sample does not is taken out, when that is allowed.
  const renders = (manifest: TemplateManifest): string => {
    if (manifest.patterns.length === 0) return '';
    const one: Template = { ...current, manifest };
    for (const items of ['min', 'max'] as const) {
      try {
        const { deck, assets } = sampleDeck(one, { items });
        renderDeck(one, deck, { assets, layout: 'adaptive' });
      } catch (error) {
        return error instanceof Error ? error.message : String(error);
      }
    }
    return '';
  };
  let failure = renders(current.manifest);
  if (failure && remove) {
    const failing = current.manifest.patterns.map((pattern) => [pattern.id, renders({ ...current.manifest, patterns: [pattern] })] as const).filter(([, message]) => message);
    const why = new Map(failing);
    takeOut(failing.map(([id]) => id), (id) => `its sample cannot be rendered (${clip(why.get(id) ?? '', 120)})`);
    failure = renders(current.manifest);
  }

  // 7. A value for each deck field: the one the app has, or the text the template itself shows there.
  const fields: Record<string, string> = {};
  const unset: string[] = [];
  for (const pattern of current.manifest.patterns) {
    const shapes = info.slides[pattern.sourceSlide - 1]?.shapes ?? [];
    for (const slot of pattern.slots) {
      if (slot.kind !== 'text' || !slot.deckField || fields[slot.deckField] !== undefined) continue;
      const given = options.fields?.[slot.deckField];
      if (typeof given === 'string' && given.trim()) {
        fields[slot.deckField] = given;
        continue;
      }
      const shape = shapes.find((candidate) => candidate.id === slot.shapeId);
      const text = shape ? shapeText(shape) : '';
      if (text) {
        fields[slot.deckField] = text;
        unset.push(slot.deckField);
      }
    }
  }

  // 8. What is left for a person.
  const now = new Set(current.manifest.patterns.map((pattern) => pattern.id));
  const added = [...now].filter((id) => !before.has(id));
  const removed = [...before].filter((id) => !now.has(id));
  if (current.manifest.patterns.length === 0) return { manifest: current.manifest, sound: false, done, left: ['The manifest has no layout. Draft the template first.'], added, removed, kept, fields };
  const remaining = checkManifest(current, info);
  const errors = remaining.filter((issue) => issue.severity === 'error');
  const warnings = remaining.filter((issue) => issue.severity === 'warning' && !issue.message.startsWith('no font was supplied'));
  const say = (issue: ManifestIssue) => `${[issue.pattern, issue.slot].filter(Boolean).join('.') || 'the manifest'} ${issue.message}`;
  const left: string[] = [];
  if (removed.length > 0) left.push(`${removed.length === 1 ? 'One layout was' : `${removed.length} layouts were`} taken out because ${removed.length === 1 ? 'it' : 'they'} cannot be used: ${removed.map((id) => `"${id}" (${reasons.get(id)})`).join('; ')}. A presentation made with ${removed.length === 1 ? 'it' : 'one of them'} can no longer be changed.`);
  if (errors.length > 0) left.push(`${errors.length === 1 ? 'One error' : `${errors.length} errors`} could not be solved: ${errors.slice(0, 3).map(say).join('; ')}${errors.length > 3 ? '; and more' : ''}.`);
  if (failure) left.push(`The sample deck does not render: ${failure}`);
  if (warnings.length > 0) left.push(`${warnings.length === 1 ? 'One warning' : `${warnings.length} warnings`} to look at: ${warnings.slice(0, 3).map(say).join('; ')}${warnings.length > 3 ? '; and more' : ''}.`);
  if (unset.length > 0) left.push(`${unset.length === 1 ? 'One text is' : `${unset.length} texts are`} the same on every slide and ${unset.length === 1 ? 'shows' : 'show'} what the template shows until you set your own: ${unset.map((name) => `${name} ("${clip(fields[name] as string)}")`).join(', ')}.`);
  if (!current.fonts) left.push("Text was not measured with the template's fonts, so budgets are estimates. Look at the sample deck for text that runs over its box.");
  else {
    const missing = remaining.find((issue) => issue.message.startsWith('no font was supplied'));
    if (missing) left.push(`Some text was not measured: ${missing.message}.`);
  }
  left.push('Read the description of each layout once. The planner chooses a layout by it, so it should say what the layout holds and not what the sample is about.');

  return { manifest: current.manifest, sound: errors.length === 0 && !failure, done, left, added, removed, kept, fields };
}
