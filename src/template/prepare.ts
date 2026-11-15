// Prepares a manifest for use without a person going through it item by item: takes what the
// slides allow, solves the issues that have one answer, and says in a few sentences what is left.
// No model is asked. It is the step a studio runs after a draft, behind one control.
import { renderDeck } from '../render/render.ts';
import type { Template } from '../render/render.ts';
import { measuredBudgets } from './budgets.ts';
import { changeWith } from './change.ts';
import type { ManifestChange } from './change.ts';
import { checkManifest } from './check.ts';
import { VARIANTS } from './derive.ts';
import type { PatternVariant } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import type { TemplateManifest } from './manifest.ts';
import { shapeText } from './measure.ts';
import { sampleDeck } from './sample.ts';
import { suggestChanges } from './suggest.ts';
import type { SuggestionKind } from './suggest.ts';

export interface PrepareOptions {
  /**
   * Whether the app has an icon set for this template, which it gives the planner as `icons`. An
   * icon slot is filled from it. Without one, a layout with a required icon slot is not offered at
   * all, so icon slots are then set to keep the template's own pictogram.
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
}

export interface PreparedManifest {
  manifest: TemplateManifest;
  /** Fit to use: `checkManifest` reports no error, and the sample deck renders at the fewest and at the most items. */
  sound: boolean;
  /** What was done, in sentences, in the order it was done. */
  done: string[];
  /** What only a person can settle, in sentences. A short list by design: it is what a studio shows first. */
  left: string[];
  /** The ids of the patterns that were added, all of them variants. */
  added: string[];
  /** The ids of the patterns that were taken out because they could not be rendered. */
  removed: string[];
  /** A value for each deck field: the text the template shows in its place. For the app to use until a person sets one. */
  fields: Record<string, string>;
}

const clip = (text: string, most = 48) => (text.length > most ? `${text.slice(0, most - 1).trimEnd()}…` : text);

/**
 * Prepares a manifest for use in one step. It takes every suggestion of `suggestChanges` that is a
 * fact of the slide (`icon`, `row`, `row-max`), lets rows widen, adds the mirrored and text-only
 * variants, and then settles the rows of those variants. It keeps the template's own pictograms
 * where the app has no icon set, sets the measured budget where a budget lets text run over its
 * box, and takes out a pattern that cannot be rendered. Nothing is renamed, so stored decks keep
 * working, and a second run changes nothing.
 *
 * The template passed in is not touched. Use it after `draftManifest`, or on a manifest that is
 * already in use. What it cannot decide, such as whether a description says the right thing, comes
 * back in `left`.
 */
export function prepareManifest(template: Template, options: PrepareOptions): PreparedManifest {
  if (typeof options?.icons !== 'boolean') throw new Error('say whether the app has an icon set for this template: icons is true or false');
  const widen = options.widen ?? 2;
  if (widen !== false && !(typeof widen === 'number' && widen > 1 && widen <= 4)) throw new Error('widen is a factor above 1 and at most 4, such as 2, or false to leave rows as they are');
  const variants = new Set<PatternVariant>(options.variants ?? ['mirror', 'text-only']);
  for (const variant of variants) if (!VARIANTS.includes(variant)) throw new Error(`"${String(variant)}" is not a variant; the variants are ${VARIANTS.join(', ')}`);

  const info = inspectTemplate(template.pkg);
  const before = new Set(template.manifest.patterns.map((pattern) => pattern.id));
  let current: Template = template;
  const done: string[] = [];
  const apply = (changes: ManifestChange[]): number => {
    if (changes.length === 0) return 0;
    const result = changeWith(current, changes, info);
    for (const entry of result.applied) if (!done.includes(entry.message)) done.push(entry.message);
    current = { ...current, manifest: result.manifest };
    return result.applied.length;
  };

  // 1. What the slides allow: the facts first, then how rows behave, then the variants, then the
  // rows once more, because a variant without its picture has room its pattern did not have.
  const rows: SuggestionKind[] = widen === false ? ['row-max'] : ['row-max', 'row-widen'];
  const kinds: SuggestionKind[] = ['icon', 'row', ...rows, ...(variants.size > 0 ? (['variant'] as const) : []), ...rows];
  for (const kind of kinds) {
    for (let turn = 0; turn < 3; turn++) {
      const offered = suggestChanges(current, { kinds: [kind], ...(widen === false ? {} : { widen }) }).filter((suggestion) => suggestion.variant === undefined || variants.has(suggestion.variant));
      if (apply(offered.flatMap((suggestion) => suggestion.changes)) === 0) break;
    }
  }

  // 2. Without an icon set, a required icon slot keeps its layout from being offered. The
  // template's own pictogram stays instead. One change reaches every item of a row.
  if (!options.icons) {
    apply(
      current.manifest.patterns.flatMap((pattern) =>
        pattern.slots
          .filter((slot) => slot.kind === 'image' && slot.imageKind === 'icon' && slot.empty !== 'keep' && slot.empty !== 'drop')
          .map((slot): ManifestChange => ({ op: 'setSlot', pattern: pattern.id, slot: slot.name, empty: 'keep' })),
      ),
    );
  }

  // 3. The issues that have one answer.
  // A budget that lets text run over its box, which is reported when the template has its fonts: the measured budget.
  if (current.fonts) {
    const over = new Set(
      checkManifest(current, info)
        .filter((issue) => issue.message.startsWith('has a budget of '))
        .map((issue) => `${issue.pattern}\n${issue.slot}`),
    );
    if (over.size > 0) {
      apply(
        measuredBudgets(current)
          .filter((budget) => over.has(`${budget.pattern}\n${budget.slot}`))
          .map((budget): ManifestChange => ({ op: 'setSlot', pattern: budget.pattern, slot: budget.slot, maxChars: budget.suggested })),
      );
    }
  }
  // A grid that is said to take more than it draws: what it draws.
  apply(
    current.manifest.patterns
      .filter((pattern) => pattern.repeat?.columns !== undefined && (pattern.repeat.max ?? 0) > pattern.repeat.shapes.length)
      .map((pattern): ManifestChange => ({ op: 'setRepeat', pattern: pattern.id, max: pattern.repeat?.shapes.length ?? 1, min: Math.min(pattern.repeat?.min ?? 1, pattern.repeat?.shapes.length ?? 1) })),
  );
  // A pattern with an error cannot be rendered. It is taken out, and said.
  const removed: string[] = [];
  for (let turn = 0; turn < 3; turn++) {
    const broken = [...new Set(checkManifest(current, info).filter((issue) => issue.severity === 'error' && issue.pattern !== undefined).map((issue) => issue.pattern as string))];
    if (broken.length === 0) break;
    const reasons = new Map(checkManifest(current, info).filter((issue) => issue.severity === 'error').map((issue) => [issue.pattern, `${issue.slot ?? 'it'} ${issue.message}`]));
    const result = changeWith(current, broken.map((pattern): ManifestChange => ({ op: 'removePattern', pattern })), info);
    if (result.applied.length === 0) break;
    for (const entry of result.applied) {
      const id = (entry.change as { pattern: string }).pattern;
      removed.push(id);
      done.push(`"${id}" was taken out, because it cannot be rendered: ${reasons.get(id) ?? 'the manifest does not match the pptx'}.`);
    }
    current = { ...current, manifest: result.manifest };
  }

  // 4. A value for each deck field: the text the template itself shows there.
  const fields: Record<string, string> = {};
  for (const pattern of current.manifest.patterns) {
    const shapes = info.slides[pattern.sourceSlide - 1]?.shapes ?? [];
    for (const slot of pattern.slots) {
      if (slot.kind !== 'text' || !slot.deckField || fields[slot.deckField] !== undefined) continue;
      const shape = shapes.find((candidate) => candidate.id === slot.shapeId);
      const text = shape ? shapeText(shape) : '';
      if (text) fields[slot.deckField] = text;
    }
  }

  // 5. Is it fit to use, and what is left for a person.
  if (current.manifest.patterns.length === 0) {
    return { manifest: current.manifest, sound: false, done, left: ['The manifest has no layout. Draft the template first.'], added: [], removed, fields };
  }
  const remaining = checkManifest(current, info);
  const errors = remaining.filter((issue) => issue.severity === 'error');
  const warnings = remaining.filter((issue) => issue.severity === 'warning' && !issue.message.startsWith('no font was supplied'));
  let renders = true;
  let failure = '';
  for (const items of ['min', 'max'] as const) {
    try {
      const { deck, assets } = sampleDeck(current, { items });
      renderDeck(current, deck, { assets, layout: 'adaptive' });
    } catch (error) {
      renders = false;
      failure = error instanceof Error ? error.message : String(error);
    }
  }
  const say = (issue: { pattern?: string; slot?: string; message: string }) => `${[issue.pattern, issue.slot].filter(Boolean).join('.') || 'the manifest'} ${issue.message}`;
  const left: string[] = [];
  if (errors.length > 0) left.push(`${errors.length === 1 ? 'One error' : `${errors.length} errors`} could not be solved: ${errors.slice(0, 3).map(say).join('; ')}${errors.length > 3 ? '; and more' : ''}.`);
  if (failure) left.push(`The sample deck does not render: ${failure}`);
  if (warnings.length > 0) left.push(`${warnings.length === 1 ? 'One warning' : `${warnings.length} warnings`} to look at: ${warnings.slice(0, 3).map(say).join('; ')}${warnings.length > 3 ? '; and more' : ''}.`);
  const named = Object.entries(fields);
  if (named.length > 0) left.push(`${named.length === 1 ? 'One text is' : `${named.length} texts are`} the same on every slide and ${named.length === 1 ? 'shows' : 'show'} what the template shows until you set your own: ${named.map(([name, text]) => `${name} ("${clip(text)}")`).join(', ')}.`);
  if (!current.fonts) left.push("Text was not measured with the template's fonts, so budgets are estimates. Look at the sample deck for text that runs over its box.");
  else {
    const missing = remaining.find((issue) => issue.message.startsWith('no font was supplied'));
    if (missing) left.push(`Some text was not measured: ${missing.message}.`);
  }
  left.push('Read the description of each layout once. The planner chooses a layout by it, so it should say what the layout holds and not what the sample is about.');

  return {
    manifest: current.manifest,
    sound: errors.length === 0 && renders,
    done,
    left,
    added: current.manifest.patterns.map((pattern) => pattern.id).filter((id) => !before.has(id)),
    removed,
    fields,
  };
}
