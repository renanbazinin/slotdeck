// The part of planning that planDeck and planEdit share: asking the model for the content of
// a set of slides, turning its answer into slide specs, and sending drafts back for repair.
import { validateDeck } from '../deck/validate.ts';
import type { Issue, LayoutMode, ValidateOptions } from '../deck/validate.ts';
import { isImageRef } from '../deck/spec.ts';
import type { DeckSpec, ImageRef, SlideSpec, SlotValue } from '../deck/spec.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { textBudget } from '../template/manifest.ts';
import type { TemplateManifest } from '../template/manifest.ts';
import type { TextMeasurer } from '../text/measure.ts';
import { readImageInfo } from '../util/image.ts';
import { budgetScale, contentSchema } from './catalog.ts';
import type { CatalogContext, UsablePattern, WriteItem } from './catalog.ts';

export function systemPrompt(language: string, source: 'brief' | 'deck'): string {
  return [
    "You write presentation decks for a fixed slide template. You do not design layouts: you choose from the template's slide patterns and write the content for their slots.",
    '',
    'Rules:',
    `- Write all slide text in the language with BCP 47 tag "${language}".`,
    '- Character limits are hard and count spaces. Stay under them. Slides carry short text: a phrase beats a sentence.',
    '- Plain text only: no markdown, no bullet characters, no line breaks inside a string, no placeholder text.',
    source === 'brief'
      ? '- Every figure on a slide must come from the brief. Never make up statistics, quotes or names.'
      : '- Every figure on a slide must come from the request or from the deck as it is. Never make up statistics, quotes or names.',
    source === 'brief' ? '- Follow the brief. Do not add topics it does not call for.' : '- Do what the request asks and nothing more.',
  ].join('\n');
}

export function cleanText(value: unknown): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Stable short id for a generated picture, so the same description maps to the same asset across edits. */
export function generatedAssetId(prompt: string): string {
  let a = 0x811c9dc5;
  let b = 0x01000193;
  for (let i = 0; i < prompt.length; i++) {
    const code = prompt.charCodeAt(i);
    a = Math.imul(a ^ code, 0x01000193);
    b = Math.imul(b ^ code, 0x85ebca6b) + 0x9e3779b9;
  }
  const hex = (value: number) => (value >>> 0).toString(16).padStart(8, '0');
  return `gen-${hex(a)}${hex(b)}`;
}

/** Turns the model's answer for one slide into a slide spec, tidying text on the way. */
export function toSlide(pattern: UsablePattern, content: unknown, context: CatalogContext): SlideSpec {
  const values = (typeof content === 'object' && content !== null ? content : {}) as Record<string, unknown>;
  const slots: Record<string, SlotValue> = {};

  for (const planned of pattern.slots) {
    const value = values[planned.slot.name];
    if (value === undefined || value === null) continue;

    if (planned.mode === 'text') {
      slots[planned.slot.name] = cleanText(Array.isArray(value) ? value.join(' ') : value);
    } else if (planned.mode === 'paragraphs') {
      const items = Array.isArray(value) ? value : String(value).split(/\r?\n/);
      slots[planned.slot.name] = items.map((item) => cleanText(item).replace(/^[•\-–]\s+/, '')).filter(Boolean);
    } else if (planned.mode === 'table') {
      const table = (typeof value === 'object' ? value : {}) as { header?: unknown; rows?: unknown };
      const header = (Array.isArray(table.header) ? table.header : []).map(cleanText);
      if (header.length === 0) continue;
      // A row the model made too short or too long is squared up rather than rejected.
      const rows = (Array.isArray(table.rows) ? table.rows : [])
        .filter((row): row is unknown[] => Array.isArray(row))
        .map((row) => header.map((_, index) => cleanText(row[index])));
      slots[planned.slot.name] = { header, rows };
    } else if (planned.mode === 'chart') {
      const chart = (typeof value === 'object' ? value : {}) as { type?: unknown; categories?: unknown; series?: unknown; format?: unknown };
      const type = chart.type === 'bar' || chart.type === 'line' || chart.type === 'pie' ? chart.type : 'column';
      const categories = (Array.isArray(chart.categories) ? chart.categories : []).map(cleanText);
      const series = (Array.isArray(chart.series) ? chart.series : []).map((entry) => {
        const item = (entry ?? {}) as { name?: unknown; values?: unknown };
        // Lengths are left as the model gave them: a short series is a problem to send back, not to pad.
        return { name: cleanText(item.name), values: (Array.isArray(item.values) ? item.values : []).map(Number).filter(Number.isFinite) };
      });
      if (categories.length === 0 || series.length === 0) continue;
      slots[planned.slot.name] = chart.format === 'percent' ? { type, categories, series, format: 'percent' } : { type, categories, series };
    } else if (planned.mode === 'photo') {
      if (typeof value !== 'object') continue;
      const photo = value as { prompt?: unknown; alt?: unknown; asset?: unknown };
      const alt = cleanText(photo.alt);
      const supplied = context.assets.find((asset) => asset.kind === 'photo' && asset.id === photo.asset);
      const prompt = cleanText(photo.prompt);
      let ref: ImageRef | undefined;
      if (supplied) ref = { asset: supplied.id };
      else if (prompt) ref = { asset: generatedAssetId(prompt), prompt };
      if (!ref) continue;
      if (alt) ref.alt = alt;
      slots[planned.slot.name] = ref;
    } else {
      slots[planned.slot.name] = { asset: cleanText(value) };
    }
  }
  return { pattern: pattern.def.id, slots };
}

/** The icon and picture lists appended to a writing prompt, when the slides being written can use them. */
export function suppliesText(items: WriteItem[], context: CatalogContext): string[] {
  const usesIcons = items.some(({ pattern }) => pattern.slots.some((planned) => planned.mode === 'icon'));
  const usesPictures = items.some(({ pattern }) => pattern.slots.some((planned) => planned.mode === 'photo' || planned.mode === 'graphic'));
  return [
    usesIcons ? `Icons to choose from:\n${context.icons.map((icon) => `- ${icon.id}: ${icon.description}`).join('\n')}` : '',
    usesPictures && context.assets.length > 0
      ? `Supplied pictures. Match a picture's proportions to the slide: a wide picture for a wide area.\n${context.assets
          .map((asset) => `- ${asset.id} (${asset.kind}, ${proportions(asset.data)}): ${asset.description}`)
          .join('\n')}`
      : '',
  ].filter(Boolean);
}

/** A picture's shape in words the model can match to a slide, e.g. "wide, 3.2:1". */
function proportions(data: Uint8Array): string {
  try {
    const { width, height } = readImageInfo(data);
    const ratio = width / height;
    const shape = ratio >= 2 ? 'wide' : ratio > 1.15 ? 'landscape' : ratio >= 0.87 ? 'square' : 'portrait';
    return `${shape}, ${ratio >= 1 ? `${Math.round(ratio * 10) / 10}:1` : `1:${Math.round((1 / ratio) * 10) / 10}`}`;
  } catch {
    return 'unknown size';
  }
}

/** Ids the renderer will be able to resolve: supplied icons and assets, plus pictures that carry a description. */
function knownAssets(deck: DeckSpec, context: CatalogContext): Set<string> {
  const known = new Set<string>([...context.icons.map((icon) => icon.id), ...context.assets.map((asset) => asset.id)]);
  for (const slide of deck.slides) {
    for (const value of Object.values(slide.slots)) {
      if (isImageRef(value) && value.prompt) known.add(value.asset);
    }
  }
  return known;
}

export interface WriteJob extends WriteItem {
  /** Slot values taken over unchanged from an existing slide. They are not offered to the model. */
  carried?: Record<string, SlotValue>;
}

export interface WriteArgs {
  ask: (prompt: string, schema: JsonSchema, room?: number) => Promise<unknown>;
  /** The writing prompt. Repair rounds append the previous answer and its problems. */
  request: string;
  jobs: WriteJob[];
  context: CatalogContext;
  /** Builds the whole deck from the freshly written slides, keyed by slide number. */
  assemble: (written: Map<number, SlideSpec>) => DeckSpec;
  manifest: TemplateManifest;
  /** Measures text with the template's fonts, so drafts are checked and repaired by the lines they take. */
  measure?: TextMeasurer | undefined;
  layout: LayoutMode;
  maxRepairs: number;
  onRepair?: (attempt: number, issues: number) => void;
}

/**
 * Asks for the content of the given slides and checks the resulting deck. Problems on the
 * written slides go back to the model, up to `maxRepairs` times. Returns the last draft with
 * every issue it still has; the caller decides what is fatal.
 */
/**
 * How many tokens the content of some slides may take: every slot at its budget, with room to
 * spare. It is a ceiling for the model's answer, so that an answer that runs away is cut short.
 */
export function contentRoom(jobs: Array<{ pattern: UsablePattern }>, context: CatalogContext): number {
  let characters = 0;
  let slots = 0;
  for (const { pattern } of jobs) {
    for (const planned of pattern.slots) {
      slots++;
      const { slot } = planned;
      if (slot.kind === 'text') characters += textBudget(slot, budgetScale(context)).maxChars ?? 400;
      else if (slot.kind === 'table') characters += (slot.maxColumns ?? 4) * ((slot.maxRows ?? 4) + 1) * (slot.maxCellChars ?? 30);
      else if (slot.kind === 'chart') characters += 400;
      else characters += planned.mode === 'photo' ? 500 : 60; // a photo is described in a sentence or two
    }
  }
  // A token is rarely shorter than a character, even in scripts that tokenise badly; names and punctuation come on top.
  return Math.ceil((characters + slots * 40) * 1.5) + 500;
}

export async function writeSlides(args: WriteArgs): Promise<{ deck: DeckSpec; issues: Issue[] }> {
  const { jobs, context } = args;
  const schema = contentSchema(jobs, context);
  const numbers = new Set(jobs.map((job) => job.number));

  const build = (answer: unknown): { deck: DeckSpec; issues: Issue[] } => {
    const content = (typeof answer === 'object' && answer !== null ? answer : {}) as Record<string, unknown>;
    const written = new Map<number, SlideSpec>();
    for (const job of jobs) {
      const slide = toSlide(job.pattern, content[`s${job.number}`], context);
      const next: SlideSpec = { pattern: slide.pattern, slots: { ...job.carried, ...slide.slots } };
      if (job.pattern.items !== undefined) next.items = job.pattern.items;
      written.set(job.number, next);
    }
    const deck = args.assemble(written);
    const options: ValidateOptions = { layout: args.layout, assets: knownAssets(deck, context) };
    if (args.measure) options.measure = args.measure;
    return { deck, issues: validateDeck(args.manifest, deck, options) };
  };
  // Only slides that were just written can be fixed by asking again.
  const repairable = (issues: Issue[]) => issues.filter((issue) => issue.slide !== undefined && numbers.has(issue.slide));

  const room = contentRoom(jobs, context);
  let answer = await args.ask(args.request, schema, room);
  let draft = build(answer);
  for (let attempt = 1; attempt <= args.maxRepairs; attempt++) {
    const problems = repairable(draft.issues);
    if (problems.length === 0) break;
    args.onRepair?.(attempt, problems.length);
    const repair = [
      args.request,
      `Your previous answer:\n${JSON.stringify(answer)}`,
      `It has these problems:\n${problems.map((issue) => `- s${issue.slide}${issue.slot ? `.${issue.slot}` : ''}: ${issue.message}`).join('\n')}`,
      'Return the full content again with these problems fixed. Change nothing else.',
    ].join('\n\n');
    answer = await args.ask(repair, schema, room);
    draft = build(answer);
  }
  return draft;
}
