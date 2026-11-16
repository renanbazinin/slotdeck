import { unstatedChartValues } from '../deck/numbers.ts';
import { DeckValidationError } from '../deck/validate.ts';
import type { Issue, LayoutMode } from '../deck/validate.ts';
import type { DeckSpec } from '../deck/spec.ts';
import type { JsonRequest, JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { TEXT_SIZES } from '../template/manifest.ts';
import { measurerFor } from '../text/measure.ts';
import type { Direction, TextSize } from '../template/manifest.ts';
import { budgetScale, buildCatalog, describeCatalog, isPlainText, itemsProperty, outlineSchema, planPattern, readItems } from './catalog.ts';
import type { AssetDef, Catalog, CatalogContext, IconDef, ImageDensity, UsablePattern, WriteItem } from './catalog.ts';
import { withLayouts } from '../template/compose.ts';
import type { PatternDesign } from '../template/compose.ts';
import { composeForPlan, usedLayouts } from './layouts.ts';
import type { ComposeLayoutsResult } from './layouts.ts';
import type { DeckReview } from './review.ts';
import { cleanText, suppliesText, systemPrompt, writeSlides } from './write.ts';

/** What the person asked for. */
export interface Brief {
  /** What the deck is about, in the person's own words. */
  prompt: string;
  /** A fixed number of slides, or a range. Defaults to 6 to 10. */
  slideCount?: number | { min: number; max: number };
  /** BCP 47 tag of the language to write in. Defaults to the template's language. */
  language?: string;
  /** Reading direction. Defaults to the language's. */
  direction?: Direction;
  /** How many slides should carry a photo. Defaults to `medium`. */
  imageDensity?: ImageDensity;
  /** Anything else to respect: audience, tone, points that must be covered. */
  instructions?: string;
  /** Deck-level values the template needs, e.g. the unit shown in the header. */
  fields?: Record<string, string>;
  /** Size of body text. Larger text means less of it fits on a slide. Defaults to `normal`. */
  textSize?: TextSize;
  /**
   * Plan every slide as plain text: a title with paragraphs or a list, and no pictures, icons,
   * cards, figures, tables or charts. Only the template's plain text patterns are offered.
   */
  textOnly?: boolean;
}

export type PlanProgress =
  | { stage: 'outline' }
  | { stage: 'layouts' }
  | { stage: 'write'; slides: number }
  | { stage: 'repair'; attempt: number; issues: number };

export interface PlanOptions {
  llm: LlmProvider;
  /** Defaults to `strict`. */
  layout?: LayoutMode;
  /** The app's icon set. Patterns with icon slots are only offered when this is given. */
  icons?: IconDef[];
  /** Ready-made pictures the deck may use, e.g. rendered diagrams. */
  assets?: AssetDef[];
  /** Restrict the deck to these pattern ids. */
  patterns?: string[];
  /**
   * When patterns built around numbers may be used. `from-brief` (default) offers them only if
   * the brief contains figures; a model given such a pattern and no data tends to invent some.
   */
  numberPatterns?: 'from-brief' | 'always' | 'never';
  /** How many times the model may be asked to fix a draft that breaks the template's limits. Defaults to 2. */
  maxRepairs?: number;
  /**
   * A second look at every draft: `designReview` for the code-only design checks, `visionReview`
   * for a model that looks at pictures of the slides, or a function of your own. What it finds is
   * sent back to the writing model in the repair rounds, and what remains is returned with the
   * warnings. Nothing reviews a draft unless this is set.
   */
  review?: DeckReview;
  /**
   * Up to this many new layouts may be made for the deck, in the template's own look, for slides
   * whose content none of its layouts shows well: three goals on a slide of running text, a plan
   * in steps. After the outline, one light call decides; each layout then takes one or two calls,
   * side by side, about 10 to 20 seconds. The designs are kept in the plan and the deck
   * (`layouts`); the template is not changed. A layout that cannot be made leaves its slides as
   * planned. Off (0) by default.
   */
  composeLayouts?: number;
  onProgress?: (event: PlanProgress) => void;
  signal?: AbortSignal;
}

export interface OutlineEntry {
  pattern: string;
  purpose: string;
  /** For a pattern with a row of items: how many items the slide has. */
  items?: number;
}

export interface PlanResult {
  deck: DeckSpec;
  /** Limits the final draft still exceeds, e.g. text a little over a slot's budget. */
  warnings: Issue[];
  outline: OutlineEntry[];
}

export class PlanError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PlanError';
  }
}

const RTL_LANGUAGES = new Set(['ar', 'dv', 'fa', 'he', 'ps', 'sd', 'ug', 'ur', 'yi']);

export function directionOf(language: string): Direction {
  return RTL_LANGUAGES.has(language.toLowerCase().split('-')[0] ?? '') ? 'rtl' : 'ltr';
}

/** Whether a text contains any digit, in any script. */
export function hasFigures(text: string): boolean {
  return /\p{Nd}/u.test(text);
}

/**
 * Binds a provider to a system prompt and an optional abort signal. `room` is how many tokens
 * the answer may reasonably take: a model sometimes runs away inside a text field, and a limit
 * is what ends that after seconds instead of minutes. `effort` is `low` for calls that choose
 * and arrange rather than write.
 */
export function asker(llm: LlmProvider, system: string, signal: AbortSignal | undefined, effort?: 'low') {
  return (prompt: string, schema: JsonSchema, room?: number): Promise<unknown> => {
    const request: JsonRequest = { system, prompt, schema };
    if (room !== undefined) request.maxOutputTokens = Math.ceil(room);
    if (effort) request.effort = effort;
    if (signal) request.signal = signal;
    return llm.generateJson(request);
  };
}

/** The caller's review, bound to the caller's abort signal. */
export function reviewWith(options: Pick<PlanOptions, 'review' | 'signal'>): ((deck: DeckSpec) => Issue[] | Promise<Issue[]>) | undefined {
  const { review, signal } = options;
  return review ? (deck) => review(deck, signal ? { signal } : {}) : undefined;
}

/** Room for a plan: a title, and for each slide a pattern id and a sentence. Generous, since it is a ceiling. */
export const planRoom = (slides: number): number => 400 + 160 * slides;

const DENSITY_GUIDANCE: Record<ImageDensity, string> = {
  none: 'Do not plan any photos.',
  low: 'Use a pattern with a photo on at most a quarter of the slides.',
  medium: 'Use a pattern with a photo on about half of the slides.',
  high: 'Prefer patterns with a photo wherever they suit the content.',
};

function briefText(brief: Brief): string {
  return [`Brief:\n${brief.prompt.trim()}`, brief.instructions?.trim() ? `Also respect:\n${brief.instructions.trim()}` : '']
    .filter(Boolean)
    .join('\n\n');
}

/**
 * A wish for columns is a wish about how items sit. The catalogue says it for each pattern:
 * "side by side in one row", or "in lines of" a number for a grid.
 */
export const columnsRule = (asked: 'brief' | 'request'): string =>
  `When the ${asked} asks for a number of columns, or for items side by side, use a pattern that takes that number of items and whose items are "side by side in one row" or "in lines of" that number or more: either sets them on one line. Items "in lines of" a smaller number are broken over several lines. When the list has no such pattern, use the one that comes closest.`;

/** What a wish about the kind of slides means, in the planner's own terms. Shared by every prompt that chooses patterns. */
export const SLIDE_KIND_RULE =
  'A wish about the kind of slides (text only, no tables, only cards) decides which patterns may be used, and comes before variety. ' +
  '"Text only" means the patterns marked "Plain text.": a title with paragraphs or a list, and no cards, figures, tables, charts, icons or pictures. ' +
  'Using one pattern for many slides is right when the wish leaves no other.';

function outlinePrompt(brief: Brief, catalog: Catalog, count: { min: number; max: number }, density: ImageDensity, textScale = 1): string {
  const size = count.min === count.max ? `exactly ${count.min}` : `${count.min} to ${count.max}`;
  return [
    briefText(brief),
    `Plan a deck of ${size} slides. Choose one pattern for each slide from this list ("?" marks an optional slot):`,
    describeCatalog(catalog, textScale),
    [
      '- Open with an opening pattern and end with a closing pattern when the list has them.',
      '- For a pattern with a row of items, set "items" to the number the content really has. Do not pad a list to fill a row.',
      `- ${columnsRule('brief')}`,
      '- Choose a pattern only when the content can fill every required slot within its limits.',
      ...(brief.textOnly
        ? ['- The deck is text only, and every pattern in the list is plain text. Using the same pattern for several slides in a row is expected.']
        : [`- ${SLIDE_KIND_RULE}`, '- Otherwise vary the patterns across the deck.', `- ${DENSITY_GUIDANCE[density]}`]),
    ].join('\n'),
  ].join('\n\n');
}

function writePrompt(brief: Brief, items: WriteItem[], context: CatalogContext): string {
  const slides = items.map(({ number, pattern, purpose }) => `s${number}: ${pattern.def.id}. ${purpose}`).join('\n');
  return [
    briefText(brief),
    `Write the content of each slide of this deck:\n${slides}`,
    ...suppliesText(items, context),
    'Fill every required property. Leave an optional property out when the slide is better without it.',
  ].join('\n\n');
}

/** A deck before its content is written: the slides in order, each with a pattern and the point it makes. */
export interface DeckPlan {
  version: 1;
  title: string;
  /** The request, with everything needed to write the deck later. */
  brief: Brief;
  slides: OutlineEntry[];
  /** Layouts made for this deck (see `composeLayouts`). Slides use them by their ids. */
  layouts?: PatternDesign[];
}

interface Planning {
  layout: LayoutMode;
  language: string;
  direction: Direction;
  density: ImageDensity;
  fields: Record<string, string>;
  context: CatalogContext;
  catalog: Catalog;
  byId: Map<string, UsablePattern>;
}

/** Everything the planning stages derive from a brief and the caller's options, so they cannot disagree. */
function planning(template: Template, brief: Brief, options: PlanOptions, layouts?: PatternDesign[]): Planning {
  const { manifest } = template;
  const layout = options.layout ?? 'strict';
  const language = brief.language ?? manifest.language ?? 'en';
  const direction = brief.direction ?? directionOf(language);
  const density = brief.imageDensity ?? 'medium';
  const fields = brief.fields ?? {};

  const numbers = options.numberPatterns ?? 'from-brief';
  const figures = numbers === 'always' || (numbers === 'from-brief' && hasFigures(`${brief.prompt} ${brief.instructions ?? ''}`));
  const textScale = TEXT_SIZES[brief.textSize ?? 'normal'] ?? 1;
  const context: CatalogContext = { layout, fields, icons: options.icons ?? [], assets: options.assets ?? [], imageDensity: density, figures, textScale, templateScale: manifest.textScale ?? 1 };
  if (options.patterns) context.allow = options.patterns;
  if (layouts?.length) {
    context.own = layouts.map((design) => design.id);
    // A deck's own layouts were made for it: a list of allowed patterns does not leave them out.
    if (context.allow) context.allow = [...context.allow, ...context.own];
  }
  const catalog = buildCatalog(manifest, context);
  if (catalog.patterns.length === 0) {
    const reasons = catalog.excluded.map((entry) => `${entry.id} ${entry.reason}`).join('; ');
    throw new PlanError(`No pattern of template "${manifest.name}" can be used${reasons ? `: ${reasons}` : ''}`);
  }
  return { layout, language, direction, density, fields, context, catalog, byId: new Map(catalog.patterns.map((pattern) => [pattern.def.id, pattern])) };
}

function readOutline(answer: unknown, byId: Map<string, UsablePattern>): { title: string; entries: OutlineEntry[]; problem?: string } {
  const raw = answer as { title?: unknown; slides?: unknown } | null;
  const list = Array.isArray(raw?.slides) ? raw.slides : [];
  const problems: string[] = [];
  const entries = list.map((item, index) => {
    const raw = item as { pattern?: unknown; purpose?: unknown; items?: unknown };
    const entry: OutlineEntry = { pattern: cleanText(raw?.pattern), purpose: cleanText(raw?.purpose) };
    const pattern = byId.get(entry.pattern);
    if (pattern) {
      const count = readItems(pattern.def, raw?.items);
      if (count.problem) problems.push(`slide ${index + 1}: ${count.problem}`);
      if (count.items !== undefined) entry.items = count.items;
    }
    return entry;
  });
  const unknown = entries.filter((entry) => !byId.has(entry.pattern)).map((entry) => entry.pattern || '(empty)');
  const outline = { title: cleanText(raw?.title), entries };
  if (entries.length === 0) return { ...outline, problem: 'the answer contained no slides' };
  if (unknown.length > 0) return { ...outline, problem: `these patterns are not in the list: ${[...new Set(unknown)].join(', ')}` };
  if (problems.length > 0) return { ...outline, problem: problems.join('; ') };
  return outline;
}

/**
 * First stage: decides which pattern each slide uses and the point it makes. One model call.
 * The result is small and quick to produce, so it is the thing to show a person and let them change.
 */
export async function planOutline(template: Template, brief: Brief, options: PlanOptions): Promise<DeckPlan> {
  const planned = planning(template, brief, options);
  const { language, density, context } = planned;
  const catalog = outlineCatalog(planned, brief);
  if (catalog.patterns.length === 0) {
    throw new PlanError(`Template "${template.manifest.name}" has no plain text pattern, so a text-only deck cannot be planned from it`);
  }
  const byId = new Map(catalog.patterns.map((pattern) => [pattern.def.id, pattern]));
  const count =
    typeof brief.slideCount === 'number'
      ? { min: brief.slideCount, max: brief.slideCount }
      : (brief.slideCount ?? { min: 6, max: 10 });
  // Choosing a pattern for each slide needs little reasoning, and is much quicker with little.
  const ask = asker(options.llm, systemPrompt(language, 'brief'), options.signal, 'low');

  options.onProgress?.({ stage: 'outline' });
  const request = outlinePrompt(brief, catalog, count, density, budgetScale(context));
  let outline = readOutline(await ask(request, outlineSchema(catalog, count), planRoom(count.max)), byId);
  if (outline.problem) {
    const retry = `${request}\n\nYour previous answer could not be used: ${outline.problem}. Answer again using only the listed patterns.`;
    outline = readOutline(await ask(retry, outlineSchema(catalog, count), planRoom(count.max)), byId);
    if (outline.problem) throw new PlanError(`The model did not produce a usable outline: ${outline.problem}`);
  }
  const plan: DeckPlan = { version: 1, title: outline.title, brief, slides: outline.entries };
  if (!options.composeLayouts) return plan;
  options.onProgress?.({ stage: 'layouts' });
  return (await composeForPlan(template, plan, catalog.patterns, { llm: options.llm, max: options.composeLayouts, ...(options.signal ? { signal: options.signal } : {}) })).plan;
}

/** What an outline may choose from. A text-only deck is offered the plain text patterns and nothing else, so the wish cannot be missed. */
function outlineCatalog(planned: Planning, brief: Brief): Catalog {
  return brief.textOnly ? { ...planned.catalog, patterns: planned.catalog.patterns.filter((pattern) => isPlainText(pattern.def) || pattern.own) } : planned.catalog;
}

/**
 * Looks at a plan for slides that a layout the template lacks would show better, makes those
 * layouts in the template's own look, and gives them to their slides. What `composeLayouts` in
 * {@link PlanOptions} does inside {@link planOutline}, as a step of its own: for a plan that was
 * already shown, or to learn what was made and what could not be. `max` defaults to 2.
 */
export async function composeLayouts(template: Template, plan: DeckPlan, options: PlanOptions & { max?: number }): Promise<ComposeLayoutsResult> {
  const planned = planning(withLayouts(template, plan.layouts), plan.brief, options, plan.layouts);
  options.onProgress?.({ stage: 'layouts' });
  return composeForPlan(template, plan, outlineCatalog(planned, plan.brief).patterns, { llm: options.llm, max: options.max ?? 2, ...(options.signal ? { signal: options.signal } : {}) });
}

/**
 * Changes a plan by instruction ("add a slide about pricing", "make it shorter", "in English").
 * One model call. The instruction is kept with the brief, so facts it adds reach the writing stage.
 */
export async function editPlan(
  template: Template,
  plan: DeckPlan,
  instruction: string,
  options: PlanOptions,
): Promise<{ plan: DeckPlan; summary: string }> {
  const request = instruction.trim();
  const brief: Brief = { ...plan.brief, instructions: [plan.brief.instructions?.trim(), request].filter(Boolean).join('\n') };
  const { language, density, catalog, byId, context } = planning(withLayouts(template, plan.layouts), brief, options, plan.layouts);

  options.onProgress?.({ stage: 'outline' });
  const prompt = [
    briefText(plan.brief),
    `The plan as it is now${plan.title ? ` ("${plan.title}")` : ''}, in the language with BCP 47 tag "${language}":\n${plan.slides
      .map((slide, index) => `${index + 1}. ${slide.pattern}${slide.items !== undefined ? ` (${slide.items} items)` : ''}: ${slide.purpose}`)
      .join('\n')}`,
    `Requested change:\n${request}`,
    `Return the whole plan after the change. Choose patterns from this list ("?" marks an optional slot):`,
    describeCatalog(catalog, budgetScale(context)),
    [
      '- Change only what the request asks for. Keep every other slide, its pattern, its purpose and its place.',
      '- Choose a pattern only when the content can fill every required slot within its limits.',
      `- ${columnsRule('request')}`,
      `- ${SLIDE_KIND_RULE}`,
      ...(plan.brief.textOnly ? ['- This deck was asked to be text only: keep to the patterns marked "Plain text." or "Made for this deck." unless the requested change asks for something else.'] : []),
      `- ${DENSITY_GUIDANCE[density]}`,
    ].join('\n'),
  ].join('\n\n');
  const schema: JsonSchema = {
    type: 'object',
    properties: {
      summary: { type: 'string', description: 'One sentence saying what you changed, written for the person who asked, in their language' },
      language: { type: 'string', description: 'BCP 47 tag of the deck language after the change' },
      title: { type: 'string', description: 'Title of the deck after the change' },
      slides: {
        type: 'array',
        minItems: 1,
        items: {
          type: 'object',
          properties: {
            pattern: { type: 'string', enum: catalog.patterns.map((pattern) => pattern.def.id) },
            purpose: { type: 'string', description: 'The one point this slide makes: a single sentence of at most 25 words' },
            ...itemsProperty(catalog),
          },
          required: ['pattern', 'purpose'],
        },
      },
    },
    required: ['summary', 'language', 'title', 'slides'],
  };

  const ask = asker(options.llm, systemPrompt(language, 'brief'), options.signal, 'low');
  // An edit may add slides, so the room is for a plan somewhat longer than the one it starts from.
  const room = planRoom(plan.slides.length + 10);
  let answer = (await ask(prompt, schema, room)) as { summary?: unknown; language?: unknown } | null;
  let outline = readOutline(answer, byId);
  if (outline.problem) {
    answer = (await ask(`${prompt}\n\nYour previous answer could not be used: ${outline.problem}. Answer again using only the listed patterns.`, schema, room)) as typeof answer;
    outline = readOutline(answer, byId);
    if (outline.problem) throw new PlanError(`The model did not produce a usable plan: ${outline.problem}`);
  }

  const nextLanguage = cleanText(answer?.language) || language;
  if (nextLanguage !== language) {
    brief.language = nextLanguage;
    delete brief.direction; // follows the new language
  }
  delete brief.slideCount; // the plan itself now says how many slides there are
  const next: DeckPlan = { version: 1, title: outline.title || plan.title, brief, slides: outline.entries };
  const layouts = usedLayouts(plan.layouts, outline.entries);
  if (layouts) next.layouts = layouts;
  return { plan: next, summary: cleanText(answer?.summary) };
}

/**
 * Second stage: writes the content of every slot of a plan's slides, against a schema built for
 * exactly that sequence. Drafts that break the template's limits go back to the model.
 */
export async function writeDeck(template: Template, plan: DeckPlan, options: PlanOptions): Promise<PlanResult> {
  const { brief } = plan;
  const layouts = usedLayouts(plan.layouts, plan.slides);
  // The deck's own layouts are built onto the template: written, checked and repaired as any other.
  template = withLayouts(template, layouts);
  const { layout, language, direction, fields, context, byId } = planning(template, brief, options, layouts);
  const missing = plan.slides.filter((slide) => !byId.has(slide.pattern)).map((slide) => slide.pattern);
  if (missing.length > 0) {
    throw new PlanError(`The plan uses patterns that cannot be filled with what was supplied: ${[...new Set(missing)].join(', ')}`);
  }
  const items: WriteItem[] = plan.slides.map((entry, index) => {
    const usable = byId.get(entry.pattern) as UsablePattern;
    // A row of items gets the slots of exactly the number the plan settled on.
    const sized = entry.items !== undefined ? planPattern(usable.def, context, undefined, entry.items) : undefined;
    return { number: index + 1, pattern: sized && 'pattern' in sized ? sized.pattern : usable, purpose: entry.purpose };
  });

  options.onProgress?.({ stage: 'write', slides: items.length });
  const draft = await writeSlides({
    ask: asker(options.llm, systemPrompt(language, 'brief'), options.signal),
    request: writePrompt(brief, items, context),
    jobs: items,
    context,
    assemble: (written) => {
      const deck: DeckSpec = {
        version: 1,
        language,
        direction,
        brief: [brief.prompt.trim(), brief.instructions?.trim()].filter(Boolean).join('\n'),
        slides: items.map((item) => written.get(item.number) ?? { pattern: item.pattern.def.id, slots: {} }),
      };
      if (context.textScale !== 1) deck.textScale = context.textScale;
      if (plan.title) deck.title = plan.title;
      if (layouts) deck.layouts = layouts;
      if (Object.keys(fields).length > 0) deck.fields = fields;
      return deck;
    },
    manifest: template.manifest,
    measure: measurerFor(template),
    review: reviewWith(options),
    layout,
    maxRepairs: options.maxRepairs ?? 2,
    onRepair: (attempt, issues) => options.onProgress?.({ stage: 'repair', attempt, issues }),
  });

  const errors = draft.issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  return { deck: draft.deck, warnings: [...draft.issues, ...unstatedChartValues(template.manifest, draft.deck)], outline: plan.slides };
}

/** Brief to deck spec in one go: {@link planOutline}, then {@link writeDeck}. */
export async function planDeck(template: Template, brief: Brief, options: PlanOptions): Promise<PlanResult> {
  return writeDeck(template, await planOutline(template, brief, options), options);
}
