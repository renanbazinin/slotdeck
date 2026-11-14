import { DeckValidationError, validateDeck } from '../deck/validate.ts';
import type { Issue } from '../deck/validate.ts';
import { unstatedChartValues } from '../deck/numbers.ts';
import { isChartValue, isImageRef } from '../deck/spec.ts';
import type { DeckSpec, SlideSpec, SlotValue } from '../deck/spec.ts';
import type { JsonSchema } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { TEXT_SIZES, effectiveSlots, itemCount, itemRange } from '../template/manifest.ts';
import type { Direction, PatternDef, TextSize } from '../template/manifest.ts';
import { measurerFor } from '../text/measure.ts';
import { budgetScale, buildCatalog, describeCatalog, planPattern, readItems } from './catalog.ts';
import type { CatalogContext, ImageDensity } from './catalog.ts';
import { PlanError, SLIDE_KIND_RULE, asker, directionOf, hasFigures, planRoom, reviewWith } from './plan.ts';
import type { PlanOptions } from './plan.ts';
import { cleanText, suppliesText, systemPrompt, writeSlides } from './write.ts';
import type { WriteJob } from './write.ts';

export interface EditOptions extends PlanOptions {
  /** Replaces the deck-level values, e.g. a translated header text. Defaults to the deck's own. */
  fields?: Record<string, string>;
  /** Reading direction after the edit. Defaults to the direction of the deck's language. */
  direction?: Direction;
  /** Whether slides added by the edit may use photo patterns. Defaults to `medium`. */
  imageDensity?: ImageDensity;
}

export interface EditResult {
  deck: DeckSpec;
  /** One sentence from the model on what it changed, for the person who asked. */
  summary: string;
  /** Slide numbers in the new deck that were written or rewritten. Every other slide is copied as it was. */
  rewritten: number[];
  warnings: Issue[];
}

interface PlanEntry {
  /** Number of the existing slide this one is based on; 0 for a new slide. */
  from: number;
  pattern: string;
  /** What to write or change. Empty when the slide stays exactly as it is. */
  change: string;
  /** Whether the pictures and icons of the slide should change. */
  images: boolean;
  /** For a pattern with a row of items: how many items the slide has after the change. */
  items?: number;
}

function describeValue(value: SlotValue | undefined): string {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return JSON.stringify(value);
  if (Array.isArray(value)) return JSON.stringify(value);
  if (!isImageRef(value)) return `${isChartValue(value) ? 'chart' : 'table'} ${JSON.stringify(value)}`;
  return value.prompt ? `picture of ${JSON.stringify(value.prompt)}` : `picture "${value.asset}"`;
}

/** The deck as the model sees it: one line per slide with every slot's content. */
function describeDeck(deck: DeckSpec): string {
  return deck.slides
    .map((slide, index) => {
      const slots = Object.entries(slide.slots)
        .map(([name, value]) => `${name}: ${describeValue(value)}`)
        .filter((line) => !line.endsWith(': '))
        .join('; ');
      return `${index + 1}. [${slide.pattern}${slide.items !== undefined ? `, ${slide.items} items` : ''}] ${slots}`;
    })
    .join('\n');
}

/** A short label for a slide that is not being rewritten, so the model knows what surrounds its work. */
function slideLabel(slide: SlideSpec): string {
  const title = slide.slots.title;
  const text = typeof title === 'string' ? title : Object.values(slide.slots).find((value): value is string => typeof value === 'string');
  return text ? JSON.stringify(text) : '';
}

function editPlanSchema(patternIds: string[], slideCount: number, fieldNames: string[], mostItems: number): JsonSchema {
  const schema = planSchemaBase(patternIds, slideCount);
  const top = schema.properties as Record<string, JsonSchema>;
  top.text_size = {
    type: 'string',
    enum: Object.keys(TEXT_SIZES),
    description: 'Size of body text after the change. Repeat the current size unless the request is about text size.',
  };
  (schema.required as string[]).push('text_size');
  if (mostItems > 0) {
    const entry = ((top.slides as JsonSchema).items as JsonSchema).properties as Record<string, JsonSchema>;
    entry.items = {
      type: 'integer',
      minimum: 1,
      maximum: mostItems,
      description: 'For a pattern with a row of items: how many items the slide has after the change. Leave it out for other patterns.',
    };
  }
  if (fieldNames.length > 0) {
    const properties = schema.properties as Record<string, JsonSchema>;
    properties.fields = {
      type: 'object',
      description: 'The deck-level texts after the change. Repeat a value exactly unless the request affects it.',
      properties: Object.fromEntries(fieldNames.map((name) => [name, { type: 'string' }])),
      required: fieldNames,
    };
    (schema.required as string[]).push('fields');
  }
  return schema;
}

function planSchemaBase(patternIds: string[], slideCount: number): JsonSchema {
  return {
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
            from: { type: 'integer', minimum: 0, maximum: slideCount, description: 'Number of the existing slide this slide is based on; 0 for a new slide' },
            pattern: { type: 'string', enum: patternIds },
            change: { type: 'string', description: 'What to write or change on this slide. Empty when the slide stays exactly as it is.' },
            images: { type: 'boolean', description: 'true only when the pictures or icons of this slide should change' },
          },
          required: ['from', 'pattern', 'change', 'images'],
        },
      },
    },
    required: ['summary', 'language', 'title', 'slides'],
  };
}

/**
 * Applies an instruction such as "shorten slide 3" or "add a slide about pricing" to a deck.
 *
 * Two model calls: the first decides what the deck looks like after the change (which slides
 * stay, move, go, get rewritten or are new); the second writes only the slides that change.
 * Untouched slides are copied by code, so they cannot drift, and pictures are kept unless the
 * instruction is about them.
 */
export async function planEdit(template: Template, deck: DeckSpec, instruction: string, options: EditOptions): Promise<EditResult> {
  const { manifest } = template;
  const measure = measurerFor(template);
  const layout = options.layout ?? 'strict';
  const fields = options.fields ?? deck.fields ?? {};
  const currentLanguage = deck.language ?? manifest.language ?? 'en';
  const definitions = new Map<string, PatternDef>(manifest.patterns.map((pattern) => [pattern.id, pattern]));

  // What new slides may use. Number patterns need figures from the instruction or the original brief.
  const numbers = options.numberPatterns ?? 'from-brief';
  const figures = numbers === 'always' || (numbers === 'from-brief' && hasFigures(`${instruction} ${deck.brief ?? ''}`));
  const context: CatalogContext = {
    layout,
    fields,
    icons: options.icons ?? [],
    assets: options.assets ?? [],
    imageDensity: options.imageDensity ?? 'medium',
    figures,
    textScale: deck.textScale ?? 1,
    templateScale: manifest.textScale ?? 1,
  };
  const currentSize = (Object.entries(TEXT_SIZES).find(([, factor]) => factor === context.textScale)?.[0] ?? 'normal') as TextSize;
  if (options.patterns) context.allow = options.patterns;
  const offer = buildCatalog(manifest, context);
  const offered = new Map(offer.patterns.map((pattern) => [pattern.def.id, pattern]));
  // An existing slide keeps its pattern even if that pattern would not be offered for a new slide.
  const existing = [...new Set(deck.slides.map((slide) => slide.pattern))].filter((id) => definitions.has(id) && !offered.has(id));
  const patternIds = [...offered.keys(), ...existing];
  if (patternIds.length === 0) throw new PlanError(`No pattern of template "${manifest.name}" can be used`);

  // 1. Plan the edit.
  // Deck-level texts can change too ("translate the deck"), unless the caller fixes them.
  const fieldNames = options.fields ? [] : Object.keys(fields);
  options.onProgress?.({ stage: 'outline' });
  const planRequest = [
    deck.brief ? `The deck was written from this brief:\n${deck.brief}` : '',
    `The deck as it is now${deck.title ? ` ("${deck.title}")` : ''}, in the language with BCP 47 tag "${currentLanguage}", body text size "${currentSize}":\n${describeDeck(deck)}`,
    fieldNames.length > 0
      ? `Deck-level texts, shown on several slides:\n${fieldNames.map((name) => `${name}: ${JSON.stringify(fields[name])}`).join('\n')}`
      : '',
    `Requested change:\n${instruction.trim()}`,
    [
      'Return the slides of the deck after the change, in their new order.',
      '- from: the number of the existing slide a slide is based on, or 0 for a new slide. Leave a slide out to delete it.',
      '- pattern: keep the pattern of an existing slide unless the request calls for a different layout.',
      `  ${SLIDE_KIND_RULE}`,
      '- change: what to write or change on that slide. Leave it empty for every slide the request does not touch.',
      '- images: true only when the pictures or icons of that slide should change.',
      '- items: for a pattern with a row of items, the number the slide has after the change.',
    ].join('\n'),
    offer.patterns.length > 0 ? `Patterns for new slides ("?" marks an optional slot):\n${describeCatalog(offer, budgetScale(context))}` : '',
  ]
    .filter(Boolean)
    .join('\n\n');

  const readPlan = (
    answer: unknown,
  ): { summary: string; language: string; title: string; textSize: TextSize; fields: Record<string, string>; entries: PlanEntry[]; problem?: string } => {
    const raw = (answer ?? {}) as { summary?: unknown; language?: unknown; title?: unknown; slides?: unknown; fields?: unknown; text_size?: unknown };
    const rawFields = (typeof raw.fields === 'object' && raw.fields !== null ? raw.fields : {}) as Record<string, unknown>;
    // A field the model leaves empty keeps its value: a header must not silently disappear.
    const nextFields = Object.fromEntries(fieldNames.map((name) => [name, cleanText(rawFields[name]) || (fields[name] as string)]));
    const countProblems: string[] = [];
    const entries: PlanEntry[] = (Array.isArray(raw.slides) ? raw.slides : []).map((item, index) => {
      const given = (item ?? {}) as { from?: unknown; pattern?: unknown; change?: unknown; images?: unknown; items?: unknown };
      const entry: PlanEntry = { from: Number(given.from) || 0, pattern: cleanText(given.pattern), change: cleanText(given.change), images: given.images === true };
      const definition = definitions.get(entry.pattern);
      if (definition?.repeat) {
        // A slide that keeps its layout keeps its number of items unless the plan says otherwise.
        const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
        const count = readItems(definition, given.items ?? (source?.pattern === entry.pattern ? source.items : undefined));
        if (count.problem) countProblems.push(`slide ${index + 1}: ${count.problem}`);
        if (count.items !== undefined) entry.items = count.items;
      }
      return entry;
    });
    const plan = {
      summary: cleanText(raw.summary),
      language: cleanText(raw.language) || currentLanguage,
      title: cleanText(raw.title),
      textSize: (typeof raw.text_size === 'string' && raw.text_size in TEXT_SIZES ? raw.text_size : currentSize) as TextSize,
      fields: { ...fields, ...nextFields },
      entries,
    };
    if (entries.length === 0) return { ...plan, problem: 'the answer contained no slides' };
    if (countProblems.length > 0) return { ...plan, problem: countProblems.join('; ') };
    for (const [index, entry] of entries.entries()) {
      const where = `slide ${index + 1}`;
      if (!Number.isInteger(entry.from) || entry.from < 0 || entry.from > deck.slides.length) {
        return { ...plan, problem: `${where}: "from" must be 0 or the number of an existing slide (1 to ${deck.slides.length})` };
      }
      if (!definitions.has(entry.pattern)) return { ...plan, problem: `${where}: unknown pattern "${entry.pattern}"` };
      const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
      if (source?.pattern !== entry.pattern && !offered.has(entry.pattern)) {
        return { ...plan, problem: `${where}: pattern "${entry.pattern}" cannot be used for a new or re-laid-out slide here` };
      }
      if (!source && !entry.change) return { ...plan, problem: `${where}: a new slide needs "change" to say what it shows` };
    }
    return plan;
  };

  const mostItems = Math.max(0, ...patternIds.map((id) => itemRange(definitions.get(id) as PatternDef)).map((range) => (range && range.max > range.min ? range.max : 0)));
  const planSchema = editPlanSchema(patternIds, deck.slides.length, fieldNames, mostItems);
  // The edit plan says a little more for each slide than an outline does, and may add slides.
  const room = planRoom(deck.slides.length + 10) * 1.5;
  let plan = readPlan(await asker(options.llm, systemPrompt(currentLanguage, 'deck'), options.signal, 'low')(planRequest, planSchema, room));
  if (plan.problem) {
    const retry = `${planRequest}\n\nYour previous answer could not be used: ${plan.problem}. Answer again.`;
    plan = readPlan(await asker(options.llm, systemPrompt(currentLanguage, 'deck'), options.signal, 'low')(retry, planSchema, room));
    if (plan.problem) throw new PlanError(`The model did not produce a usable edit plan: ${plan.problem}`);
  }

  const language = plan.language;
  const direction = options.direction ?? (language === currentLanguage ? (deck.direction ?? directionOf(language)) : directionOf(language));
  // A change of language touches every slide, whatever the plan says.
  if (language !== currentLanguage) {
    for (const entry of plan.entries) if (!entry.change) entry.change = 'Translate into the new deck language.';
  }

  // Larger text fits less: from here on, limits are those of the size the deck will have.
  const textScale = TEXT_SIZES[plan.textSize];
  const resized = textScale !== context.textScale;
  context.textScale = textScale;
  const overflows = (slide: SlideSpec) =>
    validateDeck(manifest, { ...deck, textScale, slides: [slide] }, measure ? { layout, measure } : { layout }).some((issue) => /characters|paragraphs|lines/.test(issue.message));

  // 2. Work out, per slide, whether it is copied or written, and what a written slide carries over.
  const jobs: Array<WriteJob & { source?: SlideSpec; relaid: boolean }> = [];
  const kept = new Map<number, SlideSpec>();
  plan.entries.forEach((entry, index) => {
    const number = index + 1;
    const source = entry.from > 0 ? deck.slides[entry.from - 1] : undefined;
    const definition = definitions.get(entry.pattern) as PatternDef;

    if (source && source.pattern === entry.pattern && !entry.change) {
      // Two things rewrite a slide the plan left alone: a new number of items, and text that no longer fits.
      if (entry.items !== undefined && entry.items !== itemCount(definition, source.items)) {
        entry.change = `Use ${entry.items} ${definition.repeat?.item ?? 'item'}s.`;
      } else if (resized && overflows(source)) {
        entry.change = 'Shorten the text to fit the new text size. Keep the meaning.';
      }
    }

    if (source && source.pattern === entry.pattern && !entry.change) {
      kept.set(number, structuredClone(source));
      return;
    }

    const carried: Record<string, SlotValue> = {};
    if (source && source.pattern === entry.pattern) {
      // Same layout: everything the model is not asked about stays, pictures included unless they are the point.
      for (const slot of effectiveSlots(definition, entry.items)) {
        const value = source.slots[slot.name];
        if (value === undefined) continue;
        const policy = slot.empty ?? 'error';
        if ((slot.kind === 'image' && !entry.images) || policy === 'keep' || policy === 'drop') carried[slot.name] = value;
      }
    } else if (source && !entry.images) {
      // New layout for an existing slide: its photos move over, in order, so they are not generated again.
      const oldDefinition = definitions.get(source.pattern);
      const photos = (oldDefinition?.slots ?? [])
        .filter((slot) => slot.kind === 'image' && slot.imageKind === 'photo')
        .map((slot) => source.slots[slot.name])
        .filter(isImageRef);
      for (const slot of effectiveSlots(definition, entry.items)) {
        if (slot.kind !== 'image' || slot.imageKind !== 'photo') continue;
        const photo = photos.shift();
        if (photo) carried[slot.name] = photo;
      }
    }

    // A slide that stays in its layout can always be rewritten, even if that layout is not on offer for new slides.
    const own: CatalogContext = source && source.pattern === entry.pattern ? { ...context, figures: true, imageDensity: 'high' } : context;
    let planned = planPattern(definition, own, new Set(Object.keys(carried)), entry.items);
    if ('reason' in planned && source && source.pattern === entry.pattern) {
      // Its pictures cannot be replaced with what the caller supplied (say, no icon set): keep them all.
      for (const slot of effectiveSlots(definition, entry.items)) {
        const value = source.slots[slot.name];
        if (slot.kind === 'image' && value !== undefined) carried[slot.name] = value;
      }
      planned = planPattern(definition, own, new Set(Object.keys(carried)), entry.items);
    }
    if ('reason' in planned) throw new PlanError(`Slide ${number}: pattern "${entry.pattern}" ${planned.reason}`);

    const job: WriteJob & { source?: SlideSpec; relaid: boolean } = {
      number,
      pattern: planned.pattern,
      purpose: entry.change,
      carried,
      relaid: !!source && source.pattern !== entry.pattern,
    };
    if (source) job.source = source;
    jobs.push(job);
  });

  const assemble = (written: Map<number, SlideSpec>): DeckSpec => {
    const slides = plan.entries.map((entry, index) => kept.get(index + 1) ?? written.get(index + 1) ?? { pattern: entry.pattern, slots: {} });
    // The request joins the brief: later edits, and the check of chart numbers, see what was asked for.
    const brief = [deck.brief, instruction.trim()].filter(Boolean).join('\n');
    const next: DeckSpec = { ...deck, version: 1, language, direction, brief, slides };
    if (textScale === 1) delete next.textScale;
    else next.textScale = textScale;
    if (plan.title) next.title = plan.title;
    if (Object.keys(plan.fields).length > 0) next.fields = plan.fields;
    return next;
  };

  if (jobs.length === 0) {
    // Reordering and deleting need no writing.
    return { deck: assemble(new Map()), summary: plan.summary, rewritten: [], warnings: [] };
  }

  // 3. Write the slides that change.
  options.onProgress?.({ stage: 'write', slides: jobs.length });
  const overview = plan.entries
    .map((entry, index) => {
      const number = index + 1;
      const keptSlide = kept.get(number);
      return keptSlide ? `${number}. ${entry.pattern}, unchanged ${slideLabel(keptSlide)}` : `${number}. ${entry.pattern}, to write`;
    })
    .join('\n');
  const tasks = jobs
    .map((job) => {
      const offeredSlots = new Set(job.pattern.slots.map((planned) => planned.slot.name));
      const lines = [`s${job.number} (${job.pattern.def.id}). ${job.source ? 'Change' : 'New slide'}: ${job.purpose}`];
      if (job.source) {
        // Show what is there now. After a layout change the slot names differ, so everything is shown as material.
        const current = Object.entries(job.source.slots).filter(([name, value]) => (job.relaid || offeredSlots.has(name)) && value !== null);
        const shown = Object.fromEntries(
          current.map(([name, value]) => [name, isImageRef(value) ? (value.prompt ? { prompt: value.prompt, alt: value.alt ?? '' } : value.asset) : value]),
        );
        lines.push(`${job.relaid ? 'Content of the slide it replaces' : 'Current content'}: ${JSON.stringify(shown)}`);
      }
      return lines.join('\n');
    })
    .join('\n\n');
  const request = [
    deck.brief ? `The deck was written from this brief:\n${deck.brief}` : '',
    `Requested change:\n${instruction.trim()}`,
    `The deck after the change:\n${overview}`,
    'Write only the slides below. When a slide has current content, start from it and change only what is asked: keep every other text word for word.',
    tasks,
    ...suppliesText(jobs, context),
    'Fill every required property. Leave an optional property out when the slide is better without it.',
  ]
    .filter(Boolean)
    .join('\n\n');

  const draft = await writeSlides({
    ask: asker(options.llm, systemPrompt(language, 'deck'), options.signal),
    request,
    jobs,
    context,
    assemble,
    manifest,
    measure,
    review: reviewWith(options),
    layout,
    maxRepairs: options.maxRepairs ?? 2,
    onRepair: (attempt, issues) => options.onProgress?.({ stage: 'repair', attempt, issues }),
  });

  const errors = draft.issues.filter((issue) => issue.severity === 'error');
  if (errors.length > 0) throw new DeckValidationError(errors);
  const rewritten = jobs.map((job) => job.number);
  const unstated = unstatedChartValues(manifest, draft.deck).filter((issue) => rewritten.includes(issue.slide as number));
  return { deck: draft.deck, summary: plan.summary, rewritten, warnings: [...draft.issues, ...unstated] };
}
