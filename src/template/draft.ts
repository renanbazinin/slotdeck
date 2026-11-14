// Drafts a manifest for a pptx. Code works out the facts (which shapes exist, what repeats
// across slides, how much text a box holds); a language model does the judging (what each
// shape is for, whether a slide is a layout at all). The result is a draft to review.
import type { JsonSchema, LlmProvider } from '../llm/provider.ts';
import { PptxPackage } from '../pptx/package.ts';
import { detectDirection } from '../render/text.ts';
import { changeWith } from './change.ts';
import type { ManifestChange } from './change.ts';
import { checkManifest } from './check.ts';
import type { ManifestIssue } from './check.ts';
import { deriveVariant } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import type { ShapeInfo, SlideInfo } from './inspect.ts';
import type { Direction, PatternDef, SlotDef, TemplateManifest } from './manifest.ts';
import { ROLES, describeShape, frameFacts, isPictogram, normalize, shapeText, slotFromShape, slotName, textLimits } from './measure.ts';
import { detectRepeat } from './rows.ts';
import { pictureFamilies, suggestChanges } from './suggest.ts';
import type { ManifestSuggestion, SuggestionKind } from './suggest.ts';

export { detectRepeat } from './rows.ts';

/** The roles the drafter offers. A chart in place of a picture is a choice for a person to make. */
const DRAFT_ROLES = ROLES.filter((role) => role !== 'chart');

export interface SlidePicture {
  mimeType: string;
  data: Uint8Array;
}

export interface DraftOptions {
  llm: LlmProvider;
  /** Name for the manifest. */
  name?: string;
  /**
   * A picture of a slide (1-based number), when the app can render one. Seeing the slide improves the labels.
   * May return a promise, for a picture that is rendered or read from storage when asked for.
   */
  slideImage?: (slide: number) => SlidePicture | undefined | Promise<SlidePicture | undefined>;
  /** Analyse only these slides. Defaults to all. */
  slides?: number[];
  /**
   * Also add these variants of every drafted pattern whose slide allows them: `mirror` (the
   * picture and the text beside it change sides) and `text-only` (the picture left out, the text
   * widened). Off by default: each variant is one more layout to review, and whether a mirrored
   * slide suits the design is for a person to say. Patterns that allow no variant get none.
   */
  variants?: Array<'mirror' | 'text-only'>;
  /**
   * The most items a drafted row takes. `fit`, the default: as many as the slide has room for side
   * by side at their drawn width, so a row drawn with four cards may take five or six. `drawn`:
   * only as many as the template draws. A grid always takes what it draws.
   */
  rows?: 'fit' | 'drawn';
  /**
   * Lets the items of drafted rows and grids grow wider, up to this factor (above 1, at most 4),
   * when a slide has fewer of them than the template draws on a line. It is set on every row or
   * grid where that can happen: a grid whose fewest items already fill a line gets none. Off by
   * default: the items then keep their width, and the same change is offered in `suggestions`.
   */
  widen?: number;
  /** How many slides to send to the model at once. Defaults to 4. */
  concurrency?: number;
  onProgress?: (event: { done: number; total: number; slide: number }) => void;
  signal?: AbortSignal;
}

export interface DraftResult {
  manifest: TemplateManifest;
  /** Slides that did not become patterns, and why. */
  skipped: Array<{ slide: number; reason: string }>;
  /** What {@link checkManifest} finds in the draft. */
  issues: ManifestIssue[];
  /** What the code settled from measurement on top of the model's labels, in sentences: rows that take more than they draw, pictures labelled photo that are icons. */
  notes: string[];
  /** What the slides allow beyond the draft, each with the changes for `changeManifest`: variants, wider rows. */
  suggestions: ManifestSuggestion[];
}

const SYSTEM = [
  'You label the shapes of a slide from a presentation template, so that software can reuse the slide as a layout.',
  'The software will copy the slide and replace the content of the shapes you mark as content. Everything else stays as designed.',
  "The template's text and pictures are filler. Judge each shape by its position, size, font size and the picture of the slide when you have one.",
  'The layout will carry any subject, so name and describe it by its structure, never by what the filler is about.',
].join(' ');

/**
 * Drafts a manifest for a template: one model call for the deck as a whole, then one per slide.
 * Review the draft before use: check {@link DraftResult.issues}, render a sample deck, and
 * correct budgets or roles where the render shows a problem.
 */
export async function draftManifest(pptx: Uint8Array | PptxPackage, options: DraftOptions): Promise<DraftResult> {
  const pkg = pptx instanceof PptxPackage ? pptx : PptxPackage.load(pptx);
  const info = inspectTemplate(pkg);
  const total = info.slides.length;
  if (options.rows !== undefined && options.rows !== 'fit' && options.rows !== 'drawn') throw new Error(`rows is "fit" or "drawn", not ${JSON.stringify(options.rows)}`);
  for (const variant of options.variants ?? []) if (variant !== 'mirror' && variant !== 'text-only') throw new Error(`a draft adds the variants "mirror" and "text-only", not ${JSON.stringify(variant)}`);
  if (options.widen !== undefined && !(typeof options.widen === 'number' && options.widen > 1 && options.widen <= 4)) throw new Error('widen is a factor above 1 and at most 4, such as 2 to let an item become twice as wide');
  // `room` is a ceiling on the answer in tokens: the answers are short, and a ceiling ends a model that runs away.
  const ask = async (prompt: string, schema: JsonSchema, room: number, slide?: number) => {
    const image = slide !== undefined ? await options.slideImage?.(slide) : undefined;
    const request = { system: SYSTEM, prompt, schema, maxOutputTokens: Math.ceil(room), ...(image ? { images: [image] } : {}), ...(options.signal ? { signal: options.signal } : {}) };
    return options.llm.generateJson(request);
  };
  const notes: string[] = [];

  // --- facts shared by the whole deck -----------------------------------------------------
  const frame = frameFacts(info);
  const repeatedTexts = frame.repeatedTexts;

  const letters = info.slides.flatMap((slide) => slide.shapes.map(shapeText)).map(detectDirection);
  const direction: Direction = letters.filter((d) => d === 'rtl').length > letters.filter((d) => d === 'ltr').length ? 'rtl' : 'ltr';

  // --- one call about the deck: its language, and names for texts that repeat ---------------
  const samples = frame.texts.filter((text) => text.length > 3).slice(0, 12);
  const deckAnswer = (await ask(
    [
      'Sample texts from a presentation template:',
      samples.map((text) => `- ${JSON.stringify(text.slice(0, 80))}`).join('\n'),
      repeatedTexts.length > 0
        ? `These texts appear on many slides, so each is one value for the whole deck (for example the name of a department, or a footer). Give each a short camelCase field name:\n${repeatedTexts.map((text) => `- ${JSON.stringify(text)}`).join('\n')}`
        : '',
      'Also give the BCP 47 tag of the language the template is written in.',
    ]
      .filter(Boolean)
      .join('\n\n'),
    {
      type: 'object',
      properties: {
        language: { type: 'string', description: 'BCP 47 language tag, e.g. "he" or "en"' },
        fields: {
          type: 'array',
          items: { type: 'object', properties: { text: { type: 'string' }, name: { type: 'string' } }, required: ['text', 'name'] },
        },
      },
      required: ['language', 'fields'],
    },
    // The answer repeats each of these texts, so its ceiling grows with them. Bytes bound tokens from above.
    300 + repeatedTexts.reduce((sum, text) => sum + 40 + new TextEncoder().encode(text).length, 0),
  )) as { language?: unknown; fields?: Array<{ text?: unknown; name?: unknown }> } | null;
  const fieldOf = new Map<string, string>();
  repeatedTexts.forEach((text, index) => {
    const named = (Array.isArray(deckAnswer?.fields) ? deckAnswer.fields : []).find((field) => normalize(String(field?.text ?? '')) === text);
    let name = slotName(String(named?.name ?? ''), `field${index + 1}`);
    while ([...fieldOf.values()].includes(name)) name = `${name}${index + 1}`;
    fieldOf.set(text, name);
  });

  // --- one call per slide -----------------------------------------------------------------
  const wanted = (options.slides ?? info.slides.map((slide) => slide.number)).filter((number) => number >= 1 && number <= total);
  const patterns = new Map<number, PatternDef>();
  const skipped: DraftResult['skipped'] = [];
  const slideNotes = new Map<number, string>();
  let done = 0;

  const analyse = async (number: number): Promise<void> => {
    const slide = info.slides[number - 1] as SlideInfo;
    const fixedSlots: SlotDef[] = [];
    const listed: ShapeInfo[] = [];
    for (const shape of slide.shapes) {
      if (shape.kind === 'line' || shape.kind === 'group' || shape.kind === 'other') continue;
      if (frame.isFrame(shape)) continue; // logo, rule, page number or other frame element
      if (shape.kind === 'text') {
        const field = fieldOf.get(shapeText(shape));
        if (field) {
          const limits = textLimits(shape, info, 'text');
          fixedSlots.push({ name: field, kind: 'text', shapeId: shape.id, size: limits.size, maxChars: limits.maxChars, deckField: field });
          continue;
        }
      }
      if (shape.kind === 'shape') continue; // a shape without text carries no content
      listed.push(shape);
    }
    if (listed.length === 0) {
      skipped.push({ slide: number, reason: 'it has no shapes that could carry content' });
      return;
    }

    // What is measurably a pictogram is said so, which settles icon or photo before the model guesses.
    const lines = listed.map((shape) => `${describeShape(shape, info)}${isPictogram(shape, info) ? ', small: a pictogram' : ''}`);

    const prompt = [
      `Slide ${number} of ${total}. Positions and sizes are percentages of the slide. The template reads ${direction === 'rtl' ? 'right to left, so the first of a row of items is the rightmost' : 'left to right'}.`,
      `Shapes (the logo, header and other elements that repeat on most slides are already handled and not listed):\n${lines.join('\n')}`,
      [
        'Decide:',
        '- usable: false when the slide is not a content layout: a sheet of brand assets (icons, colour swatches, graphic elements), instructions for designers, or a slide whose shapes overlap in a way that looks broken.',
        '- id: a short kebab-case name for the layout by its structure, such as "title-photo" or "icon-cards". No subject of the filler and no number of items: "icon-cards", not "four-team-cards".',
        '- description: one sentence telling an author what the layout holds and when to use it. Describe the structure only, never the subject of the filler: "a row of cards, each with an icon, a heading and a short text", not "profiles of team members". For repeated items give no number ("a row of cards", not "four cards"): the software states how many the layout takes.',
        '- figures: true when the layout is built around numbers.',
        '- a role for every shape:',
        '  title, subtitle: the headings of the slide',
        '  text: a paragraph, or a short label',
        '  list: a bulleted or multi-line list, one paragraph per item',
        '  number: a large figure',
        '  caption: a few words under a figure, icon or picture',
        '  photo: a photograph to replace',
        '  icon: a small pictogram to replace. A picture marked "small: a pictogram" is an icon, not a photo.',
        '  graphic: a chart, diagram or screenshot to replace',
        '  table: a table',
        '  decoration: an ornament drawn for the sample content, such as a sticker on a photo, to remove when the content changes',
        '  fixed: part of the design that stays exactly as it is',
        '- name: a camelCase name for what the shape holds. Number repeated items (cards, steps, columns) in reading order, each item with the same set of names: card1Heading, card1Text, card2Heading, card2Text. Name the items by what they are on the slide (card, column, step, stat), not by the subject of the filler (not person, member, product).',
        '- optional: true when the slide still works without that content, such as a subtitle or a footnote.',
      ].join('\n'),
    ].join('\n\n');
    const schema: JsonSchema = {
      type: 'object',
      properties: {
        usable: { type: 'boolean' },
        reason: { type: 'string', description: 'Why the slide is or is not usable, in a few words' },
        id: { type: 'string' },
        description: { type: 'string' },
        figures: { type: 'boolean' },
        shapes: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              shape: { type: 'string', enum: listed.map((shape) => shape.id) },
              role: { type: 'string', enum: [...DRAFT_ROLES] },
              name: { type: 'string' },
              optional: { type: 'boolean' },
            },
            required: ['shape', 'role', 'name', 'optional'],
          },
        },
      },
      required: ['usable', 'reason', 'id', 'description', 'figures', 'shapes'],
    };

    const answer = (await ask(prompt, schema, 400 + 60 * listed.length, number)) as {
      usable?: unknown;
      reason?: unknown;
      id?: unknown;
      description?: unknown;
      figures?: unknown;
      shapes?: Array<{ shape?: unknown; role?: unknown; name?: unknown; optional?: unknown }>;
    } | null;
    options.onProgress?.({ done: ++done, total: wanted.length, slide: number });
    if (answer?.usable !== true) {
      skipped.push({ slide: number, reason: normalize(String(answer?.reason ?? 'the model judged it is not a content layout')) });
      return;
    }

    const byId = new Map(listed.map((shape) => [shape.id, shape]));
    const slots: SlotDef[] = [];
    const names = new Set(fixedSlots.map((slot) => slot.name));
    const seen = new Set<string>();
    for (const entry of Array.isArray(answer.shapes) ? answer.shapes : []) {
      const shape = byId.get(String(entry?.shape ?? ''));
      const role = DRAFT_ROLES.find((candidate) => candidate === entry?.role);
      if (!shape || !role || role === 'fixed' || seen.has(shape.id)) continue;
      seen.add(shape.id);
      let name = slotName(String(entry?.name ?? ''), `${role}${slots.length + 1}`);
      while (names.has(name)) name = `${name}x`;
      names.add(name);
      const slot = slotFromShape(shape, info, role, name, entry?.optional === true, slide);
      if (slot) slots.push(slot);
    }
    if (slots.filter((slot) => (slot.empty ?? 'error') !== 'drop').length === 0) {
      skipped.push({ slide: number, reason: 'no shape was labelled as content' });
      return;
    }

    const pattern: PatternDef = {
      id: String(answer.id ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || `slide-${number}`,
      sourceSlide: number,
      description: normalize(String(answer.description ?? '')) || `Layout of slide ${number}.`,
      slots: [...slots, ...fixedSlots],
    };
    if (answer.figures === true) pattern.figures = true;
    const repeat = detectRepeat(slots, slide);
    if (repeat) pattern.repeat = repeat;
    patterns.set(number, pattern);

    // Photo or icon is settled by measurement, and for a row by all its items together: every item
    // of a row is filled alike, and items beyond the drawn ones are copies of the last. The names
    // stay as the model gave them, so the row found from them stands.
    const settledPictures: string[] = [];
    const quoted = (list: Array<{ name: string }>) => list.map((slot) => `"${slot.name}"`).join(', ');
    for (const family of pictureFamilies(pattern)) {
      const photos = family.filter((slot) => slot.imageKind === 'photo');
      const icons = family.filter((slot) => slot.imageKind === 'icon');
      if (photos.length + icons.length !== family.length || photos.length === 0) continue;
      const small = family.every((slot) => isPictogram(byId.get(slot.shapeId) as ShapeInfo, info));
      const one = photos.length === 1;
      if (small) {
        // No photograph is drawn this small.
        for (const slot of photos) slot.imageKind = 'icon';
        settledPictures.push(`${quoted(photos)} ${one ? 'was' : 'were'} labelled photo, but ${one ? 'its picture is' : 'their pictures are'} no more than 1.1 inch on a side: ${one ? 'an icon slot' : 'icon slots'}.`);
      } else if (icons.length > 0) {
        for (const slot of icons) slot.imageKind = 'photo';
        settledPictures.push(`${quoted(icons)} ${icons.length === 1 ? 'was' : 'were'} labelled icon, but not every picture of the row measures as a pictogram: all of them are photo slots, so that every item is filled alike.`);
      }
    }
    if (settledPictures.length > 0) slideNotes.set(number, `Slide ${number}: ${settledPictures.join(' ')}`);
  };

  const queue = [...wanted];
  await Promise.all(
    Array.from({ length: Math.max(1, Math.min(options.concurrency ?? 4, queue.length)) }, async () => {
      for (let number = queue.shift(); number !== undefined; number = queue.shift()) await analyse(number);
    }),
  );

  // Pattern ids must be unique: later slides with a name already taken get their slide number.
  const ids = new Set<string>();
  const ordered = [...patterns.entries()].sort(([a], [b]) => a - b).map(([number, pattern]) => {
    if (ids.has(pattern.id)) pattern.id = `${pattern.id}-${number}`;
    ids.add(pattern.id);
    return pattern;
  });

  const language = normalize(String(deckAnswer?.language ?? ''));
  let manifest: TemplateManifest = {
    version: 1,
    name: options.name ?? 'template',
    direction,
    // Text boxes rarely share both edges, so mirrored alignment is opt-in per slot.
    alignOnDirectionChange: 'keep',
    patterns: ordered,
  };
  if (language) manifest.language = language;
  for (const number of [...slideNotes.keys()].sort((a, b) => a - b)) notes.push(slideNotes.get(number) as string);

  // How many items a row takes is a fact of the slide, so it is settled here and not left for review.
  const settle: SuggestionKind[] = [...((options.rows ?? 'fit') === 'fit' ? (['row-max'] as const) : []), ...(options.widen !== undefined ? (['row-widen'] as const) : [])];
  const settleRows = () => {
    if (settle.length === 0) return;
    const wanted = suggestChanges({ pkg, manifest }, { kinds: settle, ...(options.widen !== undefined ? { widen: options.widen } : {}) });
    // One change for each row, so that one sentence says what it takes.
    const perRow = new Map<string, ManifestChange>();
    for (const change of wanted.flatMap((suggestion) => suggestion.changes)) {
      if (change.op === 'setRepeat') perRow.set(change.pattern, { ...perRow.get(change.pattern), ...change });
    }
    const settled = changeWith({ pkg, manifest }, [...perRow.values()], info);
    manifest = settled.manifest;
    notes.push(...settled.applied.map((change) => change.message));
  };
  settleRows();

  // Variants are worked out by code from the drafted patterns; no model is asked.
  const withVariants: PatternDef[] = [];
  for (const pattern of manifest.patterns) {
    withVariants.push(pattern);
    for (const variant of new Set(options.variants ?? [])) {
      const made = deriveVariant(pattern, { variant }, { pkg, info, manifest });
      if ('refused' in made) continue;
      let id = `${pattern.id}-${made.suffix}`;
      if (ids.has(id)) id = `${id}-${pattern.sourceSlide}`;
      ids.add(id);
      withVariants.push({ ...made.pattern, id, description: made.description });
    }
  }
  manifest = { ...manifest, patterns: withVariants };
  // A variant copies its pattern's row, and may have other room for it: without the picture, a row can take more.
  if (withVariants.length > ordered.length) settleRows();
  skipped.sort((a, b) => a.slide - b.slide);
  return { manifest, skipped, issues: checkManifest({ pkg, manifest }), notes, suggestions: suggestChanges({ pkg, manifest }) };
}
