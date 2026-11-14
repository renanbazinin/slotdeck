// Drafts a manifest for a pptx. Code works out the facts (which shapes exist, what repeats
// across slides, how much text a box holds); a language model does the judging (what each
// shape is for, whether a slide is a layout at all). The result is a draft to review.
import type { JsonSchema, LlmProvider } from '../llm/provider.ts';
import { PptxPackage } from '../pptx/package.ts';
import { detectDirection } from '../render/text.ts';
import { checkManifest } from './check.ts';
import type { ManifestIssue } from './check.ts';
import { deriveVariant } from './derive.ts';
import { inspectTemplate } from './inspect.ts';
import type { ShapeInfo, SlideInfo } from './inspect.ts';
import type { Direction, PatternDef, RepeatDef, SlotDef, TemplateManifest } from './manifest.ts';
import { ROLES, describeShape, frameFacts, normalize, shapeText, slotFromShape, slotName, textLimits } from './measure.ts';

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
}

const SYSTEM = [
  'You label the shapes of a slide from a presentation template, so that software can reuse the slide as a layout.',
  'The software will copy the slide and replace the content of the shapes you mark as content. Everything else stays as designed.',
  "The template's text is filler. Judge each shape by its position, size, font size and the picture of the slide when you have one.",
].join(' ');

/** Works out a row or grid of like items from numbered slot names such as card1Heading, card2Heading. */
export function detectRepeat(slots: SlotDef[], slide: SlideInfo): RepeatDef | undefined {
  const byId = new Map(slide.shapes.map((shape) => [shape.id, shape]));
  const items = new Map<string, Map<number, SlotDef[]>>();
  for (const slot of slots) {
    const match = /^([a-z][A-Za-z]*?)(\d+)([A-Z][A-Za-z0-9]*)?$/.exec(slot.name);
    // Decorations may be numbered too (sticker1, sticker2), but they are not items of a row.
    if (!match || (slot.kind === 'text' && slot.deckField) || slot.empty === 'drop' || slot.empty === 'keep') continue;
    const prefix = match[1] as string;
    const byNumber = items.get(prefix) ?? new Map<number, SlotDef[]>();
    byNumber.set(Number(match[2]), [...(byNumber.get(Number(match[2])) ?? []), slot]);
    items.set(prefix, byNumber);
  }

  for (const [prefix, byNumber] of items) {
    const count = byNumber.size;
    if (count < 2 || !Array.from({ length: count }, (_, index) => byNumber.has(index + 1)).every(Boolean)) continue;
    const suffixes = (n: number) => (byNumber.get(n) ?? []).map((slot) => slot.name.slice(prefix.length + String(n).length));
    const first = suffixes(1);
    if (!Array.from({ length: count }, (_, index) => suffixes(index + 1).slice().sort().join()).every((set) => set === first.slice().sort().join())) continue;

    // Each item moves as its outermost shapes: the group around a card, or the shapes themselves.
    const chain = (id: string): string[] => {
      const parent = byId.get(id)?.groupId;
      return parent ? [...chain(parent), id] : [id];
    };
    const chains = Array.from({ length: count }, (_, index) => (byNumber.get(index + 1) ?? []).map((slot) => chain(slot.shapeId)));
    // When every item sits inside one shared group, the items are that group's children.
    const shared = chains.flat().every((path) => path[0] === chains[0]?.[0]?.[0] && path.length > 1) ? 1 : 0;
    const shapes = chains.map((paths) => [...new Set(paths.map((path) => path[Math.min(shared, path.length - 1)] as string))]);
    if (new Set(shapes.flat()).size !== shapes.flat().length) continue; // items share a shape: not separable

    const boxOf = (ids: string[]) => {
      const boxes = ids.map((id) => (byId.get(id) as ShapeInfo).box);
      const x = Math.min(...boxes.map((box) => box.x));
      const y = Math.min(...boxes.map((box) => box.y));
      return { x, y, right: Math.max(...boxes.map((box) => box.x + box.w)), bottom: Math.max(...boxes.map((box) => box.y + box.h)) };
    };
    if (shared === 0) {
      // A card's background is often a separate shape behind its group: it belongs to the item too.
      const taken = new Set([...shapes.flat(), ...slots.map((slot) => slot.shapeId)]);
      for (const candidate of slide.shapes) {
        if (candidate.groupId || candidate.kind !== 'shape' || taken.has(candidate.id) || candidate.box.w > 0.6) continue;
        const holds = shapes.filter((ids) => {
          const box = boxOf(ids);
          const slack = 0.01;
          return candidate.box.x <= box.x + slack && candidate.box.y <= box.y + slack && candidate.box.x + candidate.box.w >= box.right - slack && candidate.box.y + candidate.box.h >= box.bottom - slack;
        });
        if (holds.length === 1) (holds[0] as string[]).unshift(candidate.id);
      }
    }

    const tops = shapes.map((ids) => boxOf(ids).y);
    const firstRow = tops.filter((top) => Math.abs(top - (tops[0] as number)) < 0.03).length;
    // Things stacked one under another (two paragraphs, say) are numbered alike but are not a row of items.
    if (firstRow < 2) continue;
    const repeat: RepeatDef = { item: prefix.toLowerCase(), slots: first.map((suffix) => `${prefix}{n}${suffix}`), shapes, min: Math.min(2, count), max: count };
    if (firstRow < count) repeat.columns = firstRow;
    return repeat;
  }
  return undefined;
}

/**
 * Drafts a manifest for a template: one model call for the deck as a whole, then one per slide.
 * Review the draft before use: check {@link DraftResult.issues}, render a sample deck, and
 * correct budgets or roles where the render shows a problem.
 */
export async function draftManifest(pptx: Uint8Array | PptxPackage, options: DraftOptions): Promise<DraftResult> {
  const pkg = pptx instanceof PptxPackage ? pptx : PptxPackage.load(pptx);
  const info = inspectTemplate(pkg);
  const total = info.slides.length;
  const ask = async (prompt: string, schema: JsonSchema, slide?: number) => {
    const image = slide !== undefined ? await options.slideImage?.(slide) : undefined;
    const request = { system: SYSTEM, prompt, schema, ...(image ? { images: [image] } : {}), ...(options.signal ? { signal: options.signal } : {}) };
    return options.llm.generateJson(request);
  };

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

    const lines = listed.map((shape) => describeShape(shape, info));

    const prompt = [
      `Slide ${number} of ${total}. Positions and sizes are percentages of the slide. The template reads ${direction === 'rtl' ? 'right to left, so the first of a row of items is the rightmost' : 'left to right'}.`,
      `Shapes (the logo, header and other elements that repeat on most slides are already handled and not listed):\n${lines.join('\n')}`,
      [
        'Decide:',
        '- usable: false when the slide is not a content layout: a sheet of brand assets (icons, colour swatches, graphic elements), instructions for designers, or a slide whose shapes overlap in a way that looks broken.',
        '- id: a short kebab-case name for the layout, such as "title-photo" or "icon-cards".',
        '- description: one sentence telling an author what the layout holds and when to use it.',
        '- figures: true when the layout is built around numbers.',
        '- a role for every shape:',
        '  title, subtitle: the headings of the slide',
        '  text: a paragraph, or a short label',
        '  list: a bulleted or multi-line list, one paragraph per item',
        '  number: a large figure',
        '  caption: a few words under a figure, icon or picture',
        '  photo: a photograph to replace',
        '  icon: a small pictogram to replace',
        '  graphic: a chart, diagram or screenshot to replace',
        '  table: a table',
        '  decoration: an ornament drawn for the sample content, such as a sticker on a photo, to remove when the content changes',
        '  fixed: part of the design that stays exactly as it is',
        '- name: a camelCase name for what the shape holds. Number repeated items (cards, steps, columns) in reading order, each item with the same set of names: card1Heading, card1Text, card2Heading, card2Text.',
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

    const answer = (await ask(prompt, schema, number)) as {
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
  // Variants are worked out by code from the drafted patterns; no model is asked.
  const withVariants: PatternDef[] = [];
  for (const pattern of ordered) {
    withVariants.push(pattern);
    for (const variant of new Set(options.variants ?? [])) {
      const made = deriveVariant(pattern, { variant }, { pkg, info, manifest: { version: 1, name: '', direction, patterns: ordered } });
      if ('refused' in made) continue;
      let id = `${pattern.id}-${made.suffix}`;
      if (ids.has(id)) id = `${id}-${pattern.sourceSlide}`;
      ids.add(id);
      withVariants.push({ ...made.pattern, id, description: made.description });
    }
  }
  const manifest: TemplateManifest = {
    version: 1,
    name: options.name ?? 'template',
    direction,
    // Text boxes rarely share both edges, so mirrored alignment is opt-in per slot.
    alignOnDirectionChange: 'keep',
    patterns: withVariants,
  };
  if (language) manifest.language = language;
  skipped.sort((a, b) => a.slide - b.slide);
  return { manifest, skipped, issues: checkManifest({ pkg, manifest }) };
}
