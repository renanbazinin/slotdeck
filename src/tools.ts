// The four tools a chat model needs to make presentations with a person in the loop:
// plan, change the plan, create the deck, change the deck.
import { parseHex } from './color/space.ts';
import { FREE_PATTERN, isFreeSlide, isImageRef } from './deck/spec.ts';
import type { DeckColors, DeckSpec, FreeBlockKind } from './deck/spec.ts';
import type { Issue, LayoutMode } from './deck/validate.ts';
import { describeBlocks } from './free/blocks.ts';
import type { ImageProvider, JsonSchema, LlmProvider } from './llm/provider.ts';
import { resolveAssets } from './plan/assets.ts';
import type { ResolveOptions } from './plan/assets.ts';
import { planPattern } from './plan/catalog.ts';
import type { AssetDef, CatalogContext, IconDef, ImageDensity } from './plan/catalog.ts';
import { planEdit, recolorDeck } from './plan/edit.ts';
import type { EditOptions } from './plan/edit.ts';
import { entryBlocks, freeOnOffer, isFreeEntry, partLabel, severalOnOneSlide } from './plan/free.ts';
import { editPlan, patternIds, planOutline, withSettings, writeDeck } from './plan/plan.ts';
import type { Brief, DeckPlan, PlanOptions, PlanSettings } from './plan/plan.ts';
import { previewExtension } from './preview/provider.ts';
import type { PreviewProvider, SlidePreview } from './preview/provider.ts';
import { renderDeck } from './render/render.ts';
import type { Template } from './render/render.ts';
import { readJson, writeJson } from './store.ts';
import type { BlobStore } from './store.ts';
import type { PatternDef } from './template/manifest.ts';
import type { GenerateProgress } from './generate.ts';

export interface ToolDefinition {
  name: string;
  description: string;
  /** JSON Schema of the arguments. */
  parameters: JsonSchema;
}

export interface PresentationToolsOptions {
  template: Template;
  llm: LlmProvider;
  /** Where plans, decks, pictures and files are kept between calls. */
  store: BlobStore;
  /** Generates photos. Without it, decks are planned with no photos. */
  images?: ImageProvider;
  icons?: IconDef[];
  assets?: AssetDef[];
  layout?: LayoutMode;
  /** Deck-level values every deck gets, e.g. the unit shown in the header. */
  fields?: Record<string, string>;
  numberPatterns?: PlanOptions['numberPatterns'];
  /** A second look at every draft the tools write, as for `planDeck`. Off unless set. */
  review?: PlanOptions['review'];
  /**
   * Up to this many new layouts may be made for a new plan, in the template's own look, as for
   * `planOutline`. They are kept with the plan and the deck. Off (0) unless set. None is made
   * while free slides are on offer (see `free`).
   */
  composeLayouts?: number;
  /**
   * Whether a slide may be a free slide, laid out by the software from the parts a person asks
   * for (six charts and a note), as `free` of `planOutline`. On unless `false`. Where free slides
   * are on offer the `layouts` argument takes `free` beside the template's layouts, and the plan
   * tool says that a slide can hold what no layout has. A plan or a presentation that has a free
   * slide is written and shown whatever this says.
   *
   * Whether the template can carry free slides is worked out when the tools are built: it reads
   * the template's slides, which takes a tenth of a second to most of a second for a large
   * template, once for a loaded template. Load the template once and build the tools from that
   * one object; an app that loads it anew for every chat turn pays this on every turn. With
   * `false`, or a `patterns` list that does not name `free`, nothing is read.
   */
  free?: boolean;
  /**
   * The only layouts the tools may use, by pattern id: the app's own list, as `patterns` of
   * `planOutline`. All four tools keep to it. A person may narrow it for one presentation (the
   * `layouts` argument, which then offers these ids alone) and cannot widen it.
   */
  patterns?: string[];
  /**
   * Whether a person may ask for another colour than the template's (the `accent_color` and
   * `color_slides` arguments). On unless `false`. When it is `false`, or the manifest says
   * `colors: false`, the tools are declared without those arguments.
   */
  colors?: boolean;
  /** Address of a picture of a pattern's layout, for showing a plan visually. */
  thumbnail?: (patternId: string) => string | undefined;
  /**
   * Draws the slides of every presentation the tools create or change. The pictures are kept in
   * the store under `preview/<presentation id>/`, and returned to the app with the result.
   */
  preview?: PreviewProvider;
  /** Makes ids for plans and decks. Defaults to a short random id. */
  newId?: (kind: 'plan' | 'deck') => string;
  onProgress?: (event: GenerateProgress, tool: string) => void;
}

export interface ToolResult {
  /** What to hand back to the chat model: small, and free of binary data. */
  output: Record<string, unknown>;
  /** The same result for the app itself, to draw a preview or offer a download. */
  plan?: DeckPlan;
  deck?: DeckSpec;
  pptx?: Uint8Array;
  /** Pictures of the slides, when the tools were given a preview provider. */
  previews?: SlidePreview[];
}

export interface PresentationTools {
  /** Tool declarations to give the chat model. They are made for this template and these options. */
  definitions: ToolDefinition[];
  /**
   * Runs a tool the chat model called. Expected failures come back as `{ error }` for the model to
   * act on, and so does a name that is none of the declared tools. Where the software itself trips
   * (a stored plan that is not what the tools saved), the model is told that in a sentence, not the
   * text of the error. What the app's own model, image provider, store or review throws is passed
   * on as its message.
   */
  call(name: string, args: unknown, context?: { signal?: AbortSignal }): Promise<ToolResult>;
}

export const TOOL_NAMES = {
  plan: 'plan_presentation',
  editPlan: 'edit_presentation_plan',
  create: 'create_presentation',
  edit: 'edit_presentation',
} as const;

/** Store keys, so an app can read what the tools saved. */
export const storeKeys = {
  plan: (id: string) => `plan/${id}.json`,
  deck: (id: string) => `deck/${id}.json`,
  pptx: (id: string) => `pptx/${id}.pptx`,
  asset: (id: string) => `asset/${id}`,
  /** A slide's picture: `extension` is `svg` or `png`, as the preview provider made it. */
  preview: (id: string, slide: number, extension: string) => `preview/${id}/slide-${String(slide).padStart(2, '0')}.${extension}`,
};

/** The first sentence of a text. A full stop inside "e.g." or "i.e." ends none. */
function firstSentence(text: string): string {
  for (const match of text.matchAll(/[.!?](?= |$)/g)) {
    if (match[0] === '.' && /\b(?:e\.g|i\.e|vs|cf)$/i.test(text.slice(0, match.index))) continue;
    return text.slice(0, match.index + 1);
  }
  return text;
}

const cutAt = (text: string, most: number): string => (text.length <= most ? text : `${text.slice(0, most - 3).trimEnd()}...`);

/**
 * One line for each layout a person may name: its id and the first sentence of its description,
 * cut at 80 characters. That is enough to tell most layouts apart. Where two lines come out the
 * same (a layout and its mirror differ in their last words), those layouts get the whole sentence,
 * and failing that the description itself.
 */
function layoutLines(layouts: Array<Pick<PatternDef, 'id' | 'description'>>): string[] {
  const texts = layouts.map((layout) => String(layout.description ?? '').replace(/\s+/g, ' ').trim());
  const sentences = texts.map(firstSentence);
  let lines = sentences.map((sentence) => cutAt(sentence, 80));
  for (const fuller of [sentences, texts.map((text) => cutAt(text, 240))]) {
    const shared = lines.filter((line, index) => lines.indexOf(line) !== index);
    lines = lines.map((line, index) => (shared.includes(line) ? (fuller[index] as string) : line));
  }
  return lines.map((line, index) => `- ${(layouts[index] as { id: string }).id}: ${line}`);
}

/** A list in words: "a", "a and b", "a, b and c". */
const inWords = (items: string[]): string => (items.length < 2 ? items.join('') : `${items.slice(0, -1).join(', ')} and ${items.at(-1)}`);

const ACCENT_COLOR =
  "Only when the user asks for a colour: the main colour as six hex digits (RRGGBB), in place of the template's. " +
  'For a colour given by name choose a shade that suits a presentation (dark green 1E6B3A, not 00FF00).';
const ACCENT_ERROR = '"accent_color" must be six hex digits such as 1E6B3A';

/** The most slides the plan tool takes in `slide_count`: what it declares, and what a larger number is held to. */
const MOST_SLIDES = 30;

// What a chat model is told when the software trips over what it read from the store. The two
// tools that start from a plan say the first, the one that starts from a presentation the second.
// The plan tool reads nothing from the store: there the request itself could not be worked out.
const PLAN_UNREADABLE = `The stored plan could not be read. Make a new plan with ${TOOL_NAMES.plan}.`;
const DECK_UNREADABLE = `The stored presentation could not be read. Make it again: a new plan with ${TOOL_NAMES.plan}, then ${TOOL_NAMES.create}.`;

/**
 * Whether a plan read from the store is one the planner can work with: slides that are objects and
 * a brief with its words. The tools store nothing else; what is not that is a plan that cannot be
 * read, and the chat model is told so, as when the store holds no JSON.
 */
function plannable(plan: unknown): boolean {
  const isObject = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);
  return isObject(plan) && Array.isArray(plan.slides) && plan.slides.every(isObject) && isObject(plan.brief) && typeof plan.brief.prompt === 'string';
}

/** The same for a presentation: slides that are objects, each with its slots. */
function editable(deck: unknown): boolean {
  const isObject = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);
  return isObject(deck) && Array.isArray(deck.slides) && deck.slides.every((slide) => isObject(slide) && typeof slide.slots === 'object' && slide.slots !== null);
}
const PLAN_FAILED = 'The presentation could not be planned: the software failed on this request. Tell the user that it did not work.';

/** How free slides are named among the layouts a person may limit a presentation to. */
const FREE_LAYOUT_LINE = `- ${FREE_PATTERN}: A slide laid out by the software from the parts asked for.`;

/**
 * The tool declarations for one template. `layouts` are the layouts a person may limit a
 * presentation to; `colors` says whether a person may ask for a colour, and `freeKinds` the kinds
 * of part a free slide may hold with what the app supplied, or `undefined` when no slide may be a
 * free slide. Only ids, descriptions and switches are read here: nothing that takes time or can
 * fail. What takes time is deciding `freeKinds`, which the caller does once for a loaded template
 * (see `presentationTools`).
 */
function toolDefinitions(layouts: Array<Pick<PatternDef, 'id' | 'description'>>, colors: boolean, freeKinds: readonly FreeBlockKind[] | undefined): ToolDefinition[] {
  const free = freeKinds !== undefined;
  // Of charts, tables and pictures, those a slide can really be given here: without an image model
  // and without pictures of the app's own, no description promises pictures.
  const several = severalOnOneSlide(freeKinds ?? []);
  // An argument is declared only when it can be used: an enum is never empty.
  const limitable = layouts.length > 0 || free;
  const ids: JsonSchema = { type: 'string', enum: [...layouts.map((layout) => layout.id), ...(free ? [FREE_PATTERN] : [])] };
  const wishes = ['slide_count', ...(limitable ? ['layouts'] : []), 'text_only', 'text_size', 'image_density', ...(colors ? ['accent_color'] : [])];

  const planProperties: Record<string, JsonSchema> = {
    topic: {
      type: 'string',
      description: 'What the presentation is about, with every fact, figure and name the user gave. Numbers not stated here will not appear on slides.',
    },
    // The planner is given 6 to 10 slides unless a number comes with the call: it cannot take one from the topic.
    slide_count: { type: 'integer', minimum: 1, maximum: MOST_SLIDES, description: 'Number of slides. Set it whenever the user names a number: it is not read from the topic.' },
    language: { type: 'string', description: 'BCP 47 tag of the language to write in, e.g. "he" or "en". Defaults to the template language.' },
    image_density: { type: 'string', enum: ['none', 'low', 'medium', 'high'], description: 'How many slides carry a photo. Defaults to medium.' },
    text_size: { type: 'string', enum: ['normal', 'large', 'x-large'], description: 'Size of body text. Larger text means less of it per slide. Defaults to normal.' },
    text_only: {
      type: 'boolean',
      description:
        'True when the user asks for a presentation of text only. Every slide is then plain text: a title with paragraphs or a list, with no pictures, icons, cards, figures, tables or charts.' +
        // Layouts named outright decide: the two together would promise plain text and plan cards.
        (limitable ? ' Not together with layouts: the layouts named then decide.' : ''),
    },
    instructions: {
      type: 'string',
      description:
        'Everything else the user asked for, in their own words: audience, tone, and every wish about how slides look or are arranged that has no argument of its own ("three columns", "no cover slide"). Pass a wish on even when it seems unusual.',
    },
  };
  if (colors) planProperties.accent_color = { type: 'string', description: ACCENT_COLOR };
  if (limitable) {
    planProperties.layouts = {
      type: 'array',
      items: ids,
      description:
        'Only when the user limits the layouts. Every slide then uses one of them, so name the opening and closing layouts too unless the user wants none. The layouts:\n' +
        [...layoutLines(layouts), ...(free ? [FREE_LAYOUT_LINE] : [])].join('\n'),
    };
  }

  // The two tools that change something take the same wishes; the list of layouts is described once, with the plan tool.
  const changeProperties = (): Record<string, JsonSchema> => {
    const properties: Record<string, JsonSchema> = {};
    if (colors) properties.accent_color = { type: 'string', description: `${ACCENT_COLOR} "template" returns to the template's colours.` };
    return properties;
  };
  const changedLayouts: JsonSchema = {
    type: 'array',
    items: ids,
    description:
      `Only when the user changes which layouts may be used: the only layouts allowed from now on. They are described under ${TOOL_NAMES.plan}. ` +
      'Slides that are there already keep their layout unless the instruction says otherwise. An empty list lifts the limit.',
  };

  const editPlanProperties: Record<string, JsonSchema> = {
    plan_id: { type: 'string', description: 'Id of the plan to change' },
    instruction: {
      type: 'string',
      // With a colour and an empty instruction the plan is recoloured by code: no model is asked, and no words about colour join the brief.
      description: `What to change, in the user's words, with any new facts or figures they gave${colors ? '. Leave it empty when only the colour changes' : ''}`,
    },
    ...changeProperties(),
  };
  if (limitable) editPlanProperties.layouts = changedLayouts;

  const editProperties: Record<string, JsonSchema> = {
    presentation_id: { type: 'string', description: 'Id of the presentation to change' },
    instruction: {
      type: 'string',
      description: `What to change, in the user's words, with any new facts or figures they gave${colors ? '. Leave it out when only the colour changes' : ''}`,
    },
    ...changeProperties(),
  };
  if (colors) {
    editProperties.color_slides = {
      type: 'array',
      items: { type: 'integer' },
      description: 'The slides to colour, by number, when the user wants the colour on some slides only. Left out, the whole presentation takes it.',
    };
  }
  if (limitable) editProperties.layouts = changedLayouts;

  return [
    {
      name: TOOL_NAMES.plan,
      description:
        'Plans a presentation: decides the slides, the layout of each and the point each makes. Nothing is written or rendered yet. ' +
        `Pass on every wish the user has about how it looks or is built: in ${inWords(wishes)} where one fits, and otherwise in instructions in the user's own words. ` +
        // Said so that the chat model does not tell the user that six charts on one slide cannot be had.
        (free ? `A slide can also hold what no layout has (${several ? `${several}, or ` : ''}any number of cards or figures): pass such a wish on in instructions, in the user's words. ` : '') +
        'Show the plan to the user and get their approval or changes before creating the presentation.',
      parameters: { type: 'object', properties: planProperties, required: ['topic'] },
    },
    {
      name: TOOL_NAMES.editPlan,
      description:
        'Changes a plan the user has not approved yet: add, remove, reorder or re-purpose slides, change the layout of a slide or what it shows, ' +
        `${colors ? 'the length, the language or the colour' : 'the length or the language'}. Returns a new plan with a new id.`,
      // With a colour, the instruction may be left out: code then sets the colour and no model is asked.
      parameters: { type: 'object', properties: editPlanProperties, required: colors ? ['plan_id'] : ['plan_id', 'instruction'] },
    },
    {
      name: TOOL_NAMES.create,
      description: 'Writes and renders the presentation of an approved plan, as a PowerPoint file. Call it only after the user approved the plan.',
      parameters: {
        type: 'object',
        properties: { plan_id: { type: 'string', description: 'Id of the approved plan' } },
        required: ['plan_id'],
      },
    },
    {
      name: TOOL_NAMES.edit,
      description:
        'Changes a presentation that was already created: rewrite or shorten a slide, add or remove slides, change the layout of a slide or the number of cards on it, ' +
        // Said for the same reason as with the plan tool: what a description does not list, a chat model takes for impossible.
        (several ? `put ${several}, ` : '') +
        `make the text larger or smaller, change a picture, translate${colors ? ', change the colour of the presentation or of one slide' : ''}. ` +
        'Only the slides the change touches are rewritten. Returns a new presentation with a new id.',
      // With a colour, the instruction may be left out: the tool then recolours and asks no model.
      parameters: { type: 'object', properties: editProperties, required: colors ? ['presentation_id'] : ['presentation_id', 'instruction'] },
    },
  ];
}

function randomId(kind: 'plan' | 'deck'): string {
  return `${kind}_${crypto.randomUUID().replace(/-/g, '').slice(0, 10)}`;
}

/** A short line per slide for the chat model: enough to talk about the deck, not the whole content. */
function slideLine(deck: DeckSpec, index: number): Record<string, unknown> {
  const slide = deck.slides[index] as DeckSpec['slides'][number];
  const title = slide.slots.title;
  const line: Record<string, unknown> = { number: index + 1, layout: slide.pattern };
  if (typeof title === 'string') line.title = title;
  // A free slide has no layout to name: the line says so, and what the slide is made of.
  if (isFreeSlide(slide)) {
    line.free_layout = true;
    line.parts = describeBlocks(entryBlocks(slide));
  }
  // A slide that has a colour of its own says so, so that the model can speak of it. So does one
  // that keeps the template's colour inside a presentation that has another: "template", the word
  // the argument takes.
  const accent = parseHex(slide.colors?.accent);
  if (accent) line.accent_color = accent;
  else if (slide.colors !== undefined && parseHex(deck.colors?.accent)) line.accent_color = 'template';
  return line;
}

/**
 * The colour a chat model gave: six hex digits with or without `#`, or three, which are doubled.
 * `null` for "template", the way back to the template's colours. `undefined` when none was given.
 */
function readAccent(value: unknown): string | null | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') throw new ToolInputError(ACCENT_ERROR);
  const given = value.trim();
  if (given === '') return undefined;
  if (given.toLowerCase() === 'template') return null;
  const digits = given.replace(/^#/, '');
  if (/^[0-9a-f]{3}$/i.test(digits)) return [...digits].map((digit) => digit + digit).join('').toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(digits)) return digits.toUpperCase();
  throw new ToolInputError(ACCENT_ERROR);
}

/**
 * The layouts a chat model named. `undefined` when the argument was left out, or names nothing:
 * a number, or a list of things that are no ids. Only an empty list is the empty list, which
 * lifts a limit, so that a slip never lifts one.
 */
function readLayouts(value: unknown): string[] | undefined {
  if (Array.isArray(value)) {
    const ids = patternIds(value);
    return ids.length > 0 || value.length === 0 ? ids : undefined;
  }
  return typeof value === 'string' && value.trim() ? patternIds([value]) : undefined;
}

/**
 * The slides a colour is for, by number. `undefined` when none are named: the whole presentation
 * is meant. A slide number is a whole number or a string of digits; anything else is an error,
 * so that `true` is never read as slide 1.
 */
function slideNumbers(value: unknown, count: number): number[] | undefined {
  if (value === undefined || value === null || (typeof value === 'string' && value.trim() === '')) return undefined;
  const given: unknown[] = Array.isArray(value) ? value : [value];
  if (given.length === 0) return undefined;
  const numbers = given.map((item) => (typeof item === 'number' ? item : typeof item === 'string' && /^\d+$/.test(item.trim()) ? Number(item.trim()) : Number.NaN));
  if (numbers.some((number) => !Number.isInteger(number) || number < 1 || number > count)) {
    throw new ToolInputError(`"color_slides" must be slide numbers from 1 to ${count}`);
  }
  return [...new Set(numbers)].sort((a, b) => a - b);
}

/**
 * Builds the presentation tools for a chat model. Give `definitions` to the model and pass its
 * tool calls to `call`. Plans, decks and pictures are kept in the store under ids, so the model
 * only ever handles ids and short summaries, never deck JSON or files.
 */
export function presentationTools(options: PresentationToolsOptions): PresentationTools {
  const { template } = options;
  // What the app's own parts throw is theirs to word, whatever kind of error it is: a network that
  // fails is a TypeError too ("fetch failed"). Their errors are remembered, so that `call` can tell
  // them from the software tripping over something it read.
  const theirs = new WeakSet<object>();
  const marked =
    <Args extends unknown[], Result>(run: (...args: Args) => Result | Promise<Result>) =>
    async (...args: Args): Promise<Result> => {
      try {
        return await run(...args);
      } catch (error) {
        if (typeof error === 'object' && error !== null) theirs.add(error);
        throw error;
      }
    };
  const llm: LlmProvider = { generateJson: marked((request) => options.llm.generateJson(request)) };
  const images: ImageProvider | undefined = options.images ? { generateImage: marked((request) => (options.images as ImageProvider).generateImage(request)) } : undefined;
  const store: BlobStore = { get: marked((key) => options.store.get(key)), set: marked((key, data) => options.store.set(key, data)) };
  const review: PlanOptions['review'] = options.review ? marked(options.review) : undefined;
  const patterns = new Map(template.manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const newId = options.newId ?? randomId;
  const canMakePhotos = !!options.images || (options.assets ?? []).some((asset) => asset.kind === 'photo');
  // A template that keeps its own colours in every deck is not offered another.
  const colorsOn = options.colors !== false && template.manifest.colors !== false;
  // What a person may limit a presentation to: the layouts these tools can plan with. A layout the
  // app does not allow, or that cannot be filled with what the app supplied (icons without an icon
  // set, a photo when none can be made), is not offered, and `unusable` keeps the reason. Figures
  // count as given here: whether a brief holds any is known only when a deck is planned.
  const supplied: CatalogContext = {
    layout: options.layout ?? 'strict',
    fields: options.fields ?? {},
    icons: options.icons ?? [],
    assets: options.assets ?? [],
    imageDensity: canMakePhotos ? 'medium' : 'none',
    figures: true,
    textScale: 1,
    templateScale: template.manifest.textScale ?? 1,
  };
  const unusable = new Map<string, string>();
  for (const pattern of template.manifest.patterns) {
    if (options.patterns && !options.patterns.includes(pattern.id)) unusable.set(pattern.id, 'is not among the layouts allowed here');
    else if (pattern.figures && options.numberPatterns === 'never') unusable.set(pattern.id, 'is built around numbers, which are turned off here');
    else {
      try {
        const planned = planPattern(pattern, supplied);
        if ('reason' in planned) unusable.set(pattern.id, planned.reason);
      } catch {
        // Declaring the tools never fails. A layout that could not be looked at stays on offer, and planning says what is wrong with it.
      }
    }
  }
  const limitable = template.manifest.patterns.filter((pattern) => !unusable.has(pattern.id));
  const offered = new Set(limitable.map((pattern) => pattern.id));
  // Free slides join the layouts a person may name when these tools can plan them: the template
  // can carry them, and neither the app's option nor its list of layouts leaves them out. Looking
  // at the template is the one thing here that takes time (up to most of a second for a large
  // one). It is remembered for the loaded template, and not done at all where the option or the
  // app's list leaves free slides out.
  // The kinds of part a free slide may hold here, with a number in every brief taken as given.
  let freeKinds: readonly FreeBlockKind[] | undefined;
  if (!patterns.has(FREE_PATTERN)) {
    try {
      const free = freeOnOffer(template, supplied, { direction: template.manifest.direction }, { option: options.free, app: options.patterns });
      freeKinds = free.offer?.kinds;
      if (!freeKinds) unusable.set(FREE_PATTERN, free.refusal ?? 'is not among the layouts allowed here');
    } catch {
      // Declaring the tools never fails: a template that cannot be looked at is offered no free slides here.
      unusable.set(FREE_PATTERN, 'cannot be made from this template');
    }
    if (freeKinds) offered.add(FREE_PATTERN);
  }
  const definitions = toolDefinitions(limitable, colorsOn, freeKinds);

  /**
   * The layouts a `layouts` argument names: `undefined` when it names none, an empty list when it
   * lifts the limit. A layout these tools do not offer is an error that says which and why, with
   * the layouts there are, so that the chat model can put it right before anything is planned.
   */
  const layoutsOf = (args: Record<string, unknown>): string[] | undefined => {
    // Where no layout can be named the argument is not declared; one that comes all the same is passed over.
    const named = offered.size > 0 ? readLayouts(args.layouts) : undefined;
    const wrong = (named ?? []).filter((id) => !offered.has(id));
    if (wrong.length > 0) {
      const why = wrong.map((id) => `${id} ${unusable.get(id) ?? 'is no layout of this template'}`).join('; ');
      throw new ToolInputError(`"layouts" names ${wrong.length === 1 ? 'a layout' : 'layouts'} that cannot be used here: ${why}. The layouts to choose from: ${[...offered].join(', ')}`);
    }
    return named;
  };

  const planOptions = (tool: string, signal: AbortSignal | undefined): PlanOptions => {
    const plan: PlanOptions = { llm };
    if (options.layout) plan.layout = options.layout;
    if (options.icons) plan.icons = options.icons;
    if (options.assets) plan.assets = options.assets;
    if (options.patterns) plan.patterns = options.patterns;
    if (options.numberPatterns) plan.numberPatterns = options.numberPatterns;
    if (review) plan.review = review;
    if (options.composeLayouts && tool === TOOL_NAMES.plan) plan.composeLayouts = options.composeLayouts;
    if (options.free !== undefined) plan.free = options.free;
    if (options.onProgress) plan.onProgress = (event) => options.onProgress?.(event, tool);
    if (signal) plan.signal = signal;
    return plan;
  };

  /**
   * What a person set for the whole presentation, said back to the model: present only when set.
   * The layouts are those of the stored list that these tools offer, which is the limit in force.
   * A layout the template has lost, or the app no longer allows, is not said back as a limit.
   */
  const echoWishes = (colors: DeckColors | undefined, limit: string[] | undefined): Record<string, unknown> => {
    const echo: Record<string, unknown> = {};
    const accent = parseHex(colors?.accent);
    if (accent) echo.accent_color = accent;
    const layouts = patternIds(limit).filter((id) => offered.has(id));
    if (layouts.length > 0) echo.layouts = layouts;
    return echo;
  };

  const describePlan = (id: string, plan: DeckPlan): Record<string, unknown> => ({
    plan_id: id,
    title: plan.title,
    language: plan.brief.language ?? template.manifest.language,
    ...echoWishes(plan.brief.colors, plan.brief.patterns),
    slides: plan.slides.map((slide, index) => {
      // A free slide is laid out for its own parts: it has no description to quote and no picture to show.
      if (isFreeEntry(slide) && !patterns.has(slide.pattern)) {
        return {
          number: index + 1,
          layout: FREE_PATTERN,
          layout_description: `A slide laid out for this presentation: ${describeBlocks(entryBlocks(slide))}.`,
          purpose: slide.purpose,
          free_layout: true,
        };
      }
      const own = plan.layouts?.find((design) => design.id === slide.pattern);
      const line: Record<string, unknown> = {
        number: index + 1,
        layout: slide.pattern,
        layout_description: own?.description ?? patterns.get(slide.pattern)?.description,
        purpose: slide.purpose,
      };
      // A layout made for this deck has no picture in the app; the model can say it is new.
      if (own) line.new_layout = true;
      const thumbnail = own ? undefined : options.thumbnail?.(slide.pattern);
      if (thumbnail) line.thumbnail = thumbnail;
      return line;
    }),
    next: `Show this plan to the user. Use ${TOOL_NAMES.editPlan} for changes and ${TOOL_NAMES.create} once they approve.`,
  });

  /** Gathers the deck's pictures (stored ones first), renders it, and saves everything under a new id. */
  const finish = async (tool: string, deck: DeckSpec, warnings: Issue[], signal: AbortSignal | undefined) => {
    const supplied = new Set([...(options.icons ?? []), ...(options.assets ?? [])].map((item) => item.id));
    const wanted = new Set<string>();
    for (const slide of deck.slides) {
      for (const value of Object.values(slide.slots)) if (isImageRef(value) && !supplied.has(value.asset)) wanted.add(value.asset);
    }
    const cache = new Map<string, Uint8Array>();
    await Promise.all(
      [...wanted].map(async (asset) => {
        const data = await store.get(storeKeys.asset(asset));
        if (data) cache.set(asset, data);
      }),
    );
    const stored = new Set(cache.keys());

    const resolve: ResolveOptions = { cache };
    if (images) resolve.images = images;
    if (options.icons) resolve.icons = options.icons;
    if (options.assets) resolve.assets = options.assets;
    if (options.onProgress) resolve.onProgress = (event) => options.onProgress?.(event, tool);
    if (signal) resolve.signal = signal;
    const assets = await resolveAssets(template, deck, resolve);
    // New pictures go to the store on their own, apart from the JSON and the file.
    await Promise.all([...cache].filter(([asset]) => !stored.has(asset)).map(([asset, data]) => store.set(storeKeys.asset(asset), data)));

    options.onProgress?.({ stage: 'render' }, tool);
    const rendered = renderDeck(template, deck, options.layout ? { assets, layout: options.layout } : { assets });
    const { pptx } = rendered;
    // What the file alone can tell of the deck as a whole (a colour close to the page it lands on)
    // is told too, once: the deck's own check may have said it already.
    const told = [
      ...warnings,
      ...rendered.warnings.filter((issue) => issue.slide === undefined && !warnings.some((known) => known.slide === undefined && known.message === issue.message)),
    ];
    const id = newId('deck');
    await writeJson(store, storeKeys.deck(id), deck);
    await store.set(storeKeys.pptx(id), pptx);

    const output: Record<string, unknown> = {
      presentation_id: id,
      title: deck.title,
      file: storeKeys.pptx(id),
      ...echoWishes(deck.colors, deck.patterns),
      slides: deck.slides.map((_, index) => slideLine(deck, index)),
    };
    if (told.length > 0) {
      // A warning about the deck as a whole has no slide to name: it is its message alone. One about
      // a part of a free slide names the part, since the name of its slot (b7Text) says nothing to a person.
      output.warnings = told.map((issue) => {
        if (issue.slide === undefined) return issue.message;
        const part = partLabel(deck.slides[issue.slide - 1]?.blocks, issue.slot);
        return part ? `slide ${issue.slide}, ${part}: ${issue.message}` : `slide ${issue.slide}${issue.slot ? ` ${issue.slot}` : ''}: ${issue.message}`;
      });
    }

    // Pictures of the slides are an extra: the presentation stands when they cannot be made.
    let previews: SlidePreview[] | undefined;
    if (options.preview) {
      try {
        previews = await options.preview.renderSlides(signal ? { pptx, signal } : { pptx });
        const keys = previews.map((picture) => storeKeys.preview(id, picture.slide, previewExtension(picture)));
        await Promise.all(previews.map((picture, index) => store.set(keys[index] as string, picture.data)));
        output.previews = keys;
      } catch (error) {
        if (signal?.aborted) throw error;
        previews = undefined;
        output.preview_error = error instanceof Error ? error.message : String(error);
      }
    }
    return { id, output, pptx, previews };
  };

  const text = (args: Record<string, unknown>, name: string): string => {
    const value = args[name];
    if (typeof value !== 'string' || value.trim() === '') throw new ToolInputError(`"${name}" is required`);
    return value.trim();
  };

  /** The colours an `accent_color` argument asks for: `null` to return to the template's, `undefined` when it says nothing. */
  const colorsOf = (args: Record<string, unknown>): DeckColors | null | undefined => {
    // Where colours are off the argument is not declared; one that comes all the same is passed over.
    const accent = colorsOn ? readAccent(args.accent_color) : undefined;
    return accent === undefined ? undefined : accent === null ? null : { accent };
  };

  const handlers: Record<string, (args: Record<string, unknown>, signal: AbortSignal | undefined) => Promise<ToolResult>> = {
    async [TOOL_NAMES.plan](args, signal) {
      const brief: Brief = { prompt: text(args, 'topic') };
      if (options.fields) brief.fields = options.fields;
      if (typeof args.language === 'string' && args.language.trim()) brief.language = args.language.trim();
      if (typeof args.instructions === 'string' && args.instructions.trim()) brief.instructions = args.instructions.trim();
      const count = Number(args.slide_count);
      // A chat model that passes more slides than the tool declares is held to what it declares:
      // the number is asked of the planning model as it stands, with room for an answer that long.
      if (Number.isInteger(count) && count > 0) brief.slideCount = Math.min(count, MOST_SLIDES);
      const density = args.image_density;
      if (density === 'none' || density === 'low' || density === 'medium' || density === 'high') brief.imageDensity = density as ImageDensity;
      if (!canMakePhotos) brief.imageDensity = 'none';
      const size = args.text_size;
      if (size === 'normal' || size === 'large' || size === 'x-large') brief.textSize = size;
      const layouts = layoutsOf(args);
      if (layouts?.length) brief.patterns = layouts;
      // Layouts named outright decide what a slide may be. The plan is not also called text only, which it would not be.
      if (args.text_only === true && !brief.patterns) brief.textOnly = true;
      const colors = colorsOf(args);
      if (colors) brief.colors = colors;

      const plan = await planOutline(template, brief, planOptions(TOOL_NAMES.plan, signal));
      const id = newId('plan');
      await writeJson(store, storeKeys.plan(id), plan);
      return { output: describePlan(id, plan), plan };
    },

    async [TOOL_NAMES.editPlan](args, signal) {
      const planId = text(args, 'plan_id');
      const current = await readJson<DeckPlan>(store, storeKeys.plan(planId));
      if (!current) throw new ToolInputError(`There is no plan with id "${planId}"`);
      const instruction = typeof args.instruction === 'string' ? args.instruction.trim() : '';
      const colors = colorsOf(args);
      if (!instruction && colors === undefined) {
        throw new ToolInputError(colorsOn ? '"instruction" is required unless a colour is given' : '"instruction" is required');
      }
      const settings: PlanSettings = {};
      if (colors !== undefined) settings.colors = colors;
      const layouts = layoutsOf(args);
      if (layouts) settings.deckPatterns = layouts;

      let plan: DeckPlan;
      const told: string[] = [];
      if (instruction) {
        if (!plannable(current)) return { output: { error: PLAN_UNREADABLE } };
        const changed = await editPlan(template, current, instruction, { ...planOptions(TOOL_NAMES.editPlan, signal), ...settings });
        plan = changed.plan;
        told.push(changed.summary);
        // What code set is said by code: the model was told that colours are not its work.
        if (colors !== undefined) told.push(planColorSummary(colors));
      } else {
        // A colour and nothing else: code writes it to the plan. No model is asked, and no words about colour join the brief.
        plan = withSettings(template, current, settings);
        told.push(planColorSummary(colors as DeckColors | null));
        if (layouts) told.push(limitSummary(layouts));
      }
      const id = newId('plan');
      await writeJson(store, storeKeys.plan(id), plan);
      return { output: { ...describePlan(id, plan), summary: told.filter(Boolean).join(' '), previous_plan_id: planId }, plan };
    },

    async [TOOL_NAMES.create](args, signal) {
      const planId = text(args, 'plan_id');
      const plan = await readJson<DeckPlan>(store, storeKeys.plan(planId));
      if (!plan) throw new ToolInputError(`There is no plan with id "${planId}"`);
      if (!plannable(plan)) return { output: { error: PLAN_UNREADABLE } };
      const written = await writeDeck(template, plan, planOptions(TOOL_NAMES.create, signal));
      const { output, pptx, previews } = await finish(TOOL_NAMES.create, written.deck, written.warnings, signal);
      return withPreviews({ output: { ...output, plan_id: planId }, deck: written.deck, pptx }, previews);
    },

    async [TOOL_NAMES.edit](args, signal) {
      const deckId = text(args, 'presentation_id');
      const deck = await readJson<DeckSpec>(store, storeKeys.deck(deckId));
      if (!deck) throw new ToolInputError(`There is no presentation with id "${deckId}"`);
      const instruction = typeof args.instruction === 'string' ? args.instruction.trim() : '';
      const colors = colorsOf(args);
      if (!instruction && colors === undefined) {
        throw new ToolInputError(colorsOn ? '"instruction" is required unless a colour is given' : '"instruction" is required');
      }
      if (!editable(deck)) return { output: { error: DECK_UNREADABLE } };
      // The slides a colour is for. Without a colour there is nothing to give them, which is said before any model is asked.
      const slides = colorsOn ? slideNumbers(args.color_slides, deck.slides.length) : undefined;
      if (slides && colors === undefined) throw new ToolInputError('"color_slides" needs "accent_color"');
      const layouts = layoutsOf(args);

      // A colour for the whole presentation is one every slide takes: a slide's own gives way to it.
      const everySlide = (given: DeckSpec): DeckSpec => (slides ? given : recolorDeck(given, null, given.slides.map((_, index) => index + 1)));
      // A slide that returns to the template's colours inside a presentation that has colours of
      // its own must say so, with an empty `colors`. Without it the slide would follow the presentation.
      const asTemplate = (given: DeckSpec, numbers: number[]): DeckSpec =>
        colors !== null || given.colors === undefined ? given : { ...given, slides: given.slides.map((slide, index) => (numbers.includes(index + 1) ? { ...slide, colors: {} } : slide)) };

      let edited: { deck: DeckSpec; summary: string; rewritten: number[]; warnings: Issue[] };
      if (instruction) {
        const editOptions: EditOptions = { ...planOptions(TOOL_NAMES.edit, signal), ...(canMakePhotos ? {} : { imageDensity: 'none' as const }) };
        if (colors !== undefined) editOptions.colors = colors;
        if (slides) editOptions.colorSlides = slides;
        if (layouts) editOptions.deckPatterns = layouts;
        const result = await planEdit(template, deck, instruction, editOptions);
        edited = result;
        if (colors !== undefined) {
          // The slides were numbered as the deck was; the edit says where they are now.
          const now = slides ? (result.recolored ?? []) : undefined;
          // What code set is said by code: the model was told that colours are not its work.
          const told = now?.length === 0 ? '' : recolorSummary(colors, now);
          edited = { ...result, deck: now ? asTemplate(result.deck, now) : everySlide(result.deck), summary: [result.summary, told].filter(Boolean).join(' ') };
        }
      } else {
        // A colour and nothing else: code sets it and the file is rendered again. No model is asked.
        let next = recolorDeck(everySlide(deck), colors as DeckColors | null, slides);
        if (slides) next = asTemplate(next, slides);
        const told = [recolorSummary(colors as DeckColors | null, slides)];
        if (layouts) {
          const { patterns: _before, ...rest } = next;
          next = layouts.length > 0 ? { ...rest, patterns: layouts } : rest;
          told.push(limitSummary(layouts));
        }
        edited = { deck: next, summary: told.join(' '), rewritten: [], warnings: [] };
      }

      const { output, pptx, previews } = await finish(TOOL_NAMES.edit, edited.deck, edited.warnings, signal);
      return withPreviews(
        {
          output: { ...output, summary: edited.summary, rewritten_slides: edited.rewritten, previous_presentation_id: deckId },
          deck: edited.deck,
          pptx,
        },
        previews,
      );
    },
  };

  return {
    definitions,
    async call(name, args, context) {
      // Only a tool that is declared is run. "constructor" and "toString" are keys of every object
      // and no tools, and a name may be anything a provider hands on: a symbol cannot stand in a text.
      const handler = typeof name === 'string' && Object.hasOwn(handlers, name) ? handlers[name] : undefined;
      if (!handler) return { output: { error: `Unknown tool "${String(name)}"` } };
      try {
        return await handler((typeof args === 'object' && args !== null ? args : {}) as Record<string, unknown>, context?.signal);
      } catch (error) {
        if (context?.signal?.aborted) throw error;
        // The software tripping over what it read: a stored plan or presentation that is no JSON, or
        // not what the tools put there. "Cannot read properties of undefined" is nothing a chat model
        // can act on, so it is told what happened. What the app's own model, image provider, store
        // or review says is passed on as it is.
        const tripped = (error instanceof TypeError || error instanceof RangeError || error instanceof SyntaxError) && !theirs.has(error);
        if (tripped) return { output: { error: name === TOOL_NAMES.plan ? PLAN_FAILED : name === TOOL_NAMES.edit ? DECK_UNREADABLE : PLAN_UNREADABLE } };
        // Told to the model in words it can act on: a missing id, a plan it must change, a limit it broke.
        return { output: { error: error instanceof Error ? error.message : String(error) } };
      }
    },
  };
}

/** What a change of colour did, in a sentence for the chat model to pass on. No model wrote it. */
function recolorSummary(colors: DeckColors | null, slides: number[] | undefined): string {
  const accent = colors ? parseHex(colors.accent) : undefined;
  if (!slides) return accent ? `The presentation now has the colour ${accent}.` : "The presentation is back in the template's colours.";
  const which = `${slides.length === 1 ? 'Slide' : 'Slides'} ${inWords(slides.map(String))}`;
  if (accent) return `${which} now ${slides.length === 1 ? 'has' : 'have'} the colour ${accent}.`;
  return `${which} ${slides.length === 1 ? 'is' : 'are'} back in the template's colours.`;
}

/** The same for a plan, whose presentation is still to be made. */
function planColorSummary(colors: DeckColors | null): string {
  const accent = colors ? parseHex(colors.accent) : undefined;
  return accent ? `The presentation will have the colour ${accent}.` : "The presentation will be in the template's colours.";
}

/** What a new list of layouts, set with nothing else to plan, means from now on. */
const limitSummary = (layouts: string[]): string =>
  layouts.length > 0
    ? `The presentation is now limited to the layouts ${layouts.join(', ')}. Slides that are there already keep their layout.`
    : 'The limit on the layouts is lifted.';

function withPreviews(result: ToolResult, previews: SlidePreview[] | undefined): ToolResult {
  if (previews) result.previews = previews;
  return result;
}

class ToolInputError extends Error {}
