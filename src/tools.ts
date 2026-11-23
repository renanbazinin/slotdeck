// The four tools a chat model needs to make presentations with a person in the loop:
// plan, change the plan, create the deck, change the deck.
import { parseHex } from './color/space.ts';
import { FREE_PATTERN, isFreeSlide, isImageRef } from './deck/spec.ts';
import type { DeckColors, DeckSpec, FreeBlockKind } from './deck/spec.ts';
import type { Issue, LayoutMode } from './deck/validate.ts';
import { describeBlocks } from './free/blocks.ts';
import type { ImageProvider, JsonSchema, LlmProvider } from './llm/provider.ts';
import { said } from './messages/en.ts';
import { formatIssue } from './messages/format.ts';
import { resolveAssets } from './plan/assets.ts';
import type { ResolveOptions } from './plan/assets.ts';
import { planPattern } from './plan/catalog.ts';
import type { AssetDef, CatalogContext, IconDef, ImageDensity } from './plan/catalog.ts';
import { MOST_DOCUMENT_CHARS, aboutWords, documentChars, documentKept, documentLimit, documentParts, documentPassages, firstWords, partsText, passageLabel, passageWords, readPassageNumbers } from './plan/document.ts';
import type { DocumentPart, DocumentPassage } from './plan/document.ts';
import { TextCutError, planEdit, recolorDeck } from './plan/edit.ts';
import type { EditOptions, EditResult } from './plan/edit.ts';
import { entryBlocks, freeOnOffer, isFreeEntry, severalOnOneSlide } from './plan/free.ts';
import { readWords, setSmall } from './plan/length.ts';
import { editPlan, freeSlidesOff, patternIds, planOutline, planTextFit, withSettings, writeDeck } from './plan/plan.ts';
import type { Brief, DeckPlan, OutlineEntry, PlanOptions, PlanSettings, SlideTextFit } from './plan/plan.ts';
import { previewExtension } from './preview/provider.ts';
import type { PreviewProvider, SlidePreview } from './preview/provider.ts';
import { renderDeck } from './render/render.ts';
import type { Template } from './render/render.ts';
import type { DeckProperties } from './render/properties.ts';
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
  /** How the app's decks are used, as `delivery` of `planDeck`: `read` for decks that are mailed and read on their own. From 0.7.0. */
  delivery?: PlanOptions['delivery'];
  /**
   * What a presentation's file says of itself (see `RenderOptions.properties`): its title,
   * language, number of slides and the version that made it, in place of the template's. Set
   * unless `false`; an object adds an author, a date or the deck kept in the file. From 0.7.0.
   */
  properties?: DeckProperties | false;
  numberPatterns?: PlanOptions['numberPatterns'];
  /**
   * Whether the writer of a layout's slot of paragraphs is told what its box holds with the
   * template's own paragraph spacing counted, as `countSpacing` of `planDeck`. On unless `false`.
   */
  countSpacing?: boolean;
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
  /**
   * Whether the plan tool takes `documents`: the names of texts the person gave in the
   * conversation, which the slides then carry word for word. The app passes the conversation's
   * texts with every call (`call(name, args, { documents })`) and lists them in the chat's system
   * prompt with `describeDocuments`. Off unless `true`: the plan tool is then declared without it.
   */
  documents?: boolean;
  /**
   * The most characters of the person's texts a presentation is made from, as `maxDocumentChars` of
   * `planOutline`: a whole number from 1000 to 30000, and 12000 (`MOST_DOCUMENT_CHARS`) unless set.
   * Named texts longer than that together are answered with a list of their parts, for the chat model
   * to choose from in `document_parts`. Give `describeDocuments` the same options. Any other value is a
   * `RangeError` when the tools are built. From 0.7.1.
   */
  maxDocumentChars?: number;
}

/** A text the person gave in a conversation, as the app lists it for the chat model. */
export interface ChatDocument {
  /** The name the chat model sees and passes in `documents`: a file name, "pasted text 1". */
  name: string;
  /** The text as the person gave it. */
  text: string;
}

/** What the app passes with a tool call. */
export interface ToolCallContext {
  signal?: AbortSignal;
  /**
   * The texts the person gave in this conversation, the list `describeDocuments` was given:
   * attached files under their names, long messages as "pasted text 1". Read by the plan tool of
   * tools built with `documents: true`, for the names the chat model passes, and nowhere else.
   */
  documents?: readonly ChatDocument[];
}

export interface ToolResult {
  /** What to hand back to the chat model: small, and free of binary data. */
  output: Record<string, unknown>;
  /**
   * What the result tells a person, for the app to show in the person's language with
   * `formatIssue(issue, language, { deck })`: every line of `output.warnings`, in its order, and a
   * note for each slide that the `note` of a plan or of a change that stops before a cut speaks
   * of, and for each text of `output.not_on_slides`. Present only when there is any. Not in
   * `output`, so that it costs the chat model nothing.
   */
  issues?: Issue[];
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
   * Lines for the chat model's system prompt, made from the same options as `definitions`: how to
   * pass on a number of slides, the person's own text and their wishes, what to tell the person,
   * and that a plan is shown before a presentation is made. Add them to the app's own prompt as
   * they are, and they change with an update of slotdeck as the tools do. From 0.7.0.
   */
  instructions: string;
  /**
   * Runs a tool the chat model called. Expected failures come back as `{ error }` for the model to
   * act on, and so does a name that is none of the declared tools. Where the software itself trips
   * (a stored plan that is not what the tools saved), the model is told that in a sentence, not the
   * text of the error. What the app's own model, image provider, store or review throws is passed
   * on as its message. `context` carries the signal that stops the call and, for tools built with
   * `documents: true`, the texts of the conversation.
   */
  call(name: string, args: unknown, context?: ToolCallContext): Promise<ToolResult>;
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

// What makes the chat model pass a person's numbers and text on as they are, and not split,
// shorten or retell them itself: the planner keeps them, and its plan says what that takes. Where
// no free slide of text is on offer nothing is set smaller, and the chat model is not told it is.
const SLIDE_COUNT =
  'Number of slides in all. Set it whenever the user names a number of slides ("2 slides", "a five-slide deck"): it is not read from the topic or the instructions. ' +
  'Leave it out when they name none: the planner then chooses the number. ';
const SLIDE_COUNT_SMALLER = 'With it the presentation has exactly that many slides: long text is set smaller to fit, and no slide is added.';
const SLIDE_COUNT_HOLDS = 'With it the presentation has exactly that many slides, and no slide is added: a slide holds what its layout holds, and the plan says in holds_words and note what is cut.';
const AS_ASKED = 'Do not split, shorten or sum up what the user asked for yourself: pass the number of slides and the amount of text as they said them. ';
const AS_ASKED_SMALLER = 'The planner keeps them, setting long text smaller where it must, and the plan says so in text_pt and note. ';
const AS_ASKED_HOLDS = 'The planner keeps them, and the plan says in holds_words and note where a slide holds less. ';
const NAME_DOCUMENTS = 'When the user wants their own text on the slides and the app lists it, name it in documents and do not copy it into topic: the planner reads it word for word. ';
const GIVEN_TEXT = '. Text the user gives for a slide goes in whole and word for word, in quotation marks';
// A short paste is not listed, so it cannot be named: this is its way to the slides.
const UNLISTED_TEXT = " A text of the user's for the slides that the app does not list (a short paste) goes here, whole and word for word, in quotation marks.";
const ACCEPT_CUT =
  "Only after the user accepted that a slide's text is cut, when an earlier call with this instruction answered with a note saying so. " +
  "Left out, a change that would cut a slide's text is not made, and the answer's note says why.";
/**
 * Said to the chat model in `next` of every result that has something for the person (a `note`,
 * `warnings`, `not_on_slides`): the tools speak English, and the person may not.
 */
const IN_THEIR_LANGUAGE = 'Tell the user what this result says in the language they write in, in your own words and with the same slide numbers and figures; never paste its sentences as they are.';
/** What the chat model is told to do when a change to a presentation would cut a slide's text: nothing was written. */
const CUT_NEXT =
  'Nothing was written. Tell the user what the note says, and ask them, in their language, whether they accept that the rest is cut. ' +
  `If they accept, call ${TOOL_NAMES.edit} again with the same instruction and accept_cut: true; otherwise change the request (more slides, a shorter text). ${IN_THEIR_LANGUAGE}`;

/** Adds {@link IN_THEIR_LANGUAGE} to `next` of an output that has something for the person, once. */
function inTheirLanguage(output: Record<string, unknown>): Record<string, unknown> {
  if (output.note === undefined && output.warnings === undefined && output.not_on_slides === undefined) return output;
  const next = typeof output.next === 'string' ? output.next : '';
  if (!next.endsWith(IN_THEIR_LANGUAGE)) output.next = next ? `${next} ${IN_THEIR_LANGUAGE}` : IN_THEIR_LANGUAGE;
  return output;
}

/** The most texts `describeDocuments` lists. Any number may be named: their characters are what is limited. */
const MOST_DOCUMENTS = 10;
/** A listed text shorter than this is most often the request itself: a topic is not checked against it. */
const CHECKED_CHARS = 1500;
/** A topic with this many runs of four words from one listed text, which make at least half of its runs, copies it. */
const COPIED_RUNS = 120;
/** A topic with this many words of three letters or more, three quarters of them found in one listed text, retells it. */
const RETOLD_WORDS = 40;

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
 * free slide. `documents` says whether the plan tool takes the names of texts the person gave.
 * Only ids, descriptions and switches are read here: nothing that takes time or can fail. What
 * takes time is deciding `freeKinds`, which the caller does once for a loaded template (see
 * `presentationTools`).
 */
/**
 * The lines of the chat model's system prompt that go with the tools: what the tool descriptions
 * cannot say, because it holds for the whole conversation. `freeKinds` and `documents` are what the
 * tools were declared with: a line about several charts on one slide, or about texts named in
 * `documents`, is there only where the tools can do it.
 */
export function toolInstructions(freeKinds: readonly FreeBlockKind[] | undefined, documents = false): string {
  const several = severalOnOneSlide(freeKinds ?? []);
  const smaller = freeKinds?.includes('text') === true;
  const lines = [
    `Making presentations with the tools ${TOOL_NAMES.plan}, ${TOOL_NAMES.editPlan}, ${TOOL_NAMES.create} and ${TOOL_NAMES.edit}:`,
    documents
      ? `- When the user wants their own text on the slides, name it in "documents" and keep "topic" to what the presentation is for. A shorter text of theirs that is not listed goes in "instructions", whole and word for word, in quotation marks. Otherwise put every fact, figure and name the user gave into "topic": a number that is in none of these is not used.`
      : `- When the user wants their own text on the slides, put it in "instructions", whole and word for word, in quotation marks. Otherwise put every fact, figure and name the user gave into "topic": a number that is in neither is not used.`,
    `- Set "slide_count" whenever the user names a number of slides, and pass an amount of text on in "instructions" as they said it. Do not split slides or shorten text yourself${smaller ? ': the planner sets long text smaller' : ''}.`,
    `- When a result has "note", "not_on_slides" or "warnings", tell the user before you go on. Set "accept_cut" only after the user accepted the cut.`,
    `- Speak the language the user writes in. What the tools say is English: retell it in your own words, in the user's language, with the same slide numbers and figures, and never paste it as it is.`,
    `- Pass wishes about the kind of slides on in the user's words ("no tables", "only cards"), in "topic" or "instructions". For a presentation of text only, set "text_only": true; "no pictures" alone ("image_density": "none") still allows cards, tables and charts.`,
    ...(several
      ? [`- A slide can hold ${several.replace(/ on one slide$/, '')}, or any number of cards or figures. When the user asks for that ("all six charts on one slide"), pass the wish on in "instructions" in their words, with the number they said. Do not answer that the template has no such layout.`]
      : []),
    `- Every number a chart is to show must be in "topic". A chart the user gave no numbers for is left off the slide, and the result says so in "warnings": pass that on.`,
    `- Show the plan to the user and wait for their approval before you call ${TOOL_NAMES.create}.`,
    `- For a change to a finished presentation, call ${TOOL_NAMES.edit} with the user's words; do not plan again from scratch. That goes for a slide of many parts too ("add a chart for the new store to slide 3"). Give the numbers of a new chart in the instruction.`,
    '- Refer to presentations and plans by their ids only. Never ask for or write deck JSON.',
  ];
  return lines.join('\n');
}

function toolDefinitions(layouts: Array<Pick<PatternDef, 'id' | 'description'>>, colors: boolean, freeKinds: readonly FreeBlockKind[] | undefined, documents = false): ToolDefinition[] {
  const free = freeKinds !== undefined;
  // Long text is set smaller only on a free slide of text: without one, a slide holds what its layout holds.
  const smaller = freeKinds?.includes('text') === true;
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
      description:
        'What the presentation is about, with every fact, figure and name the user gave. Numbers not stated here will not appear on slides.' +
        (documents ? ' When documents are named, a sentence or two on what the presentation is for is enough: the planner reads the documents themselves.' : ''),
    },
    // The planner chooses the number unless one comes with the call: it cannot take one from the topic.
    slide_count: { type: 'integer', minimum: 1, maximum: MOST_SLIDES, description: SLIDE_COUNT + (smaller ? SLIDE_COUNT_SMALLER : SLIDE_COUNT_HOLDS) },
    language: {
      type: 'string',
      description:
        'BCP 47 tag of the language to write in, e.g. "he" or "en". Defaults to the template language.' +
        // A text the person gave is carried in its own words: the chat model does not have it translated unasked.
        (documents ? ' When documents are named, the language of their text, unless the user asks for another.' : ''),
    },
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
        'Everything else the user asked for, in their own words: audience, tone, every amount of text they ask for ("500 words on each slide", "a paragraph of about 80 words on pricing"), ' +
        'and every wish about how slides look or are arranged that has no argument of its own ("three columns", "no cover slide"). Pass a wish on even when it seems unusual or long.' +
        (documents ? UNLISTED_TEXT : ''),
    },
  };
  if (documents) {
    planProperties.documents = {
      type: 'array',
      items: { type: 'string' },
      description:
        'The texts of this conversation the presentation is made from, by the names the app lists for them (a file name, "pasted text 1"). ' +
        'The slides carry every part of them in their own words and order. Name a text whenever the user wants it on the slides, and also when they want a presentation about it, ' +
        'a summary of it or some parts of it: say which in instructions. Leave it out when the presentation is not made from a text the user gave.',
    };
    planProperties.document_parts = {
      type: 'array',
      items: { type: 'string' },
      description:
        'Only with documents, for a text too long to make a presentation of whole: the parts of it the presentation is made from. ' +
        'Give each by the number or heading this tool listed for it ("2", "4-6", "Risks"), or by a heading the user named. ' +
        'With several texts, put the text\'s name before a number ("report.pdf: 2"). A named text none of whose parts is named here is taken whole. ' +
        'Leave it out to use the named texts whole: for some parts of a text that is not too long, say which in instructions.',
    };
  }
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
      description: `What to change, in the user's words, with any new facts or figures they gave${GIVEN_TEXT}${colors ? '. Leave it empty when only the colour changes' : ''}`,
    },
    ...changeProperties(),
  };
  if (limitable) editPlanProperties.layouts = changedLayouts;

  const editProperties: Record<string, JsonSchema> = {
    presentation_id: { type: 'string', description: 'Id of the presentation to change' },
    instruction: {
      type: 'string',
      description: `What to change, in the user's words, with any new facts or figures they gave${GIVEN_TEXT}${colors ? '. Leave it out when only the colour changes' : ''}`,
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
  // A change that would cut a slide's text is said before anything is written, and made only once the person accepts.
  editProperties.accept_cut = { type: 'boolean', description: ACCEPT_CUT };

  return [
    {
      name: TOOL_NAMES.plan,
      description:
        'Plans a presentation: decides the slides, the layout of each and the point each makes. Nothing is written or rendered yet. ' +
        `Pass on every wish the user has about how it looks or is built: in ${inWords(wishes)} where one fits, and otherwise in instructions in the user's own words. ` +
        // Said so that the chat model does not tell the user that six charts on one slide cannot be had.
        (free ? `A slide can also hold what no layout has (${several ? `${several}, or ` : ''}any number of cards or figures): pass such a wish on in instructions, in the user's words. ` : '') +
        // Said so that the chat model does not make 8 slides of 250 words where 2 of 500 were asked for.
        AS_ASKED +
        (smaller ? AS_ASKED_SMALLER : AS_ASKED_HOLDS) +
        (documents ? NAME_DOCUMENTS : '') +
        'Show the plan to the user in the language they write in, and get their approval or changes before creating the presentation.',
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

/** The passages of the document a plan is made from: none for a plan made from no document. */
function passagesOf(plan: DeckPlan): DocumentPassage[] {
  const document: unknown = plan.brief.document;
  return typeof document === 'string' && document.trim() ? documentPassages(document) : [];
}

/** How much of its document a plan places: the passages and words of the document, and the passages no slide carries. `undefined` for none. */
function documentPlaced(passages: readonly DocumentPassage[], slides: readonly OutlineEntry[]): { passages: number; words: number; left_out?: string } | undefined {
  if (passages.length === 0) return undefined;
  const numbers = passages.map((passage) => passage.number);
  const placed = new Set(slides.flatMap((slide) => readPassageNumbers(slide.passages, passages.length)));
  const missing = numbers.filter((number) => !placed.has(number));
  return { passages: passages.length, words: passageWords(passages, numbers), ...(missing.length > 0 ? { left_out: passageLabel(missing) } : {}) };
}

/**
 * The size a slide's text is set at, where it is well below the template's: under four fifths of
 * it. Text set down to four fifths was set so without remark before there were asked lengths, and
 * a remark on every such slide would have the chat model offer more slides on a number the person
 * gave.
 */
const smallSize = (fit: SlideTextFit): number | undefined => (fit.size !== undefined && fit.templateSize !== undefined && setSmall(fit.size, fit.templateSize) ? fit.size : undefined);

/**
 * What the chat model is to tell the person of a plan before the presentation is made: the slides
 * whose text is cut first, then those that carry all of it in type well below the template's. ''
 * when there is nothing to tell. `freeOff` says a setting turns free slides off, so that no text is
 * set smaller: a cut says so at its end.
 */
function fitNote(fits: readonly SlideTextFit[], freeOff = false): string {
  const said: string[] = [];
  const at = (fit: SlideTextFit): string => {
    const size = smallSize(fit);
    return size !== undefined ? ` (text at ${size} pt)` : '';
  };
  const [cut, ...more] = fits.filter((fit) => fit.holds < fit.words);
  if (cut) {
    said.push(
      `Not all the text fits: slide ${cut.slide} holds about ${cut.holds} of its ${cut.words} words${at(cut)}` +
        more.map((fit) => `, slide ${fit.slide} about ${fit.holds} of its ${fit.words}${at(fit)}`).join('') +
        '. The rest will be cut. Tell the user before creating the presentation, and offer more slides or a shorter text.' +
        (freeOff && [cut, ...more].some((fit) => fit.dense !== true) ? ' Free slides are off, so the text is not set smaller.' : ''),
    );
  }
  const whole = fits.filter((fit) => fit.holds >= fit.words && smallSize(fit) !== undefined);
  const [first] = whole;
  if (first) {
    const sizes = whole.map((fit) => fit.size as number);
    const [least, most] = [Math.min(...sizes), Math.max(...sizes)];
    const which = whole.length === 1 ? `Slide ${first.slide} carries all its` : `Slides ${inWords(whole.map((fit) => String(fit.slide)))} carry all their`;
    said.push(
      `${which} text in smaller type than the template's: ${least === most ? least : `${least} to ${most}`} pt, where the template's is ${first.templateSize} pt. ` +
        'Tell the user, and offer more slides if they want it larger.',
    );
  }
  return said.join(' ');
}

/**
 * The same as a note for each slide, for a person (`slide-holds-less`, then `slide-set-small`): the
 * slides the note names, in its order, each with the size the note says it is set at. A slide set
 * dense needs no word on free slides being off: its text was set smaller.
 */
function fitIssues(fits: readonly SlideTextFit[], freeOff = false): Issue[] {
  const cut = fits.filter((fit) => fit.holds < fit.words);
  const whole = fits.filter((fit) => fit.holds >= fit.words && smallSize(fit) !== undefined);
  return [
    ...cut.map((fit): Issue => ({
      severity: 'warning',
      ...said('slide-holds-less', { slide: fit.slide, words: fit.words, holdsWords: fit.holds, size: smallSize(fit), freeOff: freeOff && fit.dense !== true }),
    })),
    ...whole.map((fit): Issue => ({ severity: 'warning', ...said('slide-set-small', { slide: fit.slide, size: fit.size as number, templateSize: fit.templateSize as number }) })),
  ];
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

/** The texts of a list an app passed that can be read: each with a name and a text. */
function readableDocuments(list: unknown): ChatDocument[] {
  if (!Array.isArray(list)) return [];
  return list.filter((item): item is ChatDocument => typeof item === 'object' && item !== null && typeof item.name === 'string' && typeof item.text === 'string');
}

/** A name or a part as a chat model passed it: trimmed, without the quotes around it. */
const unquoted = (item: string): string => item.trim().replace(/^["'\u201C\u201D\u201E]+|["'\u201C\u201D\u201E]+$/g, '').trim();

/** The names a chat model passed in `documents`: trimmed, without the quotes around them, each once. */
function documentNames(value: unknown): string[] {
  const given: unknown[] = Array.isArray(value) ? value : typeof value === 'string' ? [value] : [];
  const names = given
    .filter((item): item is string => typeof item === 'string')
    .map(unquoted)
    .filter(Boolean);
  return [...new Set(names)];
}

/** Marks that set the direction of text: no part of a word. */
const BIDI_MARKS = /[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g;
/**
 * A name as names are compared. A file name may come in decomposed letters (as from a Mac) or with
 * direction marks, and a chat model gives it back composed and without them.
 */
const nameKey = (name: string): string => name.normalize('NFC').replace(BIDI_MARKS, '').trim();

/** A listed text by the name a chat model gave: that name, else the one that differs only in case, else the one name it begins ("plan 2027" finds "Plan 2027.docx"). */
function findDocument(list: readonly ChatDocument[], name: string): ChatDocument | undefined {
  const key = nameKey(name);
  const exact = list.find((document) => nameKey(document.name) === key);
  if (exact) return exact;
  const lower = key.toLowerCase();
  const lowerOf = (document: ChatDocument): string => nameKey(document.name).toLowerCase();
  for (const matches of [(document: ChatDocument) => lowerOf(document) === lower, (document: ChatDocument) => lowerOf(document).startsWith(lower)]) {
    const found = list.filter(matches);
    if (found.length === 1) return found[0];
  }
  return undefined;
}

// --- parts of a text too long to take whole ---

/** A part number, or a range of them: "2", "part 4", "4-6", "4 to 6". */
const PART_NUMBERS = /^(?:parts?\s*)?(\d+)(?:\s*(?:-|–|—|to)\s*(\d+))?$/i;
/** A part number with the label the list gave it: "4. Risks", "4) Risks (2170)". */
const NUMBERED_LABEL = /^(\d+)[.)]\s+(.+?)(?:\s+\(\d+\))?$/;
/** A list of part numbers and ranges: "2, 4-6". */
const numberList = (text: string): string[] | undefined => {
  const pieces = text.split(',').map((piece) => piece.trim());
  return pieces.every((piece) => PART_NUMBERS.test(piece)) ? pieces : undefined;
};

/** The parts a chat model named in `document_parts`: trimmed, without the quotes around them, a list of numbers split at its commas, each once. */
function partNames(value: unknown): string[] {
  const given: unknown[] = Array.isArray(value) ? value : value === undefined || value === null ? [] : [value];
  const names = given
    .map((item) => (typeof item === 'number' && Number.isInteger(item) ? String(item) : item))
    .filter((item): item is string => typeof item === 'string')
    .map(unquoted)
    .flatMap((item) => numberList(item) ?? [item])
    .filter(Boolean);
  return [...new Set(names)];
}

/** Quote marks, as headings are compared: a Hebrew gershayim is a double quote, as a chat model may give it back. */
const QUOTE_MARKS = /["'״׳“”„‘’]/g;
/** "#" marks before a heading, which stay in its name where a direction mark came before them. */
const HEADING_MARKS = /^#{1,6}\s+/;
/** A heading as headings are compared: in one form, without direction marks or "#" marks, with one kind of quote and of space, without "…" at its end. */
const headingKey = (text: string): string =>
  nameKey(text)
    .replace(HEADING_MARKS, '')
    .replace(QUOTE_MARKS, "'")
    .replace(/\s+/g, ' ')
    .replace(/\s*(?:…|\.\.\.)$/, '')
    .trim();
/** A label of the list as labels are compared: as a heading, in lower case, without the quotes and dots at its end. */
const labelKey = (text: string): string => headingKey(text).toLowerCase().replace(/(?:['\s…]|\.\.\.)+$/, '');
/** A heading as a list shows it: on one line, cut at 80 characters. */
const shownHeading = (heading: string): string => {
  const points = [...nameKey(heading).replace(HEADING_MARKS, '').replace(/\s+/g, ' ')];
  return points.length > 80 ? `${points.slice(0, 79).join('')}…` : points.join('');
};

/** A named text with its parts for the limit, and the label of each part as the list shows it. */
interface TextParts {
  document: ChatDocument;
  parts: DocumentPart[];
  labels: string[];
  /** Whether any part has a heading. */
  headed: boolean;
}

function textParts(document: ChatDocument, most: number): TextParts {
  const parts = documentParts(document.text, most);
  const headed = parts.some((part) => part.headings.length > 0);
  // The sections of each heading, so that two sections of one name are told apart by their first words.
  const sections = new Map<string, Set<number>>();
  for (const part of parts) {
    part.headings.forEach((heading, at) => {
      const key = headingKey(heading);
      const set = sections.get(key) ?? new Set<number>();
      set.add(part.sections[at] ?? -1);
      sections.set(key, set);
    });
  }
  const labels = parts.map((part) => {
    const [first, ...rest] = part.headings;
    if (first === undefined) return `${headed ? 'before the first heading, ' : ''}begins "${part.begins}"`;
    if (rest.length > 0) return `${shownHeading(first)} … ${shownHeading(rest.at(-1) as string)}, ${part.headings.length} headings`;
    const piece = part.piece ? `, ${part.piece.at} of ${part.piece.of}` : '';
    const twice = (sections.get(headingKey(first))?.size ?? 0) > 1 ? `, begins "${part.begins}"` : '';
    return `${shownHeading(first)}${piece}${twice}`;
  });
  return { document, parts, labels, headed };
}

/** The list of a text's parts, one to a line, for the chat model to choose from. */
function partList(text: TextParts, most: number): string {
  const lines = text.parts.map((part, at) => `${part.number}. ${text.labels[at]} (${part.chars}${part.chars > most ? ', too long to make a presentation of alone' : ''})`);
  return `The parts of "${text.document.name}" (${documentChars(text.document.text)} characters${text.headed ? '' : ', no headings'}), each with its characters:\n${lines.join('\n')}`;
}

/** Part numbers in words, runs in a row joined: "part 4", "parts 1-3 and 5". */
function partNumbers(numbers: readonly number[]): string {
  const sorted = [...new Set(numbers)].sort((a, b) => a - b);
  const runs: string[] = [];
  for (let at = 0; at < sorted.length; ) {
    let end = at;
    while (end + 1 < sorted.length && (sorted[end + 1] as number) === (sorted[end] as number) + 1) end++;
    runs.push(end > at ? `${sorted[at]}-${sorted[end]}` : String(sorted[at]));
    at = end + 1;
  }
  return `${sorted.length === 1 ? 'part' : 'parts'} ${inWords(runs)}`;
}

/** Parts of the named texts, by the index of the text and the numbers of its parts. */
type Picks = Array<{ text: number; numbers: number[] }>;

/** What a selector of `document_parts` found: parts, or the sections (or parts) it could mean, one choice each. */
type Found = { picks: Picks } | { could: Picks };

/**
 * The parts the selectors of `document_parts` name, in the named texts: by number or range (in a
 * text named before it, "report.pdf: 2", where more than one text is named), by the label the list
 * gave a part, by a heading (as it is, then without regard to case, then by its beginning, and at
 * the first of these that finds any), by a number with its label as the list gave it, or by a
 * text's name. A heading of two sections or more, in any of the texts, and a number where several
 * texts are named and none before it, is `ambiguous`, the first such selector; selectors that find
 * nothing are `unknown`.
 */
function pickParts(texts: readonly TextParts[]): (selectors: readonly string[]) => { chosen: Array<Set<number>>; ambiguous?: { selector: string; could?: Picks }; unknown: string[] } {
  const all = texts.map((_, at) => at);
  /** The text a name before a colon names: the one of that name, else the one name that differs only in case. */
  const textNamed = (prefix: string): number => {
    const key = nameKey(unquoted(prefix));
    const exact = texts.findIndex((one) => nameKey(one.document.name) === key);
    if (exact >= 0) return exact;
    const cased = all.filter((at) => nameKey((texts[at] as TextParts).document.name).toLowerCase() === key.toLowerCase());
    return cased.length === 1 ? (cased[0] as number) : -1;
  };
  /** The texts a selector speaks of, and what it says of them: one text where it begins with that text's name and a colon. */
  const scopeOf = (selector: string): { scope: number[]; rest: string } => {
    for (let colon = selector.indexOf(':'); colon > 0; colon = selector.indexOf(':', colon + 1)) {
      const text = textNamed(selector.slice(0, colon));
      if (text >= 0) return { scope: [text], rest: unquoted(selector.slice(colon + 1)) };
    }
    return { scope: all, rest: selector };
  };
  /** Every part of a text that holds a section. */
  const partsWith = (text: number, section: number): number[] => (texts[text] as TextParts).parts.filter((part) => part.sections.includes(section)).map((part) => part.number);
  /** Whether a part's label is its heading alone: the heading then decides, as it would in another text. */
  const bare = (text: number, number: number): boolean => {
    const one = texts[text] as TextParts;
    const part = one.parts[number - 1] as DocumentPart;
    return part.headings.length === 1 && !part.piece && one.labels[number - 1] === shownHeading(part.headings[0] as string);
  };
  /** A label, a heading or a text's name, in the texts of the scope. `prefix` lets a heading be found by its beginning. */
  const find = (selector: string, scope: readonly number[], prefix = true): Found | undefined => {
    if (selector.trim() === '') return undefined;
    const label = labelKey(selector);
    const labelled = scope.flatMap((text) =>
      (texts[text] as TextParts).parts.filter((_, at) => labelKey((texts[text] as TextParts).labels[at] as string) === label).map((part) => ({ text, numbers: [part.number] })),
    );
    if (labelled.length > 1) return { could: labelled };
    const only = labelled[0];
    if (only && !bare(only.text, only.numbers[0] as number)) return { picks: [only] };
    const key = headingKey(selector);
    const levels: Array<(heading: string) => boolean> = [
      (heading) => headingKey(heading) === key,
      (heading) => headingKey(heading).toLowerCase() === key.toLowerCase(),
      ...(prefix ? [(heading: string) => key !== '' && headingKey(heading).toLowerCase().startsWith(key.toLowerCase())] : []),
    ];
    for (const matches of levels) {
      const hits: Array<{ text: number; section: number; number: number }> = [];
      for (const text of scope) {
        for (const part of (texts[text] as TextParts).parts) {
          part.headings.forEach((heading, at) => {
            const names = heading.includes(' / ') ? [heading, ...heading.split(' / ')] : [heading];
            if (names.some(matches)) hits.push({ text, section: part.sections[at] ?? -1, number: part.number });
          });
        }
      }
      if (hits.length === 0) continue;
      const first = hits[0] as { text: number; section: number; number: number };
      const sections = new Map(hits.map((hit) => [`${hit.text}:${hit.section}`, hit]));
      if (sections.size === 1) return { picks: [{ text: first.text, numbers: partsWith(first.text, first.section) }] };
      if (new Set(hits.map((hit) => `${hit.text}:${hit.number}`)).size === 1) return { picks: [{ text: first.text, numbers: [first.number] }] };
      // One choice for each section: a section in pieces is one choice of several parts.
      return { could: [...sections.values()].map((hit) => ({ text: hit.text, numbers: partsWith(hit.text, hit.section) })) };
    }
    if (only) return { picks: [only] };
    const named = findDocument(
      scope.map((text) => (texts[text] as TextParts).document),
      selector,
    );
    const text = named ? texts.findIndex((one) => one.document === named) : -1;
    return text >= 0 ? { picks: [{ text, numbers: (texts[text] as TextParts).parts.map((part) => part.number) }] } : undefined;
  };
  /** A list of numbers and ranges in one text: `undefined` where any is not a part of it. */
  const numbered = (pieces: readonly string[], text: number): number[] | undefined => {
    const count = (texts[text] as TextParts).parts.length;
    const numbers: number[] = [];
    for (const piece of pieces) {
      const match = PART_NUMBERS.exec(piece) as RegExpExecArray;
      const a = Number(match[1]);
      const b = match[2] === undefined ? a : Number(match[2]);
      const [low, high] = a <= b ? [a, b] : [b, a];
      if (low < 1 || high > count) return undefined;
      for (let number = low; number <= high; number++) numbers.push(number);
    }
    return numbers;
  };
  /**
   * A number with the label the list gave it ("4. Risks (2170)"). The whole of it is read first, as a
   * heading may begin with a number; then the label, found as it is or without regard to case, and a
   * number that disagrees with the label gives way to it.
   */
  const withLabel = (rest: string, scope: readonly number[]): Found | undefined => {
    const match = NUMBERED_LABEL.exec(rest);
    if (!match) return undefined;
    const whole = find(rest, scope, false);
    if (whole) return whole;
    const number = Number(match[1]);
    const byLabel = find(match[2] as string, scope, false);
    if (!byLabel) return undefined;
    const listed = 'picks' in byLabel ? byLabel.picks : byLabel.could;
    const agree = listed.filter((pick) => pick.numbers.includes(number));
    if (agree.length === 1) return { picks: [{ text: (agree[0] as { text: number }).text, numbers: [number] }] };
    return byLabel;
  };
  return (selectors) => {
    const chosen = texts.map(() => new Set<number>());
    const unknown: string[] = [];
    let ambiguous: { selector: string; could?: Picks } | undefined;
    for (const selector of selectors) {
      const { scope, rest } = scopeOf(selector);
      let found: Found | undefined;
      const pieces = numberList(rest);
      if (pieces) {
        if (scope.length > 1) {
          ambiguous ??= { selector };
          continue;
        }
        const numbers = numbered(pieces, scope[0] as number);
        if (numbers) found = { picks: [{ text: scope[0] as number, numbers }] };
      }
      found ??= withLabel(rest, scope) ?? find(rest, scope);
      if (!found && /\s*\(\d+\)$/.test(rest)) found = find(rest.replace(/\s*\(\d+\)$/, ''), scope);
      if (!found) unknown.push(selector);
      else if ('could' in found) ambiguous ??= { selector, could: found.could };
      else for (const pick of found.picks) for (const number of pick.numbers) chosen[pick.text]?.add(number);
    }
    return { chosen, ...(ambiguous ? { ambiguous } : {}), unknown };
  };
}

/** The words of a text as texts are compared (a topic with a listed text, two named texts): runs of letters and digits, in lower case. */
const wordsOf = (text: string): string[] => text.normalize('NFC').replace(BIDI_MARKS, '').toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
/** Every run of four words in a row. */
const runsOf = (words: readonly string[]): string[] => words.slice(3).map((_, index) => words.slice(index, index + 4).join(' '));
/** A Hebrew word of four letters or more without the one-letter prefix it may have, so that "and the plan" finds "the plan". */
const stemOf = (word: string): string => (/^[א-ת]{4,}$/.test(word) && /^[והבלמשכ]/.test(word) ? word.slice(1) : word);

/** Whether `said` copies `text`: many of its runs of four words are the text's, and they make at least half of its runs. */
function copies(said: string, text: string): boolean {
  const theirs = new Set(runsOf(wordsOf(text)));
  const runs = runsOf(wordsOf(said));
  const found = runs.filter((run) => theirs.has(run)).length;
  return found >= COPIED_RUNS && found >= runs.length / 2;
}

/** Whether `said` retells `text`: it is long, and three quarters of its words of three letters or more are words of the text. */
function retells(said: string, text: string): boolean {
  const theirs = new Set(wordsOf(text).flatMap((word) => [word, stemOf(word)]));
  const mine = wordsOf(said).filter((word) => word.length >= 3 && !/^\d+$/.test(word));
  return mine.length >= RETOLD_WORDS && mine.filter((word) => theirs.has(word) || theirs.has(stemOf(word))).length >= 0.75 * mine.length;
}

/**
 * The line for a chat model's system prompt that lists the texts a person gave in a conversation,
 * each with its words and first words, and says how to make a presentation of one. '' when there
 * are none. Give it the list the tools are given with each call: it lists the last ten of them
 * that have a name and a text. `options` are the options the tools were built with, or any object
 * with their `maxDocumentChars`: a text longer than that is marked too long. Left out, 12000.
 */
export function describeDocuments(documents: readonly ChatDocument[], options?: { maxDocumentChars?: number } | null): string;
/** As a callback of `map`: the index it is given is passed over. */
export function describeDocuments(documents: readonly ChatDocument[], index: number, ...rest: unknown[]): string;
export function describeDocuments(documents: readonly ChatDocument[], options?: { maxDocumentChars?: number } | number | null): string {
  // A second argument that is no object (an index, as from `map`) is left out.
  const most = documentLimit(typeof options === 'object' && options !== null ? options.maxDocumentChars : undefined);
  const listed = readableDocuments(documents)
    .filter((document) => document.name.trim() && document.text.trim())
    .slice(-MOST_DOCUMENTS);
  if (listed.length === 0) return '';
  const long = listed.filter((document) => documentChars(document.text) > most);
  const lines = listed.map((document) => {
    const words = document.text.split(/\s+/).filter(Boolean).length;
    // A text the planner cannot take whole is listed all the same: the chat model names it, and the plan tool lists its parts to choose from.
    const whole = long.includes(document) ? ', too long to make a presentation of whole' : '';
    // The first words, so that the chat model can tell which text is which.
    return `"${document.name}" (${words} words${whole}, begins "${firstWords(document.text)}")`;
  });
  return (
    `Texts the user gave in this conversation, which the presentation tools can read by name: ${lines.join(', ')}. ` +
    `When the user wants a presentation of one of them, name it in "documents" of ${TOOL_NAMES.plan} and say in "topic" in a sentence or two what the presentation is for: ` +
    'the planner reads the text itself and puts all of it on the slides, word for word. Do not copy or retell a listed text in "topic". ' +
    'Name it also when the user wants a summary of it or some parts of it, and say which in "instructions". ' +
    'A message that only asks for a presentation is not such a text: pass what it asks in "topic" and "instructions".' +
    (long.length > 0
      ? ` A text marked too long is named all the same: put the part the user wants in "document_parts", by its heading as the user said it, or leave that out and ${TOOL_NAMES.plan} answers with the parts to choose from. Never ask the user to paste a part of a text they gave.`
      : '')
  );
}

/**
 * Builds the presentation tools for a chat model. Give `definitions` to the model and pass its
 * tool calls to `call`. Plans, decks and pictures are kept in the store under ids, so the model
 * only ever handles ids and short summaries, never deck JSON or files.
 */
export function presentationTools(options: PresentationToolsOptions): PresentationTools {
  const { template } = options;
  // The most characters of named texts, read once: a wrong value is the app's, said when the tools are built (in a call it would be PLAN_FAILED).
  const most = documentLimit(options.maxDocumentChars);
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
  const definitions = toolDefinitions(limitable, colorsOn, freeKinds, options.documents === true);
  const instructions = toolInstructions(freeKinds, options.documents === true);

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
    if (options.delivery) plan.delivery = options.delivery;
    if (options.layout) plan.layout = options.layout;
    if (options.icons) plan.icons = options.icons;
    if (options.assets) plan.assets = options.assets;
    if (options.patterns) plan.patterns = options.patterns;
    if (options.numberPatterns) plan.numberPatterns = options.numberPatterns;
    if (options.countSpacing !== undefined) plan.countSpacing = options.countSpacing;
    if (review) plan.review = review;
    if (options.composeLayouts && tool === TOOL_NAMES.plan) plan.composeLayouts = options.composeLayouts;
    if (options.free !== undefined) plan.free = options.free;
    if (options.maxDocumentChars != null) plan.maxDocumentChars = most;
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

  // What a plan's slides hold is worked out with the options the planner has, and asks no model.
  const fitOptions: Pick<PlanOptions, 'patterns' | 'free' | 'icons' | 'assets' | 'layout' | 'numberPatterns' | 'countSpacing'> = {};
  if (options.layout) fitOptions.layout = options.layout;
  if (options.icons) fitOptions.icons = options.icons;
  if (options.assets) fitOptions.assets = options.assets;
  if (options.patterns) fitOptions.patterns = options.patterns;
  if (options.numberPatterns) fitOptions.numberPatterns = options.numberPatterns;
  if (options.free !== undefined) fitOptions.free = options.free;
  if (options.countSpacing !== undefined) fitOptions.countSpacing = options.countSpacing;

  /** Whether a setting turns free slides off for a plan. The plan stands without it: a plan that trips it is said to have them. */
  const freeOffFor = (plan: DeckPlan): boolean => {
    try {
      return freeSlidesOff(template, plan, fitOptions);
    } catch {
      return false;
    }
  };

  /**
   * What each slide of a plan given words or a document's text will carry, worked out by code
   * before anything is written: [] for any other plan. The plan stands without it, so a slide that
   * cannot be worked out, or a plan that trips the working out, is said nothing of.
   */
  const textFits = (plan: DeckPlan): SlideTextFit[] => {
    try {
      return planTextFit(template, plan, fitOptions);
    } catch {
      return [];
    }
  };

  /**
   * A plan as the chat model is shown it. A plan given words or a document's text also says, by
   * slide, the passages it carries, a size well below the template's and the words it keeps where
   * it is cut, and in `note` what to tell the person. `countGiven` is false when the plan tool was
   * called without `slide_count`: the chat model is then asked to plan again with the number the
   * person named, if they named one. `issues` are the note for a person, slide by slide.
   */
  const describePlan = (id: string, plan: DeckPlan, countGiven = true): { output: Record<string, unknown>; issues: Issue[] } => {
    const fits = textFits(plan);
    const fitOf = new Map(fits.map((fit) => [fit.slide, fit]));
    const passages = passagesOf(plan);
    // What a slide carries of the person's text and how it is set, after its purpose and words.
    const carried = (slide: OutlineEntry, number: number): Record<string, unknown> => {
      const line: Record<string, unknown> = {};
      const own = readPassageNumbers(slide.passages, passages.length);
      if (own.length > 0) {
        line.passages = passageLabel(own);
        line.text_words = passageWords(passages, own);
      }
      const fit = fitOf.get(number);
      const size = fit ? smallSize(fit) : undefined;
      if (size !== undefined) line.text_pt = size;
      if (fit && fit.holds < fit.words) line.holds_words = fit.holds;
      return line;
    };
    const output: Record<string, unknown> = {
      plan_id: id,
      title: plan.title,
      language: plan.brief.language ?? template.manifest.language,
      ...echoWishes(plan.brief.colors, plan.brief.patterns),
      slides: plan.slides.map((slide, index) => {
        // The words asked for a slide, so that the model can tell the user how long its text will be.
        const words = readWords(slide.words);
        // A free slide is laid out for its own parts: it has no description to quote and no picture to show.
        if (isFreeEntry(slide) && !patterns.has(slide.pattern)) {
          return {
            number: index + 1,
            layout: FREE_PATTERN,
            layout_description: `A slide laid out for this presentation: ${describeBlocks(entryBlocks(slide))}.`,
            purpose: slide.purpose,
            ...(words !== undefined ? { words } : {}),
            ...carried(slide, index + 1),
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
        if (words !== undefined) line.words = words;
        Object.assign(line, carried(slide, index + 1));
        // A layout made for this deck has no picture in the app; the model can say it is new.
        if (own) line.new_layout = true;
        const thumbnail = own ? undefined : options.thumbnail?.(slide.pattern);
        if (thumbnail) line.thumbnail = thumbnail;
        return line;
      }),
    };
    const placed = documentPlaced(passages, plan.slides);
    if (placed) output.document = placed;
    // A cut is said with what keeps its text from being set smaller, where a setting does.
    const cut = fits.some((fit) => fit.holds < fit.words);
    const freeOff = cut && freeOffFor(plan);
    const note = fitNote(fits, freeOff);
    if (note) output.note = note;
    // Without a number the planner chose one, and the words asked for a slide never make it more:
    // where the person named a number the chat model left out, it is to plan again with it.
    const chosen = !countGiven && plan.slides.some((slide) => readWords(slide.words) !== undefined);
    output.next =
      (chosen ? `No slide_count was given, so the planner chose ${plan.slides.length} ${plan.slides.length === 1 ? 'slide' : 'slides'}. If the user named a number of slides, plan again with slide_count. ` : '') +
      `Show this plan to the user. Use ${TOOL_NAMES.editPlan} for changes and ${TOOL_NAMES.create} once they approve.`;
    return { output, issues: note ? fitIssues(fits, freeOff) : [] };
  };

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
    const rendered = renderDeck(template, deck, { assets, ...(options.layout ? { layout: options.layout } : {}), ...(options.properties === false ? {} : { properties: options.properties ?? {} }) });
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
    // A presentation made from the person's text says how much of it its slides carry.
    const kept = documentKept(deck);
    if (kept) output.document = kept;
    if (told.length > 0) {
      // A warning about the deck as a whole has no slide to name: it is its message alone. One about
      // a part of a free slide names the part, since the name of its slot (b7Text) says nothing to a
      // person. The line is the one an app makes of the issue with formatIssue in English.
      output.warnings = told.map((issue) => formatIssue(issue, 'en', { deck }));
    }

    // Pictures of the slides are an extra: the presentation stands when they cannot be made.
    let previews: SlidePreview[] | undefined;
    if (options.preview) {
      try {
        const direction = deck.direction ?? template.manifest.direction;
        previews = await options.preview.renderSlides(signal ? { pptx, direction, signal } : { pptx, direction });
        const keys = previews.map((picture) => storeKeys.preview(id, picture.slide, previewExtension(picture)));
        await Promise.all(previews.map((picture, index) => store.set(keys[index] as string, picture.data)));
        output.previews = keys;
      } catch (error) {
        if (signal?.aborted) throw error;
        previews = undefined;
        output.preview_error = error instanceof Error ? error.message : String(error);
      }
    }
    return { id, output, pptx, previews, issues: told };
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

  /**
   * The texts a plan is made from, joined, for the names the chat model passed in `documents`;
   * no `text` when it named none. A name the conversation does not have, an empty text and too much
   * text are errors the chat model can act on: texts longer than the most a presentation is made
   * from are answered with a list of their parts, and the parts named in `document_parts` are what
   * the plan is made from, with `parts` saying which and `partOf` the texts of which only some are
   * used. With no text named, the topic and the instructions are checked against each long listed
   * text: a copy of one is an error that names the argument that copies, and the names of those they
   * retell come back in `retold`, for the output to speak of.
   */
  const documentsOf = (args: Record<string, unknown>, list: unknown, topic: string, instructions: string): { text?: string; names: string[]; retold: string[]; parts?: string[]; partOf?: string[] } => {
    // Where the tools take no documents the arguments are not declared; those that come all the same are passed over.
    if (options.documents !== true) return { names: [], retold: [] };
    const listed = readableDocuments(list);
    const names = documentNames(args.documents);
    const asked = partNames(args.document_parts);
    const quoted = (items: readonly string[]): string => inWords(items.map((name) => `"${name}"`));
    if (names.length === 0) {
      if (asked.length > 0) throw new ToolInputError('"document_parts" needs "documents"');
      // A shorter listed message is most often the request itself, which a topic may well retell.
      const checked = listed.filter((document) => documentChars(document.text) >= CHECKED_CHARS);
      // Each argument on its own, so that the error names the one that copies.
      for (const [argument, said] of Object.entries({ topic, instructions })) {
        const copied = checked.find((document) => copies(said, document.text));
        if (copied) {
          throw new ToolInputError(
            `"${argument}" copies the text "${copied.name}". Name it in documents instead and keep topic to what the presentation is for: the planner reads the text word for word. ` +
              'If the user wants a presentation about it or a summary of it, name it all the same and say so in instructions.',
          );
        }
      }
      return { names: [], retold: checked.filter((document) => retells(`${topic}\n${instructions}`, document.text)).map((document) => document.name) };
    }
    const unknown = names.filter((name) => !findDocument(listed, name));
    if (unknown.length > 0) {
      // A short paste is not listed, and has a way of its own; a file that is not listed is one the app could not read.
      const there = listed.length > 0 ? `The texts the app lists for this conversation are ${quoted(listed.map((document) => document.name))}: use those names.` : 'The app lists no text for this conversation.';
      // A name with a part after it ("report.pdf, Risks"): the part has an argument of its own. Names
      // of several listed texts in one ("report.pdf, notes.docx") are each named on their own.
      const keyOf = (name: string): string => nameKey(name).toLowerCase();
      const after = unknown
        .map((name) => {
          const matches = listed.filter((document) => keyOf(name).length > keyOf(document.name).length && keyOf(name).startsWith(keyOf(document.name)) && ' ,:;(-–#'.includes(keyOf(name)[keyOf(document.name).length] as string));
          if (matches.length !== 1) return undefined;
          const text = matches[0] as ChatDocument;
          const rest = nameKey(name).slice(nameKey(text.name).length).replace(/^[\s,:;(\-–#]+/, '');
          const others = rest.split(',').map(unquoted).filter(Boolean);
          const names = others.length > 0 && others.every((other) => findDocument(listed, other)) ? [text.name, ...others.map((other) => (findDocument(listed, other) as ChatDocument).name)] : undefined;
          return { text, names };
        })
        .find((found) => found !== undefined);
      const hint = !after
        ? ''
        : after.names
          ? ` Name each text on its own in documents: ${quoted(after.names)}.`
          : ` To make it from a part of "${after.text.name}", name "${after.text.name}" in documents and the part in document_parts.`;
      throw new ToolInputError(
        `"documents" names ${unknown.length === 1 ? 'a text' : 'texts'} the app does not have: ${quoted(unknown)}. ${there} ` +
          'A text the user pasted in a short message is not listed: pass it in instructions, whole and word for word, in quotation marks. A file that is not listed could not be read: tell the user.' +
          hint,
      );
    }
    // Two names may find the same text: it is carried once.
    const found = [...new Set(names.map((name) => findDocument(listed, name) as ChatDocument))];
    const empty = found.find((document) => document.text.trim() === '');
    if (empty) throw new ToolInputError(`The text "${empty.name}" is empty: tell the user it could not be read.`);
    // So is a text under two names, such as a file and the message it came in: a text whose words
    // all come in a row in another named text is left out, and an equal one named later. Words
    // that are only in the same order, with others between them, make a text of its own.
    const flat = found.map((document) => ` ${wordsOf(document.text).join(' ')} `);
    const texts = found.filter((_, at) => {
      const mine = flat[at] as string;
      return !flat.some((other, by) => by !== at && other.includes(mine) && (other.length > mine.length || by < at));
    });
    // Said with a list of parts, so that the chat model knows why a text it named is not listed.
    const dropped = found
      .map((document, at) => ({ document, keeper: texts.includes(document) ? undefined : texts.find((text) => flat[found.indexOf(text)]?.includes(flat[at] as string)) }))
      .filter((pair): pair is { document: ChatDocument; keeper: ChatDocument } => pair.keeper !== undefined)
      .map(({ document, keeper }) => ` "${document.name}" is in "${keeper.name}" word for word, so only "${keeper.name}" is used.`)
      .join('');
    const it = texts.length === 1 ? 'it' : 'them';
    const words = aboutWords(most);
    if (asked.length === 0) {
      const joined = texts.map((document) => document.text.replace(/^\s+|\s+$/g, '')).join('\n\n');
      // Counted as the planner counts it, so that a text the tool takes is never one the planner refuses.
      const chars = documentChars(joined);
      if (chars <= most) return { text: joined, names: texts.map((document) => document.name), retold: [] };
      const lists = texts.map((document) => partList(textParts(document, most), most)).join('\n');
      throw new ToolInputError(
        `${texts.length === 1 ? `The text "${(texts[0] as ChatDocument).name}" is ${chars} characters` : `The texts ${quoted(texts.map((document) => document.name))} are ${chars} characters together`}, ` +
          `and a presentation is made from at most ${most} (about ${words} words). ` +
          `Do not ask the user to paste a part of ${it}: call ${TOOL_NAMES.plan} again with the same documents, and name in document_parts the parts the presentation is about, by number or heading, up to ${most} characters in all.` +
          (texts.length > 1 ? ` With several texts, put the text's name before a number ("${(texts[0] as ChatDocument).name}: 2"). A named text none of whose parts is named is taken whole.` : '') +
          ' If the user did not say which parts they want, ask them first, and say what the parts are about. The slides carry only the parts named.' +
          `${dropped}\n${lists}`,
      );
    }
    const each = texts.map((document) => textParts(document, most));
    const lists = each.map((text) => partList(text, most)).join('\n');
    const picked = pickParts(each)(asked);
    if (picked.ambiguous) {
      const { selector, could } = picked.ambiguous;
      if (!could) {
        throw new ToolInputError(
          `"document_parts" names "${selector}", and more than one text is named: put the text's name before the number, as "${(texts[0] as ChatDocument).name}: ${selector}".${dropped}\n${lists}`,
        );
      }
      // One choice for each section it could mean, so that a section in pieces does not read as a range of two.
      const choices = could.map((pick) => `${partNumbers(pick.numbers)}${texts.length > 1 ? ` of "${(texts[pick.text] as ChatDocument).name}"` : ''}`);
      const which = `${choices.slice(0, -1).join(', ')}${choices.length > 1 ? ' or ' : ''}${choices.at(-1)}`;
      throw new ToolInputError(
        `"document_parts" names "${selector}", which could be ${which}: name the parts meant by their numbers${texts.length > 1 ? `, with the text's name before each ("${(texts[0] as ChatDocument).name}: 2")` : ''}.${dropped}\n${lists}`,
      );
    }
    if (picked.unknown.length > 0) {
      throw new ToolInputError(
        `"document_parts" names ${picked.unknown.length === 1 ? 'a part' : 'parts'} the ${texts.length === 1 ? 'text does' : 'texts do'} not have: ${quoted(picked.unknown)}. ` +
          `Name parts by the numbers or headings below, with the same documents.${dropped}\n${lists}`,
      );
    }
    const joined = each
      .map((text, at) => {
        const chosen = picked.chosen[at] as Set<number>;
        return chosen.size > 0 ? partsText(text.document.text, text.parts.filter((part) => chosen.has(part.number))) : text.document.text.replace(/^\s+|\s+$/g, '');
      })
      .join('\n\n');
    const size = documentChars(joined);
    if (size > most) {
      // What was used, text by text: some parts, all of them, or the whole text where no part of it was named.
      const used = each.map((text, at) => {
        const chosen = picked.chosen[at] as Set<number>;
        const name = `"${text.document.name}"`;
        if (chosen.size === 0) return { said: `all of ${name}, since no part of it is named`, one: true };
        if (chosen.size === text.parts.length) return { said: `all of ${name}`, one: true };
        return { said: `${partNumbers([...chosen])} of ${name}`, one: chosen.size === 1 };
      });
      const said = inWords(used.map((use) => use.said));
      const comes = used.length === 1 && (used[0] as { one: boolean }).one ? 'comes' : 'come';
      // A part that is longer than the limit alone could not be cut: a table, or a line with no sentence end.
      const uncut = each.flatMap((text, at) => text.parts.filter((part) => picked.chosen[at]?.has(part.number) && part.chars > most).map((part) => `part ${part.number} of "${text.document.name}"`));
      const one = uncut.length === 1;
      throw new ToolInputError(
        `${said.charAt(0).toUpperCase()}${said.slice(1)} ${comes} to ${size} characters, and a presentation is made from at most ${most} (about ${words} words). ` +
          `Name fewer parts, up to ${most} characters in all, and call ${TOOL_NAMES.plan} again with the same documents. ` +
          `If the user wants all of ${it}, tell them it is too much for one presentation and ask which parts to leave out.` +
          (uncut.length > 0
            ? ` ${inWords(uncut).charAt(0).toUpperCase()}${inWords(uncut).slice(1)} cannot be cut and ${one ? 'is' : 'are'} more than a presentation is made from: leave ${one ? 'it' : 'them'} out, or ask the user which rows or lines of ${one ? 'it' : 'them'} they want.`
            : '') +
          `${dropped}\n${lists}`,
      );
    }
    const parts = each.flatMap((text, at) =>
      text.parts.filter((part) => picked.chosen[at]?.has(part.number)).map((part) => `${text.document.name}: ${part.number}. ${text.labels[part.number - 1]}`),
    );
    // The texts of which only some parts are used: the person is told which.
    const partOf = each.filter((text, at) => (picked.chosen[at]?.size ?? 0) > 0 && (picked.chosen[at]?.size ?? 0) < text.parts.length).map((text) => text.document.name);
    return { text: joined, names: texts.map((document) => document.name), retold: [], parts, partOf };
  };

  const handlers: Record<string, (args: Record<string, unknown>, signal: AbortSignal | undefined, context: ToolCallContext) => Promise<ToolResult>> = {
    async [TOOL_NAMES.plan](args, signal, context) {
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
      // Checked before any model is asked: a wrong name costs the chat model a turn, not a plan.
      const given = documentsOf(args, context.documents, brief.prompt, brief.instructions ?? '');
      if (given.text !== undefined) brief.document = given.text;

      let plan = await planOutline(template, brief, planOptions(TOOL_NAMES.plan, signal));
      // A named text of which no part goes on a slide only asked for the presentation (the planner
      // left it all out): it is no text to carry. The plan is made again with it as instructions,
      // where the request belongs. Sending the chat model to put it there itself would send a long
      // request back to the copy check, which sends it here. Parts chosen of a text are what the
      // presentation is about, not a request; and a text longer than a request ever is stays out of
      // the instructions, where a larger limit would let it go whole.
      const placedNone = given.text !== undefined && !plan.slides.some((slide) => Array.isArray(slide.passages) && slide.passages.length > 0);
      const request = placedNone && given.parts === undefined && documentChars(given.text as string) <= MOST_DOCUMENT_CHARS;
      if (request) {
        const { document: _request, ...asked } = brief;
        plan = await planOutline(template, { ...asked, instructions: [brief.instructions, given.text].filter(Boolean).join('\n\n') }, planOptions(TOOL_NAMES.plan, signal));
      }
      const id = newId('plan');
      await writeJson(store, storeKeys.plan(id), plan);
      const { output, issues } = describePlan(id, plan, brief.slideCount !== undefined);
      if (request) {
        const one = given.names.length === 1;
        output.next =
          `${inWords(given.names.map((name) => `"${name}"`))} only ${one ? 'asks' : 'ask'} for a presentation, so the planner took ${one ? 'it' : 'them'} as instructions, ` +
          `and no slide carries ${one ? 'it' : 'them'} word for word. ${String(output.next)}`;
      } else if (placedNone) {
        output.next = `The planner put none of ${given.parts ? 'the parts named' : inWords(given.names.map((name) => `"${name}"`))} on a slide. Ask the user what the presentation should carry, and plan again. ${String(output.next)}`;
      }
      if (given.parts) {
        if (output.document) output.document = { ...(output.document as Record<string, unknown>), parts_used: given.parts };
        if (!placedNone && given.partOf && given.partOf.length > 0) {
          output.next = `Only some parts of ${inWords(given.partOf.map((name) => `"${name}"`))} are in this plan, as document.parts_used lists: tell the user which. ${String(output.next)}`;
        }
      }
      if (given.retold.length > 0) {
        // Whether the person wanted that text on the slides the chat model knows, and the tools do not.
        const which = inWords(given.retold.map((name) => `"${name}"`));
        output.not_on_slides = given.retold;
        output.next =
          `The topic retells ${which}, and the slides will not carry ${given.retold.length === 1 ? 'that text' : 'those texts'} word for word. ` +
          `If the user wanted their text on the slides, plan again and name it in documents. ${String(output.next)}`;
        issues.push(...given.retold.map((name): Issue => ({ severity: 'warning', ...said('text-retold', { name }) })));
      }
      return withIssues({ output: inTheirLanguage(output), plan }, issues);
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
      const described = describePlan(id, plan);
      return withIssues({ output: inTheirLanguage({ ...described.output, summary: told.filter(Boolean).join(' '), previous_plan_id: planId }), plan }, described.issues);
    },

    async [TOOL_NAMES.create](args, signal) {
      const planId = text(args, 'plan_id');
      const plan = await readJson<DeckPlan>(store, storeKeys.plan(planId));
      if (!plan) throw new ToolInputError(`There is no plan with id "${planId}"`);
      if (!plannable(plan)) return { output: { error: PLAN_UNREADABLE } };
      const written = await writeDeck(template, plan, planOptions(TOOL_NAMES.create, signal));
      const { output, pptx, previews, issues } = await finish(TOOL_NAMES.create, written.deck, written.warnings, signal);
      return withIssues(withPreviews({ output: inTheirLanguage({ ...output, plan_id: planId }), deck: written.deck, pptx }, previews), issues);
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
        // A slide's text is cut only once the person has accepted it. Until then the change stops
        // after its plan, before anything is written, and the chat model is told what would be cut.
        editOptions.stopBeforeCut = args.accept_cut !== true;
        let result: EditResult;
        try {
          result = await planEdit(template, deck, instruction, editOptions);
        } catch (error) {
          if (error instanceof TextCutError) return withIssues({ output: { note: error.message, next: CUT_NEXT } }, error.issues);
          throw error;
        }
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

      const { output, pptx, previews, issues } = await finish(TOOL_NAMES.edit, edited.deck, edited.warnings, signal);
      return withIssues(
        withPreviews(
          {
            output: inTheirLanguage({ ...output, summary: edited.summary, rewritten_slides: edited.rewritten, previous_presentation_id: deckId }),
            deck: edited.deck,
            pptx,
          },
          previews,
        ),
        issues,
      );
    },
  };

  return {
    definitions,
    instructions,
    async call(name, args, context) {
      // Only a tool that is declared is run. "constructor" and "toString" are keys of every object
      // and no tools, and a name may be anything a provider hands on: a symbol cannot stand in a text.
      const handler = typeof name === 'string' && Object.hasOwn(handlers, name) ? handlers[name] : undefined;
      if (!handler) return { output: { error: `Unknown tool "${String(name)}"` } };
      try {
        return await handler((typeof args === 'object' && args !== null ? args : {}) as Record<string, unknown>, context?.signal, context ?? {});
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

/** A result with what it tells a person, where it tells anything. */
function withIssues(result: ToolResult, issues: Issue[]): ToolResult {
  if (issues.length > 0) result.issues = issues;
  return result;
}

class ToolInputError extends Error {}
