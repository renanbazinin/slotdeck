// A free slide of many parts is written by calls of its own, with an array for each run of like
// parts where a small slide has a property for each slot. A schema of two hundred charts as
// properties is too large to send, and one answer that holds them all too long to wait for. This
// file cuts such a slide into calls that each fit, builds the schema and the task of each call,
// and reads an answer back into the values of the slide's slots. When a deck is changed, some
// parts of a slide stay as they are: the calls then write the others, and nothing else.
import { isChartValue, isImageRef, isTableValue, toParagraphs } from '../deck/spec.ts';
import type { FreeBlockKind, SlotValue } from '../deck/spec.ts';
import { describeBlocks } from '../free/blocks.ts';
import type { JsonSchema } from '../llm/provider.ts';
import { textBudget } from '../template/manifest.ts';
import { CHART_FORMAT_NOTE, CHART_RULE, CHART_TYPE_NOTE, PHOTO_ALT_NOTE, PHOTO_ASSET_NOTE, PHOTO_PROMPT_NOTE, budgetScale, slotNote } from './catalog.ts';
import type { CatalogContext, PlannedSlot } from './catalog.ts';
import type { WriteJob } from './write.ts';

/** A job that writes a free slide. */
export type FreeJob = WriteJob & { free: NonNullable<WriteJob['free']> };

/**
 * The most tokens the answer of one such call may take. With what an adapter adds for reasoning
 * this stays inside the output limit of any current model, and an answer of this size comes back
 * well inside a request's time limit. A slide that needs more is written by more calls.
 */
export const WRITE_CALL_TOKENS = 24000;

/**
 * The most arrays the schema of one such call has. Parts of two kinds that take turns (a figure
 * and its chart, twelve times) are an array each, and a schema of a hundred arrays is as large as
 * the one of properties that arrays are there to avoid. A slide with more is written by more calls.
 */
export const MOST_ARRAYS = 12;

/** Like parts that follow one another among the parts to write, as one call writes them: one array of its answer. */
export interface PartRun {
  /** The name of the array: `r01` for the slide's first run of like parts. A run cut over two calls has one name in both. */
  name: string;
  kind: FreeBlockKind;
  /** The parts the array holds in this call, in order, numbered over the whole slide from 1. */
  parts: number[];
}

/** The parts of a slide that one call writes. */
export interface PartRange {
  /** Their numbers, in order. None for a call that writes the title alone. */
  parts: number[];
  runs: PartRun[];
  /** Whether the slide's title is written with these parts: it is, in the slide's first call. */
  title: boolean;
  /** The ceiling of the answer, in tokens. */
  room: number;
  /** Charts and tables only: numbers to set down, which takes little reasoning. */
  light: boolean;
}

/** The slot of a part as its name says it: `b12Caption` is the caption of part 12. `undefined` for a slot that is no part. */
export function partOf(slot: string | undefined): { block: number; part: string } | undefined {
  const match = /^b(\d+)([A-Z][A-Za-z]*)$/.exec(slot ?? '');
  return match ? { block: Number(match[1]), part: match[2] as string } : undefined;
}

/** The slots a model writes for each part of a slide, by the part's number. */
function slotsByPart(job: FreeJob): Map<number, Array<{ part: string; planned: PlannedSlot }>> {
  const byPart = new Map<number, Array<{ part: string; planned: PlannedSlot }>>();
  for (const planned of job.pattern.slots) {
    const at = partOf(planned.slot.name);
    if (!at) continue;
    byPart.set(at.block, [...(byPart.get(at.block) ?? []), { part: at.part, planned }]);
  }
  return byPart;
}

/** How many tokens one part may take in an answer: its slots at their budgets, with room to spare. */
function partRoom(slots: Array<{ planned: PlannedSlot }>, context: CatalogContext): number {
  let tokens = 12; // its number, and the brackets around it
  for (const { planned } of slots) {
    const { slot } = planned;
    // A chart is as long as its numbers: a label and a value for each category of each series, as
    // many as its slot says it shows. The chart of a free slide states both for the size it is drawn at.
    if (slot.kind === 'chart') tokens += (40 + 12 * (slot.maxCategories ?? 8) * (1 + (slot.maxSeries ?? 2))) * 1.5;
    else if (slot.kind === 'text') tokens += ((textBudget(slot, budgetScale(context)).maxChars ?? 400) + 40) * 1.5;
    else if (slot.kind === 'table') tokens += ((slot.maxColumns ?? 4) * ((slot.maxRows ?? 4) + 1) * (slot.maxCellChars ?? 30) + 40) * 1.5;
    else tokens += ((planned.mode === 'photo' ? 500 : 60) + 40) * 1.5;
  }
  return tokens;
}

/** Parts by their numbers, in words: "part 5", "parts 5 and 6", "parts 1 to 12", "parts 3, 7 and 12 to 20". */
export function partsWords(numbers: readonly number[]): string {
  const spans: string[] = [];
  for (let at = 0; at < numbers.length; ) {
    let end = at;
    while (end + 1 < numbers.length && numbers[end + 1] === (numbers[end] as number) + 1) end++;
    // Three or more neighbours are a span. Two are named one by one: "5 and 6".
    if (end - at >= 2) spans.push(`${numbers[at]} to ${numbers[end]}`);
    else for (let index = at; index <= end; index++) spans.push(String(numbers[index]));
    at = end + 1;
  }
  return `${numbers.length === 1 ? 'part' : 'parts'} ${spans.length < 2 ? spans.join('') : `${spans.slice(0, -1).join(', ')} and ${spans.at(-1)}`}`;
}

const kindWords = (kind: FreeBlockKind, number: number): string => describeBlocks(Array<FreeBlockKind>(number).fill(kind));

/** A number as a place in a row: 1st, 2nd, 3rd, 11th, 21st. */
function ordinal(number: number): string {
  const teens = number % 100 >= 11 && number % 100 <= 13;
  return `${number}${teens ? 'th' : (['th', 'st', 'nd', 'rd'][number % 10] ?? 'th')}`;
}

/**
 * Which of their kind some parts are, counted over the whole slide: "the 31st to the 60th chart
 * of the slide", "the 2nd and the 5th figure of the slide". `places` are those counts, in order.
 */
function placesWords(kind: FreeBlockKind, places: readonly number[]): string {
  const spans: string[] = [];
  for (let at = 0; at < places.length; ) {
    let end = at;
    while (end + 1 < places.length && places[end + 1] === (places[end] as number) + 1) end++;
    // As for parts by their numbers: three or more neighbours are a span.
    if (end - at >= 2) spans.push(`the ${ordinal(places[at] as number)} to the ${ordinal(places[end] as number)}`);
    else for (let index = at; index <= end; index++) spans.push(`the ${ordinal(places[index] as number)}`);
    at = end + 1;
  }
  return `${spans.length < 2 ? spans.join('') : `${spans.slice(0, -1).join(', ')} and ${spans.at(-1)}`} ${kind} of the slide`;
}

/**
 * Cuts the parts of a slide into the calls that write them: consecutive parts whose answer fits
 * under {@link WRITE_CALL_TOKENS} and whose schema has at most {@link MOST_ARRAYS} arrays. A slide that needs several calls is cut into even shares, so
 * that calls made side by side take about as long as one another. Only the parts that have
 * something to write are counted: a part whose every slot is kept from the slide as it was is in
 * no call. A slide of which the title alone is written has one call, for the title.
 */
export function partRanges(job: FreeJob, context: CatalogContext): PartRange[] {
  const { blocks } = job.free;
  const byPart = slotsByPart(job);
  const written = blocks.map((_, index) => index + 1).filter((number) => byPart.has(number));
  const rooms = written.map((number) => partRoom(byPart.get(number) ?? [], context));
  const title = job.pattern.slots.find((planned) => planned.slot.name === 'title');
  const overhead = 300 + (title && title.slot.kind === 'text' ? ((textBudget(title.slot, budgetScale(context)).maxChars ?? 100) + 40) * 1.5 : 0);
  const most = WRITE_CALL_TOKENS - overhead;
  if (written.length === 0) return title ? [{ parts: [], runs: [], title: true, room: Math.ceil(overhead), light: false }] : [];

  // Like parts that follow one another among the parts to write are one array, whatever stands between them on the slide.
  const runs: Array<{ kind: FreeBlockKind; parts: number[] }> = [];
  // For each part to write, in order: whether it opens an array.
  const opens: boolean[] = [];
  for (const number of written) {
    const kind = blocks[number - 1] as FreeBlockKind;
    const last = runs.at(-1);
    opens.push(!last || last.kind !== kind);
    if (last && last.kind === kind) last.parts.push(number);
    else runs.push({ kind, parts: [number] });
  }

  // Where each call starts and ends, as places in the list of parts to write, from 1. A call ends
  // where the next part would not fit: by the length of its answer, or by the number of arrays
  // its schema would have. `share` adds that it ends where it has its share and others are to come.
  const cutting = (share?: { tokens: number; calls: number }): Array<[first: number, last: number]> => {
    const cuts: Array<[first: number, last: number]> = [];
    let first = 1;
    let sum = 0;
    let arrays = 0;
    rooms.forEach((room, index) => {
      const full = sum + room > most || (opens[index] === true && arrays >= MOST_ARRAYS);
      if (index + 1 > first && (full || (share !== undefined && sum >= share.tokens - 0.5 && cuts.length < share.calls - 1))) {
        cuts.push([first, index]);
        first = index + 1;
        sum = 0;
        arrays = 0;
      }
      // The first part of a call opens an array of that call, also where its run began in the call before.
      if (opens[index] === true || arrays === 0) arrays++;
      sum += room;
    });
    cuts.push([first, written.length]);
    return cuts;
  };
  const calls = cutting().length;
  const cuts = cutting({ tokens: rooms.reduce((total, room) => total + room, 0) / calls, calls });
  const digits = Math.max(2, String(runs.length).length);
  return cuts.map(([from, to], at) => {
    const parts = written.slice(from - 1, to);
    const here = new Set(parts);
    const held = runs.flatMap((run, index): PartRun[] => {
      const mine = run.parts.filter((number) => here.has(number));
      return mine.length > 0 ? [{ name: `r${String(index + 1).padStart(digits, '0')}`, kind: run.kind, parts: mine }] : [];
    });
    const tokens = rooms.slice(from - 1, to).reduce((total, room) => total + room, 0);
    return {
      parts,
      runs: held,
      title: at === 0 && title !== undefined,
      room: Math.min(WRITE_CALL_TOKENS, Math.ceil(tokens + overhead)),
      light: held.every((run) => run.kind === 'chart' || run.kind === 'table'),
    };
  });
}

/** A text as one line, for telling two texts apart: no stray spaces. */
const oneLine = (value: unknown): string => (typeof value === 'string' || Array.isArray(value) ? toParagraphs(value as string | string[]).join(' ').replace(/\s+/g, ' ').trim() : '');

/**
 * Whether the charts of a run each show one series, by the limit their slots state: a chart drawn
 * as a line or a few bars alone. The numbers of such a chart are one list, named by the chart's
 * caption when it is read back.
 *
 * A chart the slide holds already, and that is shown to be written again, may hold more than
 * that: several series, or a series with a name of its own. One list would lose them, with no
 * word said. A run with such a chart is asked for in the shape that has the series.
 */
function oneSeries(job: FreeJob, run: PartRun, byPart: ReturnType<typeof slotsByPart>): boolean {
  return run.parts.every((number) => {
    const slots = byPart.get(number) ?? [];
    return slots.every(({ planned }) => {
      if (planned.slot.kind !== 'chart') return true;
      if (planned.slot.maxSeries !== 1) return false;
      const now = job.current?.[planned.slot.name];
      if (!isChartValue(now) || now.series.length === 0) return true;
      const caption = slots.find((entry) => entry.part === 'Caption');
      return now.series.length === 1 && oneLine(now.series[0]?.name) === oneLine(caption ? job.current?.[caption.planned.slot.name] : undefined);
    });
  });
}

const BLOCK: JsonSchema = { type: 'integer', description: 'The number of the part this item is, counting every part of the slide from 1.' };
const strings = (description: string): JsonSchema => ({ type: 'array', items: { type: 'string' }, description });

/** Which of two slots of like parts holds less, by the limits they state: the first limit in which they differ decides. */
function tighter(a: PlannedSlot, b: PlannedSlot): boolean {
  const less = (...limits: Array<[mine: number | undefined, other: number | undefined]>): boolean => {
    const differs = limits.find(([mine, other]) => (mine ?? Infinity) !== (other ?? Infinity));
    return differs !== undefined && (differs[0] ?? Infinity) < (differs[1] ?? Infinity);
  };
  if (a.slot.kind === 'text' && b.slot.kind === 'text') return less([a.slot.maxChars, b.slot.maxChars], [a.slot.maxParagraphs, b.slot.maxParagraphs]);
  // The limits of a chart follow the size it is drawn at: a chart drawn without axes shows one series.
  if (a.slot.kind === 'chart' && b.slot.kind === 'chart') return less([a.slot.maxSeries, b.slot.maxSeries], [a.slot.maxCategories, b.slot.maxCategories]);
  if (a.slot.kind === 'table' && b.slot.kind === 'table') return less([a.slot.maxRows, b.slot.maxRows], [a.slot.maxColumns, b.slot.maxColumns], [a.slot.maxCellChars, b.slot.maxCellChars]);
  return false;
}

/** The lowest of a limit that like slots state, and no lower than what some part holds (`held`). `undefined` when none states it. */
function leastLimit(stated: ReadonlyArray<number | undefined>, held: readonly number[] = []): number | undefined {
  const limits = stated.filter((limit): limit is number => limit !== undefined);
  return limits.length > 0 ? Math.max(Math.min(...limits), ...held) : undefined;
}

/**
 * What is said of the like parts of a run, once for them all: the description of the slot that
 * holds least, and of each limit the lowest that any of them states. The writer reads one number
 * for the whole array, and what fits the smallest part fits every one.
 *
 * Where a deck is changed, a chart or a table the slide holds may be larger than its place reads
 * well at: it was written when the slide had fewer parts, or by code. It is shown to the writer
 * as it is, so the limit said is no lower than what a part of the run holds. A change that is not
 * about its numbers (another language) must not cost a chart a series for a limit it never met.
 */
function saidOf(job: FreeJob, like: readonly PlannedSlot[]): PlannedSlot {
  const least = like.reduce((best, planned) => (tighter(planned, best) ? planned : best));
  const held = like.map(({ slot }) => job.current?.[slot.name]);
  if (least.mode === 'chart') {
    const slots = like.flatMap(({ slot }) => (slot.kind === 'chart' ? [slot] : []));
    const charts = held.filter(isChartValue);
    const slot = { ...least.slot };
    const categories = leastLimit(slots.map((each) => each.maxCategories), charts.map((chart) => chart.categories.length));
    const series = leastLimit(slots.map((each) => each.maxSeries), charts.map((chart) => chart.series.length));
    if (categories !== undefined) slot.maxCategories = categories;
    if (series !== undefined) slot.maxSeries = series;
    return { ...least, slot };
  }
  if (least.mode === 'table') {
    const slots = like.flatMap(({ slot }) => (slot.kind === 'table' ? [slot] : []));
    const tables = held.filter(isTableValue);
    const slot = { ...least.slot };
    const columns = leastLimit(slots.map((each) => each.maxColumns), tables.map((table) => table.header.length));
    const rows = leastLimit(slots.map((each) => each.maxRows), tables.map((table) => table.rows.length));
    const cells = leastLimit(slots.map((each) => each.maxCellChars));
    if (columns !== undefined) slot.maxColumns = columns;
    if (rows !== undefined) slot.maxRows = rows;
    if (cells !== undefined) slot.maxCellChars = cells;
    return { ...least, slot };
  }
  // A text is told by the slot that holds fewest characters: its description says how much to write there.
  return least;
}

/**
 * The schema of one run of like parts: an array of items, each with its part's number and what
 * that kind of part holds. No array has a stated length, at any depth: a model's API refuses
 * nested arrays with limits. The number of items is said in the description and checked when
 * the answer is read. Every property is asked of every item, with "" or [] for none, except the
 * two that choose from a list of ids, an icon and a supplied photo: a list has no value that
 * means none, so an item leaves such a property out.
 *
 * What is said of a property is what its slot says: the slot's description, then its limits,
 * each once (see {@link saidOf} for the slot that speaks for the run). A chart or a table has no
 * property of its own in an item, so what its slot says stands in the description of the array.
 */
function runSchema(job: FreeJob, run: PartRun, byPart: ReturnType<typeof slotsByPart>, context: CatalogContext): JsonSchema {
  const like = new Map<string, PlannedSlot[]>();
  for (const number of run.parts) {
    for (const { part, planned } of byPart.get(number) ?? []) like.set(part, [...(like.get(part) ?? []), planned]);
  }
  const single = oneSeries(job, run, byPart);

  const properties: Record<string, JsonSchema> = { block: BLOCK };
  const optional = new Set<string>();
  const said: string[] = [];
  for (const [part, slots] of like) {
    const planned = saidOf(job, slots);
    const note = slotNote(planned.slot, context);
    const name = part.toLowerCase();
    switch (planned.mode) {
      case 'text':
        properties[name] = { type: 'string', description: planned.required ? note || 'Text' : `${note || 'Text'}. "" for none` };
        break;
      case 'paragraphs':
        properties[name] = strings(note || 'One string per paragraph');
        break;
      case 'icon':
        properties.icon = { type: 'string', enum: context.icons.map((icon) => icon.id), description: note || 'Id of the icon that best matches the item' };
        // Its description says "to every card of the slide, or to none", and a choice of ids has no value for none.
        optional.add('icon');
        break;
      case 'graphic':
        properties.asset = { type: 'string', enum: context.assets.filter((asset) => asset.kind === 'graphic').map((asset) => asset.id), description: note || 'Id of the supplied graphic to show here' };
        break;
      case 'photo': {
        properties.prompt = { type: 'string', description: PHOTO_PROMPT_NOTE };
        properties.alt = { type: 'string', description: PHOTO_ALT_NOTE };
        const photos = context.assets.filter((asset) => asset.kind === 'photo');
        if (photos.length > 0) {
          properties.asset = { type: 'string', enum: photos.map((asset) => asset.id), description: PHOTO_ASSET_NOTE };
          optional.add('asset');
        }
        if (note) said.push(note);
        break;
      }
      case 'table':
        properties.header = strings('Column headings, a word or two each');
        properties.rows = { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Body rows. Every row has one cell per column, in column order.' };
        if (note) said.push(note);
        break;
      case 'chart': {
        properties.type = { type: 'string', enum: ['column', 'bar', 'line', 'pie'], description: CHART_TYPE_NOTE };
        properties.categories = strings('Labels along the axis, or the names of the pie slices');
        properties.format = { type: 'string', enum: ['number', 'percent'], description: CHART_FORMAT_NOTE };
        // An item with one list of numbers has no series to count: the shape says that there is one.
        const { maxSeries: _one, ...listed } = planned.slot;
        const limits = single ? slotNote(listed, context) : note;
        if (single) {
          properties.values = { type: 'array', items: { type: 'number' }, description: 'One number per category, in category order' };
        } else {
          properties.series = {
            type: 'array',
            description: 'One or more series. A pie takes exactly one.',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string', description: 'What the series measures, a word or two' },
                values: { type: 'array', items: { type: 'number' }, description: 'One number per category, in category order' },
              },
              required: ['name', 'values'],
            },
          };
        }
        said.push(`${limits ? `${limits}. ` : ''}${CHART_RULE}`);
        break;
      }
    }
  }

  // In alphabetical order, so that a model that writes properties as declared and one that sorts them agree.
  const names = Object.keys(properties).sort();
  const number = run.parts.length;
  return {
    type: 'array',
    description: [`Exactly ${number} ${run.kind}${number === 1 ? '' : 's'}: ${partsWords(run.parts)}.`, ...said.map((note) => (/[.!?]$/.test(note) ? note : `${note}.`))].join(' '),
    items: {
      type: 'object',
      properties: Object.fromEntries(names.map((name) => [name, properties[name] as JsonSchema])),
      required: names.filter((name) => !optional.has(name)),
    },
  };
}

/** Whether the slide a job changes has a title now. Such a title stays unless a call gives another. */
export const titled = (job: Pick<WriteJob, 'current'>): boolean => oneLine(job.current?.title) !== '';

/**
 * Said of the title wherever a change offers the title of a slide that has one: with its parts,
 * by itself, in a call of its own or with the deck. Without it a writer that is asked for two
 * new charts gives the slide a new title too, or none.
 */
export const REPEAT_TITLE = 'Repeat the current title unless the request is about it';

/** The schema of one call: an array for each run of like parts it writes, and the slide's title in the slide's first call. */
export function rangeSchema(job: FreeJob, range: PartRange, context: CatalogContext): JsonSchema {
  const byPart = slotsByPart(job);
  const properties: Record<string, JsonSchema> = {};
  for (const run of range.runs) properties[run.name] = runSchema(job, run, byPart, context);
  if (range.title) {
    const title = job.pattern.slots.find((planned) => planned.slot.name === 'title');
    const note = title ? slotNote(title.slot, context) : '';
    // The property is asked of every answer. A slide that has a title is told to say it again: an
    // empty answer must not be how a call that adds two charts takes the title off the slide.
    const last = titled(job) ? REPEAT_TITLE : '"" for a slide without one';
    properties.title = { type: 'string', description: `The title of the slide${note ? `. ${note}` : ''}. ${last}` };
  }
  return { type: 'object', properties, required: Object.keys(properties) };
}

/**
 * What a call is asked to write, in the lines its prompt is built around. Where the slide is one
 * that is being changed, the content its parts have now is shown after them. `calls` is how many
 * calls write the slide: one of several is told that it is, since it sees its own parts alone.
 * And where like parts stand before its parts on the slide, the line of a run says which of
 * their kind its parts are: part 31 is the 31st chart only on a slide of nothing but charts, and
 * a call that takes "the charts" from the top of the brief would repeat the first call's.
 */
export function rangeTask(job: FreeJob, range: PartRange, calls = 1): string {
  const { blocks } = job.free;
  const first = range.parts[0];
  const earlier = new Set(calls > 1 && first !== undefined ? blocks.slice(0, first - 1) : []);
  // Which of its kind each part is, counted over the whole slide from 1.
  const seen = new Map<FreeBlockKind, number>();
  const places = blocks.map((kind) => {
    seen.set(kind, (seen.get(kind) ?? 0) + 1);
    return seen.get(kind) as number;
  });
  const lines = [
    `Write ${range.parts.length > 0 ? partsWords(range.parts) : 'the title'} of slide ${job.number} (free: ${describeBlocks(blocks)}): ${job.purpose}`,
    ...range.runs.map((run) => {
      const which = earlier.has(run.kind) ? ` (${placesWords(run.kind, run.parts.map((number) => places[number - 1] as number))})` : '';
      return `- ${run.name}: ${partsWords(run.parts)}, ${kindWords(run.kind, run.parts.length)}${which}`;
    }),
  ];
  if (range.runs.length > 0) lines.push('Part N is the item whose "block" is N. Where an item has nothing for a property, give "" or [].');
  if (calls > 1) {
    // Without it two calls may each show the brief's first things. In a change the parts are told by what they hold now.
    lines.push(`The other parts of the slide are written apart in the same way.${job.task ? '' : " Where the brief lists what the parts show, like parts take it in the brief's order."}`);
  }
  const current = rangeCurrent(job, range);
  if (current) lines.push(`Current content: ${JSON.stringify(current)}`);
  return lines.join('\n');
}

/** A slot's value as an item of an array states it: the properties the schema of its run asks for. */
function asItem(name: string, planned: PlannedSlot, value: SlotValue, single: boolean): Record<string, unknown> {
  switch (planned.mode) {
    case 'text':
      return typeof value === 'string' || Array.isArray(value) ? { [name]: toParagraphs(value).filter(Boolean).join(' ') } : {};
    case 'paragraphs':
      return typeof value === 'string' || Array.isArray(value) ? { [name]: toParagraphs(value).filter(Boolean) } : {};
    case 'icon':
      return isImageRef(value) ? { icon: value.asset } : {};
    case 'graphic':
      return isImageRef(value) ? { asset: value.asset } : {};
    case 'photo':
      // A picture that was generated is known by its description, a supplied one by its id.
      return !isImageRef(value) ? {} : value.prompt ? { alt: value.alt ?? '', prompt: value.prompt } : { alt: value.alt ?? '', asset: value.asset };
    case 'table':
      return isTableValue(value) ? { header: value.header, rows: value.rows } : {};
    case 'chart':
      if (!isChartValue(value)) return {};
      return { type: value.type, categories: value.categories, format: value.format ?? 'number', ...(single ? { values: value.series[0]?.values ?? [] } : { series: value.series }) };
  }
}

/**
 * The content the parts of a call have now, in the shape the call answers in: the same arrays,
 * and items with the same numbers. `undefined` where the job states none (`current`), as for a
 * slide that is new.
 */
function rangeCurrent(job: FreeJob, range: PartRange): Record<string, unknown> | undefined {
  const { current } = job;
  if (!current) return undefined;
  const byPart = slotsByPart(job);
  const shown: Record<string, unknown> = {};
  for (const run of range.runs) {
    const single = oneSeries(job, run, byPart);
    const items: Array<Record<string, unknown>> = [];
    for (const number of run.parts) {
      const item: Record<string, unknown> = { block: number };
      for (const { part, planned } of byPart.get(number) ?? []) {
        const value = current[planned.slot.name];
        if (value !== undefined && value !== null) Object.assign(item, asItem(part.toLowerCase(), planned, value, single));
      }
      // In the order the schema declares the properties.
      if (Object.keys(item).length > 1) items.push(Object.fromEntries(Object.entries(item).sort(([a], [b]) => (a < b ? -1 : 1))));
    }
    if (items.length > 0) shown[run.name] = items;
  }
  if (range.title && typeof current.title === 'string' && current.title.trim() !== '') shown.title = current.title;
  return Object.keys(shown).length > 0 ? shown : undefined;
}

/** What one call's answer holds. */
export interface RangeAnswer {
  /** The content of the call's parts by slot name, in the shape a slide written as properties has it. */
  values: Record<string, unknown>;
  /** What is wrong with the number of items: a line for each array that is short of items it has to hold. */
  counts: string[];
  /** The parts no item was given for. */
  absent: number[];
}

const some = (value: unknown): value is string => typeof value === 'string' && value.trim() !== '';
/** Whether an item says something for a text: a string, a number or a list, and not one that is empty. What it is made into is the reader's work. */
const said = (value: unknown): boolean => (typeof value === 'number' ? Number.isFinite(value) : Array.isArray(value) ? value.length > 0 : some(value));

type Item = Record<string, unknown>;
const record = (value: unknown): Item => (typeof value === 'object' && value !== null && !Array.isArray(value) ? (value as Item) : {});
const itemsOf = (list: unknown): Item[] => (Array.isArray(list) ? list : []).filter((item): item is Item => typeof item === 'object' && item !== null && !Array.isArray(item));

/** The number of the part an item says it is. A number written as a string is a number all the same. */
function blockOf(item: Item): number {
  if (typeof item.block === 'number') return item.block;
  return typeof item.block === 'string' && /^\d+$/.test(item.block.trim()) ? Number(item.block.trim()) : Number.NaN;
}

/**
 * Gives the items of an array to the parts of its run. An item goes to the part its `block`
 * names; one whose number is missing, repeated or not of the run takes the next part that has no
 * item, in order, and items beyond that are dropped. `taken` names parts that have an item
 * already, from an earlier answer: no item is given to them by its place alone.
 *
 * `stays` says which parts may stay as they are (see {@link staying}). In an array that is short
 * such a part is given an item only when the item names it: an item without a number of its own
 * could be any part, and must not write over one that was to keep what it has. An array with an
 * item for every part is taken in its order, as always.
 */
function placeItems(run: PartRun, items: Item[], taken: ReadonlySet<number> = new Set(), stays?: (part: number) => boolean): Map<number, Item> {
  const asked = new Set(run.parts);
  const placed = new Map<number, Item>();
  const unplaced: Item[] = [];
  for (const item of items) {
    const number = blockOf(item);
    if (asked.has(number) && !placed.has(number)) placed.set(number, item);
    else unplaced.push(item);
  }
  const short = items.length < run.parts.length;
  for (const number of run.parts) {
    if (placed.has(number) || taken.has(number) || (short && stays?.(number))) continue;
    const next = unplaced.shift();
    if (!next) break;
    placed.set(number, next);
  }
  return placed;
}

/**
 * A repair answer laid over the answer it mends. A model that is asked to put three items right
 * may return those three and nothing else: what it leaves out is then what it gave before, and
 * no part is lost for having been right the first time. An array that comes back whole is taken
 * as it is. The title of the answer before stays when the new one has none. Items are given
 * to parts as {@link readRange} gives them, so that the merged answer reads as the two did.
 */
export function mergeRange(job: FreeJob, range: PartRange, before: unknown, next: unknown): unknown {
  const old = record(before);
  const now = record(next);
  const stays = staying(job);
  // In the order of the answer before, which is the order the schema asks for.
  const merged: Item = { ...old, ...now };
  if (range.title && typeof now.title !== 'string' && typeof old.title === 'string') merged.title = old.title;
  for (const run of range.runs) {
    const items = itemsOf(now[run.name]);
    if (items.length >= run.parts.length) continue;
    const had = placeItems(run, itemsOf(old[run.name]), undefined, stays);
    const got = placeItems(run, items, new Set(had.keys()), stays);
    // Each item says which part it is, so that it is read as it was placed here.
    merged[run.name] = run.parts.flatMap((number) => {
      const item = got.get(number) ?? had.get(number);
      return item ? [{ ...item, block: number }] : [];
    });
  }
  return merged;
}

/**
 * Which parts may be left out of an answer: those the slide that is being changed has content
 * for, and that were not named outright to be written again (`due`). Such a part is offered with
 * what it holds, and keeps that where nothing comes for it. A new part has no content to keep,
 * and on a slide that is new no part has.
 */
function staying(job: FreeJob): (part: number) => boolean {
  const { current, due } = job;
  if (!current) return () => false;
  const byPart = slotsByPart(job);
  return (part) => {
    const names = (byPart.get(part) ?? []).map(({ planned }) => planned.slot.name);
    return names.some((name) => current[name] !== undefined) && !names.some((name) => due?.has(name));
  };
}

/**
 * Reads the answer of one call. An item goes to the part its `block` names; one whose number is
 * missing, repeated or not of its array takes the next part of the array that has none. Items
 * beyond the number asked for are dropped. An array that is short is a problem to send back,
 * unless the parts it leaves out are parts of a slide being changed that may stay as they are.
 */
export function readRange(job: FreeJob, range: PartRange, answer: unknown, context: CatalogContext): RangeAnswer {
  const raw = record(answer);
  const byPart = slotsByPart(job);
  const read: RangeAnswer = { values: {}, counts: [], absent: [] };
  // A slide without a title is given none by an empty one. A slide that has a title keeps it
  // through an empty answer: the call was told to repeat it, and a slip must not take it off.
  if (range.title && typeof raw.title === 'string' && (raw.title.trim() !== '' || !titled(job))) read.values.title = raw.title;

  const stays = staying(job);
  for (const run of range.runs) {
    const items = itemsOf(raw[run.name]);
    const placed = placeItems(run, items, undefined, stays);
    for (const number of run.parts) if (!placed.has(number)) read.absent.push(number);
    const needed = run.parts.length;
    if (items.length < needed) {
      // Where a deck is changed, a part that may stay can be left out of an answer: it keeps
      // what it had. The number of items is then a problem only for the parts that are new or
      // due, and they alone are asked for. An item that found no part (it names none, and no
      // part that has to be written was left for it) could be any of them: the array is then
      // asked for whole, as it is where every part without an item is needed.
      const owed = run.parts.filter((number) => !placed.has(number) && !stays(number));
      const lost = placed.size < items.length;
      const few = `${items.length} ${items.length === 1 ? 'item' : 'items'}`;
      if (lost || owed.length === needed - placed.size) read.counts.push(`s${job.number}.${run.name}: ${few}; it needs exactly ${needed}, ${partsWords(run.parts)}`);
      else if (owed.length > 0) read.counts.push(`s${job.number}.${run.name}: ${few}, and none for ${partsWords(owed)}; it needs ${owed.length === 1 ? 'that one' : 'those'} too`);
    }

    for (const [number, item] of placed) {
      for (const { part, planned } of byPart.get(number) ?? []) {
        const name = planned.slot.name;
        const given = item[part.toLowerCase()];
        switch (planned.mode) {
          // A value of another type than the schema asks for is passed on, as it is for a slide
          // written with the deck: a figure given as the number 87 is the text "87".
          case 'text':
          case 'paragraphs':
            if (said(given)) read.values[name] = given;
            break;
          case 'icon':
            if (some(item.icon)) read.values[name] = item.icon;
            break;
          case 'graphic':
            if (some(item.asset)) read.values[name] = item.asset;
            break;
          case 'photo':
            read.values[name] = { prompt: item.prompt, alt: item.alt, asset: item.asset };
            break;
          case 'table':
            read.values[name] = { header: item.header, rows: item.rows };
            break;
          case 'chart':
            // One list of numbers is one series, named by the chart's caption where it has one.
            read.values[name] = { type: item.type, categories: item.categories, format: item.format, series: Array.isArray(item.series) ? item.series : [{ name: some(item.caption) ? item.caption : '', values: item.values }] };
            break;
        }
      }
    }
  }
  return read;
}
