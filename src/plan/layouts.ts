// Layouts made for one deck. After the outline, one light call looks for slides whose content none
// of the template's layouts shows well (three goals on a slide of running text, a plan in steps),
// and says what layout each would need. Those are composed in the template's own look, side by
// side, and given to the slides they were asked for. The designs go into the plan and then the
// deck; the template is not changed. A layout that cannot be made leaves its slides as they were.
import type { JsonSchema, LlmProvider } from '../llm/provider.ts';
import type { Template } from '../render/render.ts';
import { DesignError, composePattern, withLayouts } from '../template/compose.ts';
import type { PatternDesign } from '../template/compose.ts';
import type { PatternDef } from '../template/manifest.ts';
import type { UsablePattern } from './catalog.ts';
import { isPlainText } from './catalog.ts';
import type { DeckPlan, OutlineEntry } from './plan.ts';

/** A layout made for a deck, and the slides that use it. */
export interface MadeLayout {
  pattern: string;
  /** Slide numbers in the plan, from 1. */
  slides: number[];
  /** What was made, in a sentence or two. */
  summary: string;
}

/** A layout that was asked for and could not be made. Its slides keep the layout they had. */
export interface FailedLayout {
  /** What the layout was to be, as the deciding call wrote it. Empty when the call itself failed. */
  instruction: string;
  slides: number[];
  reason: string;
}

export interface ComposeLayoutsResult {
  /** The plan with `layouts`, and the slides they were made for using them. */
  plan: DeckPlan;
  made: MadeLayout[];
  failed: FailedLayout[];
}

interface Ask {
  instruction: string;
  slides: number[];
}

const SYSTEM =
  'You review the plan of a slide deck against the layouts its template offers, and say where a new layout, made in the template\'s own style, would show a slide\'s content better. Answer only with JSON.';

/**
 * Whether a slide with this layout may get a new one: a slide of content, with running text or a
 * row of items. A cover, a section divider, a quotation or a closing slide keeps its own: a new
 * layout on its frame would carry its background and lose its point.
 */
export function holdsContent(pattern: PatternDef | undefined): boolean {
  if (!pattern) return false;
  if (pattern.repeat) return true;
  return pattern.slots.some((slot) => slot.kind === 'text' && !slot.deckField && slot.empty !== 'drop' && slot.empty !== 'keep' && ((slot.maxParagraphs ?? 1) > 1 || (slot.maxChars ?? 0) >= 200));
}

function decisionPrompt(plan: DeckPlan, offered: UsablePattern[], max: number, candidates: number[]): string {
  const { brief } = plan;
  const counts = new Map<string, number>();
  for (const slide of plan.slides) counts.set(slide.pattern, (counts.get(slide.pattern) ?? 0) + 1);
  const repeated = [...counts].filter(([, count]) => count >= 3).map(([id, count]) => `${id} on ${count} slides`);
  return [
    `The request:\n${brief.prompt.trim()}${brief.instructions?.trim() ? `\n${brief.instructions.trim()}` : ''}`,
    `The deck, slide by slide (layout: what the slide says):\n${plan.slides.map((slide, index) => `${index + 1}. ${slide.pattern}${slide.items !== undefined ? ` (${slide.items} items)` : ''}: ${slide.purpose}`).join('\n')}`,
    repeated.length > 0 ? `The same layout is used for many slides: ${repeated.join(', ')}.` : '',
    `The layouts the template offers for this deck:\n${offered.map((pattern) => `- ${pattern.def.id}: ${pattern.def.description.slice(0, 200)}${isPlainText(pattern.def) ? ' Plain text.' : ''}`).join('\n')}`,
    [
      'Find the slides whose content has a shape that none of these layouts shows well: several points or goals of equal weight, the steps of a process, a comparison of two or three options, an agenda, a timeline, one statement or quote, a few key figures the request gives.',
      `- Name at most ${max} new layout${max === 1 ? '' : 's'}. A layout may serve several slides whose content has the same shape.`,
      `- Only slides ${candidates.join(', ')} may get one: the others open, divide or close the deck.`,
      '- Not a slide that one of the listed layouts already shows well. Running text with a title is shown well by a layout for running text.',
      '- A deck that uses one layout for many slides gains most from a new one.',
      '- "instruction" tells a designer what the layout holds: its parts, and how many items, such as "three cards side by side, each with a heading and two lines of text". Text and the template\'s drawn shapes only: no pictures, icons, tables or charts.',
      '- Ask for figures (85%, 2.4M) only when the request gives them. The number of a step is not a figure.',
      '- Answer with an empty list when the listed layouts suit every slide.',
    ].join('\n'),
  ]
    .filter(Boolean)
    .join('\n\n');
}

const decisionSchema: JsonSchema = {
  type: 'object',
  properties: {
    layouts: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          instruction: { type: 'string', description: 'What the new layout holds, for the designer, in English' },
          slides: { type: 'array', items: { type: 'integer' }, description: 'Numbers of the slides that would use it' },
        },
        required: ['instruction', 'slides'],
      },
    },
  },
  required: ['layouts'],
};

/** The layouts asked for, cleaned: slides that may get one, each slide once. */
function readAsks(answer: unknown, candidates: number[], max: number): Ask[] {
  const list = (answer as { layouts?: unknown } | null)?.layouts;
  const taken = new Set<number>();
  const asks: Ask[] = [];
  for (const item of Array.isArray(list) ? list : []) {
    const raw = item as { instruction?: unknown; slides?: unknown };
    const instruction = typeof raw?.instruction === 'string' ? raw.instruction.trim() : '';
    const slides = (Array.isArray(raw?.slides) ? raw.slides : [])
      .map((value) => (typeof value === 'number' ? Math.round(value) : Number.NaN))
      .filter((number) => candidates.includes(number) && !taken.has(number));
    if (!instruction || slides.length === 0) continue;
    for (const number of slides) taken.add(number);
    asks.push({ instruction, slides: [...new Set(slides)].sort((a, b) => a - b) });
    if (asks.length >= max) break;
  }
  return asks;
}

/** An id the template and the other layouts do not have. */
function freeId(id: string, taken: Set<string>): string {
  let free = id;
  for (let n = 2; taken.has(free); n++) free = `${id}-${n}`;
  taken.add(free);
  return free;
}

const isAbort = (error: unknown, signal?: AbortSignal) => signal?.aborted === true || (error instanceof Error && error.name === 'AbortError');
const reasonOf = (error: unknown) => (error instanceof DesignError ? error.problems.join(' ') : error instanceof Error ? error.message : String(error));

/**
 * Looks at a plan for slides that a layout the template lacks would show better, composes those
 * layouts and gives them to their slides. One light model call decides, then one or two per layout
 * compose, side by side. `offered` is what the outline could choose from.
 */
export async function composeForPlan(
  template: Template,
  plan: DeckPlan,
  offered: UsablePattern[],
  options: { llm: LlmProvider; max: number; signal?: AbortSignal },
): Promise<ComposeLayoutsResult> {
  const { llm, signal } = options;
  const max = Math.max(0, Math.floor(options.max));
  const failed: FailedLayout[] = [];
  const patterns = new Map(template.manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const last = plan.slides.length;
  const candidates = plan.slides.map((slide, index) => index + 1).filter((number) => number > 1 && number < last && holdsContent(patterns.get((plan.slides[number - 1] as OutlineEntry).pattern)));
  if (max === 0 || candidates.length === 0) return { plan, made: [], failed };

  let asks: Ask[];
  try {
    const answer = await llm.generateJson({ system: SYSTEM, prompt: decisionPrompt(plan, offered, max, candidates), schema: decisionSchema, maxOutputTokens: 1500, effort: 'low', ...(signal ? { signal } : {}) });
    asks = readAsks(answer, candidates, max);
  } catch (error) {
    if (isAbort(error, signal)) throw error;
    return { plan, made: [], failed: [{ instruction: '', slides: [], reason: reasonOf(error) }] };
  }
  if (asks.length === 0) return { plan, made: [], failed };

  // Each stands on the frame of the layout its first slide had, so the title stays where the deck has it.
  const settled = await Promise.allSettled(
    asks.map((ask) => composePattern(template, ask.instruction, { llm, effort: 'low', base: (plan.slides[(ask.slides[0] as number) - 1] as OutlineEntry).pattern, ...(signal ? { signal } : {}) })),
  );
  if (signal?.aborted) throw signal.reason ?? new Error('aborted');

  const prior = plan.layouts ?? [];
  const ids = new Set([...template.manifest.patterns, ...prior].map((pattern) => pattern.id));
  const designs: PatternDesign[] = [];
  const made: MadeLayout[] = [];
  settled.forEach((result, index) => {
    const ask = asks[index] as Ask;
    if (result.status === 'rejected') return void failed.push({ ...ask, reason: reasonOf(result.reason) });
    const design = { ...result.value.design, id: freeId(result.value.design.id, ids) };
    // Built here once more on top of the others, as rendering will build it.
    try {
      withLayouts(template, [...prior, ...designs, design]);
    } catch (error) {
      return void failed.push({ ...ask, reason: reasonOf(error) });
    }
    designs.push(design);
    // The slide number in the summary is the template's; a deck's layout has none of its own.
    const summary = result.value.summary.replaceAll(`"${result.value.design.id}"`, `"${design.id}"`).replace(/ is a new layout on slide \d+,/, ' is a new layout');
    made.push({ pattern: design.id, slides: ask.slides, summary });
  });
  if (designs.length === 0) return { plan, made, failed };

  const slideLayout = new Map(made.flatMap((layout) => layout.slides.map((number) => [number, layout.pattern] as const)));
  const slides = plan.slides.map((entry, index): OutlineEntry => {
    const pattern = slideLayout.get(index + 1);
    // A row made for the slide holds as many items as it was drawn with.
    return pattern ? { pattern, purpose: entry.purpose } : entry;
  });
  return { plan: { ...plan, slides, layouts: [...prior, ...designs] }, made, failed };
}

/** The layouts a list of slides uses, of those given. */
export function usedLayouts(layouts: PatternDesign[] | undefined, slides: Array<{ pattern: string }>): PatternDesign[] | undefined {
  const used = (layouts ?? []).filter((design) => slides.some((slide) => slide.pattern === design.id));
  return used.length > 0 ? used : undefined;
}
