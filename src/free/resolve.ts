// The one place that says what pattern a slide is rendered and checked with: the manifest's, or,
// for a free slide, the one worked out for it.
import { isFreeSlide, isTableValue, isTextValue, plainText, toParagraphs } from '../deck/spec.ts';
import type { DeckSpec, FreeBlockKind, SlideSpec } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import type { PatternDef } from '../template/manifest.ts';
import { isFreeBlockKind } from './blocks.ts';
import { freeBaseRefusal, freeKitOrReason } from './kit.ts';
import type { FreeKit, FreeTextStyle } from './kit.ts';
import { freePlan } from './plan.ts';
import type { FreeContext, FreePlan } from './plan.ts';

/**
 * A free slide as it was worked out: what it is made with, and its plan. The rest is how the
 * checks and the renderer learn what to say and do: working data of the engine, which the
 * package does not export.
 */
export interface FreeResolved {
  kit: FreeKit;
  plan: FreePlan;
  /** The slide names a frame (`base`) that cannot carry a free slide and stands on another: why, as a sentence. */
  refused?: string;
  /**
   * The deck asks for larger text and the slide has no room for all of it: it is made with half
   * of what was asked, or with none. Left out when the slide has the text size the deck asks for.
   */
  enlarged?: 'half' | 'none';
  /**
   * For a deck that asks for larger text: the plan of the slide without that wish. Its sizes are
   * the least the texts of the slide are set at, also where the renderer sets a text smaller to
   * make it fit (the manifest's `shrinkText`): see `freeTextFloor` in deck/validate.ts.
   */
  plain?: FreePlan;
}

/** The size of every text of a plan, by the name of its shape or, for a chart and a table, of its slot. */
function textSizes(plan: FreePlan): Map<string, number> {
  const sizes = new Map<string, number>();
  for (const shape of plan.shapes) if (shape.kind === 'text' && shape.size !== undefined) sizes.set(shape.name, shape.size);
  for (const slot of plan.pattern.slots) {
    if (slot.kind === 'chart' && slot.text?.size !== undefined) sizes.set(slot.name, slot.text.size);
    if (slot.kind === 'table' && slot.bodyText?.size !== undefined) sizes.set(slot.name, slot.bodyText.size);
  }
  return sizes;
}

/**
 * What of a slide's content its plan has no room for, by slot: a text over the budget of its
 * slot, with that budget, and a table that its slot reports. The counts are the ones the deck's
 * check reports by.
 */
function unfitted(plan: FreePlan, slots: SlideSpec['slots'] | undefined): Map<string, { chars: number; paragraphs: number }> {
  const over = new Map<string, { chars: number; paragraphs: number }>();
  for (const slot of plan.pattern.slots) {
    const value = slots?.[slot.name];
    if (slot.kind === 'text' && isTextValue(value)) {
      const paragraphs = toParagraphs(value).filter((paragraph) => paragraph !== '');
      const chars = paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
      if (chars > (slot.maxChars ?? Infinity) || paragraphs.length > (slot.maxParagraphs ?? Infinity)) over.set(slot.name, { chars: slot.maxChars ?? Infinity, paragraphs: slot.maxParagraphs ?? Infinity });
    } else if (slot.kind === 'table' && isTableValue(value)) {
      let longest = 0;
      for (const row of [value.header, ...value.rows]) for (const cell of row) longest = Math.max(longest, plainText(cell).trim().length);
      if (value.rows.length > (slot.maxRows ?? Infinity) || value.header.length > (slot.maxColumns ?? Infinity) || longest > (slot.maxCellChars ?? Infinity)) over.set(slot.name, { chars: 0, paragraphs: 0 });
    }
  }
  return over;
}

/** Kinds that take the room there is, or bring sizes of their own: a slide with one keeps its sizes. */
const FILLING: ReadonlySet<FreeBlockKind> = new Set<FreeBlockKind>(['chart', 'table', 'photo', 'graphic']);
/** Running text, lists and quotations: a slide of these alone hangs from the top line, and has its own rules for a text alone. */
const PROSE: ReadonlySet<FreeBlockKind> = new Set<FreeBlockKind>(['text', 'list', 'quote']);
/** A slide is set larger only when its blocks take less than this share of the room's height. */
const GROW_UNDER = 2 / 3;
/** How much larger its text may be set, the largest first: the first that suits is taken. */
const GROW_STEPS: readonly number[] = [1.5, 1.4, 1.3, 1.2, 1.1];
/**
 * The texts of a slide grow together: the one that grows most grows no more than this many times
 * what the one that grows least does. A heading over a text that its block's width holds at its
 * size may be set a little larger than before, not half again as large over the same small text.
 */
const GROW_APART = 1.3;

/** The kit with every text style of the template so many times its size: what a slide that leaves room is planned with. */
function largerKit(kit: FreeKit, factor: number): FreeKit {
  const larger = (style: FreeTextStyle): FreeTextStyle => ({ ...style, set: { ...style.set, size: (style.set?.size ?? style.size) * factor } });
  const { note, ...styles } = kit.styles;
  return {
    ...kit,
    styles: { body: larger(styles.body), heading: larger(styles.heading), figure: larger(styles.figure), caption: larger(styles.caption), list: larger(styles.list), quote: larger(styles.quote), ...(note ? { note: larger(note) } : {}) },
    ...(kit.notePanel ? { notePanel: { ...kit.notePanel, text: larger(kit.notePanel.text) } } : {}),
  };
}

/** How much of the slide's height the blocks of a plan take, from the top of the highest to the foot of the lowest. */
const heightUsed = (plan: FreePlan): number => (plan.layout.blocks.length === 0 ? 0 : Math.max(...plan.layout.blocks.map((block) => block.box.y + block.box.h)) - Math.min(...plan.layout.blocks.map((block) => block.box.y)));
/** How the blocks of a plan are set out, as far as growing may not change it: so many on a line, in so many lines. */
const shapeOf = (plan: FreePlan): string => `${plan.layout.columns}|${new Set(plan.layout.blocks.filter((block) => block.kind !== 'note').map((block) => block.box.y)).size}|${plan.layout.tiles}`;
/** A plan that grew for want of room, or was set out in a way a slide with room is not: it is not grown further, nor taken as grown. */
const strained = (plan: FreePlan): boolean => plan.layout.dense || plan.layout.relaxed !== undefined || plan.layout.lower !== undefined;

/**
 * The plan of a slide whose blocks leave much of its room free, with its text set larger (from
 * 0.9.0): cards, steps, figures, phases, milestones, options and notes under a title, where three
 * of them at the template's own sizes are a low row over an empty half. Every text style is tried
 * at up to one and a half times its size, the largest first, and the first plan is taken in which
 * the blocks are set out as they were (as many on a line, in as many lines), no text is smaller
 * than it was, some text is larger, the texts grew together (`GROW_APART`), a text of one line is
 * still one line, and nothing that fitted is now over its box. The size of a
 * text still follows the width of its block (a line of a card holds its twenty characters), so a
 * row of many narrow blocks grows little or not at all, and the title stays the largest text:
 * the limits that hold a text under the title's size are the plan's own.
 *
 * A slide with a chart, a table or a picture keeps its sizes: these take the room there is. So
 * does a slide of running text, lists or quotations alone, which hangs from the top line as the
 * template's own text layouts do, and one whose deck says `arrange.grow: false`. A slide is planned
 * for a writer at the template's own sizes, so what a writer is told always fits.
 */
function grown(kit: FreeKit, slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'headings'>, context: FreeContext, plain: FreePlan): FreePlan {
  const kinds = slide.blocks as FreeBlockKind[];
  if (slide.arrange?.grow === false || slide.arrange?.dense === true) return plain;
  if (kinds.some((kind) => FILLING.has(kind)) || kinds.every((kind) => PROSE.has(kind))) return plain;
  if (strained(plain) || !(heightUsed(plain) < GROW_UNDER * (kit.band.bottom - kit.band.top))) return plain;
  const sizes = textSizes(plain);
  const over = unfitted(plain, slide.slots);
  const shape = shapeOf(plain);
  const several = new Map(plain.shapes.filter((entry) => entry.kind === 'text').map((entry) => [entry.name, entry.lines === 'several']));
  for (const factor of GROW_STEPS) {
    const plan = freePlan(largerKit(kit, factor), slide, context);
    if (strained(plan) || shapeOf(plan) !== shape) continue;
    const now = textSizes(plan);
    const ratios = [...sizes].filter(([, size]) => size > 0).map(([name, size]) => (now.get(name) ?? 0) / size);
    const least = Math.min(...ratios);
    const most = Math.max(...ratios);
    if (least < 1 - 1e-6) continue;
    // Nothing larger at this factor: the width of the blocks holds no more, and a smaller factor gives no more either.
    if (!(most > 1 + 1e-6)) return plain;
    if (most > GROW_APART * least + 1e-6) continue;
    // A text of one line stays one line: a heading that grew until it is counted in two lines has
    // a line kept free under it that PowerPoint, which sets it in one, leaves empty.
    if (plan.shapes.some((shape) => shape.kind === 'text' && shape.lines === 'several' && several.get(shape.name) === false)) continue;
    if ([...unfitted(plan, slide.slots).keys()].some((name) => !over.has(name))) continue;
    return plan;
  }
  return plain;
}

/**
 * The plan of a free slide: its pattern and the shapes to make. `template` is the template as it
 * was loaded, before a deck's own layouts are built onto it, so that a free slide stands on the
 * same frame whatever else the deck holds. A sentence instead when the template cannot carry
 * a free slide.
 */
export function freeFor(template: Template, deck: Pick<DeckSpec, 'direction' | 'textScale'>, slide: Pick<SlideSpec, 'blocks' | 'arrange' | 'slots' | 'base' | 'headings'>, planning?: FreeContext['planning']): FreeResolved | string {
  // Blocks that are no blocks give no slide to work out. The deck's check says what is wrong
  // with them; whoever asks without that check (the design checks, the pictures of a deck) is
  // told here, in a sentence, and passes the slide over.
  const blocks: unknown = slide.blocks;
  if (!Array.isArray(blocks) || blocks.length === 0 || !blocks.every(isFreeBlockKind)) return 'its "blocks" are not a list of kinds of block';
  if (blocks.length > 2000) return `it names ${blocks.length} blocks, and a free slide holds at most 2000`;
  const kit = freeKitOrReason(template, slide.base);
  if (typeof kit === 'string') return kit;
  const context: FreeContext = { direction: deck.direction ?? template.manifest.direction };
  // A number written as text is read as that number; null is a size that is not set.
  if (deck.textScale !== undefined && deck.textScale !== null) context.textScale = Number(deck.textScale);
  if (planning) context.planning = planning;
  const resolved: Omit<FreeResolved, 'plan'> = { kit };
  // A frame that is gone needs no reason; one that is there and cannot carry a free slide has one.
  if (slide.base !== undefined && kit.base !== slide.base && template.manifest.patterns.some((pattern) => pattern.id === slide.base)) {
    const refused = freeBaseRefusal(template, slide.base);
    if (refused !== undefined) resolved.refused = refused;
  }
  const asked = context.textScale ?? 1;
  if (planning) return { ...resolved, plan: freePlan(kit, slide, context) };
  if (!(asked > 1)) {
    const plain = freePlan(kit, slide, context);
    return { ...resolved, plan: grown(kit, slide, context, plain) };
  }
  // A deck that asks for larger text must not get smaller text for it. Larger text takes room
  // from its neighbours, and a part that is then fitted to what is left can end up under the
  // size it has without the wish. So the slide is also planned without the wish, and that plan
  // says the least every text may be. The slide takes all of what was asked, or half of it, or
  // none: the most at which no text is smaller. That least is counted as 0.6.6 counted
  // (`FreeContext.asThen`): a text that the count of headings run in, or of the template's own
  // width, sets larger without the wish would otherwise cost the rest of the slide the wish that
  // 0.6.6 gave it. Where none is given, the slide is set as without the wish, counted as now.
  //
  // How much of the wish it takes is chosen as 0.6.8 chose it, every plan counted at the
  // template's own spacing (`FreeContext.ownSpacing`): where spacing gives way (0.6.9), a text lifted
  // just to the least let the slide take the whole wish, and left that text, or a note beside it,
  // smaller than 0.6.8 set it. The share chosen is then made with spacing that gives way, no text
  // under the size it has in the plan chosen (`atLeast`): closer spacing only sets a text larger.
  const own: FreeContext = { ...context, ownSpacing: true };
  const plain = freePlan(kit, slide, { ...own, textScale: 1 });
  resolved.plain = plain;
  const least = textSizes(freePlan(kit, slide, { ...own, textScale: 1, asThen: true }));
  const smaller = (plan: FreePlan): boolean => [...textSizes(plan)].some(([name, size]) => size < (least.get(name) ?? 0) - 1e-6);
  const chosen = ((): { scale: number; plan: FreePlan; enlarged?: 'half' | 'none' } => {
    for (const [scale, enlarged] of [[asked, undefined], [1 + (asked - 1) / 2, 'half']] as const) {
      const plan = freePlan(kit, slide, { ...own, textScale: scale });
      if (!smaller(plan)) return { scale, plan, ...(enlarged ? { enlarged } : {}) };
      // Often all that is smaller is a half point lost to rounding, or to a block that got a
      // little narrower in the larger room a fuller slide takes. That must not cost the slide its
      // larger text: such a text is held at the size it has without the wish (`atLeast`), and
      // the slide is made so when nothing is then left with less room than it needs, measured
      // against the plan as it was and against the plan without the wish.
      const held = freePlan(kit, slide, { ...own, textScale: scale, atLeast: least });
      const before = [unfitted(plan, slide.slots), unfitted(plain, slide.slots)];
      const crowded = [...unfitted(held, slide.slots)].some(([name, has]) => !before.some((other) => {
        const had = other.get(name);
        return had !== undefined && had.chars <= has.chars && had.paragraphs <= has.paragraphs;
      }));
      if (!smaller(held) && !crowded) return { scale, plan: held, ...(enlarged ? { enlarged } : {}) };
    }
    return { scale: 1, plan: plain, enlarged: 'none' };
  })();
  const plan = freePlan(kit, slide, { ...context, textScale: chosen.scale, atLeast: textSizes(chosen.plan) });
  const made = { ...resolved, plan, ...(chosen.enlarged ? { enlarged: chosen.enlarged } : {}) };
  // A slide that leaves room has its text set larger without the wish too (`grown`, from 0.9.0).
  // A slide that does not grow is made as it always was. One that does grows from the size the
  // wish gave it, and a wish for larger text must not get smaller text than no wish does: where
  // the slide as it is without the wish still has some text larger, that slide is the one made.
  const unasked = { ...context, textScale: 1 };
  const base = freePlan(kit, slide, unasked);
  const without = grown(kit, slide, unasked, base);
  if (without === base) return made;
  const further = grown(kit, slide, { ...context, textScale: chosen.scale }, plan);
  const unaskedSizes = textSizes(without);
  if ([...textSizes(further)].some(([name, size]) => size < (unaskedSizes.get(name) ?? 0) - 1e-6)) return { ...resolved, plan: without };
  return { ...made, plan: further };
}

/**
 * The pattern of a slide: the one the manifest (with the deck's own layouts built in) has under
 * its id, or the one worked out for a free slide. `undefined` when there is none.
 */
export function slidePattern(original: Template, built: Template, deck: Pick<DeckSpec, 'direction' | 'textScale'>, slide: SlideSpec): PatternDef | undefined {
  if (isFreeSlide(slide)) {
    const free = freeFor(original, deck, slide);
    return typeof free === 'string' ? undefined : free.plan.pattern;
  }
  return built.manifest.patterns.find((pattern) => pattern.id === slide.pattern);
}
