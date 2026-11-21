// The one place that says what pattern a slide is rendered and checked with: the manifest's, or,
// for a free slide, the one worked out for it.
import { isFreeSlide, isTableValue, isTextValue, plainText, toParagraphs } from '../deck/spec.ts';
import type { DeckSpec, SlideSpec } from '../deck/spec.ts';
import type { Template } from '../render/render.ts';
import type { PatternDef } from '../template/manifest.ts';
import { isFreeBlockKind } from './blocks.ts';
import { freeBaseRefusal, freeKitOrReason } from './kit.ts';
import type { FreeKit } from './kit.ts';
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
  if (planning || !(asked > 1)) return { ...resolved, plan: freePlan(kit, slide, context) };
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
  return { ...resolved, plan, ...(chosen.enlarged ? { enlarged: chosen.enlarged } : {}) };
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
