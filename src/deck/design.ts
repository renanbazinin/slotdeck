// Design problems of a deck that code can see without a picture: every one of them follows from
// how the deck's text measures in the template's boxes.
import type { Template } from '../render/render.ts';
import { withLayouts } from '../template/compose.ts';
import { effectiveSlots, itemCount, rowLayout, slotSources } from '../template/manifest.ts';
import type { PatternDef, TextSlotDef } from '../template/manifest.ts';
import { textMeasurer } from '../text/measure.ts';
import type { TextFit } from '../text/measure.ts';
import type { DeckSpec } from './spec.ts';
import { fitOptions, overflowMessage, resolveSlot } from './validate.ts';
import type { Issue, LayoutMode } from './validate.ts';

/**
 * - `overflow`: the text runs over its box.
 * - `empty`: the text takes far less room than the template's own text does, so its box looks empty.
 * - `uneven`: one item of a row of cards holds far less text than another.
 * - `orphan`: a heading ends with a single word on a line of its own.
 */
export type DesignKind = 'overflow' | 'empty' | 'uneven' | 'orphan';

/** A design problem on one slot of one slide. Its message is written so that it can be sent to the model that wrote the text. */
export interface DesignIssue extends Issue {
  severity: 'warning';
  slide: number;
  slot: string;
  kind: DesignKind;
}

export interface DesignOptions {
  /** As for `renderDeck`. Defaults to `strict`. */
  layout?: LayoutMode;
  /** Which problems to look for. Defaults to all of them. */
  kinds?: DesignKind[];
  /**
   * A text leaves its box empty when it takes this share, or less, of the lines the template's own
   * text takes there; a card is uneven when it takes this share or less of its fullest neighbour.
   * Defaults to a third.
   */
  emptyBelow?: number;
  /** The font size in points from which a single paragraph counts as a heading. Defaults to 20. */
  headingSize?: number;
}

interface Measured {
  /** The slot name the deck used. */
  source: string;
  slot: TextSlotDef;
  fit: TextFit;
  chars: number;
  paragraphs: number;
}

/** The slot names of each group of like slots across a row of items: every card's text, every card's heading. */
function rowGroups(pattern: PatternDef, items: number | undefined): string[][] {
  const groups: string[][] = [];
  const rows: string[][][] = [...(pattern.sequences ?? [])];
  if (pattern.repeat) {
    const repeat = pattern.repeat;
    let next = 1;
    for (const size of rowLayout(repeat, itemCount(pattern, items))) {
      rows.push(
        Array.from({ length: size }, () => {
          const n = next++;
          return repeat.slots.map((name) => name.replace('{n}', String(n)));
        }),
      );
    }
  }
  for (const row of rows) {
    const width = Math.max(0, ...row.map((item) => item.length));
    for (let position = 0; position < width; position++) {
      const names = row.map((item) => item[position]).filter((name): name is string => name !== undefined);
      if (names.length > 1) groups.push(names);
    }
  }
  return groups;
}

const quote = (text: string) => JSON.stringify(text.length > 30 ? `${text.slice(0, 29)}…` : text);
const lines = (count: number) => `${count} ${count === 1 ? 'line' : 'lines'}`;

/**
 * Lists the design problems of a deck that follow from measurement alone: text that runs over its
 * box, boxes and cards left mostly empty, and headings that end in a lone word. The template must
 * be loaded with fonts. No picture is made and no model is called.
 *
 * A slot that cannot be measured (its typeface was not supplied) is passed over, as is a slide
 * that cannot be rendered: `validateDeck` reports those.
 */
export function checkDesign(template: Template, deck: DeckSpec, options: DesignOptions = {}): DesignIssue[] {
  template = withLayouts(template, deck.layouts);
  const measure = textMeasurer(template);
  const { manifest } = template;
  const layout = options.layout ?? 'strict';
  const kinds = new Set<DesignKind>(options.kinds ?? ['overflow', 'empty', 'uneven', 'orphan']);
  const emptyBelow = options.emptyBelow ?? 1 / 3;
  const headingSize = options.headingSize ?? 20;
  const mirrored = (deck.direction ?? manifest.direction) !== manifest.direction;
  const issues: DesignIssue[] = [];

  (deck.slides ?? []).forEach((slide, index) => {
    const number = index + 1;
    const pattern = manifest.patterns.find((entry) => entry.id === slide.pattern);
    if (!pattern) return;
    const report = (kind: DesignKind, slot: string, message: string) => {
      if (kinds.has(kind)) issues.push({ severity: 'warning', slide: number, slot, kind, message });
    };
    const sources = slotSources(pattern, mirrored, slide.items);
    const measured = new Map<string, Measured>();

    for (const slot of effectiveSlots(pattern, slide.items)) {
      if (slot.kind !== 'text') continue;
      const source = sources.get(slot.name) ?? slot.name;
      const resolved = resolveSlot(slot, { ...slide, slots: slide.slots ?? {} }, deck, layout, source);
      if ('error' in resolved || resolved.action.type !== 'text') continue;
      const { paragraphs } = resolved.action;
      const fit = measure.fit(pattern, slot, paragraphs, fitOptions(manifest, deck, slide));
      if (!fit) continue;
      const entry: Measured = { source, slot, fit, chars: paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0), paragraphs: paragraphs.length };
      measured.set(source, entry);

      const over = overflowMessage(fit);
      if (over) {
        report('overflow', source, over);
        continue; // the other checks are about text that fits
      }

      // A heading whose last line holds one word. The text that would end one line sooner is the
      // text without that word. An address on a line of its own is meant to stand there.
      const last = fit.paragraphs[0]?.at(-1)?.trim() ?? '';
      const lastWidth = fit.widths[0]?.at(-1) ?? 0;
      const widest = Math.max(...(fit.widths[0] ?? [0]));
      const lone = last !== '' && !/\s/.test(last) && !/[@/]|\.\p{L}/u.test(last);
      if (paragraphs.length === 1 && fit.size >= headingSize && fit.lines >= 2 && lone && lastWidth <= widest * 0.4) {
        const sooner = entry.chars - last.length - 1;
        report('orphan', source, `the last of its ${lines(fit.lines)} holds one word, ${quote(last)}. Text of about ${sooner} characters ends a line sooner; a few words more fill the line.`);
      }

      // A box the template fills with several lines, holding a fraction of that. Where the
      // manifest moved or resized the box there is no text of the template's to compare with,
      // and the box is taken to be meant half full. A deck field is the app's text, not the model's.
      const designed = fit.sampleLines > 0 ? fit.sampleLines : Math.ceil(fit.capacity / 2);
      if (!slot.deckField && designed >= 3 && fit.lines <= Math.max(1, Math.floor(designed * emptyBelow))) {
        const wanted = Math.round((entry.chars * designed) / fit.lines);
        const against = fit.sampleLines > 0 ? `where the template's own text takes ${fit.sampleLines}` : `in a box that holds ${fit.capacity}`;
        report('empty', source, `text takes ${lines(fit.lines)} ${against}, so its box will look empty. Text of about ${wanted} characters fills it as designed.`);
      }
    }

    // Cards of one row: a card with a fraction of its fullest neighbour's text.
    for (const names of rowGroups(pattern, slide.items)) {
      const row = names.map((name) => measured.get(sources.get(name) ?? name)).filter((entry): entry is Measured => entry !== undefined && entry.fit.fits);
      const fullest = row.reduce<Measured | undefined>((most, entry) => (!most || entry.fit.lines > most.fit.lines ? entry : most), undefined);
      if (!fullest || fullest.fit.lines < 3) continue;
      for (const entry of row) {
        if (entry === fullest || entry.fit.lines > Math.max(1, Math.floor(fullest.fit.lines * emptyBelow))) continue;
        const wanted = Math.round((fullest.chars * 2) / 3);
        report('uneven', entry.source, `text takes ${lines(entry.fit.lines)} where "${fullest.source}" beside it takes ${fullest.fit.lines}, so the row will look uneven. Text of about ${wanted} characters or more evens it out.`);
      }
    }
  });

  return issues;
}
