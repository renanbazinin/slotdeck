import type { TemplateManifest } from '../template/manifest.ts';
import { freeSlots, isFreeBlockKind } from '../free/blocks.ts';
import { isChartValue, isFreeSlide, knownValues } from './spec.ts';
import type { DeckSpec } from './spec.ts';
import type { Issue } from './validate.ts';
import { said } from '../messages/en.ts';

/**
 * The numbers a text states. "3K" counts as both 3 and 3000, "1,200" as 1200, "45%" as 45.
 * Only Western digits are read.
 */
export function statedNumbers(text: string): Set<number> {
  const numbers = new Set<number>();
  for (const match of text.matchAll(/(\d[\d,]*(?:\.\d+)?)\s*([KkMm])?/g)) {
    const value = Number((match[1] as string).replace(/,/g, ''));
    if (!Number.isFinite(value)) continue;
    numbers.add(value);
    const suffix = match[2]?.toLowerCase();
    if (suffix === 'k') numbers.add(value * 1000);
    if (suffix === 'm') numbers.add(value * 1_000_000);
  }
  return numbers;
}

/**
 * Chart values that the deck's brief, or the document it was made from, does not state. A model
 * asked for a chart can fill it with plausible numbers; this flags them so a person can check.
 * Values worked out from the brief (a sum, a growth rate applied) are flagged too: the check
 * cannot tell them from invented ones.
 */
export function unstatedChartValues(manifest: TemplateManifest, deck: DeckSpec): Issue[] {
  // A deck made from a document may chart the figures the document states.
  const stated = statedNumbers(deck.document?.trim() ? `${deck.brief ?? ''}\n${deck.document}` : (deck.brief ?? ''));
  const issues: Issue[] = [];
  deck.slides.forEach((slide, index) => {
    const pattern = manifest.patterns.find((candidate) => candidate.id === slide.pattern);
    // A free slide names its charts itself: block 3 of kind chart is the slot b3Chart.
    const charts = isFreeSlide(slide) ? freeSlots((slide.blocks ?? []).filter(isFreeBlockKind)).filter((entry) => entry.part.holds === 'chart').map((entry) => ({ name: entry.name, kind: 'chart' as const })) : (pattern?.slots ?? []);
    for (const slot of charts) {
      const value = slide.slots[slot.name];
      if (slot.kind !== 'chart' || !isChartValue(value)) continue;
      const unknown = [...new Set(knownValues(value))].filter((number) => !stated.has(number));
      if (unknown.length > 0) {
        issues.push({
          severity: 'warning',
          slide: index + 1,
          slot: slot.name,
          ...said('chart-values-unstated', { numbers: unknown.join(', '), count: unknown.length }),
        });
      }
    }
  });
  return issues;
}
