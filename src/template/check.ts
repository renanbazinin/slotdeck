import type { Template } from '../render/render.ts';
import { measurerFor } from '../text/measure.ts';
import { inspectTemplate } from './inspect.ts';
import type { ShapeInfo, TemplateInfo } from './inspect.ts';
import { effectiveSlots, textBudget, widenFactor } from './manifest.ts';
import type { SlotDef } from './manifest.ts';
import { filler } from './sample.ts';

export interface ManifestIssue {
  severity: 'error' | 'warning';
  /** Pattern id; absent for manifest-level issues. */
  pattern?: string;
  slot?: string;
  message: string;
}

/**
 * Checks a manifest against the pptx it describes: every pattern's slide, every slot's shape
 * and every row item must exist and be of the right kind. Run it when a template or manifest
 * is loaded, so a template that was changed shows up here and not in the middle of a render.
 *
 * For a template loaded with fonts it also measures every text slot's budget: sample text of
 * the budget's length that would run over its box is reported, with the length that fits.
 */
export function checkManifest(template: Template, inspected?: TemplateInfo): ManifestIssue[] {
  const { manifest } = template;
  const info = inspected ?? inspectTemplate(template.pkg);
  const issues: ManifestIssue[] = [];
  const ids = new Set<string>();
  const measure = measurerFor(template);

  for (const pattern of manifest.patterns) {
    const report = (severity: ManifestIssue['severity'], message: string, slot?: string) => {
      const issue: ManifestIssue = { severity, pattern: pattern.id, message };
      if (slot) issue.slot = slot;
      issues.push(issue);
    };
    if (ids.has(pattern.id)) report('error', 'another pattern has the same id');
    ids.add(pattern.id);

    const slide = info.slides[pattern.sourceSlide - 1];
    if (!slide) {
      report('error', `points at slide ${pattern.sourceSlide}, but the template has ${info.slides.length} slides`);
      continue;
    }
    const shapes = new Map<string, ShapeInfo>(slide.shapes.map((shape) => [shape.id, shape]));
    const names = new Set<string>();
    const usedShapes = new Map<string, string>();

    // With fonts, a budget is tried out: the slot's own sample words, as many as the budget allows.
    // `items` tries the budget a slot of a row has when the slide holds that many items.
    const tryBudget = (slot: SlotDef, shape: ShapeInfo, items?: number) => {
      if (!measure || slot.kind !== 'text' || shape.kind !== 'text' || slot.maxChars === undefined || slot.deckField || slot.empty === 'drop' || slot.empty === 'keep') return;
      const scale = manifest.textScale ?? 1;
      const budget = textBudget(slot, scale);
      const words = (shape.paragraphs ?? []).join(' ').split(/\s+/).filter(Boolean);
      const paragraphs = (slot.maxParagraphs ?? 1) > 1 ? (budget.maxParagraphs ?? 2) : 1;
      const each = Math.max(1, Math.floor((budget.maxChars ?? 1) / paragraphs));
      const text = Array.from({ length: paragraphs }, (_, index) => filler(words, each, index));
      const fit = measure.fit(pattern, slot, text, items === undefined ? { textScale: scale } : { textScale: scale, items });
      if (fit && !fit.fits) {
        const what = fit.overflow > 0 ? `takes ${fit.lines} lines where ${fit.capacity} fit` : `is ${Math.round((fit.widthRatio - 1) * 100)}% wider than the box`;
        const when = items === undefined ? '' : ` on a slide with ${items} of its ${pattern.repeat?.shapes.length ?? items} items`;
        report('warning', `has a budget of ${budget.maxChars} characters${when}, but sample text of that length ${what}; about ${fit.fitsChars} characters fit`, slot.name);
      }
    };

    for (const slot of pattern.slots) {
      if (names.has(slot.name)) report('error', 'another slot of the pattern has the same name', slot.name);
      names.add(slot.name);
      const shape = shapes.get(slot.shapeId);
      if (!shape) {
        report('error', `points at shape #${slot.shapeId}, which is not on slide ${pattern.sourceSlide}`, slot.name);
        continue;
      }
      const other = usedShapes.get(slot.shapeId);
      if (other) report('warning', `points at the same shape as "${other}"`, slot.name);
      usedShapes.set(slot.shapeId, slot.name);

      const grouped = shape.groupId !== undefined;
      // A slot that only ever removes its shape may point at any kind of shape, and so may one
      // that keeps the shape as designed and only shows it in another place.
      const dropped = slot.empty === 'drop' || (slot.empty === 'keep' && 'box' in slot && slot.box !== undefined);
      if (slot.kind === 'text' && shape.kind !== 'text' && !dropped) {
        report(shape.kind === 'shape' ? 'warning' : 'error', `is a text slot, but shape #${slot.shapeId} is ${describe(shape)}`, slot.name);
      }
      if (slot.kind === 'image' && shape.kind !== 'picture' && !dropped) report('error', `is an image slot, but shape #${slot.shapeId} is ${describe(shape)}`, slot.name);
      if (slot.kind === 'table' && shape.kind !== 'table') report('error', `is a table slot, but shape #${slot.shapeId} is ${describe(shape)}`, slot.name);
      if (slot.kind === 'chart' && grouped) report('error', 'is a chart slot, but its shape is inside a group', slot.name);
      if ((slot.kind === 'text' || slot.kind === 'image') && slot.box && grouped) report('error', 'has a box, but its shape is inside a group', slot.name);
      if (slot.kind === 'text' && slot.maxChars === undefined && !slot.deckField && !dropped) {
        report('warning', 'has no character budget, so nothing stops text that does not fit', slot.name);
      }
      tryBudget(slot, shape);
    }

    const repeat = pattern.repeat;
    if (repeat) {
      const drawn = repeat.shapes.length;
      if (drawn === 0) report('error', 'has a repeat with no items');
      for (let n = 1; n <= drawn; n++) {
        for (const name of repeat.slots) {
          if (!names.has(name.replace('{n}', String(n)))) report('error', `repeat expects a slot named "${name.replace('{n}', String(n))}"`);
        }
      }
      const parents = new Set<string>();
      repeat.shapes.forEach((item, index) => {
        for (const id of item) {
          const shape = shapes.get(id);
          if (!shape) report('error', `repeat item ${index + 1} points at shape #${id}, which is not on slide ${pattern.sourceSlide}`);
          else parents.add(shape.groupId ?? '');
        }
      });
      if (parents.size > 1) report('error', 'repeat items must all sit directly on the slide, or all in the same group');
      const min = repeat.min ?? drawn;
      const max = repeat.max ?? drawn;
      if (min < 1 || min > max) report('error', `repeat range ${min} to ${max} is not valid`);
      if (repeat.columns !== undefined && max > drawn) report('warning', 'a grid cannot hold more items than it draws');
      if (repeat.widen !== undefined && !(typeof repeat.widen === 'number' && repeat.widen >= 1)) report('error', 'repeat.widen is a factor of 1 or more, such as 1.5');
      // A row that widens gives its texts a larger budget when it holds fewer items. With fonts,
      // that budget is tried out in the wider box, on the first item: the items of a row are alike.
      if (measure && min >= 1 && min <= max) {
        for (let count = min; count < drawn; count++) {
          if (!(widenFactor(repeat, count) > 1)) continue;
          const slots = effectiveSlots(pattern, count);
          for (const name of repeat.slots) {
            const slot = slots.find((entry) => entry.name === name.replace('{n}', '1'));
            const shape = slot && shapes.get(slot.shapeId);
            if (slot && shape) tryBudget(slot, shape, count);
          }
        }
      }
    }
  }
  if (measure) {
    const missing = measure.missingFonts().map((font) => `${font.typeface}${font.bold ? ' bold' : ''}${font.italic ? ' italic' : ''}`);
    if (missing.length > 0) {
      issues.push({ severity: 'warning', message: `no font was supplied for ${missing.join(', ')}; text set in ${missing.length === 1 ? 'it' : 'them'} is checked by its character budget` });
    }
  }
  return issues;
}

function describe(shape: ShapeInfo): string {
  return shape.kind === 'picture' ? 'a picture' : shape.kind === 'table' ? 'a table' : shape.kind === 'text' ? 'a text box' : `a ${shape.kind} without text`;
}
