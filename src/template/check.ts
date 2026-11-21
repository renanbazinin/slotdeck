import { parseHex } from '../color/space.ts';
import { FREE_STYLE_ROLES, freeBaseRefusal, freeKitOrReason } from '../free/kit.ts';
import type { Template } from '../render/render.ts';
import { measurerFor } from '../text/measure.ts';
import { inspectTemplate } from './inspect.ts';
import type { ShapeInfo, TemplateInfo } from './inspect.ts';
import { effectiveSlots, templateScaleProblem, templateTextProblem, textBudget, widenFactor } from './manifest.ts';
import type { SlotDef } from './manifest.ts';
import { filler } from './sample.ts';

export interface ManifestIssue {
  severity: 'error' | 'warning';
  /** Pattern id; absent for manifest-level issues. */
  pattern?: string;
  slot?: string;
  message: string;
}

/** What makes a manifest no list of patterns with a list of slots each, as errors; empty when it is one. */
function shapeIssues(manifest: unknown): ManifestIssue[] {
  const isObject = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null && !Array.isArray(value);
  const patterns = isObject(manifest) ? manifest.patterns : undefined;
  if (!Array.isArray(patterns)) return [{ severity: 'error', message: 'patterns is no list of patterns, such as [{ "id": "cover", "sourceSlide": 1, "slots": [...] }]' }];
  const issues: ManifestIssue[] = [];
  patterns.forEach((pattern: unknown, index) => {
    if (!isObject(pattern)) {
      issues.push({ severity: 'error', message: `pattern ${index + 1} is no object with an id, a slide and slots` });
      return;
    }
    const at = typeof pattern.id === 'string' ? { pattern: pattern.id } : {};
    if (!Array.isArray(pattern.slots)) issues.push({ severity: 'error', ...at, message: 'slots is no list of slots, such as [{ "name": "title", "kind": "text", "shapeId": "2" }]' });
    else if (!pattern.slots.every(isObject)) issues.push({ severity: 'error', ...at, message: 'slots holds something that is no slot: each is an object with a name, a kind and a shape' });
  });
  return issues;
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
  // A manifest written by hand may be of any shape: patterns that are no list of objects, or slots
  // that are none, are said, and nothing more of such a manifest is read.
  const shape = shapeIssues(manifest);
  if (shape.length > 0) return shape;
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
  // Settings of the whole template: its text size, the colours it names and the frame of its free
  // slides. A manifest written by hand may hold anything here; null is a setting that is not given.
  const scaleProblem = templateScaleProblem(manifest.textScale);
  if (scaleProblem) issues.push({ severity: 'error', message: scaleProblem });
  // The direction its charts read: anything but rtl reads left to right.
  const chartDirection: unknown = manifest.chartDirection;
  if (chartDirection !== undefined && chartDirection !== null && chartDirection !== 'rtl' && chartDirection !== 'ltr') {
    issues.push({ severity: 'warning', message: `chartDirection is ${JSON.stringify(chartDirection)}; it is "rtl" or "ltr", and charts read left to right` });
  }
  // How wide its text is set and the least size its running text takes: a setting that is wrong is not used.
  for (const name of ['characterWidth', 'leastTextSize'] as const) {
    const problem = templateTextProblem(name, manifest[name]);
    if (problem) issues.push({ severity: 'warning', message: problem });
  }
  const colors: unknown = manifest.colors;
  if (colors !== undefined && colors !== null && colors !== false) {
    if (typeof colors !== 'object' || Array.isArray(colors)) {
      issues.push({ severity: 'error', message: `colors is ${JSON.stringify(colors)}; it is an object such as { "accent": "0E7C7B" }, or false for a template that keeps its own colours` });
    } else {
      for (const [name, value] of Object.entries(colors)) {
        if (name !== 'accent' && name !== 'accent2') issues.push({ severity: 'warning', message: `colors has no setting called "${name}"; it is ignored` });
        else if (value !== undefined && parseHex(value) === undefined) issues.push({ severity: 'error', message: `colors.${name} is ${JSON.stringify(value)}; a colour is six hex digits such as "0E7C7B"` });
      }
    }
  }
  if (manifest.free) {
    const base = manifest.free.base;
    const gone = base !== undefined && !ids.has(base);
    if (gone) issues.push({ severity: 'error', message: `free.base names the pattern "${base}", which the manifest does not have` });
    const parts = [...FREE_STYLE_ROLES, 'panel'] as readonly string[];
    // A manifest written by hand may hold anything here.
    const stated = Object.entries((manifest.free.styles ?? {}) as Record<string, unknown>);
    const usable: Array<[string, { slide: number; shape: string }]> = [];
    for (const [role, given] of stated) {
      const source = given as { slide?: unknown; shape?: unknown } | null;
      if (!parts.includes(role)) issues.push({ severity: 'warning', message: `free.styles names "${role}"; a free slide is made of ${parts.join(', ')}` });
      else if (!source || typeof source !== 'object' || typeof source.slide !== 'number' || (typeof source.shape !== 'string' && typeof source.shape !== 'number')) {
        issues.push({ severity: 'error', message: `free.styles.${role} is ${JSON.stringify(given)}; it names a shape of the template, such as { "slide": 7, "shape": "7" }` });
      } else if (!info.slides[source.slide - 1]?.shapes.some((shape) => shape.id === String(source.shape))) {
        issues.push({ severity: 'error', message: `free.styles.${role} points at shape #${source.shape} of slide ${source.slide}, which the template does not have` });
      } else usable.push([role, { slide: source.slide, shape: String(source.shape) }]);
    }
    // A layout that is there is not yet a frame, and a shape that is there not yet one a free slide
    // can copy. What cannot be used is passed over when a slide is made, so it is said here.
    const named = base !== undefined && !gone;
    if (named || usable.length > 0) {
      const kit = freeKitOrReason(template);
      if (typeof kit !== 'string') {
        if (named && kit.base !== base) {
          const why = freeBaseRefusal(template, base) ?? `the layout "${base}" gives no frame of its own`;
          issues.push({ severity: 'warning', message: `free.base names the pattern "${base}", which cannot be the frame of free slides: ${why}; they stand on "${kit.base}"` });
        }
        for (const [role, source] of usable) {
          const used = role === 'panel' ? kit.panel : kit.styles[role as (typeof FREE_STYLE_ROLES)[number]];
          if (used && used.slide === source.slide && used.shapeId === source.shape) continue;
          const what = role === 'panel' ? 'no layout of the manifest sets text on' : 'no layout of the manifest uses as a text that can be copied';
          issues.push({ severity: 'warning', message: `free.styles.${role} points at shape #${source.shape} of slide ${source.slide}, which ${what}; the software's own choice is used` });
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
