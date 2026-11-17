import { isFreeSlide, isImageRef, toParagraphs } from '../deck/spec.ts';
import type { DeckSpec } from '../deck/spec.ts';
import { PptxPackage, readRels, resolveTarget } from '../pptx/package.ts';
import { readPresentation } from '../pptx/presentation.ts';
import { NS, descendants, parseXml, path, rootElement } from '../pptx/xml.ts';
import { freeFor } from '../free/resolve.ts';
import { effectiveSlots, slotSources } from '../template/manifest.ts';
import type { PatternDef } from '../template/manifest.ts';
import { indexShapes } from './render.ts';
import type { Template } from './render.ts';

const normalize = (text: string) => text.replace(/\s+/g, ' ').trim();

/**
 * Recovers the pictures of a deck from a pptx that was rendered from it, keyed by asset id.
 * With this, a deck's JSON and its pptx are all that has to be stored between edits: generated
 * photos are read back instead of being generated, and paid for, again.
 *
 * Only slides that still match the deck are used. If the file was changed by hand (slides
 * added, removed or reordered, or a title retyped), those slides are skipped.
 */
export function extractAssets(template: Template, deck: DeckSpec, pptx: Uint8Array): Record<string, Uint8Array> {
  const assets: Record<string, Uint8Array> = {};
  const pkg = PptxPackage.load(pptx);
  const slides = readPresentation(pkg).slides;
  if (slides.length !== deck.slides.length) return assets;

  const patterns = new Map(template.manifest.patterns.map((pattern) => [pattern.id, pattern]));
  const mirrored = (deck.direction ?? template.manifest.direction) !== template.manifest.direction;

  deck.slides.forEach((slide, index) => {
    // A slide the software laid out is laid out again the same way, which says where its pictures are.
    const free = isFreeSlide(slide) ? freeFor(template, deck, slide) : undefined;
    const pattern: PatternDef | undefined = free === undefined ? patterns.get(slide.pattern) : typeof free === 'string' ? undefined : free.plan.pattern;
    const part = slides[index]?.part;
    if (!pattern || !part || !pkg.has(part)) return;
    const shapes = indexShapes(rootElement(parseXml(pkg.readText(part))));
    const sources = slotSources(pattern, mirrored, slide.items);
    const slots = effectiveSlots(pattern, slide.items);
    const valueOf = (slotName: string) => slide.slots[sources.get(slotName) ?? slotName];

    // Check one text slot to be sure this is still the slide the deck describes.
    const probe = slots.find((slot) => {
      const value = valueOf(slot.name);
      return slot.kind === 'text' && (typeof value === 'string' || Array.isArray(value));
    });
    if (probe) {
      const shape = shapes.get(probe.shapeId);
      const expected = normalize(toParagraphs(valueOf(probe.name) as string | string[]).join(' '));
      const actual = shape
        ? normalize(
            descendants(shape, NS.a, 'p')
              .map((paragraph) =>
                descendants(paragraph, NS.a, 't')
                  .map((node) => node.textContent ?? '')
                  .join(''),
              )
              .join(' '),
          )
        : undefined;
      if (actual !== expected) return;
    }

    const targets = new Map(readRels(pkg, part).map((rel) => [rel.id, resolveTarget(part, rel.target)]));
    for (const slot of slots) {
      const value = valueOf(slot.name);
      if (slot.kind !== 'image' || !isImageRef(value)) continue;
      // A pictogram of a deck in other colours was drawn in those colours: the icon set has the original.
      if (slot.imageKind === 'icon' && (deck.colors !== undefined || slide.colors !== undefined)) continue;
      const shape = shapes.get(slot.shapeId);
      const blip = shape && path(shape, [NS.p, 'blipFill'], [NS.a, 'blip']);
      const target = targets.get(blip?.getAttributeNS(NS.r, 'embed') ?? '');
      if (target && pkg.has(target)) assets[value.asset] = pkg.read(target);
    }
  });

  return assets;
}
