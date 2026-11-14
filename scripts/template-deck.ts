// Writes a deck that shows every pattern of a manifest once, in manifest order, filled with the
// template's own sample texts and pictures. A derived pattern (a mirror image, a text-only
// version, a native chart) has no slide of its own in the template, so this is how to get a
// picture of it, for a plan preview or for a look at what a variant does.
//
//   node scripts/template-deck.ts <template.pptx> <manifest.json> <out.pptx>
//   powershell -File scripts/render-pptx.ps1 -Path <out.pptx> -OutDir out/patterns-png -Width 480
//   node scripts/pattern-thumbnails.ts <manifest.json> out/template-png templates/demo.thumbnails out/patterns-png
//
// Charts and tables have no sample content the code can read, so they get the values of `sampleDeck`.
import { readFileSync, writeFileSync } from 'node:fs';
import { inspectTemplate, loadTemplate, renderDeck, sampleDeck } from '../src/index.ts';
import type { SlotValue, TemplateManifest } from '../src/index.ts';

const [templatePath, manifestPath, outPath] = process.argv.slice(2);
if (!templatePath || !manifestPath || !outPath) {
  console.error('usage: node scripts/template-deck.ts <template.pptx> <manifest.json> <out.pptx>');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as TemplateManifest;
const template = loadTemplate(readFileSync(templatePath), manifest);
const info = inspectTemplate(template.pkg);
const { deck, assets } = sampleDeck(template);

manifest.patterns.forEach((pattern, index) => {
  const shapes = new Map((info.slides[pattern.sourceSlide - 1]?.shapes ?? []).map((shape) => [shape.id, shape]));
  const slots: Record<string, SlotValue> = deck.slides[index]?.slots ?? {};
  for (const slot of pattern.slots) {
    const shape = shapes.get(slot.shapeId);
    if (!shape || !(slot.name in slots)) continue;
    if (slot.kind === 'text') {
      const paragraphs = (shape.paragraphs ?? []).filter((paragraph) => paragraph.trim() !== '');
      if (paragraphs.length > 0) slots[slot.name] = paragraphs;
    } else if (slot.kind === 'image' && shape.image && template.pkg.has(shape.image.part)) {
      assets[shape.image.part] = template.pkg.read(shape.image.part);
      slots[slot.name] = { asset: shape.image.part };
    }
  }
});

const { pptx, warnings } = renderDeck(template, deck, { assets, layout: 'adaptive' });
writeFileSync(outPath, pptx);
console.log(`${outPath}: ${deck.slides.length} slides, one per pattern`);
for (const warning of warnings) console.log(`  warning: slide ${warning.slide ?? '-'} ${warning.slot ?? ''}: ${warning.message}`);
