// Compares the text measurer with PowerPoint on a rendered deck: for every text slot, the lines
// the measurer predicts against the lines PowerPoint reports (scripts/measure-pptx.ps1).
// Dev-only; needs the template's fonts on disk.
//
//   node scripts/compare-text-fit.ts <deck.json> <deck.lines.json> [--open] [--verbose] [--record <name>] [--template <pptx>]
//
// --open measures with the open stand-in fonts the tests use, in place of the installed ones.
// --template names the pptx the deck was rendered from, when it is not templates/demo.pptx.
// --record adds PowerPoint's line counts for this deck to test/fixtures/powerpoint-lines.json,
// which the tests compare the measurer with.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as fontkit from 'fontkit';
import { fitOptions } from '../src/deck/validate.ts';
import { effectiveSlots, slotSources } from '../src/template/manifest.ts';
import { fontkitMetrics, loadTemplate, textMeasurer } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { officeFonts, openFonts } from '../test/fonts.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const recordAs = args.includes('--record') ? args[args.indexOf('--record') + 1] : undefined;
const templateFile = args.includes('--template') ? (args[args.indexOf('--template') + 1] as string) : join(root, 'templates', 'demo.pptx');
const [deckFile, linesFile] = args.filter((arg, index) => !arg.startsWith('--') && args[index - 1] !== '--record' && args[index - 1] !== '--template');
if (!deckFile || !linesFile) {
  console.error('usage: node scripts/compare-text-fit.ts <deck.json> <deck.lines.json> [--open] [--verbose]');
  process.exit(1);
}
const verbose = args.includes('--verbose');
const files = args.includes('--open') ? openFonts() : officeFonts();
if (!files) {
  console.error('Calibri and Cambria are not installed here. Set SLOTDECK_FONT_DIR to their folder, or pass --open.');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const fonts = fontkitMetrics(fontkit, files);
const template = loadTemplate(readFileSync(templateFile), manifest, { fonts });
const measure = textMeasurer(template);
const deck = JSON.parse(readFileSync(deckFile, 'utf8')) as DeckSpec;
interface Measured {
  slides: Array<{ slide: number; shapes: Array<{ id: number; kind: string; lines?: string[]; paragraphs?: number[]; boundHeight?: number; fontSize?: number }> }>;
}
const truth = JSON.parse(readFileSync(linesFile, 'utf8')) as Measured;

/** What PowerPoint reported, by slide number and shape id: for each paragraph the length of each of its lines, and the height of the text in points. */
const recorded: Record<string, Record<string, { lines: number[][]; height: number }>> = {};
let slots = 0;
let same = 0;
let sameBreaks = 0;
let paragraphs = 0;
let sameParagraphs = 0;
const mirrored = (deck.direction ?? manifest.direction) !== manifest.direction;
deck.slides.forEach((slide, index) => {
  const pattern = manifest.patterns.find((entry) => entry.id === slide.pattern);
  const shapes = new Map((truth.slides[index]?.shapes ?? []).map((shape) => [String(shape.id), shape]));
  if (!pattern) return;
  const sources = slotSources(pattern, mirrored, slide.items);
  for (const slot of effectiveSlots(pattern, slide.items)) {
    if (slot.kind !== 'text') continue;
    const value = slide.slots[sources.get(slot.name) ?? slot.name] ?? (slot.deckField ? deck.fields?.[slot.deckField] : undefined);
    const shape = shapes.get(slot.shapeId);
    if (typeof value !== 'string' && !Array.isArray(value)) continue;
    if (!shape?.lines) continue;
    let taken = 0;
    const byParagraph = (shape.paragraphs ?? []).map((count) => (shape.lines ?? []).slice(taken, (taken += count)).map((line) => line.trimEnd().length));
    ((recorded[String(index + 1)] ??= {})[slot.shapeId] = { lines: byParagraph, height: shape.boundHeight ?? 0 });
    // Measured as validation and the renderer measure: at the deck's text size and the slide's number of items.
    const fit = measure.fit(pattern, slot, value as string | string[], fitOptions(manifest, deck, slide));
    if (!fit) {
      console.log(`slide ${index + 1} ${slot.name}: not measured`);
      continue;
    }
    slots++;
    const predicted = fit.paragraphs.map((lines) => lines.length);
    const actual = shape.paragraphs ?? [];
    paragraphs += actual.length;
    sameParagraphs += actual.filter((count, at) => predicted[at] === count).length;
    const linesMatch = fit.lines === (shape.lines?.length ?? 0);
    if (linesMatch) same++;
    const breaksMatch = JSON.stringify(fit.paragraphs.flat().map((line) => line.trimEnd())) === JSON.stringify((shape.lines ?? []).map((line) => line.trimEnd()));
    if (breaksMatch) sameBreaks++;
    const heightOff = Math.abs(fit.height - (shape.boundHeight ?? 0));
    if (verbose || !linesMatch || heightOff > 0.6) {
      console.log(`slide ${index + 1} ${slot.name}: measurer ${fit.lines} lines ${fit.height} pt at ${fit.size} pt${fit.fits ? '' : ` (over by ${fit.overflow})`}; PowerPoint ${shape.lines?.length} lines ${shape.boundHeight} pt at ${shape.fontSize} pt${linesMatch ? '' : '  <-- differs'}`);
      if (!breaksMatch && (verbose || !linesMatch)) {
        console.log(`    measurer:   ${JSON.stringify(fit.paragraphs.flat())}`);
        console.log(`    PowerPoint: ${JSON.stringify(shape.lines)}`);
      }
    }
  }
});
console.log(`${slots} text slots: ${same} with the same number of lines (${Math.round((same / Math.max(slots, 1)) * 100)}%), ${sameBreaks} with every break in the same place; ${sameParagraphs}/${paragraphs} paragraphs with the same number of lines`);
const missing = measure.missingFonts();
if (missing.length > 0) console.log(`fonts not supplied: ${missing.map((entry) => `${entry.typeface}${entry.bold ? ' bold' : ''}${entry.italic ? ' italic' : ''}`).join(', ')}`);

if (recordAs) {
  const file = join(root, 'test', 'fixtures', 'powerpoint-lines.json');
  mkdirSync(dirname(file), { recursive: true });
  const fixture = existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as { cases: Array<{ name: string }> }) : { cases: [] };
  // A deck rendered from the template with its font sizes halved says so, and the tests halve them too.
  const entry = { name: recordAs, small: args.includes('--template'), deck, shapes: recorded };
  fixture.cases = [...fixture.cases.filter((item) => item.name !== recordAs), entry].sort((a, b) => a.name.localeCompare(b.name));
  writeFileSync(file, `${JSON.stringify(fixture)}
`);
  console.log(`recorded as "${recordAs}" in ${file}`);
}
