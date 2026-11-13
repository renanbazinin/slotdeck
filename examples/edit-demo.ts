// Applies an instruction to a deck that was generated earlier, and renders the result.
//
//   node --env-file=.env examples/edit-demo.ts out/dogs.deck.json "Shorten slide 3 to two bullets"
//
// The pictures of the earlier version are read back from its pptx (same name, .pptx), so photos
// that do not change are not generated again.
//
// Options:
//   --pptx <file>     the earlier pptx, when it is not next to the deck JSON
//   --out <name>      base name of the output in out/ (default: <deck>.edited)
//   --adaptive        allow optional slots to be left out
//   --charts          supply the two sample charts (needed for a deck that was generated with them)
//   --real-images     generate any new photo with the image model (paid); the default draws stand-ins
//
// Environment: GEMINI_API_KEY, and optionally SLOTDECK_TEXT_MODEL and SLOTDECK_IMAGE_MODEL.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { parseArgs } from 'node:util';
import { editDeck, extractAssets } from '../src/index.ts';
import type { DeckSpec } from '../src/index.ts';
import { demoTemplate, icons, logProgress, providers, root, sampleCharts } from './shared.ts';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    pptx: { type: 'string' },
    out: { type: 'string' },
    adaptive: { type: 'boolean' },
    charts: { type: 'boolean' },
    'real-images': { type: 'boolean' },
  },
});
const [deckPath, ...rest] = positionals;
const instruction = rest.join(' ').trim();
if (!deckPath || !instruction) {
  console.error('usage: node --env-file=.env examples/edit-demo.ts <deck.json> "what to change" [options]');
  process.exit(1);
}

const template = demoTemplate();
const deck = JSON.parse(readFileSync(deckPath, 'utf8')) as DeckSpec;
const pptxPath = values.pptx ?? deckPath.replace(/\.deck\.json$/, '.pptx');
const cache = new Map<string, Uint8Array>();
if (existsSync(pptxPath)) {
  for (const [id, data] of Object.entries(extractAssets(template, deck, readFileSync(pptxPath)))) cache.set(id, data);
  console.log(`reusing ${cache.size} picture(s) from ${basename(pptxPath)}`);
}

const started = performance.now();
const result = await editDeck(template, deck, instruction, {
  ...providers(values['real-images'] === true),
  icons,
  assets: values.charts ? sampleCharts : [],
  cache,
  layout: values.adaptive ? 'adaptive' : 'strict',
  onProgress: logProgress,
});

const name = values.out ?? `${basename(deckPath).replace(/\.deck\.json$/, '')}.edited`;
const outDir = join(root, 'out');
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, `${name}.pptx`), result.pptx);
writeFileSync(join(outDir, `${name}.deck.json`), JSON.stringify(result.deck, null, 2));

console.log(`\n${result.summary}`);
console.log(`${result.deck.slides.length} slides (was ${deck.slides.length}), rewritten: ${result.rewritten.join(', ') || 'none'}, ${((performance.now() - started) / 1000).toFixed(1)} s`);
for (const warning of result.warnings) console.log(`  warning: slide ${warning.slide} ${warning.slot}: ${warning.message}`);
console.log(`\nout/${name}.pptx\nout/${name}.deck.json`);
