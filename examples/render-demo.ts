// Renders the example decks from the demo template. No model calls and no API key: the decks
// are JSON files, and the pictures are drawn stand-ins or files saved next to a deck.
//
//   node examples/render-demo.ts
//
// - examples/decks/*.deck.json are hand-written. Their photos are drawn stand-ins.
// - examples/showcase/<name>/deck.json was planned and written by a model from a one-line brief
//   (see generate-demo.ts --save); the pictures it was given are in the folder's assets/.
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, extname, join } from 'node:path';
import { renderDeck } from '../src/index.ts';
import type { DeckSpec } from '../src/index.ts';
import { placeholderChart, placeholderPhoto } from './placeholder.ts';
import { demoTemplate, icons, root } from './shared.ts';

const outDir = join(root, 'out');
const template = demoTemplate();
const readJson = <T>(file: string): T => JSON.parse(readFileSync(file, 'utf8')) as T;

const standIns: Record<string, Uint8Array> = {
  'photo-tall': placeholderPhoto(1100, 1500, [14, 124, 123], [244, 162, 89]),
  'photo-square': placeholderPhoto(1400, 1400, [27, 42, 65], [14, 124, 123]),
  'photo-wide': placeholderPhoto(1600, 1350, [244, 162, 89], [232, 241, 240]),
  'chart-months': placeholderChart(1500, 980, [0.3, 0.38, 0.45, 0.55, 0.7, 0.92, 0.85, 0.6]),
  'chart-wide': placeholderChart(2400, 760, [0.35, 0.5, 0.42, 0.7, 0.62, 0.9, 0.75, 0.82, 0.58, 0.95]),
  ...Object.fromEntries(icons.map((icon) => [icon.id, icon.data])),
};

interface Job {
  deck: string;
  out: string;
  assets: Record<string, Uint8Array>;
  textScale?: number;
}
const jobs: Job[] = [];
const decksDir = join(root, 'examples', 'decks');
for (const file of readdirSync(decksDir).filter((name) => name.endsWith('.deck.json')).sort()) {
  jobs.push({ deck: join(decksDir, file), out: `${file.replace(/\.deck\.json$/, '').replace(/\./g, '-')}.pptx`, assets: standIns });
}
jobs.push({ deck: join(decksDir, 'patterns.en.deck.json'), out: 'patterns-en-large.pptx', assets: standIns, textScale: 1.25 });

const showcaseDir = join(root, 'examples', 'showcase');
for (const name of existsSync(showcaseDir) ? readdirSync(showcaseDir).sort() : []) {
  const deck = join(showcaseDir, name, 'deck.json');
  if (!existsSync(deck)) continue;
  const saved: Record<string, Uint8Array> = { ...standIns };
  const assetsDir = join(showcaseDir, name, 'assets');
  for (const file of existsSync(assetsDir) ? readdirSync(assetsDir) : []) saved[basename(file, extname(file))] = new Uint8Array(readFileSync(join(assetsDir, file)));
  jobs.push({ deck, out: `showcase-${name}.pptx`, assets: saved });
}

mkdirSync(outDir, { recursive: true });
for (const job of jobs) {
  const deck = readJson<DeckSpec>(job.deck);
  if (job.textScale) deck.textScale = job.textScale;
  const started = performance.now();
  const result = renderDeck(template, deck, { assets: job.assets, layout: 'adaptive' });
  try {
    writeFileSync(join(outDir, job.out), result.pptx);
  } catch (error) {
    // Usually the file is open in PowerPoint, which locks it.
    console.log(`${job.out}: rendered, but could not be written (${(error as NodeJS.ErrnoException).code ?? error})`);
    continue;
  }
  const size = (result.pptx.length / 1024).toFixed(0);
  console.log(`${job.out}: ${deck.slides.length} slides, ${size} KB, ${Math.round(performance.now() - started)} ms`);
  for (const warning of result.warnings) {
    console.log(`  warning: slide ${warning.slide ?? '-'} ${warning.slot ?? ''}: ${warning.message}`);
  }
}
