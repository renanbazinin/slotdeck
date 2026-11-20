// Measures how wide PowerPoint sets a character of a template's running text, as a share of its
// size: the value of the manifest's `characterWidth` (see docs/text-fit.md). Renders free slides of
// running text from the template, asks the locally installed PowerPoint how it breaks their lines
// (scripts/measure-pptx.ps1), and reads every full line: its width over its characters and its
// size. Prints what most lines stay under, and the value to state. Dev-only; needs Windows and
// PowerPoint, and the template's fonts installed or embedded in it.
//
//   node scripts/measure-character-width.ts <template.pptx> <manifest.json> [--text <paragraphs.txt>] [--language he] [--out out/width]
//
// --text is a file of paragraphs of the kind the template's decks hold, one a line: in the
// language they are written in, with their figures and their names. Left out, the text of the
// template's own slides is used, where it has enough. The value is the 95th percentile of the
// lines, rounded up to two decimals: 19 lines of 20 are set no wider.
import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { loadTemplate, renderDeck } from '../src/index.ts';
import type { DeckSpec, SlideSpec, TemplateManifest } from '../src/index.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const option = (name: string): string | undefined => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const [pptxFile, manifestFile] = args.filter((arg, index) => !arg.startsWith('--') && !args[index - 1]?.startsWith('--'));
if (!pptxFile || !manifestFile) {
  console.error('usage: node scripts/measure-character-width.ts <template.pptx> <manifest.json> [--text <paragraphs.txt>] [--language he] [--out out/width]');
  process.exit(1);
}
const out = resolve(option('--out') ?? join(root, 'out', 'width'));
mkdirSync(out, { recursive: true });
const pptx = new Uint8Array(readFileSync(pptxFile));
const manifest = JSON.parse(readFileSync(manifestFile, 'utf8')) as TemplateManifest;
const template = loadTemplate(pptx, manifest);

// The paragraphs to set: the file given, or the sentences on the template's own slides.
let paragraphs: string[];
const textFile = option('--text');
if (textFile) {
  paragraphs = readFileSync(textFile, 'utf8')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length >= 40);
} else {
  const parts = unzipSync(pptx);
  paragraphs = Object.keys(parts)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .flatMap((name) => [...strFromU8(parts[name] as Uint8Array).matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((match) => [...(match[1] as string).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((run) => run[1]).join('')))
    .map((text) => text.trim())
    .filter((text) => text.length >= 40);
}
const total = paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
if (paragraphs.length < 4 || total < 1500) {
  console.error(`only ${total} characters of text in ${paragraphs.length} paragraphs to measure with: give some 3000 characters of the template's kind of text with --text`);
  process.exit(1);
}
/** So many characters of the paragraphs, from the paragraph at `from` on, again from the first where they run out. */
const take = (chars: number, from: number): string[] => {
  const taken: string[] = [];
  let used = 0;
  for (let at = from; used < chars; at++) {
    const paragraph = paragraphs[at % paragraphs.length] as string;
    taken.push(paragraph);
    used += paragraph.length;
  }
  return taken;
};
// Slides of one running text, short to long, as one column and set dense: their text is set at
// sizes from the template's own down to the least that reads.
const slides: SlideSpec[] = [];
for (const [chars, dense] of [[500, false], [900, false], [1400, false], [1200, true], [2000, true], [3000, true]] as const) {
  slides.push({ pattern: 'free', blocks: ['text'], ...(dense ? { arrange: { dense: true } } : {}), slots: { title: 'Text', b1Text: take(chars, slides.length * 3) } });
}
const language = option('--language') ?? manifest.language;
const deck: DeckSpec = { version: 1, ...(language ? { language } : {}), slides };
// Every text of the deck as a whole gets a word, so that none the template needs is missing.
const fields = manifest.patterns.flatMap((pattern) => pattern.slots).flatMap((slot) => (slot.kind === 'text' && slot.deckField !== undefined ? [slot.deckField] : []));
if (fields.length > 0) deck.fields = Object.fromEntries(fields.map((field) => [field, 'Text']));
const file = join(out, 'width.pptx');
writeFileSync(file, renderDeck(template, deck, { layout: 'adaptive' }).pptx);
const linesFile = join(out, 'width.lines.json');
execFileSync('powershell', ['-ExecutionPolicy', 'Bypass', '-File', join(root, 'scripts', 'measure-pptx.ps1'), '-Path', file, '-Out', linesFile], { stdio: 'inherit' });

interface Measured {
  kind: string;
  name: string;
  fontSize: number | null;
  lines: string[];
  lineWidths: number[];
  paragraphs: number[];
}
const measured = JSON.parse(readFileSync(linesFile, 'utf8').replace(/^﻿/, '')) as { slides: Array<{ shapes: Measured[] }> };
const shares: number[] = [];
for (const slide of measured.slides) {
  for (const shape of slide.shapes) {
    if (shape.kind !== 'text' || !shape.name.startsWith('Block 1') || !(typeof shape.fontSize === 'number' && shape.fontSize > 0)) continue;
    // The last line of a paragraph is not full; nor is a line of a few characters.
    const ends = new Set<number>();
    let line = 0;
    for (const count of shape.paragraphs) ends.add((line += count) - 1);
    shape.lines.forEach((text, at) => {
      const characters = text.trimEnd().length;
      if (ends.has(at) || characters < 15) return;
      shares.push((shape.lineWidths[at] as number) / (characters * (shape.fontSize as number)));
    });
  }
}
if (shares.length < 40) {
  console.error(`PowerPoint set only ${shares.length} full lines: give more text with --text`);
  process.exit(1);
}
shares.sort((a, b) => a - b);
const at = (share: number): number => shares[Math.min(shares.length - 1, Math.floor(share * shares.length))] as number;
const mean = shares.reduce((sum, share) => sum + share, 0) / shares.length;
// Rounded up: a value under the width most lines take lets text run past its box.
const value = Math.ceil(at(0.95) * 100 - 1e-9) / 100;
console.log(`${shares.length} full lines: mean ${mean.toFixed(3)}, half under ${at(0.5).toFixed(3)}, 19 of 20 under ${at(0.95).toFixed(3)}, 99 of 100 under ${at(0.99).toFixed(3)}, widest ${(shares.at(-1) as number).toFixed(3)}`);
console.log(`characterWidth: ${value}`);
