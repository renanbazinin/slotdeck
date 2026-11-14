// Writes decks that fill every text slot of the demo template with running text of many lengths,
// in English and in Hebrew, for checking the text measurer against PowerPoint. Dev-only.
//
//   node scripts/text-fit-stress.ts            writes out/stress-<language>-<n>.pptx and .deck.json
//   node scripts/text-fit-stress.ts --small    the same with every font size halved, as out/small-*:
//                                              small text is where PowerPoint's own rounding shows most
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { loadTemplate, renderDeck, sampleDeck } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { effectiveSlots, textBudget } from '../src/template/manifest.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'out');
mkdirSync(outDir, { recursive: true });
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const small = process.argv.includes('--small');
const prefix = small ? 'small' : 'stress';
const templateBytes = small ? halveFontSizes(readFileSync(join(root, 'templates', 'demo.pptx'))) : readFileSync(join(root, 'templates', 'demo.pptx'));
if (small) writeFileSync(join(outDir, 'small-template.pptx'), templateBytes);
const template = loadTemplate(templateBytes, manifest);

/** The template with every stated font size halved: 18 pt body text becomes 9 pt in the same boxes. */
function halveFontSizes(pptx: Uint8Array): Uint8Array {
  const files = unzipSync(pptx);
  for (const [name, data] of Object.entries(files)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    files[name] = strToU8(strFromU8(data).replace(/ sz="(\d+)"/g, (_, size: string) => ` sz="${Math.round(Number(size) / 100) * 50}"`));
  }
  return zipSync(files);
}

/** Every word of the text slots of an example deck, in order. */
function wordsOf(file: string): string[] {
  const deck = JSON.parse(readFileSync(join(root, 'examples', 'decks', file), 'utf8')) as DeckSpec;
  const words: string[] = [];
  for (const slide of deck.slides) {
    for (const value of Object.values(slide.slots)) {
      const texts = typeof value === 'string' ? [value] : Array.isArray(value) ? value : [];
      for (const text of texts) words.push(...text.split(/\s+/).filter(Boolean));
    }
  }
  return words;
}

/** Running text of about `length` characters, starting at a different word each time. */
function running(words: string[], length: number, start: number): string {
  let text = '';
  for (let index = start % words.length; ; index = (index + 1) % words.length) {
    const next = text ? `${text} ${words[index]}` : (words[index] as string);
    if (next.length > length && text) break;
    text = next;
    if (text.length >= length) break;
  }
  return text;
}

const languages = [
  { tag: 'en', direction: 'ltr' as const, words: wordsOf('patterns.en.deck.json') },
  { tag: 'he', direction: 'rtl' as const, words: wordsOf('patterns.he.deck.json') },
];
// Halved text leaves four times the room, so the small decks are filled further to get as many lines.
const fills = small ? [1.5, 2.5, 3.5, 4.5, 6] : [0.5, 0.8, 1.0, 1.3, 1.7];

let seed = 7;
for (const language of languages) {
  fills.forEach((fill, round) => {
    const { deck, assets } = sampleDeck(template);
    deck.language = language.tag;
    deck.direction = language.direction;
    deck.title = `Stress ${language.tag} ${fill}`;
    for (const slide of deck.slides) {
      const pattern = manifest.patterns.find((entry) => entry.id === slide.pattern);
      if (!pattern) continue;
      for (const slot of effectiveSlots(pattern, slide.items)) {
        if (slot.kind !== 'text' || slot.deckField || slide.slots[slot.name] === undefined) continue;
        const budget = textBudget(slot).maxChars ?? 60;
        const several = (slot.maxParagraphs ?? 1) > 1;
        const count = several ? Math.max(2, (slot.maxParagraphs ?? 2) - (round % 2)) : 1;
        const each = Math.max(3, Math.round((budget * fill) / count));
        const paragraphs = Array.from({ length: count }, () => running(language.words, each, (seed = (seed * 31 + 17) % 9973)));
        slide.slots[slot.name] = several ? paragraphs : (paragraphs[0] as string);
      }
    }
    write(`${prefix}-${language.tag}-${round + 1}`, deck, assets);
    if (small) return;
    // The same text set larger, as a deck asks for with its text size.
    if (round === 2) write(`stress-${language.tag}-large`, { ...deck, textScale: 1.25 }, assets);
    if (round === 1) write(`stress-${language.tag}-xlarge`, { ...deck, textScale: 1.5 }, assets);
  });
}

// The studio's own sample deck: every slot filled to its budget.
const sample = sampleDeck(template);
if (!small) write('sample', sample.deck, sample.assets);

function write(name: string, deck: DeckSpec, assets: Record<string, Uint8Array>): void {
  const { pptx } = renderDeck(template, deck, { assets, layout: 'adaptive' });
  writeFileSync(join(outDir, `${name}.pptx`), pptx);
  writeFileSync(join(outDir, `${name}.deck.json`), JSON.stringify(deck, null, 2));
  console.log(`${name}: ${deck.slides.length} slides`);
}
