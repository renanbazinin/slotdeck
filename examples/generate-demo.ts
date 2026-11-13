// Brief in, pptx out, using Gemini and the demo template.
//
//   node --env-file=.env examples/generate-demo.ts "A deck about adopting a rescue dog"
//
// Options:
//   --lang he|en|...        language to write in (default: the template's)
//   --slides 8              exact number of slides (default: 6 to 10)
//   --images none|low|medium|high   how many slides carry a photo (default: medium)
//   --org "Name | Department"       text of the line at the top of every slide
//   --text-size large|x-large       larger body text
//   --adaptive              allow optional slots to be left out
//   --charts                supply two sample charts, as an app would supply its own diagrams
//   --real-images           generate photos with the image model (paid); the default draws stand-ins
//   --save <name>           also keep the deck and its pictures in examples/showcase/<name>, so that
//                           render-demo.ts can render it again without a model
//   --dry-run               call no model: print the prompts and schema sizes that would be sent
//
// Environment: GEMINI_API_KEY, and optionally SLOTDECK_TEXT_MODEL and SLOTDECK_IMAGE_MODEL.
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { DeckValidationError, generateDeck, planDeck } from '../src/index.ts';
import { readImageInfo } from '../src/util/image.ts';
import type { Brief, ImageDensity, LlmProvider, TextSize } from '../src/index.ts';
import { demoTemplate, icons, logProgress, providers, root, sampleCharts } from './shared.ts';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    lang: { type: 'string' },
    slides: { type: 'string' },
    images: { type: 'string' },
    org: { type: 'string' },
    'text-size': { type: 'string' },
    save: { type: 'string' },
    adaptive: { type: 'boolean' },
    charts: { type: 'boolean' },
    'real-images': { type: 'boolean' },
    'dry-run': { type: 'boolean' },
  },
});
const prompt = positionals.join(' ').trim();
if (!prompt) {
  console.error('usage: node --env-file=.env examples/generate-demo.ts "what the deck is about" [options]');
  process.exit(1);
}

const template = demoTemplate();
const brief: Brief = { prompt, fields: { organizationDepartment: values.org ?? 'slotdeck  |  Demo' } };
if (values['text-size']) brief.textSize = values['text-size'] as TextSize;
if (values.lang) brief.language = values.lang;
if (values.slides) brief.slideCount = Number(values.slides);
if (values.images) brief.imageDensity = values.images as ImageDensity;
const layout = values.adaptive ? 'adaptive' : 'strict';
const assets = values.charts ? sampleCharts : [];

if (values['dry-run']) {
  // A stand-in model that prints what it is asked and answers with the first patterns on offer.
  let call = 0;
  const llm: LlmProvider = {
    async generateJson(request) {
      call++;
      const schema = JSON.stringify(request.schema);
      if (call === 1) console.log(`--- system prompt\n${request.system}\n`);
      console.log(`--- request ${call}: prompt\n${request.prompt}\n`);
      console.log(`--- request ${call}: schema is ${schema.length} characters, ${(schema.match(/"type"/g) ?? []).length} typed nodes\n`);
      if (call > 1) return {};
      const offered = (request.schema as { properties: { slides: { items: { properties: { pattern: { enum: string[] } } } } } }).properties.slides
        .items.properties.pattern.enum;
      return { title: 'Dry run', slides: offered.map((pattern) => ({ pattern, purpose: `(the point of the ${pattern} slide)` })) };
    },
  };
  try {
    await planDeck(template, brief, { llm, icons, assets, layout, maxRepairs: 0 });
  } catch (error) {
    if (!(error instanceof DeckValidationError)) throw error; // the empty stand-in answer is expected to fail validation
  }
  process.exit(0);
}

const started = performance.now();
const result = await generateDeck(template, brief, { ...providers(values['real-images'] === true), icons, assets, layout, onProgress: logProgress });

const slug =
  prompt
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 40) || 'deck';
const outDir = join(root, 'out');
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, `${slug}.pptx`), result.pptx);
writeFileSync(join(outDir, `${slug}.deck.json`), JSON.stringify(result.deck, null, 2));

console.log(`\n${result.deck.title ?? slug}: ${result.deck.slides.length} slides in ${((performance.now() - started) / 1000).toFixed(1)} s`);
result.outline.forEach((entry, index) => console.log(`  ${index + 1}. ${entry.pattern}: ${entry.purpose}`));
for (const warning of result.warnings) console.log(`  warning: slide ${warning.slide} ${warning.slot}: ${warning.message}`);
console.log(`\nout/${slug}.pptx\nout/${slug}.deck.json`);

if (values.save) {
  // The deck with the pictures it was given, apart from the icon set and sample charts, which the demos already have.
  const dir = join(root, 'examples', 'showcase', values.save);
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(join(dir, 'assets'), { recursive: true });
  writeFileSync(join(dir, 'deck.json'), `${JSON.stringify(result.deck, null, 2)}\n`);
  const known = new Set([...icons.map((icon) => icon.id), ...sampleCharts.map((chart) => chart.id)]);
  for (const [id, data] of Object.entries(result.assets)) {
    if (!known.has(id)) writeFileSync(join(dir, 'assets', `${id}.${readImageInfo(data).extension}`), data);
  }
  console.log(`examples/showcase/${values.save}/`);
}
