// Plays the part of a chat app: calls the four presentation tools in the order a conversation
// would, and prints what each hands back to the chat model.
//
//   node --env-file=.env examples/tools-demo.ts "A deck about adopting a rescue dog"
//
// Options:
//   --lang he|en|...   language to write in
//   --real-images      generate photos with the image model (paid); the default draws stand-ins
//
// Everything is kept under out/store: plans and decks as small JSON, pictures on their own,
// rendered files apart, and a picture of every slide under preview/ (SVG, drawn by the Pagus
// library: open one in a browser). out/plan-preview.html shows the plan as layout thumbnails.
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import { TOOL_NAMES, pagusPreviews, presentationTools } from '../src/index.ts';
import { directoryStore } from '../src/node.ts';
import { planPreviewHtml } from './plan-preview.ts';
import { demoTemplate, icons, providers, root } from './shared.ts';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { lang: { type: 'string' }, 'real-images': { type: 'boolean' } },
});
const topic = positionals.join(' ').trim();
if (!topic) {
  console.error('usage: node --env-file=.env examples/tools-demo.ts "what the deck is about" [--lang he]');
  process.exit(1);
}

const outDir = join(root, 'out');
mkdirSync(outDir, { recursive: true });
const tools = presentationTools({
  template: demoTemplate(),
  ...providers(values['real-images'] === true),
  icons,
  store: directoryStore(join(outDir, 'store')),
  fields: { organizationDepartment: values.lang === 'he' ? 'slotdeck  |  הדגמה' : 'slotdeck  |  Demo' },
  thumbnail: (pattern) => `/thumbnails/demo/${pattern}.png`,
  preview: pagusPreviews(pagusCore, pagusRenderer),
  onProgress: (event, tool) => {
    if (event.stage === 'images') console.log(`    [${tool}] images ${event.done}/${event.total}`);
    else console.log(`    [${tool}] ${event.stage}`);
  },
});

async function call(name: string, args: Record<string, unknown>) {
  const started = performance.now();
  console.log(`\n> ${name}(${JSON.stringify(args)})`);
  const result = await tools.call(name, args);
  console.log(`< ${((performance.now() - started) / 1000).toFixed(1)} s`);
  console.log(JSON.stringify(result.output, null, 2));
  if (result.output.error) process.exit(1);
  return result;
}

// 1. The user asks for a deck: the model plans it and shows the plan.
const plan = await call(TOOL_NAMES.plan, values.lang ? { topic, slide_count: 6, language: values.lang } : { topic, slide_count: 6 });

// 2. The user asks for a change to the plan.
const instruction = values.lang === 'he' ? 'הוסף לפני שקף הסיום שקף עם טבלה שמשווה שלוש אפשרויות' : 'Add a slide that compares three options in a table, before the closing slide';
const revised = await call(TOOL_NAMES.editPlan, { plan_id: plan.output.plan_id, instruction });
if (revised.plan) writeFileSync(join(outDir, 'plan-preview.html'), planPreviewHtml(revised.plan, join(root, 'templates', 'demo.thumbnails'), 'en'));

// 3. The user approves: the deck is written, its pictures made and stored, the file rendered.
const created = await call(TOOL_NAMES.create, { plan_id: revised.output.plan_id });

// 4. The user asks for a change to the finished deck.
const change = values.lang === 'he' ? 'קצר את שקף 2 ותן לו כותרת קליטה יותר' : 'Shorten slide 2 and make its title punchier';
const edited = await call(TOOL_NAMES.edit, { presentation_id: created.output.presentation_id, instruction: change });

writeFileSync(join(outDir, 'tools-demo.pptx'), edited.pptx as Uint8Array);
console.log('\nout/plan-preview.html\nout/tools-demo.pptx\nout/store/');
