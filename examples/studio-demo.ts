// The template studio from the command line: draft a manifest for a pptx with a model, change it
// by instruction, and render a sample deck that shows whether its budgets hold.
//
//   node --env-file=.env examples/studio-demo.ts draft [template.pptx] [--thumbnails dir]
//   node --env-file=.env examples/studio-demo.ts edit "let the cards take up to six"
//   node examples/studio-demo.ts sample
//
// `draft` writes out/studio.manifest.json; `edit` and `sample` work on that file. `--thumbnails`
// names a folder of slide pictures (slide-01.png, ...), which improves the draft; without it the
// model judges from measurements alone. Text calls only: no picture is generated.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { checkManifest, draftManifest, editManifest, loadTemplate, renderDeck, sampleDeck } from '../src/index.ts';
import type { TemplateManifest } from '../src/index.ts';
import { providers, root } from './shared.ts';

const [command, ...rest] = process.argv.slice(2);
const flag = (name: string): string | undefined => {
  const index = rest.indexOf(`--${name}`);
  return index >= 0 ? rest.splice(index, 2)[1] : undefined;
};
const thumbnails = flag('thumbnails');
const outDir = join(root, 'out');
const manifestFile = join(outDir, 'studio.manifest.json');
const sourceFile = join(outDir, 'studio.source.txt');
mkdirSync(outDir, { recursive: true });

const report = (manifest: TemplateManifest, pptx: Uint8Array): void => {
  const issues = checkManifest(loadTemplate(pptx, manifest));
  console.log(`${manifest.patterns.length} patterns; checker: ${issues.filter((issue) => issue.severity === 'error').length} errors, ${issues.filter((issue) => issue.severity === 'warning').length} warnings`);
  for (const issue of issues) console.log(`  ${issue.severity} ${issue.pattern ?? ''}${issue.slot ? `/${issue.slot}` : ''}: ${issue.message}`);
};

if (command === 'draft') {
  const templatePath = rest[0] ?? join(root, 'templates', 'demo.pptx');
  const pptx = readFileSync(templatePath);
  const { llm } = providers(false);
  let calls = 0;
  const started = performance.now();
  const result = await draftManifest(pptx, {
    llm: { generateJson: (request) => (calls++, llm.generateJson(request)) },
    name: basename(templatePath).replace(/\.pptx$/i, ''),
    slideImage: (slide) => {
      const file = thumbnails ? join(thumbnails, `slide-${String(slide).padStart(2, '0')}.png`) : '';
      return file && existsSync(file) ? { mimeType: 'image/png', data: new Uint8Array(readFileSync(file)) } : undefined;
    },
    onProgress: ({ done, total }) => process.stdout.write(`\r  slide ${done}/${total}`),
  });
  console.log(`\n${calls} model calls, ${((performance.now() - started) / 1000).toFixed(0)} s; written in ${result.manifest.language ?? '?'}, reads ${result.manifest.direction}`);
  for (const pattern of result.manifest.patterns) {
    const repeat = pattern.repeat ? ` [${pattern.repeat.columns ? 'grid' : 'row'} of ${pattern.repeat.shapes.length} ${pattern.repeat.item}s, takes ${pattern.repeat.min}-${pattern.repeat.max}]` : '';
    console.log(`  slide ${pattern.sourceSlide}: ${pattern.id}${pattern.figures ? ' (figures)' : ''}${repeat}\n    ${pattern.slots.map((slot) => `${slot.name}${slot.empty === 'drop' ? ':drop' : slot.empty === 'remove' ? '?' : ''}${slot.kind === 'text' && slot.deckField ? '@deck' : ''}`).join(' ')}`);
  }
  for (const skip of result.skipped) console.log(`  slide ${skip.slide} skipped: ${skip.reason}`);
  report(result.manifest, pptx);
  writeFileSync(manifestFile, JSON.stringify(result.manifest, null, 2));
  writeFileSync(sourceFile, templatePath);
  console.log(`wrote ${manifestFile}`);
} else if (command === 'edit' || command === 'sample') {
  if (!existsSync(manifestFile) || !existsSync(sourceFile)) {
    console.error('Nothing drafted yet. Run: node --env-file=.env examples/studio-demo.ts draft');
    process.exit(1);
  }
  const pptx = readFileSync(readFileSync(sourceFile, 'utf8'));
  let manifest = JSON.parse(readFileSync(manifestFile, 'utf8')) as TemplateManifest;

  if (command === 'edit') {
    const instruction = rest.join(' ').trim();
    if (!instruction) {
      console.error('usage: node --env-file=.env examples/studio-demo.ts edit "<what to change>"');
      process.exit(1);
    }
    const result = await editManifest(loadTemplate(pptx, manifest), instruction, { llm: providers(false).llm });
    for (const entry of result.applied) console.log(`  done    ${entry.message}`);
    for (const entry of result.rejected) console.log(`  refused ${entry.reason}`);
    if (result.note) console.log(`  note    ${result.note}`);
    manifest = result.manifest;
    writeFileSync(manifestFile, JSON.stringify(manifest, null, 2));
  }

  report(manifest, pptx);
  const template = loadTemplate(pptx, manifest);
  const { deck, assets } = sampleDeck(template, { items: 'max' });
  const rendered = renderDeck(template, deck, { assets, layout: 'adaptive' });
  const sampleFile = join(outDir, 'studio.sample.pptx');
  writeFileSync(sampleFile, rendered.pptx);
  console.log(`wrote ${sampleFile}: ${deck.slides.length} slides, every slot filled to its budget`);
} else {
  console.error('usage: node --env-file=.env examples/studio-demo.ts draft|edit|sample ...');
  process.exit(1);
}
