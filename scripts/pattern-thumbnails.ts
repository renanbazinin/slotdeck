// Makes one picture for each pattern of a template, named after the pattern, so a plan can be
// shown as a row of layout thumbnails.
//
// Without PowerPoint: the patterns are drawn from the template's sample deck by a preview
// provider (here the Pagus library), as SVG. Give the template itself. A derived pattern (a
// mirror image, a text-only version) is drawn as it will look, since each pattern has a slide
// of its own in the sample deck:
//
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json templates/demo.pptx out/thumbnails
//
// With PowerPoint: export the template's slides to PNG first, and give the folder. Each pattern
// gets the picture of the slide it is made from:
//
//   powershell -File scripts/render-pptx.ps1 -Path template.pptx -OutDir out/template-png -Width 480
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails
//
// A derived pattern looks different from the slide it is made from. Give a fourth folder with
// pictures of a deck that has one slide per pattern, in manifest order (see
// scripts/template-deck.ts), and every pattern that is not the first of its slide takes its
// picture from there:
//
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails out/patterns-png
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import { loadTemplate, pagusPreviews, patternPreviews, previewExtension } from '../src/index.ts';
import type { TemplateManifest } from '../src/index.ts';

const [manifestPath, source, outDir, deckDir] = process.argv.slice(2);
if (!manifestPath || !source || !outDir) {
  console.error('usage: node scripts/pattern-thumbnails.ts <manifest.json> <template.pptx | rendered-slides-dir> <out-dir> [rendered-pattern-deck-dir]');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as TemplateManifest;
const picture = (dir: string, number: number) => join(dir, `slide-${String(number).padStart(2, '0')}.png`);
mkdirSync(outDir, { recursive: true });

if (source.toLowerCase().endsWith('.pptx')) {
  const template = loadTemplate(readFileSync(source), manifest);
  const pictures = await patternPreviews(template, pagusPreviews(pagusCore, pagusRenderer));
  for (const [id, preview] of Object.entries(pictures)) {
    const name = `${id}.${previewExtension(preview)}`;
    writeFileSync(join(outDir, name), preview.data);
    console.log(name);
  }
} else {
  const seen = new Set<number>();
  manifest.patterns.forEach((pattern, index) => {
    const derived = seen.has(pattern.sourceSlide) && deckDir !== undefined;
    seen.add(pattern.sourceSlide);
    const from = derived ? picture(deckDir as string, index + 1) : picture(source, pattern.sourceSlide);
    if (!existsSync(from)) {
      console.error(`missing ${from} for pattern "${pattern.id}"`);
      process.exitCode = 1;
      return;
    }
    copyFileSync(from, join(outDir, `${pattern.id}.png`));
    console.log(`${pattern.id}.png  <- ${derived ? `slide ${index + 1} of the pattern deck` : `slide ${pattern.sourceSlide}`}`);
  });
}
