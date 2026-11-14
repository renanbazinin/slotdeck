// Makes one picture for each pattern of a template, named after the pattern, so a plan can be
// shown as a row of layout thumbnails.
//
// Without PowerPoint: the patterns are drawn from the template's sample deck by a preview
// provider (here the Pagus library), as SVG. Give the template itself:
//
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json templates/demo.pptx out/thumbnails
//
// With PowerPoint: export the template's slides to PNG first, and give the folder. Each pattern
// gets the picture of the slide it is made from:
//
//   powershell -File scripts/render-pptx.ps1 -Path template.pptx -OutDir out/template-png -Width 480
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import * as pagusCore from '@pagus-kit/core';
import * as pagusRenderer from '@pagus-kit/renderer';
import { loadTemplate, pagusPreviews, patternPreviews, previewExtension } from '../src/index.ts';
import type { TemplateManifest } from '../src/index.ts';

const [manifestPath, source, outDir] = process.argv.slice(2);
if (!manifestPath || !source || !outDir) {
  console.error('usage: node scripts/pattern-thumbnails.ts <manifest.json> <template.pptx | rendered-slides-dir> <out-dir>');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as TemplateManifest;
mkdirSync(outDir, { recursive: true });

if (source.toLowerCase().endsWith('.pptx')) {
  const template = loadTemplate(readFileSync(source), manifest);
  const pictures = await patternPreviews(template, pagusPreviews(pagusCore, pagusRenderer));
  for (const [id, picture] of Object.entries(pictures)) {
    const name = `${id}.${previewExtension(picture)}`;
    writeFileSync(join(outDir, name), picture.data);
    console.log(name);
  }
} else {
  for (const pattern of manifest.patterns) {
    const slide = join(source, `slide-${String(pattern.sourceSlide).padStart(2, '0')}.png`);
    if (!existsSync(slide)) {
      console.error(`missing ${slide} for pattern "${pattern.id}"`);
      process.exitCode = 1;
      continue;
    }
    copyFileSync(slide, join(outDir, `${pattern.id}.png`));
    console.log(`${pattern.id}.png  <- slide ${pattern.sourceSlide}`);
  }
}
