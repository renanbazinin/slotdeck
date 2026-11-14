// Names the pictures of a template's sample slides after the patterns they belong to, so a
// plan can be shown as a row of layout thumbnails.
//
//   powershell -File scripts/render-pptx.ps1 -Path template.pptx -OutDir out/template-png -Width 480
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails
//
// A derived pattern (a mirror image, a text-only version) looks different from the slide it is
// made from. Give a fourth folder with pictures of a deck that has one slide per pattern, in
// manifest order (see scripts/template-deck.ts), and every pattern that is not the first of its
// slide takes its picture from there:
//
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails out/patterns-png
import { copyFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { TemplateManifest } from '../src/index.ts';

const [manifestPath, slidesDir, outDir, deckDir] = process.argv.slice(2);
if (!manifestPath || !slidesDir || !outDir) {
  console.error('usage: node scripts/pattern-thumbnails.ts <manifest.json> <rendered-slides-dir> <out-dir> [rendered-pattern-deck-dir]');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as TemplateManifest;
const picture = (dir: string, number: number) => join(dir, `slide-${String(number).padStart(2, '0')}.png`);
mkdirSync(outDir, { recursive: true });
const seen = new Set<number>();
manifest.patterns.forEach((pattern, index) => {
  const derived = seen.has(pattern.sourceSlide) && deckDir !== undefined;
  seen.add(pattern.sourceSlide);
  const source = derived ? picture(deckDir as string, index + 1) : picture(slidesDir, pattern.sourceSlide);
  if (!existsSync(source)) {
    console.error(`missing ${source} for pattern "${pattern.id}"`);
    process.exitCode = 1;
    return;
  }
  copyFileSync(source, join(outDir, `${pattern.id}.png`));
  console.log(`${pattern.id}.png  <- ${derived ? `slide ${index + 1} of the pattern deck` : `slide ${pattern.sourceSlide}`}`);
});
