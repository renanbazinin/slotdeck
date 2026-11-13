// Names the pictures of a template's sample slides after the patterns they belong to, so a
// plan can be shown as a row of layout thumbnails.
//
//   powershell -File scripts/render-pptx.ps1 -Path template.pptx -OutDir out/template-png -Width 480
//   node scripts/pattern-thumbnails.ts templates/demo.manifest.json out/template-png templates/demo.thumbnails
import { copyFileSync, existsSync, mkdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { TemplateManifest } from '../src/index.ts';

const [manifestPath, slidesDir, outDir] = process.argv.slice(2);
if (!manifestPath || !slidesDir || !outDir) {
  console.error('usage: node scripts/pattern-thumbnails.ts <manifest.json> <rendered-slides-dir> <out-dir>');
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as TemplateManifest;
mkdirSync(outDir, { recursive: true });
for (const pattern of manifest.patterns) {
  const source = join(slidesDir, `slide-${String(pattern.sourceSlide).padStart(2, '0')}.png`);
  if (!existsSync(source)) {
    console.error(`missing ${source} for pattern "${pattern.id}"`);
    process.exitCode = 1;
    continue;
  }
  copyFileSync(source, join(outDir, `${pattern.id}.png`));
  console.log(`${pattern.id}.png  <- slide ${pattern.sourceSlide}`);
}
