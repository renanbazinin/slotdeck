// Prints every shape on the chosen slides of a pptx, with the ids a manifest needs.
//
//   node scripts/inspect.ts <template.pptx> [--slides 1,2,10] [--json]
import { readFileSync } from 'node:fs';
import { PptxPackage } from '../src/pptx/package.ts';
import { inspectTemplate } from '../src/template/inspect.ts';

const args = process.argv.slice(2);
const file = args.find((arg) => !arg.startsWith('--'));
if (!file) {
  console.error('usage: node scripts/inspect.ts <template.pptx> [--slides 1,2,10] [--json]');
  process.exit(1);
}
const slidesArg = args[args.indexOf('--slides') + 1];
const only = args.includes('--slides') && slidesArg ? new Set(slidesArg.split(',').map(Number)) : undefined;

const info = inspectTemplate(PptxPackage.load(readFileSync(file)));
const slides = info.slides.filter((slide) => !only || only.has(slide.number));

if (args.includes('--json')) {
  console.log(JSON.stringify({ ...info, slides }, null, 2));
} else {
  const pct = (value: number) => `${Math.round(value * 100)}`.padStart(3);
  for (const slide of slides) {
    console.log(`\nslide ${slide.number}  (${slide.part}, ${slide.layout.split('/').pop()})`);
    for (const shape of slide.shapes) {
      const box = `x${pct(shape.box.x)} y${pct(shape.box.y)} w${pct(shape.box.w)} h${pct(shape.box.h)}`;
      let detail = '';
      if (shape.paragraphs) {
        const text = shape.paragraphs.join(' / ');
        const chars = shape.paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
        detail = `${shape.fontSizes?.join('/') ?? '?'}pt ${shape.paragraphs.length}p ${chars}ch "${text.slice(0, 48)}${text.length > 48 ? '...' : ''}"`;
      } else if (shape.image) {
        const size = shape.image.width ? `${shape.image.width}x${shape.image.height}` : '?';
        detail = `${shape.image.part.split('/').pop()} ${size}${shape.image.cropped ? ' cropped' : ''} ${shape.image.geometry}`;
      }
      const flags = [shape.groupId ? `in #${shape.groupId}` : '', shape.placeholder ? `ph:${shape.placeholder}` : '']
        .filter(Boolean)
        .join(' ');
      console.log(`  #${shape.id.padEnd(5)} ${shape.kind.padEnd(8)} ${box}  ${detail} ${flags}`.trimEnd());
    }
  }
}
