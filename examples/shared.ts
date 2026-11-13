// Setup shared by the demos: the demo template, its icon set and the model providers.
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gemini, geminiImages, loadTemplate } from '../src/index.ts';
import type { AssetDef, GenerateProgress, IconDef, ImageProvider, LlmProvider, Template, TemplateManifest } from '../src/index.ts';
import { placeholderChart, placeholderPhoto } from './placeholder.ts';

export const root = join(dirname(fileURLToPath(import.meta.url)), '..');

/** The demo template with its manifest. Both were made for this repository and are free to reuse. */
export function demoTemplate(): Template {
  const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
  return loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), manifest);
}

/** The icon set an app would register: here twelve drawn pictograms. */
export const icons: IconDef[] = (JSON.parse(readFileSync(join(root, 'examples', 'icons', 'icons.json'), 'utf8')) as Array<{ id: string; description: string }>).map(
  ({ id, description }) => ({ id, description, data: new Uint8Array(readFileSync(join(root, 'examples', 'icons', `${id}.png`))) }),
);

/** Two drawn charts standing in for pictures an app would supply, such as rendered Mermaid diagrams. */
export const sampleCharts: AssetDef[] = [
  {
    id: 'chart-by-month',
    kind: 'graphic',
    description: 'Bar chart of a monthly figure over the year, rising towards the summer',
    data: placeholderChart(1500, 980, [0.3, 0.38, 0.45, 0.55, 0.7, 0.92, 0.85, 0.6]),
  },
  {
    id: 'chart-wide',
    kind: 'graphic',
    description: 'Wide bar chart of a figure per month, growing through the year',
    data: placeholderChart(2400, 760, [0.35, 0.5, 0.42, 0.7, 0.62, 0.9, 0.75, 0.82, 0.58, 0.95]),
  },
];

/** Draws a stand-in at the frame's shape. Costs nothing. */
export const placeholderImages: ImageProvider = {
  async generateImage({ aspectRatio }) {
    const height = 900;
    return placeholderPhoto(Math.round(height * aspectRatio), height, [14, 124, 123], [244, 162, 89]);
  },
};

/** The text model, and either stand-in photos (default) or the paid image model. */
export function providers(realImages: boolean): { llm: LlmProvider; images: ImageProvider } {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.error('GEMINI_API_KEY is not set. Put it in .env and run with: node --env-file=.env examples/<demo>.ts ...');
    process.exit(1);
  }
  return {
    llm: gemini({ model: process.env.SLOTDECK_TEXT_MODEL ?? 'gemini-3.8-flash', apiKey }),
    images: realImages
      ? geminiImages({
          model: process.env.SLOTDECK_IMAGE_MODEL ?? 'gemini-3.1-flash-image',
          apiKey,
          style: 'Natural-light editorial photograph. No text, captions or logos.',
        })
      : placeholderImages,
  };
}

export function logProgress(event: GenerateProgress): void {
  if (event.stage === 'images') console.log(`  images ${event.done}/${event.total}`);
  else if (event.stage === 'repair') console.log(`  repair ${event.attempt}: ${event.issues} issue(s) sent back`);
  else console.log(`  ${event.stage}`);
}
