// Shows a plan as a page of layout thumbnails: what a chat app could show in place of a text
// list. The page is self-contained (the thumbnails are embedded), so it opens anywhere.
//
//   node examples/plan-preview.ts out/store/plan/<id>.json [out/plan-preview.html]
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { directionOf } from '../src/index.ts';
import type { DeckPlan } from '../src/index.ts';

const escape = (text: unknown) => String(text ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char] as string);

/**
 * `thumbnailsDir` holds one picture per pattern, `<pattern id>.png` or `<pattern id>.svg`
 * (see scripts/pattern-thumbnails.ts, which makes either).
 */
export function planPreviewHtml(plan: DeckPlan, thumbnailsDir: string, fallbackLanguage = 'en'): string {
  const language = plan.brief.language ?? fallbackLanguage;
  const direction = plan.brief.direction ?? directionOf(language);
  const cards = plan.slides
    .map((slide, index) => {
      const found = [
        { file: join(thumbnailsDir, `${slide.pattern}.png`), type: 'image/png' },
        { file: join(thumbnailsDir, `${slide.pattern}.svg`), type: 'image/svg+xml' },
      ].find((candidate) => existsSync(candidate.file));
      const picture = found
        ? `<img src="data:${found.type};base64,${readFileSync(found.file).toString('base64')}" alt="Layout: ${escape(slide.pattern)}">`
        : `<div class="missing">${escape(slide.pattern)}</div>`;
      return `<li>${picture}<div class="text"><b>${index + 1}</b><span>${escape(slide.pattern)}</span><p>${escape(slide.purpose)}</p></div></li>`;
    })
    .join('\n');

  return `<!doctype html>
<html lang="${escape(language)}" dir="${direction}">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escape(plan.title || 'Plan preview')}</title>
<style>
  :root { --ground: #f4f5f4; --card: #ffffff; --line: #dcddde; --ink: #0c0c0c; --muted: #5f6368; --accent: #00683a; }
  @media (prefers-color-scheme: dark) {
    :root { --ground: #141615; --card: #1e2120; --line: #343837; --ink: #f1f3f2; --muted: #a3a8a6; --accent: #61ff74; }
  }
  body { margin: 0; padding: 24px 16px; font: 15px/1.45 system-ui, sans-serif; background: var(--ground); color: var(--ink); }
  h1 { font-size: 20px; font-weight: 600; margin: 0 auto 16px; max-width: 1100px; }
  ol { list-style: none; margin: 0 auto; padding: 0; max-width: 1100px; display: grid; gap: 16px; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); }
  li { background: var(--card); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; }
  img, .missing { display: block; width: 100%; aspect-ratio: 16 / 9; object-fit: cover; border-bottom: 1px solid var(--line); }
  .missing { display: grid; place-items: center; color: var(--muted); }
  .text { padding: 10px 12px 12px; }
  b { display: inline-block; min-width: 22px; color: var(--accent); font-variant-numeric: tabular-nums; }
  span { color: var(--muted); font-size: 12px; }
  p { margin: 6px 0 0; }
</style>
<h1>${escape(plan.title)}</h1>
<ol>
${cards}
</ol>
</html>
`;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const [planPath, outPath = 'out/plan-preview.html'] = process.argv.slice(2);
  if (!planPath) {
    console.error('usage: node examples/plan-preview.ts <plan.json> [out.html]');
    process.exit(1);
  }
  const root = join(fileURLToPath(import.meta.url), '..', '..');
  const plan = JSON.parse(readFileSync(planPath, 'utf8')) as DeckPlan;
  writeFileSync(outPath, planPreviewHtml(plan, join(root, 'templates', 'demo.thumbnails'), 'en'));
  console.log(outPath);
}
