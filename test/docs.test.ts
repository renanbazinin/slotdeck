// The documentation in docs/ ships inside the package and is read where there is no network, so
// it must not drift from the code: every export is described, every import it shows exists, its
// examples are valid, and it links to nothing outside the folder.
import assert from 'node:assert/strict';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { ROLES, TOOL_NAMES, changeManifest, presentationTools, memoryStore, loadTemplate, validateDeck } from '../src/index.ts';
import type { DeckSpec, ManifestChange, TemplateManifest } from '../src/index.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const docsDir = join(root, 'docs');
const pages = readdirSync(docsDir).filter((name) => name.endsWith('.md'));
const read = (name: string) => readFileSync(join(docsDir, name), 'utf8');

/** Names exported by an entry file, values and types alike. */
function exportsOf(file: string): string[] {
  const source = readFileSync(join(root, 'src', file), 'utf8');
  return [...source.matchAll(/export (?:type )?\{([^}]+)\}/g)].flatMap((match) =>
    (match[1] as string)
      .split(',')
      .map((name) => name.trim())
      .filter(Boolean),
  );
}
const mainExports = exportsOf('index.ts');
const nodeExports = ['directoryStore'];

test('the documentation covers the pages its index promises', () => {
  const index = read('README.md');
  const linked = [...index.matchAll(/\]\(([a-z-]+\.md)\)/g)].map((match) => match[1] as string);
  assert.ok(linked.length >= 12);
  for (const page of linked) assert.ok(pages.includes(page), `${page} is linked from the index but missing`);
  for (const page of pages) assert.ok(page === 'README.md' || linked.includes(page), `${page} is not listed in the index`);
});

test('every export is described in the API reference', () => {
  const reference = read('api-reference.md');
  for (const name of [...mainExports, ...nodeExports]) {
    assert.match(reference, new RegExp(`\\b${name}\\b`), `${name} is exported but not in docs/api-reference.md`);
  }
});

test('every import shown in the documentation exists', () => {
  for (const file of [...pages.map((page) => join(docsDir, page)), join(root, 'README.md')]) {
    const text = readFileSync(file, 'utf8');
    for (const match of text.matchAll(/import (?:type )?\{([^}]+)\} from '(slotdeck(?:\/node)?)'/g)) {
      const known = match[2] === 'slotdeck' ? mainExports : nodeExports;
      for (const name of (match[1] as string).split(',').map((part) => part.replace(/\/\*.*?\*\//g, '').trim()).filter(Boolean)) {
        assert.ok(known.includes(name), `${file} imports ${name} from '${match[2]}', which is not exported`);
      }
    }
  }
});

test('the documentation links only to its own pages', () => {
  for (const page of pages) {
    const text = read(page);
    assert.doesNotMatch(text, /\]\(https?:/, `${page} links to the internet`);
    for (const match of text.matchAll(/\]\(([^)#]+)(#[^)]*)?\)/g)) {
      assert.ok(existsSync(join(docsDir, match[1] as string)), `${page} links to ${match[1]}, which does not exist`);
    }
  }
});

test('the complete deck example is a valid deck for the demo template', () => {
  const example = /## Complete example\s+```json\n([\s\S]*?)\n```/.exec(read('deck-format.md'))?.[1];
  assert.ok(example, 'deck-format.md has a complete example');
  const deck = JSON.parse(example) as DeckSpec;
  const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
  assert.deepEqual(validateDeck(manifest, deck, { layout: 'adaptive', assets: ['photo-commute', 'clock', 'shield', 'people'] }), []);
});

test('the studio recipes work on the demo template as written', () => {
  const recipes = read('template-studio.md').split('### Recipes')[1]?.split('\n## ')[0] ?? '';
  const blocks = [...recipes.matchAll(/```ts\n([\s\S]*?)\n```/g)].map((match) => match[1] as string);
  assert.equal(blocks.length, 8);
  const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
  const shipped = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
  // The recipes make derived patterns the demo manifest already has, so start without them.
  const start = changeManifest(loadTemplate(pptx, shipped), [
    { op: 'removePattern', pattern: 'chart-native' },
    { op: 'removePattern', pattern: 'text-only' },
    { op: 'removePattern', pattern: 'photo-stat-mirrored' },
  ]).manifest;
  for (const block of blocks) {
    // A recipe is a change or a list of changes, written as a JavaScript literal.
    const value = new Function(`return (${block.replace(/\/\/.*$/gm, '')});`)() as ManifestChange | ManifestChange[];
    const result = changeManifest(loadTemplate(pptx, start), Array.isArray(value) ? value : [value]);
    assert.deepEqual(result.rejected.map((entry) => entry.reason), [], block);
    assert.deepEqual(result.issues.filter((issue) => issue.severity === 'error'), [], block);
  }
});

test('operations, roles and tools are documented under their real names', () => {
  const studio = read('template-studio.md');
  for (const op of ['setManifest', 'addPattern', 'copyPattern', 'derivePattern', 'removePattern', 'setPattern', 'addSlot', 'removeSlot', 'setSlot', 'setRepeat', 'removeRepeat']) {
    assert.match(studio, new RegExp(`#### \`${op}\``), `${op} has no section in template-studio.md`);
  }
  for (const role of ROLES) assert.match(studio, new RegExp(`\`${role}\``), `role ${role} is not in template-studio.md`);

  const chat = read('chat-tools.md');
  const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
  const tools = presentationTools({
    template: loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), manifest),
    llm: { generateJson: async () => ({}) },
    store: memoryStore(),
  });
  assert.deepEqual(tools.definitions.map((tool) => tool.name), Object.values(TOOL_NAMES));
  for (const tool of tools.definitions) {
    assert.match(chat, new RegExp(`### \`${tool.name}\``), `${tool.name} has no section in chat-tools.md`);
    // Every argument a tool accepts is named in its page.
    for (const argument of Object.keys((tool.parameters as { properties: Record<string, unknown> }).properties)) {
      assert.match(chat, new RegExp(`\`${argument}\``), `argument ${argument} of ${tool.name} is not in chat-tools.md`);
    }
  }
});
