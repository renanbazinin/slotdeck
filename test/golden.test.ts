// What slotdeck 0.5.2 rendered. test/fixtures/golden-0.5.2.json holds, for eleven decks made from
// the repository's own material, a hash of the parts of the file and the number of warnings. It was
// taken from the code of 0.5.2, and so were the hashes written further down. A deck that an app has
// stored must render the same after an upgrade: a change to the fixture or to one of those hashes
// is a decision to change what stored decks look like, and is never made to get this test to pass.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readdirSync, readFileSync } from 'node:fs';
import { basename, dirname, extname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { placeholderChart, placeholderPhoto } from '../examples/placeholder.ts';
import { checkManifest, designKit, loadTemplate, prepareManifest, renderDeck, sampleDeck, templateFonts } from '../src/index.ts';
import type { DeckSpec, DesignElement, PatternDesign, RenderOptions, RenderResult, SlideSpec, Template, TemplateManifest } from '../src/index.ts';
import { buildFixturePptx, cardsManifest, fixtureManifest, tableManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const golden = JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'golden-0.5.2.json'), 'utf8')) as Record<string, { parts: string; warnings: number }>;

const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = (): Template => loadTemplate(pptx, manifest);
const readDeck = (...path: string[]): DeckSpec => JSON.parse(readFileSync(join(root, ...path), 'utf8')) as DeckSpec;

/**
 * A hash of the parts of a pptx, names and bytes, in the order of the file. It is taken from the
 * unzipped parts, so it does not depend on how the file was packed. Pictures are parts too, and
 * the stand-in pictures are packed by zlib when they are drawn: another zlib can move a hash with
 * no change in the renderer.
 */
function partsHash(file: Uint8Array): string {
  const hash = createHash('sha256');
  for (const [name, data] of Object.entries(unzipSync(file))) {
    hash.update(name);
    hash.update(data);
  }
  return hash.digest('hex').slice(0, 16);
}

/** A hash of a value as JSON, for what the template answers without rendering. */
const valueHash = (value: unknown): string => createHash('sha256').update(JSON.stringify(value)).digest('hex').slice(0, 16);

// The pictures the example decks ask for, drawn in code, and the icons of the examples.
const icons = (JSON.parse(readFileSync(join(root, 'examples', 'icons', 'icons.json'), 'utf8')) as Array<{ id: string }>).map(
  ({ id }): [string, Uint8Array] => [id, new Uint8Array(readFileSync(join(root, 'examples', 'icons', `${id}.png`)))],
);
const standIns: Record<string, Uint8Array> = {
  'photo-tall': placeholderPhoto(1100, 1500, [14, 124, 123], [244, 162, 89]),
  'photo-square': placeholderPhoto(1400, 1400, [27, 42, 65], [14, 124, 123]),
  'photo-wide': placeholderPhoto(1600, 1350, [244, 162, 89], [232, 241, 240]),
  'chart-months': placeholderChart(1500, 980, [0.3, 0.38, 0.45, 0.55, 0.7, 0.92, 0.85, 0.6]),
  'chart-wide': placeholderChart(2400, 760, [0.35, 0.5, 0.42, 0.7, 0.62, 0.9, 0.75, 0.82, 0.58, 0.95]),
  ...Object.fromEntries(icons),
};
const withStandIns: RenderOptions = { assets: standIns, layout: 'adaptive' };

/** The showcase deck with the pictures saved beside it, each under its file name without the ending. */
function showcase(): RenderResult {
  const dir = join('examples', 'showcase', 'rescue-dog');
  const saved: Record<string, Uint8Array> = { ...standIns };
  for (const file of readdirSync(join(root, dir, 'assets'))) saved[basename(file, extname(file))] = new Uint8Array(readFileSync(join(root, dir, 'assets', file)));
  return renderDeck(demo(), readDeck(dir, 'deck.json'), { assets: saved, layout: 'adaptive' });
}

/** A deck that fills every pattern of the demo template, with the fewest or the most items in each row. */
function sample(items: 'min' | 'max'): RenderResult {
  const template = demo();
  const { deck, assets } = sampleDeck(template, { items });
  return renderDeck(template, deck, { assets, layout: 'adaptive' });
}

/** A deck with a layout of its own, as test/deck-layouts.test.ts builds it: three goals on cards under the title of text-only. */
function ownLayout(): RenderResult {
  const template = demo();
  const kit = designKit(template);
  assert.ok(kit, 'the demo template has a design kit');
  const styleFor = (use: string): string => {
    const style = kit.styles.find((candidate) => candidate.text.includes(`used for ${use}`));
    assert.ok(style, `the kit has a style used for ${use}`);
    return style.key;
  };
  const panel = kit.styles.find((style) => style.kind === 'decoration' && style.shape === 'area' && style.text.includes('panel'));
  assert.ok(panel, 'the kit has a panel');
  const elements: DesignElement[] = [];
  for (const [n, column] of [[1, 1], [2, 5], [3, 9]] as const) {
    elements.push({ name: '', style: panel.key, role: 'decoration', optional: false, sample: '', column, row: 6, width: 4, height: 7 });
    elements.push({ name: `goal${n}Heading`, style: styleFor('card1Heading'), role: 'subtitle', optional: false, sample: `Goal number ${n}`, column, row: 6, width: 4, height: 2 });
    elements.push({ name: `goal${n}Text`, style: styleFor('card1Text'), role: 'text', optional: false, sample: 'A sentence or two on what reaching this goal means.', column, row: 8, width: 4, height: 5 });
  }
  const design: PatternDesign = { id: 'three-goals', description: 'Three goals side by side on cards, each with a heading and a sentence.', keep: ['title'], base: 'text-only', elements };
  const deck: DeckSpec = {
    version: 1,
    fields: { organizationDepartment: 'Strategy office' },
    layouts: [design],
    slides: [
      { pattern: 'cover', slots: { title: 'The year ahead' } },
      { pattern: 'three-goals', slots: { title: 'Three goals for the year', goal1Heading: 'Grow', goal1Text: 'Win two new regions.', goal2Heading: 'Keep', goal2Text: 'Keep every client we have.', goal3Heading: 'Learn', goal3Text: 'Train every team twice.' } },
      { pattern: 'closing', slots: { title: 'Thank you' } },
    ],
  };
  return renderDeck(template, deck, { layout: 'adaptive' });
}

// The small template of the tests: plain, turned left to right, a row of five, a table.
const fixture = buildFixturePptx();
const photo = placeholderPhoto(100, 100, [0, 0, 0], [255, 255, 255]);
const content: SlideSpec = { pattern: 'content', slots: { title: 'First', body: 'Body', photo: { asset: 'photo' }, note: 'Second' } };

/** Every deck of the fixture, under the name the fixture gives it. */
const decks: Record<string, () => RenderResult> = {
  'patterns.en.deck.json': () => renderDeck(demo(), readDeck('examples', 'decks', 'patterns.en.deck.json'), withStandIns),
  'patterns.en.deck.json textScale 1.25': () => renderDeck(demo(), { ...readDeck('examples', 'decks', 'patterns.en.deck.json'), textScale: 1.25 }, withStandIns),
  'patterns.he.deck.json': () => renderDeck(demo(), readDeck('examples', 'decks', 'patterns.he.deck.json'), withStandIns),
  'showcase rescue-dog': showcase,
  'sampleDeck demo min': () => sample('min'),
  'sampleDeck demo max': () => sample('max'),
  'deck with its own layout': ownLayout,
  'fixture rtl': () =>
    renderDeck(loadTemplate(fixture, fixtureManifest), { version: 1, slides: [content, { pattern: 'closing', slots: { title: 'Bye' } }] }, { assets: { photo } }),
  'fixture ltr': () =>
    renderDeck(loadTemplate(fixture, fixtureManifest), { version: 1, direction: 'ltr', fields: { closing: 'Bye' }, slides: [content, { pattern: 'closing', slots: {} }] }, { assets: { photo } }),
  'fixture cards x5': () =>
    renderDeck(loadTemplate(fixture, cardsManifest), { version: 1, slides: [{ pattern: 'cards', items: 5, slots: { title: 'T', card1Text: 'a', card2Text: 'b', card3Text: 'c', card4Text: 'd', card5Text: 'e' } }] }),
  'fixture table': () =>
    renderDeck(loadTemplate(fixture, tableManifest), { version: 1, slides: [{ pattern: 'table', slots: { title: 'T', table: { header: ['A', 'B'], rows: [['1', '2'], ['3', '4']] } } }] }),
};

/**
 * Decks whose file was changed on purpose since 0.5.2: the version, what changed, and the hash
 * the deck has since. Each is a decision written in docs/changelog.md, and the 0.5.2 hash stays
 * in the fixture as the record of what was there before.
 */
const changedSince: Record<string, { version: string; change: string; parts: string }> = {
  // The English title "First" is written into the template's run marked he-IL. One attribute
  // changes, lang="he-IL" to lang="en-US"; the rest of the file is as 0.5.2 wrote it.
  'fixture rtl': { version: '0.6.5', change: 'English text in a run the template marks Hebrew is marked English', parts: 'cd85de46fd310f5c' },
  'fixture ltr': { version: '0.6.5', change: 'English text in a run the template marks Hebrew is marked English', parts: '87d825a0520d2578' },
};

test('the fixture holds the decks rendered here, and no other', () => {
  assert.deepEqual(Object.keys(decks).sort(), Object.keys(golden).sort());
  for (const [name, changed] of Object.entries(changedSince)) assert.notEqual(changed.parts, golden[name]?.parts, `"${name}" is listed as changed but has its 0.5.2 hash`);
});

for (const [name, render] of Object.entries(decks)) {
  test(`renders as 0.5.2 did: ${name}`, () => {
    const { pptx: file, warnings } = render();
    const said = warnings.map((issue) => issue.message).join('; ');
    const changed = changedSince[name];
    const expected = changed ? { parts: changed.parts, warnings: golden[name]?.warnings } : golden[name];
    const what = changed ? `what ${changed.version} renders (${changed.change})` : 'what 0.5.2 rendered';
    assert.deepEqual({ parts: partsHash(file), warnings: warnings.length }, expected, `"${name}" is not ${what}${said ? `; it warns: ${said}` : ''}`);
  });
}

test('the design kit of the demo template is the one 0.5.2 gave', () => {
  const kit = designKit(demo());
  assert.equal(kit?.base, 'list-with-photo');
  assert.equal(valueHash(kit), '86af9ec327d461d0');
});

test('the sample deck of the demo template is the one 0.5.2 made', () => {
  const { deck } = sampleDeck(demo());
  assert.equal(deck.slides.length, 18);
  assert.equal(valueHash(deck), '6f7497d2e3552faf');
});

test('preparing the demo manifest, icons on and nothing removed, gives the manifest 0.5.2 gave', () => {
  assert.equal(valueHash(prepareManifest(demo(), { icons: true, remove: false }).manifest), '0322d945d565590f');
});

test('the demo manifest has nothing to report, and the template asks for the same five fonts', () => {
  const template = demo();
  assert.deepEqual(checkManifest(template), []);
  assert.deepEqual(
    templateFonts(template).map((font) => `${font.typeface}${font.bold ? ' bold' : ''}${font.italic ? ' italic' : ''}`),
    ['Calibri', 'Calibri italic', 'Calibri bold', 'Cambria italic', 'Cambria bold'],
  );
});

// Decks made from a person's document, as 0.6.5 wrote them (test/fixtures/documents-0.6.5.json):
// dense text slides and a layout slide whose text code set from passages, a document heading
// among its paragraphs as a line of plain text. A deck stored then renders the same after an
// upgrade, as the decks above.
const documentDecks = JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'documents-0.6.5.json'), 'utf8')) as Record<string, DeckSpec>;
const documentHashes: Record<string, string> = {
  'document en dense x2': 'd2bff389bf5fdb32',
  'document he dense': '4c92531341d3865b',
  'document en text-only': '444161b5989321e0',
};

for (const [name, deck] of Object.entries(documentDecks)) {
  test(`renders as 0.6.5 did: ${name}`, () => {
    const { pptx: file, warnings } = renderDeck(demo(), deck, { layout: 'adaptive' });
    assert.deepEqual({ parts: partsHash(file), warnings: warnings.length }, { parts: documentHashes[name], warnings: 0 }, `"${name}" is not what 0.6.5 rendered`);
  });
}
