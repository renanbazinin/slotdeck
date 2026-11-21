// What a rendered deck's file says of itself (0.7.0): its title, language, number of slides and
// the version that made it, in place of its template's; its template's picture of its cover gone;
// and the deck kept in the file where that is asked for. No test calls a model.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { test } from 'node:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, SLOTDECK_VERSION, loadTemplate, readEmbeddedDeck, renderDeck, sampleDeck } from '../src/index.ts';
import type { DeckSpec, TemplateManifest } from '../src/index.ts';
import { readRels } from '../src/pptx/package.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const root = new URL('..', import.meta.url);
const demoPptx = new Uint8Array(readFileSync(new URL('templates/demo.pptx', root)));
const demoManifest = JSON.parse(readFileSync(new URL('templates/demo.manifest.json', root), 'utf8')) as TemplateManifest;
const demo = loadTemplate(demoPptx, demoManifest);
const { deck, assets } = sampleDeck(demo);
const parts = (file: Uint8Array) => Object.fromEntries(Object.entries(unzipSync(file)).map(([name, bytes]) => [name, name.match(/\.(xml|rels)$/) ? strFromU8(bytes) : bytes]));
const photo = placeholderPhoto(100, 100, [10, 20, 30], [200, 210, 220]);

test('package.json and SLOTDECK_VERSION name the same version', () => {
  assert.equal(JSON.parse(readFileSync(new URL('package.json', root), 'utf8')).version, SLOTDECK_VERSION);
});

test('without properties a file keeps its template\'s, and with them it says what it is as a deck', () => {
  const plain = renderDeck(demo, deck, { assets, layout: 'adaptive' }).pptx;
  const template = parts(demoPptx);
  const before = parts(plain);
  // As before 0.7.0: the template's properties, which name its 19 sample slides.
  assert.equal(before['docProps/app.xml'], template['docProps/app.xml']);
  assert.equal(before['docProps/core.xml'], template['docProps/core.xml']);
  assert.match(String(template['docProps/app.xml']), /<Slides>\d+<\/Slides>/);

  const file = renderDeck(demo, deck, { assets, layout: 'adaptive', properties: {} }).pptx;
  const after = parts(file);
  const core = String(after['docProps/core.xml']);
  // The deck's title, else the title of its first slide; its language; no author and no date of the template's.
  const title = deck.title ?? (deck.slides[0]?.slots.title as string);
  assert.ok(core.includes(`<dc:title>${title}</dc:title>`), core);
  assert.ok(core.includes('<dc:language>en</dc:language>') && !/dcterms:created|dcterms:modified|dc:creator|lastModifiedBy/.test(core), core);
  const app = String(after['docProps/app.xml']);
  assert.ok(app.includes(`<Slides>${deck.slides.length}</Slides>`) && app.includes('<Notes>0</Notes>') && !/TitlesOfParts|HeadingPairs|<Words>/.test(app), app);
  const custom = String(after['docProps/custom.xml']);
  assert.match(custom, new RegExp(`<property fmtid="\\{D5CDD505-2E9C-101B-9397-08002B2CF9AE\\}" pid="2" name="slotdeck"><vt:lpwstr>${SLOTDECK_VERSION.replace(/\./g, '\\.')}</vt:lpwstr></property>`));
  assert.match(String(after['[Content_Types].xml']), /PartName="\/docProps\/custom\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.custom-properties\+xml"/);
  assert.ok(readRels(PptxPackage.load(file), '').some((rel) => rel.type.endsWith('/custom-properties') && rel.target === 'docProps/custom.xml'));
  // The slides themselves are the same, and so is the same deck rendered again: no date, no time.
  for (const name of Object.keys(before).filter((name) => name.startsWith('ppt/'))) assert.deepEqual(after[name], before[name], name);
  assert.deepEqual(renderDeck(demo, deck, { assets, layout: 'adaptive', properties: {} }).pptx, file);
  // A title, an author and a date given are written, and a date that is none is refused.
  const given = String(parts(renderDeck(demo, deck, { assets, layout: 'adaptive', properties: { title: 'Q3 & beyond', author: 'Dana', date: new Date('2026-11-21T09:30:12.345Z'), language: 'en-GB' } }).pptx)['docProps/core.xml']);
  for (const part of ['<dc:title>Q3 &amp; beyond</dc:title>', '<dc:creator>Dana</dc:creator>', '<cp:lastModifiedBy>Dana</cp:lastModifiedBy>', '<dc:language>en-GB</dc:language>', '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-21T09:30:12Z</dcterms:created>']) assert.ok(given.includes(part), part);
  assert.throws(() => renderDeck(demo, deck, { assets, layout: 'adaptive', properties: { date: 'soon' } }), /properties\.date is "soon"; it is a Date, or a date such as/);
});

test('a template\'s picture of its cover is taken out and its own custom properties stay; one with no properties gets them', () => {
  // The fixture with a thumbnail and a label of how confidential its files are, as an enterprise template has.
  const files = unzipSync(buildFixturePptx());
  const label = '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MSIP_Label_General_Enabled"><vt:lpwstr>true</vt:lpwstr></property>';
  files['docProps/thumbnail.jpeg'] = new Uint8Array([0xff, 0xd8, 0xff, 0xd9]);
  files['docProps/custom.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${label}</Properties>`);
  files['_rels/.rels'] = strToU8(
    strFromU8(files['_rels/.rels'] as Uint8Array).replace(
      '</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>',
    ),
  );
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml'] as Uint8Array).replace('</Types>', '<Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'));
  const template = loadTemplate(zipSync(files), fixtureManifest);
  const slides: DeckSpec = { version: 1, slides: [{ pattern: 'content', slots: { title: 'שלום', body: 'גוף', photo: { asset: 'p' }, note: 'הערה' } }] };
  assert.ok(parts(renderDeck(template, slides, { assets: { p: photo } }).pptx)['docProps/thumbnail.jpeg'], 'kept without properties');
  const after = parts(renderDeck(template, slides, { assets: { p: photo }, properties: {} }).pptx);
  assert.equal(after['docProps/thumbnail.jpeg'], undefined);
  assert.ok(!String(after['_rels/.rels']).includes('thumbnail'));
  assert.ok(String(after['docProps/custom.xml']).includes(label) && String(after['docProps/custom.xml']).includes('pid="3" name="slotdeck"'));
  // A template with no core or app properties gets both; a Hebrew deck is told in Hebrew.
  const core = String(after['docProps/core.xml']);
  assert.ok(core.includes('<dc:title>שלום</dc:title>') && core.includes('<dc:language>he</dc:language>'), core);
  assert.ok(String(after['docProps/app.xml']).includes('<Slides>1</Slides>'));
  for (const [part, type] of [['core', 'application/vnd.openxmlformats-package.core-properties+xml'], ['app', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']] as const) {
    assert.ok(String(after['[Content_Types].xml']).includes(`PartName="/docProps/${part}.xml" ContentType="${type}"`), String(part));
  }
});

test('a deck kept in its file comes out again with its pictures, and renders to the same file', () => {
  const template = loadTemplate(buildFixturePptx(), fixtureManifest);
  const asked: DeckSpec = {
    version: 1,
    brief: 'A private request',
    document: 'A private document',
    slides: [{ pattern: 'content', slots: { title: 'שלום', body: 'גוף', photo: { asset: 'p', alt: 'A photo' }, note: 'הערה' } }],
  };
  const file = renderDeck(template, asked, { assets: { p: photo }, properties: { embedDeck: true } }).pptx;
  const found = readEmbeddedDeck(file);
  assert.ok(found);
  assert.equal(found.version, SLOTDECK_VERSION);
  // What a person asked and gave stays out of the file.
  const { brief: _brief, document: _document, ...rest } = asked;
  assert.deepEqual(found.deck, rest);
  assert.deepEqual(found.assets, { p: photo });
  assert.deepEqual(renderDeck(template, found.deck, { assets: found.assets, properties: { embedDeck: true } }).pptx, renderDeck(template, rest, { assets: { p: photo }, properties: { embedDeck: true } }).pptx);
  // The part is the presentation's, as PowerPoint keeps custom XML.
  const all = parts(file);
  assert.match(String(all['ppt/_rels/presentation.xml.rels']), /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/customXml" Target="\.\.\/customXml\/item1\.xml"/);
  assert.ok(String(all['customXml/itemProps1.xml']).includes('ds:itemID="{6B1D3F0A-5C2E-4E8B-9A7D-3F1C2B4D5E6F}"'));
  assert.ok(!String(all['customXml/item1.xml']).includes('private'));
  assert.equal(readEmbeddedDeck(renderDeck(template, asked, { assets: { p: photo }, properties: {} }).pptx), undefined);
  assert.equal(readEmbeddedDeck(renderDeck(template, asked, { assets: { p: photo } }).pptx), undefined);
});
