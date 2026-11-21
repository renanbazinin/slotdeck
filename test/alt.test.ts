// What a screen reader is told of what the renderer draws (0.7.0): a pictogram no one described is
// decoration, a table says its headings and how many rows it has. Charts are in chart.test.ts.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { PptxPackage, loadTemplate, renderDeck } from '../src/index.ts';
import type { DeckSpec, ImageRef } from '../src/index.ts';
import { tableAlt } from '../src/render/alt.ts';
import { buildFixturePptx, fixtureManifest, fixtureWithSlide, tableManifest } from './fixture.ts';

const photo = placeholderPhoto(100, 100, [10, 20, 30], [200, 210, 220]);
const DECORATIVE = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative val="1" xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative"/></a:ext></a:extLst>';
const slideOf = (file: Uint8Array) => PptxPackage.load(file).readText('ppt/slides/slide1.xml');
const content = (icon: ImageRef, photoAlt?: string): DeckSpec => ({
  version: 1,
  slides: [{ pattern: 'content', slots: { title: 'שלום', body: 'גוף', photo: { asset: 'p', ...(photoAlt ? { alt: photoAlt } : {}) }, icon } }],
});

test('a pictogram no one described is marked as decoration; one with words, and any photo, is not', () => {
  const template = loadTemplate(buildFixturePptx(), fixtureManifest);
  const plain = slideOf(renderDeck(template, content({ asset: 'p' }), { assets: { p: photo }, layout: 'adaptive' }).pptx);
  assert.ok(plain.includes(`<p:cNvPr id="8" name="Icon" descr="">${DECORATIVE}</p:cNvPr>`), plain);
  assert.equal(plain.split('adec:decorative').length - 1, 1, 'one mark: the photo is not decoration');
  const told = slideOf(renderDeck(template, content({ asset: 'p', alt: 'A shield' }, 'The pier'), { assets: { p: photo }, layout: 'adaptive' }).pptx);
  assert.ok(told.includes('<p:cNvPr id="8" name="Icon" descr="A shield"/>') && told.includes('descr="The pier"') && !told.includes('decorative'));
  // A template that marks its own pictogram as decoration loses the mark where the deck describes it.
  const marked = loadTemplate(fixtureWithSlide(1, (xml) => xml.replace('<p:cNvPr id="8" name="Icon"/>', `<p:cNvPr id="8" name="Icon">${DECORATIVE}</p:cNvPr>`)), fixtureManifest);
  assert.ok(!slideOf(renderDeck(marked, content({ asset: 'p', alt: 'A shield' }), { assets: { p: photo }, layout: 'adaptive' }).pptx).includes('decorative'));
  assert.equal(slideOf(renderDeck(marked, content({ asset: 'p' }), { assets: { p: photo }, layout: 'adaptive' }).pptx).split('adec:decorative').length - 1, 1, 'marked once, not twice');
});

test('a table says its headings and how many rows it has, in the language of the deck', () => {
  assert.equal(tableAlt({ header: ['Station', 'Docks'], rows: [['A', '1'], ['B', '2']] }), 'Table: Station, Docks; 2 rows');
  assert.equal(tableAlt({ header: ['Station', ''], rows: [['A', '1']] }), 'Table: Station; 1 row');
  assert.equal(tableAlt({ header: ['תחנה', 'עגינות'], rows: [['א', '1'], ['ב', '2']] }, 'he', 'rtl'), 'טבלה: תחנה, עגינות; 2 שורות');
  assert.equal(tableAlt({ header: ['תחנה'], rows: [['א']] }, 'he-IL', 'rtl'), 'טבלה: תחנה; שורה אחת');
  const template = loadTemplate(buildFixturePptx(), tableManifest);
  const deck = (language: string, direction: 'rtl' | 'ltr'): DeckSpec => ({ version: 1, language, direction, slides: [{ pattern: 'table', slots: { title: 'T', table: { header: ['Station', 'Docks'], rows: [['A', '1']] } } }] });
  assert.match(slideOf(renderDeck(template, deck('en', 'ltr')).pptx), /descr="Table: Station, Docks; 1 row"/);
  assert.match(slideOf(renderDeck(template, deck('he', 'rtl')).pptx), /descr="טבלה: Station, Docks; שורה אחת"/);
});
