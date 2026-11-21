// The language and direction a filled paragraph is written in. PowerPoint orders the words of a
// run by the run's language as well as by the paragraph's direction: English in a run marked
// Hebrew comes out as "reached landing just ,74% 75% shy of our ambitioustarget" for "reached 74%,
// landing just shy of our ambitious 75% target". It also checks spelling in the run's language.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { directionOf, freeSlide, loadTemplate, renderDeck } from '../src/index.ts';
import type { ChartValue, DeckSpec, TemplateManifest } from '../src/index.ts';
import { NS, childElement, descendants, parseXml, rootElement } from '../src/pptx/xml.ts';
import { filledLanguage, marksRightToLeft, paragraphDirection, setBodyText } from '../src/render/text.ts';
import type { TextOptions } from '../src/render/text.ts';
import { isRtlLanguage, leftToRightTag } from '../src/util/language.ts';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { buildFixturePptx, fixtureManifest, tableManifest } from './fixture.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Fills one paragraph whose template run states `lang` (none when empty) and reads back what was written. */
function fill(text: string, lang: string, options: Partial<TextOptions> = {}, endLang = lang, pPr = '<a:pPr algn="r" rtl="1"/>') {
  const stated = lang ? ` lang="${lang}"` : '';
  const end = endLang ? ` lang="${endLang}"` : '';
  const doc = parseXml(`<a:txBody xmlns:a="${A}"><a:bodyPr/><a:p>${pPr}<a:r><a:rPr${stated} sz="1600"/><a:t>sample</a:t></a:r><a:endParaRPr${end} sz="1600"/></a:p></a:txBody>`);
  setBodyText(rootElement(doc), [text], { flipAlign: false, ...options });
  const paragraph = descendants(doc, NS.a, 'p')[0];
  const run = descendants(doc, NS.a, 'rPr')[0];
  const endProps = paragraph ? childElement(paragraph, NS.a, 'endParaRPr') : undefined;
  return {
    rtl: paragraph ? childElement(paragraph, NS.a, 'pPr')?.getAttribute('rtl') : undefined,
    lang: run?.getAttribute('lang') ?? null,
    end: endProps?.getAttribute('lang') ?? null,
  };
}

test('right-to-left languages are known by one rule, whatever the tag looks like', () => {
  for (const tag of ['he', 'he-IL', 'iw', 'iw-IL', 'HE-il', 'yi', 'ar-SA', 'fa', 'ur-PK', 'ps', 'ug', 'sd', 'ckb', 'ckb-IQ', 'syr', 'syr-SY', 'dv-MV', 'he_IL']) {
    assert.equal(isRtlLanguage(tag), true, tag);
  }
  for (const tag of ['', 'en', 'en-US', 'en-IL', 'hello', 'heb', 'arn', 'fr-FR', 'ru']) assert.equal(isRtlLanguage(tag), false, tag);
  // The planner's direction of a deck language follows the same rule.
  assert.equal(directionOf('iw'), 'rtl');
  assert.equal(directionOf('ckb'), 'rtl');
  assert.equal(directionOf('syr'), 'rtl');
  assert.equal(directionOf('en-US'), 'ltr');
});

test('the left-to-right language a run is given is the deck language when it reads left to right', () => {
  assert.equal(leftToRightTag(undefined), 'en-US');
  assert.equal(leftToRightTag('he'), 'en-US');
  assert.equal(leftToRightTag('en'), 'en-US');
  assert.equal(leftToRightTag('en-GB'), 'en-GB');
  assert.equal(leftToRightTag('fr'), 'fr-FR');
  assert.equal(leftToRightTag('de-CH'), 'de-CH');
  assert.equal(leftToRightTag(' '), 'en-US');
});

test('English written into a run marked Hebrew gets a left-to-right language, on its run and its paragraph end', () => {
  const english = 'The digital transaction share reached 74%, landing just shy of our ambitious 75% target.';
  assert.deepEqual(fill(english, 'he-IL', { direction: 'ltr', language: 'en' }), { rtl: '0', lang: 'en-US', end: 'en-US' });
  assert.deepEqual(fill('2027 Strategic Roadmap', 'he-IL', { direction: 'ltr', language: 'en-GB' }), { rtl: '0', lang: 'en-GB', end: 'en-GB' });
  // Without a deck language, or in a Hebrew deck, English text is taken to be English.
  assert.deepEqual(fill('Microsoft Teams', 'he-IL'), { rtl: '0', lang: 'en-US', end: 'en-US' });
  assert.deepEqual(fill('KPI', 'he-IL', { direction: 'rtl', language: 'he' }), { rtl: '0', lang: 'en-US', end: 'en-US' });
  // An end of paragraph marked with a language of its own that is not right to left keeps it.
  assert.deepEqual(fill('Revenue', 'he-IL', {}, 'en-GB'), { rtl: '0', lang: 'en-US', end: 'en-GB' });
});

test('a figure without letters reads left to right in a left-to-right deck, and as the template has it in a right-to-left one', () => {
  // In a paragraph marked right to left, "+12%" shows as "12%+" and "4.6 / 5" as "5 / 4.6".
  const english = { direction: 'ltr', language: 'en', templateDirection: 'rtl' } as const;
  for (const text of ['+12%', '-3.5%', '4.6 / 5', '2027 – 2029', '68%', '₪120']) {
    assert.deepEqual(fill(text, 'he-IL', english), { rtl: '0', lang: 'en-US', end: 'en-US' }, text);
    assert.deepEqual(fill(text, 'he-IL', { direction: 'rtl', language: 'he', templateDirection: 'rtl' }), { rtl: '1', lang: 'he-IL', end: 'he-IL' }, text);
  }
  // A paragraph that inherits its direction from a right-to-left template is turned too.
  assert.deepEqual(fill('+12%', 'he-IL', english, 'he-IL', '<a:pPr algn="r"/>'), { rtl: '0', lang: 'en-US', end: 'en-US' });
  // On a left-to-right template a paragraph that states no direction is left as it is.
  assert.deepEqual(fill('+12%', 'en-US', { direction: 'ltr', language: 'en', templateDirection: 'ltr' }, 'en-US', '<a:pPr algn="l"/>'), { rtl: null, lang: 'en-US', end: 'en-US' });
  // Text with neither letters nor digits keeps what the template gave it.
  assert.deepEqual(fill('—', 'he-IL', english), { rtl: '1', lang: 'he-IL', end: 'he-IL' });
  // Without a deck direction nothing is turned.
  assert.deepEqual(fill('+12%', 'he-IL'), { rtl: '1', lang: 'he-IL', end: 'he-IL' });
});

test('Hebrew text, and text that is mostly Hebrew, keeps its Hebrew language', () => {
  assert.deepEqual(fill('תוכנית העבודה לרבעון הבא', 'he-IL', { direction: 'rtl', language: 'he' }), { rtl: '1', lang: 'he-IL', end: 'he-IL' });
  assert.deepEqual(fill('AI בשירות הלקוחות', 'he-IL', { direction: 'ltr', language: 'en' }), { rtl: '0', lang: 'he-IL', end: 'he-IL' });
});

test('in a right-to-left deck, a paragraph with any Hebrew letter reads right to left and is marked Hebrew', () => {
  // It starts with a Latin word: the first letter used to decide, and the bullet went to the left.
  assert.deepEqual(fill('AI בשירות הלקוחות', 'he-IL', { direction: 'rtl', language: 'he' }), { rtl: '1', lang: 'he-IL', end: 'he-IL' });
  assert.deepEqual(fill('Microsoft Teams ככלי עבודה', 'he-IL', { direction: 'rtl', language: 'he' }), { rtl: '1', lang: 'he-IL', end: 'he-IL' });
  // On a template whose runs are English, such a paragraph gets the Hebrew language too.
  assert.deepEqual(fill('Microsoft Teams ככלי עבודה', 'en-US', { direction: 'rtl', language: 'he' }), { rtl: '1', lang: 'he-IL', end: 'en-US' });
  // In a left-to-right deck the same paragraph keeps the rule of its first letter and of most of its letters.
  assert.deepEqual(fill('Microsoft Teams ככלי עבודה', 'en-US', { direction: 'ltr', language: 'en' }), { rtl: '0', lang: 'en-US', end: 'en-US' });
});

test('a run that states no language of its own is left to the language it inherits', () => {
  assert.deepEqual(fill('Revenue grew', '', { direction: 'ltr', language: 'en' }, ''), { rtl: '0', lang: null, end: null });
  assert.deepEqual(fill('הכנסות', '', { direction: 'rtl', language: 'he' }, ''), { rtl: '1', lang: 'he-IL', end: null });
});

test('filledLanguage, the one rule the renderer and the measurer use, gives the language of a filled run', () => {
  // The measurer's side is tested in text-fit.test.ts, through the measurer itself.
  assert.equal(filledLanguage('Revenue 2026', 'he-IL', { direction: 'ltr', language: 'en' }), 'en-US');
  assert.equal(filledLanguage('2026', 'he-IL', { direction: 'ltr', language: 'en' }), 'en-US');
  assert.equal(filledLanguage('2026', 'he-IL', { direction: 'rtl', language: 'he' }), 'he-IL');
  assert.equal(filledLanguage('—', 'he-IL', { direction: 'ltr', language: 'en' }), 'he-IL');
  assert.equal(filledLanguage('Revenue 2026', '', { direction: 'ltr', language: 'en' }), '');
  assert.equal(filledLanguage('הכנסות 2026', 'en-US', {}), 'he-IL');
  assert.equal(filledLanguage('AI בשירות', 'en-US', { direction: 'rtl' }), 'he-IL');
  assert.equal(filledLanguage('مرحبا', '', {}), 'ar-SA');
  assert.equal(filledLanguage('+12%', 'he-IL', { direction: 'ltr' }), 'en-US');
  assert.equal(filledLanguage('+12%', 'he-IL', { direction: 'rtl' }), 'he-IL');
  assert.equal(marksRightToLeft('Microsoft Teams ככלי עבודה'), false);
  assert.equal(marksRightToLeft('Microsoft Teams ככלי עבודה', 'rtl'), true);
  assert.equal(paragraphDirection('Microsoft Teams ככלי עבודה'), 'ltr');
  assert.equal(paragraphDirection('Microsoft Teams ככלי עבודה', 'rtl'), 'rtl');
  assert.equal(paragraphDirection('78%', 'rtl'), undefined);
});

test('an English deck on a right-to-left template writes English runs, and its Hebrew deck is as before', () => {
  const template = loadTemplate(buildFixturePptx(), fixtureManifest);
  const assets = { photo: placeholderPhoto(100, 100, [0, 0, 0], [255, 255, 255]) };
  const slideXml = (deck: Parameters<typeof renderDeck>[1]) => strFromU8(unzipSync(renderDeck(template, deck, { assets }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  const english = slideXml({ version: 1, language: 'en', direction: 'ltr', slides: [{ pattern: 'content', slots: { title: 'Reached 74% of 75%', body: 'Body', photo: { asset: 'photo' }, note: 'n' } }] });
  const title = /<p:cNvPr id="2" name="Title"\/>[\s\S]*?<\/p:sp>/.exec(english)?.[0] ?? '';
  assert.match(title, /<a:rPr lang="en-US" sz="2600"\/><a:t>Reached 74% of 75%<\/a:t>/);
  assert.doesNotMatch(title, /he-IL/);

  const hebrew = slideXml({ version: 1, slides: [{ pattern: 'content', slots: { title: 'הגענו ל-74%', body: 'גוף', photo: { asset: 'photo' }, note: 'n' } }] });
  const hebrewTitle = /<p:cNvPr id="2" name="Title"\/>[\s\S]*?<\/p:sp>/.exec(hebrew)?.[0] ?? '';
  assert.match(hebrewTitle, /<a:rPr lang="he-IL" sz="2600"\/><a:t>הגענו ל-74%<\/a:t>/);
});

test('the renderer writes slots and table cells by the deck\u2019s direction and language', () => {
  const assets = { photo: placeholderPhoto(100, 100, [0, 0, 0], [255, 255, 255]) };
  const content = (title: string) => ({ pattern: 'content', slots: { title, body: 'Body', photo: { asset: 'photo' }, note: 'n' } });
  const render = (deck: DeckSpec, manifest: TemplateManifest = fixtureManifest) => {
    const xml = strFromU8(unzipSync(renderDeck(loadTemplate(buildFixturePptx(), manifest), deck, { assets }).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
    return { xml, title: /<p:cNvPr id="2" name="Title"\/>[\s\S]*?<\/p:sp>/.exec(xml)?.[0] ?? '' };
  };
  // A deck in the template's direction (right to left) states none: a Latin-led Hebrew title reads right to left.
  assert.match(render({ version: 1, slides: [content('AI בשירות הלקוחות')] }).title, /<a:pPr algn="r" rtl="1"\/><a:r><a:rPr lang="he-IL" sz="2600"\/><a:t>AI בשירות הלקוחות<\/a:t>/);
  // The deck's own language names English text.
  assert.match(render({ version: 1, language: 'en-GB', direction: 'ltr', slides: [content('Reached')] }).title, /<a:rPr lang="en-GB" sz="2600"\/><a:t>Reached<\/a:t>/);
  // A table cell follows the same rule.
  const table = render({ version: 1, slides: [{ pattern: 'table', slots: { title: 'T', table: { header: ['AI בשירות הלקוחות', 'B'], rows: [['1', '2']] } } }] }, tableManifest).xml;
  const cell = /<a:tc>[\s\S]*?<\/a:tc>/.exec(table)?.[0] ?? '';
  assert.match(cell, /rtl="1"/);
  assert.match(cell, /<a:rPr lang="he-IL"/);
});

test('a label the free-slide code writes, such as a legend, follows the deck\u2019s direction', () => {
  const demo = loadTemplate(readFileSync(join(root, 'templates', 'demo.pptx')), JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest);
  const chart: ChartValue = { type: 'column', categories: ['א', 'ב'], series: [{ name: 'AI בשירות', values: [1, 2] }, { name: 'הכנסות', values: [3, 4] }] };
  const deck: DeckSpec = { version: 1, language: 'he', direction: 'rtl', fields: { organizationDepartment: 'צוות' }, slides: [freeSlide({ title: 'שתים עשרה', blocks: Array.from({ length: 12 }, () => ({ chart })) })] };
  const xml = strFromU8(unzipSync(renderDeck(demo, deck).pptx)['ppt/slides/slide1.xml'] as Uint8Array);
  const legend = /<p:cNvPr id="\d+" name="Legend 1"\/>[\s\S]*?<\/p:sp>/.exec(xml)?.[0] ?? '';
  assert.match(legend, /<a:t>AI בשירות<\/a:t>/);
  assert.match(legend, /rtl="1"/);
  assert.match(legend, /lang="he-IL"/);
});
