// The headings of a person's document in a slide's text (`SlideSpec.headings`): which paragraphs
// they are, how they are set, how they are measured, and how a written slide is settled with its passages.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { loadTemplate, planOutline, renderDeck, slideAlternatives, textMeasurer, writeDeck } from '../src/index.ts';
import type { DeckSpec, FontMetrics, SlideSpec, TemplateManifest } from '../src/index.ts';
import { headingPlaces, leadsInto } from '../src/deck/headings.ts';
import { fitOptions } from '../src/deck/validate.ts';
import { freeKit } from '../src/free/kit.ts';
import { freeTitleLines } from '../src/free/plan.ts';
import { readDocument, settleDocumentSlide } from '../src/plan/document.ts';
import { NS, childElement, childElements, descendants, parseXml, rootElement } from '../src/pptx/xml.ts';
import { setHeadings, setShapeText } from '../src/render/text.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pptx = readFileSync(join(root, 'templates', 'demo.pptx'));
const manifest = JSON.parse(readFileSync(join(root, 'templates', 'demo.manifest.json'), 'utf8')) as TemplateManifest;
const demo = loadTemplate(pptx, manifest);
const stored = JSON.parse(readFileSync(join(root, 'test', 'fixtures', 'documents-0.6.5.json'), 'utf8')) as Record<string, DeckSpec>;
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const EM = ' ';

test('headingPlaces names paragraphs in order, never the last, and nothing for what is no list of strings', () => {
  const text = ['Opening text.', 'Riders', 'Text on riders.', 'Costs', 'Text on costs.', 'Plans'];
  assert.deepEqual(headingPlaces(text, ['Riders', 'Costs']), [1, 3]);
  // The last paragraph stands over nothing.
  assert.deepEqual(headingPlaces(text, ['Riders', 'Plans']), [1]);
  // An entry that matches nothing is passed over; the next is still found.
  assert.deepEqual(headingPlaces(text, ['Gone', 'Costs']), [3]);
  // Entries match in order, each after the one before.
  assert.deepEqual(headingPlaces(text, ['Costs', 'Riders']), [3]);
  assert.deepEqual(headingPlaces(['A', 'x', 'A', 'y'], ['A', 'A']), [0, 2]);
  assert.deepEqual(headingPlaces(text, [' Riders ']), [1]);
  // One paragraph has none, and neither has anything but a list of strings.
  assert.deepEqual(headingPlaces(['Riders'], ['Riders']), []);
  for (const bad of [undefined, null, 'Riders', 7, [1, 'Riders'], { 0: 'Riders' }]) assert.deepEqual(headingPlaces(text, bad), []);
  // A heading over an item of a list or a row of a table keeps its own line.
  assert.equal(leadsInto('Text on riders.'), true);
  assert.equal(leadsInto('• Docks are full'), false);
  assert.equal(leadsInto('1. Docks are full'), false);
  assert.equal(leadsInto('Item | 2023 | 2024'), false);
});

const points = (space: ReturnType<typeof childElement>): string | null => (space ? (childElement(space, NS.a, 'spcPts')?.getAttribute('val') ?? null) : null);

/** A text shape of the given width (EMU), filled with paragraphs and their headings set. */
function shapeWith(paragraphs: string[], places: number[], options: { width?: number; columns?: number; gap?: number; bullet?: boolean; lang?: string; rtl?: boolean; margin?: number } = {}) {
  const width = options.width ?? 9144000;
  const bullet = options.bullet ? '<a:buChar char="•"/>' : '';
  const lang = options.lang ?? 'en-US';
  const columns = options.columns ? ` numCol="${options.columns}" spcCol="228600"` : '';
  const doc = parseXml(
    `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="3000000"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr${columns}/><a:lstStyle/><a:p><a:pPr marL="${options.margin ?? 0}"${options.rtl ? ' rtl="1"' : ''}><a:spcBef><a:spcPts val="0"/></a:spcBef>${options.gap !== undefined ? `<a:spcAft><a:spcPts val="${options.gap * 100}"/></a:spcAft>` : ''}${bullet}</a:pPr>` +
      `<a:r><a:rPr lang="${lang}" sz="1600"/><a:t>sample</a:t></a:r></a:p></p:txBody></p:sp>`,
  );
  const shape = rootElement(doc);
  setShapeText(shape, paragraphs, { flipAlign: false, ...(options.rtl ? { direction: 'rtl' as const, language: 'he' } : {}) });
  setHeadings(shape, places, paragraphs, options.rtl ? 'rtl' : 'ltr');
  const body = childElement(shape, NS.p, 'txBody') as NonNullable<ReturnType<typeof childElement>>;
  return childElements(body, NS.a, 'p').map((paragraph) => ({
    runs: childElements(paragraph, NS.a, 'r').map((run) => ({ text: childElement(run, NS.a, 't')?.textContent ?? '', bold: childElement(run, NS.a, 'rPr')?.getAttribute('b') === '1', lang: childElement(run, NS.a, 'rPr')?.getAttribute('lang') })),
    before: points(descendants(paragraph, NS.a, 'spcBef')[0]),
    after: points(descendants(paragraph, NS.a, 'spcAft')[0]),
    rtl: childElement(paragraph, NS.a, 'pPr')?.getAttribute('rtl') ?? null,
    bullet: childElement(childElement(paragraph, NS.a, 'pPr') ?? paragraph, NS.a, 'buNone') ? 'none' : childElement(childElement(paragraph, NS.a, 'pPr') ?? paragraph, NS.a, 'buChar') ? 'char' : null,
  }));
}

test('a heading over text is set run in: bold, an em space opening the text, its space after moved before the paragraph', () => {
  const set = shapeWith(['Opening text.', 'Riders', 'Students and commuters make up most members.'], [1], { gap: 6 });
  assert.equal(set.length, 2);
  assert.deepEqual(set[1]?.runs, [
    { text: 'Riders', bold: true, lang: 'en-US' },
    { text: `${EM}Students and commuters make up most members.`, bold: false, lang: 'en-US' },
  ]);
  assert.equal(set[1]?.before, '600');
  // No paragraph is a heading alone.
  assert.ok(!set.some((paragraph) => paragraph.runs.length === 1 && paragraph.runs[0]?.text === 'Riders'));
});

test('a heading run in keeps its language; the em space is the paragraph’s, so a heading in another language is not glued to its text', () => {
  const he = shapeWith(['פתיחה.', 'תוכנית', 'התקציב לשנה הבאה גדל.'], [1], { gap: 6, lang: 'he-IL', rtl: true });
  assert.deepEqual(he[1]?.runs.map((run) => [run.text, run.bold, run.lang]), [
    ['תוכנית', true, 'he-IL'],
    [`${EM}התקציב לשנה הבאה גדל.`, false, 'he-IL'],
  ]);
  assert.equal(he[1]?.rtl, '1');
  // An English heading over Hebrew text: the space opens the Hebrew run, in Hebrew. In the
  // heading's English run, PowerPoint drew "KPIהלקוחות", the heading glued to its text.
  const mixed = shapeWith(['פתיחה.', 'KPI', 'הלקוחות מקבלים מענה מהיר יותר.'], [1], { gap: 6, lang: 'he-IL', rtl: true });
  assert.deepEqual(mixed[1]?.runs.map((run) => [run.text, run.lang]), [
    ['KPI', 'en-US'],
    [`${EM}הלקוחות מקבלים מענה מהיר יותר.`, 'he-IL'],
  ]);
  assert.equal(mixed[1]?.rtl, '1');
  // A Hebrew heading over an English paragraph of a Hebrew deck would make it a Hebrew paragraph: it keeps its own line.
  const turned = shapeWith(['פתיחה.', 'סיכונים', 'The main risk is a shortage of staff.'], [1], { gap: 6, lang: 'he-IL', rtl: true });
  assert.equal(turned.length, 3);
  assert.deepEqual(turned[1]?.runs.map((run) => [run.text, run.bold]), [['סיכונים', true]]);
  assert.equal(turned[2]?.rtl, '0');
});

test('a heading keeps a line of its own over a list item, a table row, another heading, or where it is too long to run in', () => {
  const overItem = shapeWith(['Opening text.', 'Tasks', '• Repair the docks'], [1], { gap: 6, bullet: true });
  assert.equal(overItem.length, 3);
  assert.deepEqual(overItem[1], { runs: [{ text: 'Tasks', bold: true, lang: 'en-US' }], before: '600', after: '0', rtl: '0', bullet: 'none' });
  assert.equal(overItem[2]?.bullet, 'char');
  // An item whose bullet the paragraph draws, with no marker in its text, as a layout's list sets it.
  const drawn = shapeWith(['Opening text.', 'Tasks', 'Repair the docks'], [1], { gap: 6, bullet: true });
  assert.equal(drawn.length, 3);
  assert.deepEqual([drawn[1]?.runs.map((run) => [run.text, run.bold]), drawn[1]?.bullet, drawn[2]?.bullet], [[['Tasks', true]], 'none', 'char']);
  const overRow = shapeWith(['Opening text.', 'Figures', 'Bikes | 120 | 140'], [1], { gap: 6 });
  assert.equal(overRow[1]?.runs[0]?.text, 'Figures');
  // Two headings in a row: the first keeps its line, the second runs in.
  const stacked = shapeWith(['Opening text.', 'Part one', 'Riders', 'Text on riders.'], [1, 2], { gap: 6 });
  assert.deepEqual(stacked.map((paragraph) => paragraph.runs.map((run) => run.text)), [['Opening text.'], ['Part one'], ['Riders', `${EM}Text on riders.`]]);
  // Run in only where its bold letters and the space take no more than a line, as the column is:
  // 9 letters run in across a box 216 pt wide at 16 pt (18 counted to a line), and not in one of
  // its two columns (8), nor under a paragraph whose margin takes most of the line.
  const runsIn = (options: Parameters<typeof shapeWith>[2]) => shapeWith(['Opening text.', 'Ridership', 'Text under it.'], [1], { gap: 6, width: 2743200, ...options }).length === 2;
  assert.equal(runsIn({}), true);
  assert.equal(runsIn({ columns: 2 }), false);
  assert.equal(runsIn({ margin: 1371600 }), false);
  // A heading of a full line of a wide box keeps its line too.
  const long = 'A heading long enough to need most of a line of its own here, and more';
  assert.deepEqual(shapeWith(['Opening text.', long, 'Text under it.'], [1], { gap: 6 })[1]?.runs.map((run) => [run.text, run.bold]), [[long, true]]);
});

test('settleDocumentSlide marks the headings of its passages, sets a bracketed one plain, and drops a first line that says its title again', () => {
  const doc = readDocument(['Background', 'Opening text of the report, a sentence long.', '', 'Riders', 'Students and commuters make up most members.', '', 'Costs', 'A yearly pass costs sixty euros.'].join('\n'), 'en');
  assert.deepEqual(
    doc.passages.map((passage) => passage.section),
    ['Background', 'Riders', 'Costs'],
  );
  const slide = (slots: SlideSpec['slots'], more: Partial<SlideSpec> = {}): SlideSpec => ({ pattern: 'free', blocks: ['text'], slots, ...more });
  // Text set by code: the headings inside it are marked.
  const filled = settleDocumentSlide(slide({ title: 'Background', b1Text: ['Opening text of the report, a sentence long.', 'Riders', 'Students and commuters make up most members.', 'Costs', 'A yearly pass costs sixty euros.'] }), doc.passages, [1, 2, 3]);
  assert.deepEqual(filled.headings, ['Riders', 'Costs']);
  // Text a writer wrote, with a heading in brackets and its first line saying the title again.
  const written = settleDocumentSlide(slide({ title: 'Background', b1Text: ['Background:', 'Opening text of the report, a sentence long.', '[Riders]', 'Students and commuters make up most members.'] }), doc.passages, [1, 2]);
  assert.deepEqual(written.slots.b1Text, ['Opening text of the report, a sentence long.', 'Riders', 'Students and commuters make up most members.']);
  assert.deepEqual(written.headings, ['Riders']);
  // A Hebrew title and its line with a colon.
  const he = readDocument(['רקע', 'הצוות עובד יחד כל שבוע כדי לשפר את השירות.', '', 'תוכנית', 'התקציב לשנה הבאה גדל בעשרה אחוזים.'].join('\n'), 'he');
  const hebrew = settleDocumentSlide(slide({ title: 'רקע', b1Text: ['רקע:', 'הצוות עובד יחד כל שבוע כדי לשפר את השירות.', 'תוכנית', 'התקציב לשנה הבאה גדל בעשרה אחוזים.'] }), he.passages, [1, 2]);
  assert.deepEqual(hebrew.slots.b1Text, ['הצוות עובד יחד כל שבוע כדי לשפר את השירות.', 'תוכנית', 'התקציב לשנה הבאה גדל בעשרה אחוזים.']);
  assert.deepEqual(hebrew.headings, ['תוכנית']);
  // A heading no longer in the text is not marked, and a slide with none has no field.
  const reworded = settleDocumentSlide(slide({ title: 'Background', b1Text: ['A shorter text.', 'Another line.'] }, { headings: ['Riders'] }), doc.passages, [1, 2]);
  assert.equal('headings' in reworded, false);
  // A line that says the title again, but is the whole text, stays.
  const alone = settleDocumentSlide(slide({ title: 'Riders', b1Text: ['Riders'] }), doc.passages, [2]);
  assert.deepEqual(alone.slots.b1Text, ['Riders']);
  // Other slots are left as they were.
  assert.deepEqual(settleDocumentSlide(slide({ title: 'Costs', b1Text: ['x', 'y'], b2Chart: { type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }] } }), doc.passages, [3]).slots.b2Chart, { type: 'bar', categories: ['a'], series: [{ name: 's', values: [1] }] });
});

test('settleDocumentSlide: a heading over a list whose items lost their markers stays plain; a translated deck’s bracketed headings are marked', () => {
  const doc = readDocument(['Background', 'Opening text of the report.', '', 'Tasks', '- Repair the docks', '- Hire two mechanics', '', 'Costs', 'A yearly pass costs sixty euros.'].join('\n'), 'en');
  assert.deepEqual(
    doc.passages.map((passage) => [passage.section, passage.kind]),
    [
      ['Background', 'text'],
      ['Tasks', 'list'],
      ['Costs', 'text'],
    ],
  );
  const slide = (b1Text: string[]): SlideSpec => ({ pattern: 'text-only', slots: { title: 'Background', b1Text } });
  // A layout that draws bullets of its own was given the items without theirs: run in, the heading would lead the first item alone.
  const stripped = settleDocumentSlide(slide(['Opening text of the report.', 'Tasks', 'Repair the docks', 'Hire two mechanics', 'Costs', 'A yearly pass costs sixty euros.']), doc.passages, [1, 2, 3]);
  assert.deepEqual(stripped.headings, ['Costs']);
  // With their markers the items say what they are, and the heading over them keeps its own line when it is set.
  const marked = settleDocumentSlide(slide(['Opening text of the report.', 'Tasks', '- Repair the docks', '- Hire two mechanics', 'Costs', 'A yearly pass costs sixty euros.']), doc.passages, [1, 2, 3]);
  assert.deepEqual(marked.headings, ['Tasks', 'Costs']);
  // A deck in another language than its document: the writer translated the headings, kept in brackets.
  const translated = settleDocumentSlide(slide(['Texte d’ouverture du rapport.', '[Coûts]', 'Un abonnement annuel coûte soixante euros.', '[Fin]']), doc.passages, [1, 3], { translated: true });
  assert.deepEqual(translated.slots.b1Text, ['Texte d’ouverture du rapport.', 'Coûts', 'Un abonnement annuel coûte soixante euros.', '[Fin]']);
  assert.deepEqual(translated.headings, ['Coûts']);
  // In the document's own language a bracketed line that is none of its headings is the person's text, and stays.
  assert.deepEqual(settleDocumentSlide(slide(['Text.', '[Note]', 'More text.']), doc.passages, [1]).slots.b1Text, ['Text.', '[Note]', 'More text.']);
});

test('a stored deck renders as before; with its headings marked, no heading stands alone as a paragraph', () => {
  for (const [name, deck] of Object.entries(stored)) {
    const plain = renderDeck(demo, deck, { layout: 'adaptive' }).pptx;
    // Marks that match nothing, or are no list of strings, change nothing.
    for (const headings of [[], ['Not in the text'], 'Riders', [3]] as unknown[]) {
      const marked = { ...deck, slides: deck.slides.map((slide) => ({ ...slide, headings })) } as DeckSpec;
      assert.deepEqual(unzipSync(renderDeck(demo, marked, { layout: 'adaptive' }).pptx), unzipSync(plain), `${name}: ${JSON.stringify(headings)}`);
    }
  }
  const deck = structuredClone(stored['document en dense x2'] as DeckSpec);
  (deck.slides[0] as SlideSpec).headings = ['Riders'];
  (deck.slides[1] as SlideSpec).headings = ['Plans'];
  const parts = unzipSync(renderDeck(demo, deck, { layout: 'adaptive' }).pptx);
  for (const [number, heading] of [
    [1, 'Riders'],
    [2, 'Plans'],
  ] as const) {
    const xml = strFromU8(parts[`ppt/slides/slide${number}.xml`] as Uint8Array);
    assert.match(xml, new RegExp(`<a:rPr [^>]*b="1"[^>]*>(?:(?!</a:r>).)*<a:t>${heading}</a:t></a:r><a:r>(?:(?!</a:r>).)*<a:t>${EM}`));
    assert.doesNotMatch(xml, new RegExp(`<a:t>${heading}</a:t></a:r><a:endParaRPr`));
  }
});

test('headings are measured in bold, as they are drawn', () => {
  // A font whose bold letters are a fifth wider than its regular ones.
  const metrics: FontMetrics = { face: (_name, style) => ({ has: () => true, advances: (text) => [...text].map(() => (style.bold ? 0.6 : 0.5)) }) };
  const measure = textMeasurer(loadTemplate(pptx, manifest), metrics);
  const text = ['Opening text.', 'Riders and their trips', 'Students and commuters make up most members.'];
  const widths = (bold?: number[]) => measure.fit('text-and-photo', 'body', text, bold ? { bold } : {})?.widths[1]?.[0];
  assert.equal(widths(), 22 * 0.5 * 18);
  // Each glyph on PowerPoint's grid of an eighth of a point: 10.8 pt is set at 10.75.
  assert.equal(widths([1]), 22 * 10.75);
  // The checks pass the headings a slide marks, by their places in the text.
  assert.deepEqual(fitOptions(manifest, {}, { headings: ['Riders and their trips'] }, undefined, text).bold, [1]);
  assert.equal(fitOptions(manifest, {}, { headings: ['Riders and their trips'] }).bold, undefined);
  assert.equal(fitOptions(manifest, {}, {}, undefined, text).bold, undefined);
});

test('a free slide’s title is counted in lines as its plan counts them', () => {
  const kit = freeKit(demo);
  assert.ok(kit?.title);
  assert.equal(freeTitleLines(kit, 10), 1);
  assert.ok(freeTitleLines(kit, 400) > 1);
});

test('a layout the slide is offered instead keeps its headings', () => {
  const body = manifest.patterns.find((pattern) => pattern.id === 'text-and-photo')?.slots.filter((slot) => slot.kind === 'text' && !slot.deckField) ?? [];
  const slots = Object.fromEntries(body.map((slot) => [slot.name, slot.kind === 'text' && (slot.maxParagraphs ?? 1) > 1 ? ['One.', 'Riders', 'Two.'] : 'Short']));
  const deck: DeckSpec = { version: 1, fields: { organizationDepartment: 'Retail' }, slides: [{ pattern: 'text-and-photo', slots, headings: ['Riders'] }] };
  const alternatives = slideAlternatives(manifest, deck, 1, { layout: 'adaptive' });
  assert.ok(alternatives.length > 0);
  for (const alternative of alternatives) assert.deepEqual(alternative.slide.headings, ['Riders'], alternative.pattern);
});

test('a slide is titled with the heading its text opens under where the title holds it; a longer one is the writer’s to title', async () => {
  const sentence = 'The city opened its bike share in 2019 with forty stations and four hundred bikes.';
  const asked: unknown[] = [];
  const llm = (answers: unknown[]) => ({
    async generateJson(request: { schema: unknown }) {
      asked.push(request.schema);
      const next = answers.shift();
      return typeof next === 'function' ? (next as (r: unknown) => unknown)(request) : next;
    },
  });
  const fields = { organizationDepartment: 'Retail' };
  const make = async (heading: string) => {
    asked.length = 0;
    const text = [heading, sentence, '', 'Riders', sentence].join('\n');
    const outline = { title: 'Review', left_out: [], asked_slides: 0, slides: [{ pattern: 'free', purpose: 'All', passages: [1, 2], shows: [{ kind: 'text', count: 1 }], shows_main: 'none', side_by_side: 0 }] };
    const plan = await planOutline(demo, { prompt: 'A deck', language: 'en', fields, imageDensity: 'none', document: text, slideCount: 1 }, { llm: llm([outline]), layout: 'adaptive' });
    const writer = (request: { schema: { properties: Record<string, { properties: Record<string, unknown> }> } }) => Object.fromEntries(Object.entries(request.schema.properties).map(([key, slide]) => [key, 'title' in slide.properties ? { title: 'Written title' } : {}]));
    const { deck } = await writeDeck(demo, plan, { llm: llm([writer]), layout: 'adaptive' });
    return { slide: deck.slides[0] as SlideSpec, writerAsked: asked.length > 1 };
  };
  const short = await make('Background');
  assert.deepEqual([short.slide.slots.title, short.slide.headings, short.writerAsked], ['Background', ['Riders'], false]);
  // Over the title's limit on this template (41 characters): the writer titles the slide.
  const long = await make('Background and the first five years of the network');
  assert.deepEqual([long.slide.slots.title, long.slide.headings, long.writerAsked], ['Written title', ['Riders'], true]);
});
