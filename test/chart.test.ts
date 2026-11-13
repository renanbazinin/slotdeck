import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { DeckValidationError, PptxPackage, loadTemplate, planDeck, renderDeck, validateDeck } from '../src/index.ts';
import type { ChartValue, DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { statedNumbers } from '../src/deck/numbers.ts';
import { readRels, resolveTarget } from '../src/pptx/package.ts';
import { NS, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import { buildChartWorkbook, buildChartXml } from '../src/render/chart.ts';
import { buildFixturePptx, fixtureManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** The fixture's first slide with its photo (shape 4) used as a chart slot. */
const chartManifest = (extra: Record<string, unknown> = {}, shapeId = '4'): TemplateManifest => ({
  ...fixtureManifest,
  patterns: [
    ...fixtureManifest.patterns,
    {
      id: 'chart',
      sourceSlide: 1,
      description: 'A chart.',
      figures: true,
      slots: [
        { name: 'title', kind: 'text', shapeId: '2' },
        { name: 'chart', kind: 'chart', shapeId, maxCategories: 4, maxSeries: 2, colors: ['00A252', '6924E8'], text: { size: 9, font: 'Assistant', color: '0C0C0C' }, ...extra },
      ],
    },
  ],
});

const column: ChartValue = { type: 'column', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Adoptions', values: [62, 74, 95] }] };
const slide = (chart: unknown): SlideSpec => ({ pattern: 'chart', slots: { title: 'T', chart: chart as ChartValue } });
const deckOf = (...slides: SlideSpec[]): DeckSpec => ({ version: 1, slides });

function scripted(...answers: unknown[]): { llm: LlmProvider; requests: JsonRequest[] } {
  const requests: JsonRequest[] = [];
  return {
    requests,
    llm: {
      async generateJson(request) {
        requests.push(request);
        return answers.shift();
      },
    },
  };
}

test('a chart slot becomes a native chart with its workbook', () => {
  const result = renderDeck(loadTemplate(pptx, chartManifest()), deckOf(slide(column)));
  const pkg = PptxPackage.load(result.pptx);
  const root = rootElement(parseXml(pkg.readText('ppt/slides/slide1.xml')));

  // The sample picture is gone; a chart frame with the same id sits in its place and area.
  assert.ok(!descendants(root, NS.p, 'pic').some((pic) => path(pic, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr'])?.getAttribute('id') === '4'));
  const frame = descendants(root, NS.p, 'graphicFrame').find((el) => path(el, [NS.p, 'nvGraphicFramePr'], [NS.p, 'cNvPr'])?.getAttribute('id') === '4');
  assert.ok(frame, 'chart frame present');
  assert.equal(path(frame, [NS.p, 'xfrm'], [NS.a, 'ext'])?.getAttribute('cx'), '4000000');
  assert.equal(path(frame, [NS.p, 'xfrm'], [NS.a, 'off'])?.getAttribute('x'), '1000000');

  // Slide -> chart -> workbook, each reachable through relationships.
  const relId = descendants(frame, C, 'chart')[0]?.getAttributeNS(NS.r, 'id');
  const chartRel = readRels(pkg, 'ppt/slides/slide1.xml').find((rel) => rel.id === relId);
  assert.ok(chartRel?.type.endsWith('/chart'));
  const chartPart = resolveTarget('ppt/slides/slide1.xml', chartRel?.target ?? '');
  assert.equal(chartPart, 'ppt/charts/chart1.xml');
  const workbookRel = readRels(pkg, chartPart)[0];
  assert.equal(resolveTarget(chartPart, workbookRel?.target ?? ''), 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx');
  assert.ok(pkg.has('ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'));
  assert.equal(descendants(rootElement(parseXml(pkg.readText(chartPart))), C, 'externalData')[0]?.getAttributeNS(NS.r, 'id'), workbookRel?.id);

  const types = pkg.readText('[Content_Types].xml');
  assert.ok(types.includes('PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"'));
  assert.ok(types.includes('Extension="xlsx"'));
  assert.ok(!pkg.has('ppt/media/image1.png') || readRels(pkg, 'ppt/slides/slide1.xml').some((rel) => rel.type.endsWith('/image')), 'unused sample image is pruned');
});

test('each chart in a deck gets its own parts', () => {
  const line: ChartValue = { type: 'line', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] };
  const pkg = PptxPackage.load(renderDeck(loadTemplate(pptx, chartManifest()), deckOf(slide(column), slide(line))).pptx);
  for (const part of ['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml', 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx', 'ppt/embeddings/Microsoft_Excel_Sheet2.xlsx']) {
    assert.ok(pkg.has(part), part);
  }
  assert.ok(pkg.readText('ppt/charts/chart2.xml').includes('<c:lineChart>'));
});

test('chart XML: types, data references, colours and formats', () => {
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const build = (chart: ChartValue, direction: 'rtl' | 'ltr' = 'ltr') => {
    const xml = buildChartXml(chart, slot, 'rId1', direction);
    rootElement(parseXml(xml)); // well-formed
    return xml;
  };

  const col = build(column);
  assert.ok(col.includes('<c:barDir val="col"/>'));
  assert.ok(col.includes('<c:f>Sheet1!$A$2:$A$4</c:f>') && col.includes('<c:f>Sheet1!$B$2:$B$4</c:f>') && col.includes('<c:f>Sheet1!$B$1</c:f>'));
  assert.ok(col.includes('<c:pt idx="2"><c:v>95</c:v></c:pt>'), 'values are cached for viewers that do not open the workbook');
  assert.ok(col.includes('<a:srgbClr val="00A252"/>'));
  assert.ok(col.includes('<c:min val="0"/>'), 'bars start at zero');
  assert.ok(col.includes('typeface="Assistant"') && col.includes('sz="900"'));
  assert.ok(!col.includes('<c:legend>'), 'one series needs no legend');

  const bar = build({ ...column, type: 'bar' });
  assert.ok(bar.includes('<c:barDir val="bar"/>') && bar.includes('<c:orientation val="maxMin"/>') && bar.includes('<c:crosses val="max"/>'));

  const two = build({ type: 'line', categories: ['a', 'b'], series: [{ name: 'x', values: [1, 2] }, { name: 'y', values: [3, 4] }] });
  assert.ok(two.includes('<c:lineChart>') && two.includes('<c:f>Sheet1!$C$2:$C$3</c:f>') && two.includes('<a:srgbClr val="6924E8"/>'));
  assert.ok(two.includes('<c:legend>') && !two.includes('<c:min val="0"/>'));

  const pie = build({ type: 'pie', categories: ['a', 'b', 'c'], series: [{ name: 's', values: [50, 30, 20] }], format: 'percent' });
  assert.equal(pie.split('<c:dPt>').length - 1, 3, 'one coloured slice per category');
  assert.ok(pie.includes('<c:formatCode>0"%"</c:formatCode>') && pie.includes('<c:legend>') && !pie.includes('<c:catAx>'));

  assert.ok(!build({ ...column, series: [{ name: 's', values: [-5, 10, 3] }] }).includes('<c:min val="0"/>'), 'negative values keep the automatic axis');
  assert.ok(build(column, 'rtl').includes('<a:pPr rtl="1">') && !col.includes('rtl="1"'));
  assert.ok(build({ ...column, categories: ['A & B', '<C>', 'D'] }).includes('A &amp; B'), 'text is escaped');
});

test('the workbook holds the chart data where the chart expects it', () => {
  const files = unzipSync(buildChartWorkbook({ type: 'column', categories: ['Q1', 'Q2'], series: [{ name: 'Adoptions', values: [62, 74] }, { name: 'Returns', values: [4, 3] }] }));
  assert.deepEqual(Object.keys(files).sort(), ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  const sheet = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array);
  rootElement(parseXml(sheet));
  assert.ok(sheet.includes('<c r="B1" t="inlineStr"><is><t xml:space="preserve">Adoptions</t></is></c>'));
  assert.ok(sheet.includes('<c r="C1" t="inlineStr"><is><t xml:space="preserve">Returns</t></is></c>'));
  assert.ok(sheet.includes('<c r="A2" t="inlineStr"><is><t xml:space="preserve">Q1</t></is></c><c r="B2"><v>62</v></c><c r="C2"><v>4</v></c>'));
  assert.ok(strFromU8(files['xl/workbook.xml'] as Uint8Array).includes('name="Sheet1"'));
});

test('a chart can be given its own area, and cannot sit inside a group', () => {
  const boxed = renderDeck(loadTemplate(pptx, chartManifest({ box: { x: 0.1, y: 0.2, w: 0.5, h: 0.5 } })), deckOf(slide(column)));
  const root = rootElement(parseXml(PptxPackage.load(boxed.pptx).readText('ppt/slides/slide1.xml')));
  const frame = descendants(root, NS.p, 'graphicFrame')[0];
  assert.equal(path(frame as never, [NS.p, 'xfrm'], [NS.a, 'off'])?.getAttribute('x'), '914400');
  assert.equal(path(frame as never, [NS.p, 'xfrm'], [NS.a, 'ext'])?.getAttribute('cy'), '2571750');

  // Shape 8 is the picture inside the fixture's group.
  assert.throws(() => renderDeck(loadTemplate(pptx, chartManifest({}, '8')), deckOf(slide(column))), /not inside a group/);
});

test('chart values are validated', () => {
  const manifest = chartManifest();
  const check = (chart: unknown) => validateDeck(manifest, deckOf(slide(chart)));
  const first = (chart: unknown) => check(chart)[0]?.message ?? '';

  assert.deepEqual(check(column), []);
  assert.match(first({ ...column, series: [{ name: 'Adoptions', values: [1, 2] }] }), /series "Adoptions" has 2 values; the chart has 3 categories/);
  assert.match(first({ ...column, series: [{ name: 'A', values: [1, 'x', 3] }] }), /not a number/);
  assert.match(first({ ...column, type: 'radar' }), /unknown chart type "radar"/);
  assert.match(first({ ...column, series: [] }), /no series/);
  assert.match(first({ type: 'pie', categories: ['a', 'b', 'c'], series: [column.series[0], column.series[0]] }), /exactly one series/);
  assert.match(first('a chart please'), /expected a chart/);
  assert.match(validateDeck(manifest, deckOf({ pattern: 'chart', slots: { title: column, chart: column } }))[0]?.message ?? '', /expected text/);

  const big: ChartValue = {
    type: 'column',
    categories: ['a', 'b', 'c', 'd', 'e'],
    series: [1, 2, 3].map((n) => ({ name: `s${n}`, values: [1, 2, 3, 4, 5] })),
  };
  assert.deepEqual(
    check(big).map((issue) => [issue.severity, issue.message]),
    [
      ['warning', 'the chart has 5 categories; the slot fits 4'],
      ['warning', 'the chart has 3 series; the slot fits 2'],
    ],
  );
  assert.throws(() => renderDeck(loadTemplate(pptx, manifest), deckOf(slide({ ...column, series: [] }))), DeckValidationError);
});

test('the planner asks for chart data and flags numbers the brief does not state', async () => {
  const template = loadTemplate(pptx, chartManifest());
  const outline = { title: 'T', slides: [{ pattern: 'chart', purpose: 'Adoptions per quarter' }] };
  const answer = { s1: { title: 'Adoptions', chart: { type: 'line', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Adoptions', values: [62, 74, 89] }], format: 'number' } } };

  const { llm, requests } = scripted(outline, answer);
  const brief = { prompt: 'Adoptions were 62 in Q1 and 74 in Q2.', language: 'en', imageDensity: 'none' as const };
  const { deck, warnings } = await planDeck(template, brief, { llm });

  const schema = (requests[1]?.schema as { properties: { s1: { properties: { chart: { properties: Record<string, { enum?: string[]; maxItems?: number }>; required: string[] } } } } }).properties.s1
    .properties.chart;
  assert.deepEqual(schema.properties.type?.enum, ['column', 'bar', 'line', 'pie']);
  assert.equal(schema.properties.categories?.maxItems, 4);
  assert.equal(schema.properties.series?.maxItems, 2);
  assert.deepEqual(schema.required, ['type', 'categories', 'series', 'format']);

  assert.deepEqual(deck.slides[0]?.slots.chart, { type: 'line', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Adoptions', values: [62, 74, 89] }] });
  assert.equal(requests.length, 2, 'an unstated number is reported, not sent back for repair');
  assert.deepEqual(
    warnings.map((issue) => [issue.slot, issue.message]),
    [['chart', 'chart values not stated in the brief: 89. Check them before using the deck.']],
  );

  // A brief without figures is never offered the chart pattern at all.
  const none = scripted({ title: 'T', slides: [{ pattern: 'closing', purpose: 'p' }] }, { s1: {} });
  await planDeck(template, { prompt: 'A deck about dogs', imageDensity: 'none', fields: { closing: 'Bye' } }, { llm: none.llm });
  assert.match(none.requests[0]?.prompt ?? '', /^(?![\s\S]*- chart:)/);
});

test('numbers stated in a text', () => {
  assert.deepEqual([...statedNumbers('320 adoptions, 3K hours, 1,200 visits, 45% and 7.5 days')].sort((a, b) => a - b), [3, 7.5, 45, 320, 1200, 3000]);
  assert.deepEqual([...statedNumbers('no figures here')], []);
});
