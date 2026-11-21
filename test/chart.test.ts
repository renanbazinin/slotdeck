import assert from 'node:assert/strict';
import { test } from 'node:test';
import { strFromU8, unzipSync } from 'fflate';
import { DeckValidationError, PptxPackage, changeManifest, checkManifest, editManifest, loadTemplate, planDeck, renderDeck, validateDeck } from '../src/index.ts';
import type { ChartValue, DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { keptGaps, statedNumbers } from '../src/deck/numbers.ts';
import { readRels, resolveTarget } from '../src/pptx/package.ts';
import { NS, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import { chartAlt } from '../src/render/alt.ts';
import { axisLabels, buildChartWorkbook, buildChartXml, formatCode, numberLook, plotMargins, roundAxis, roundSpan, seriesColor, shownAs } from '../src/render/chart.ts';
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

// ---- A chart whose place is known (the charts of a slide the software lays out itself) --------

test('a chart slot that states no place is drawn as it always was', () => {
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const xml = buildChartXml({ ...column, series: [...column.series, { name: 'Returns', values: [4, 3, 6] }] }, slot, 'rId1');
  // PowerPoint lays the plot out and picks the end of the axis; numbers are written as they are given; colours repeat.
  assert.ok(xml.includes('<c:layout/>') && xml.includes('<c:formatCode>General</c:formatCode>') && xml.includes('<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'));
  assert.ok(!xml.includes('<c:max ') && !xml.includes('<c:majorUnit ') && !xml.includes('<c:tickLblSkip ') && !xml.includes('rot="0"'));
});

test('the numbers of a chart are written with thousands apart, in the unit and with the decimals its values have', () => {
  const written = (values: number[], percent = false) => {
    const look = numberLook(values, percent);
    return [formatCode(look), ...values.map((value) => shownAs(value, look))];
  };
  assert.deepEqual(written([16440, 6028]), ['#,##0', '16,440', '6,028']);
  // A million and more is written in millions, with the decimals the values have there, up to two; others are
  // rounded to two, as a reader says them, or to three where two would write two values alike. Thousands only
  // where that loses nothing of a value.
  assert.deepEqual(written([1250000, 1410000, 1385000]), ['#,##0.00,,"M"', '1.25M', '1.41M', '1.39M']);
  assert.deepEqual(written([1381000, 1384000]), ['#,##0.000,,"M"', '1.381M', '1.384M']);
  assert.deepEqual(written([2400000, 1250000]), ['#,##0.00,,"M"', '2.40M', '1.25M']);
  assert.deepEqual(written([2000000, 1000000]), ['#,##0,,"M"', '2M', '1M']);
  assert.deepEqual(written([12500, 18000]), ['#,##0', '12,500', '18,000']);
  assert.deepEqual(written([1234567]), ['#,##0.00,,"M"', '1.23M']);
  assert.deepEqual(written([12000, 250000]), ['#,##0,"K"', '12K', '250K']);
  // A value keeps the decimals it has, up to two: 46.5 is not shown as 47.
  assert.deepEqual(written([46.5, 22.3, 18], true), ['0.0"%"', '46.5%', '22.3%', '18.0%']);
  assert.deepEqual(written([45, 30], true), ['0"%"', '45%', '30%']);
  assert.deepEqual(written([0.5, 0.25, -3]), ['#,##0.00', '0.50', '0.25', '-3.00']);
  // An axis in a unit writes nothing as a bare 0.
  assert.equal(formatCode(numberLook([12000, 18000], false), 0, true), '#,##0,"K";-#,##0,"K";0');
});

test('a value axis ends on a round number that is near the largest value, in few steps of whole numbers', () => {
  const axis = (most: number, steps: number, grain = 1) => roundAxis(most, steps, grain);
  assert.deepEqual(axis(18, 8), { max: 20, step: 5 });
  assert.deepEqual(axis(44, 3), { max: 50, step: 25 });
  assert.deepEqual(axis(167, 8), { max: 175, step: 25 });
  assert.deepEqual(axis(167, 6), { max: 200, step: 50 });
  assert.deepEqual(axis(291, 5), { max: 300, step: 100 });
  assert.deepEqual(axis(100, 4), { max: 100, step: 25 });
  // Steps are whole numbers of what the values are written to: thousands, tenths.
  assert.deepEqual(axis(1790000, 6, 1000), { max: 2000000, step: 500000 });
  assert.deepEqual(axis(46.5, 6, 0.1), { max: 50, step: 10 });
  assert.deepEqual(axis(0.37, 5, 0.01), { max: 0.4, step: 0.1 });
  for (const most of [0.003, 0.7, 3, 7, 19, 64, 99, 101, 250, 999, 12345, 7.7e8]) {
    for (const steps of [2, 3, 4, 6, 8]) {
      const { max, step } = roundAxis(most, steps);
      assert.ok(max >= most - 1e-9 && max / step <= steps + 1e-9 && max < 2.5 * most + 1e-9, `${most} in ${steps} steps: ${max} by ${step}`);
    }
  }
});

test('a value axis has three to five round labels and ends at the half step over the largest value', () => {
  const axis = (low: number, high: number, steps: number, more: Partial<Parameters<typeof roundSpan>[2]> = {}) => roundSpan(low, high, { steps, grain: 1, ...more });
  // 105 in a low plot: labels at 0, 50 and 100, and the plot ends at 125, not at 200 with labels at 0, 100, 200.
  assert.deepEqual(axis(0, 105, 2), { min: 0, max: 125, step: 50 });
  assert.deepEqual(axisLabels({ min: 0, max: 125, step: 50 }), [0, 50, 100]);
  // With room for more labels, 0 to 100 by 25, and the plot to 112.5.
  assert.deepEqual(axis(0, 105, 4), { min: 0, max: 112.5, step: 25 });
  // Whole steps before steps that write a decimal on every label: 0, 2, 4, 6 for 7.4, not 0.0, 2.5, 5.0, 7.5.
  assert.deepEqual(axis(3.2, 7.4, 3, { grain: 0.1 }), { min: 0, max: 7.4, step: 2 });
  assert.deepEqual(axis(1250000, 1790000, 4, { grain: 10000, unit: 1000000 }), { min: 0, max: 2000000, step: 1000000 });
  // A line that stays far above nothing starts at the step under its lowest value.
  assert.deepEqual(axis(32, 44, 4, { from: 32 }), { min: 30, max: 45, step: 5 });
  // Values under nothing start at the step under the lowest, and a step lower only where the number at the end
  // of the lowest bar needs more of the plot than that leaves it.
  assert.deepEqual(axis(-31, 4, 4), { min: -40, max: 5, step: 10 });
  assert.deepEqual(axis(-31, 4, 4, { under: 0.15 }), { min: -40, max: 5, step: 10 });
  assert.deepEqual(axis(-8.7, 10.5, 4, { grain: 0.1 }), { min: -10, max: 12.5, step: 5 });
  assert.deepEqual(axis(-8.7, 10.5, 4, { grain: 0.1, under: 0.2 }), { min: -20, max: 15, step: 10 });
  for (const high of [0.003, 0.7, 3, 7, 19, 64, 99, 101, 250, 999, 12345, 7.7e8]) {
    for (const low of [0, -0.3 * high, 0.7 * high]) {
      for (const steps of [2, 3, 4, 6, 8]) {
        const from = low > 0 ? low : undefined;
        const { min, max, step } = roundSpan(low, high, { steps, from });
        const labels = axisLabels({ min, max, step }).length;
        assert.ok(max >= high - 1e-9 && min <= (from ?? Math.min(0, low)) + 1e-9, `${low} to ${high}: ${min} to ${max}`);
        assert.ok(labels >= 3 && labels <= Math.min(5, steps + 1), `${low} to ${high} in ${steps} steps: ${labels} labels`);
      }
    }
  }
});

test('colours past the end of the chart colours are tints of them, never the first colour again', () => {
  const colors = ['0E7C7B', 'F4A259', '1B2A41'];
  const seven = Array.from({ length: 7 }, (_, index) => seriesColor(colors, index, 'FFFFFF'));
  assert.deepEqual(seven.slice(0, 3), colors);
  assert.equal(new Set(seven).size, 7);
  // A chart whose place is known draws its slices so; another repeats its colours, as it always did.
  const pie: ChartValue = { type: 'pie', categories: ['a', 'b', 'c', 'd'], series: [{ name: 's', values: [40, 30, 20, 10] }] };
  const slot = { ...(chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1]), colors };
  const fills = (xml: string) => [...xml.matchAll(/<c:dPt>.*?<a:srgbClr val="([0-9A-F]{6})"\/>/g)].map((match) => match[1]);
  assert.deepEqual(fills(buildChartXml(pie, slot, 'rId1')), ['0E7C7B', 'F4A259', '1B2A41', '0E7C7B']);
  const placed = fills(buildChartXml(pie, { ...slot, drawn: { width: 300, height: 200, ground: 'FFFFFF' } }, 'rId1'));
  assert.deepEqual(placed, [...colors, seven[3]]);
});

test('a chart whose place is known: a counted plot, upright names, the axis it is given, bars no thicker than they read', () => {
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const year: ChartValue = { type: 'column', categories: months, series: [{ name: 'Rides', values: months.map((_, index) => 1000 + 150 * index) }] };
  const at = (chart: ChartValue, drawn: NonNullable<typeof slot.drawn>, more: Partial<typeof slot> = {}) => {
    const xml = buildChartXml(chart, { ...slot, ...more, drawn }, 'rId1');
    rootElement(parseXml(xml));
    return xml;
  };
  const layoutOf = (xml: string) => {
    const found = /<c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"\/><c:xMode val="edge"\/><c:yMode val="edge"\/><c:x val="([\d.]+)"\/><c:y val="([\d.]+)"\/><c:w val="([\d.]+)"\/><c:h val="([\d.]+)"\/>/.exec(xml);
    assert.ok(found, 'the plot has a place of its own');
    return found.slice(1).map(Number) as [number, number, number, number];
  };
  // A small chart 150 pt wide, its numbers on an axis that ends where it is told: text of 9 pt.
  const small = at(year, { width: 150, height: 90, axis: { max: 3000, step: 1000 }, numbers: false }, { detail: 'compact' });
  assert.ok(small.includes('<c:max val="3000"/><c:min val="0"/>') && small.includes('<c:majorUnit val="1000"/>'));
  // Names stand upright, and every second one is written: twelve of them share the plot's 112 pt, 9.3 pt each, and "Jan" asks for 18.
  assert.ok(small.includes('<a:bodyPr rot="0" vert="horz"/>') && small.includes('<c:tickLblSkip val="2"/>'));
  assert.ok(small.includes('<c:formatCode>#,##0</c:formatCode>') && small.includes('formatCode="#,##0"'));
  // The plot: the axis labels (5 characters) at the left, the names under it, a little air above and at the right.
  const [x, y, w, h] = layoutOf(small);
  const margins = plotMargins({ horizontal: false, size: 9, names: 3, axis: 5, legend: false });
  assert.deepEqual(margins, { left: (5 * 0.55 + 1) * 9, right: 4.5, top: 0.7 * 9, bottom: 2.1 * 9 });
  // With a legend of its own under the names, the names keep less: the legend's line gives them air.
  assert.equal(plotMargins({ horizontal: false, size: 9, names: 3, axis: 5, legend: true }).bottom, 1.7 * 9 + 1.9 * 9);
  assert.ok(Math.abs(x - margins.left / 150) < 1e-3 && Math.abs(y - margins.top / 90) < 1e-3 && Math.abs(w - (150 - margins.left - margins.right) / 150) < 1e-3 && Math.abs(h - (90 - margins.top - margins.bottom) / 90) < 1e-3);
  // Numbers at the bars in place of the axis: the axis is not shown, the largest bar ends where the plot ends, and the number has its room above.
  const numbered = at({ ...column }, { width: 150, height: 90, axis: { max: 95 }, numbers: true, valueAxis: false }, { detail: 'compact' });
  assert.ok(/<c:valAx>(?:(?!<\/c:valAx>).)*?<c:delete val="1"\/>/.test(numbered) && /<c:catAx>(?:(?!<\/c:catAx>).)*?<c:delete val="0"\/>/.test(numbered));
  assert.ok(numbered.includes('<c:showVal val="1"/>') && numbered.includes('<c:max val="95"/>') && !numbered.includes('<c:majorUnit'));
  assert.ok(Math.abs(layoutOf(numbered)[1] - (1.3 * 9) / 90) < 1e-3);
  // Three columns in a wide frame are no wider than 11% of it; four horizontal bars no thicker than 2.6 times the text.
  const gap = (xml: string) => Number(/<c:gapWidth val="(\d+)"\/>/.exec(xml)?.[1]);
  assert.ok(gap(at(column, { width: 600, height: 300 })) > 150 && gap(at(year, { width: 600, height: 300 })) === 60);
  assert.ok(gap(at({ ...column, type: 'bar' }, { width: 600, height: 300 })) > 150);
  // Beside horizontal bars every category is named; with a value under nothing the names stand at the foot of the plot.
  assert.ok(at({ ...column, type: 'bar' }, { width: 300, height: 120 }).includes('<c:tickLblSkip val="1"/>'));
  assert.ok(at({ ...column, series: [{ name: 's', values: [-5, 10, 3] }] }, { width: 300, height: 200 }).includes('<c:tickLblPos val="low"/>'));
  // A legend of its own stands on the line kept for it at the foot; a chart told that its slide has one draws none.
  const pair: ChartValue = { ...column, series: [...column.series, { name: 'Returns', values: [4, 3, 6] }] };
  assert.match(at(pair, { width: 300, height: 200 }), /<c:legend><c:legendPos val="b"\/><c:layout><c:manualLayout><c:xMode val="edge"\/><c:yMode val="edge"\/><c:x val="[\d.]+"\/><c:y val="0\.92\d*"\/><\/c:manualLayout><\/c:layout><c:overlay val="0"\/><\/c:legend>/);
  assert.ok(!at(pair, { width: 300, height: 200, legend: false }).includes('<c:legend>'));
  // A bare chart keeps the whole frame for its plot, and its bars the height the plan gives them.
  const bare = at(column, { width: 60, height: 30, axis: { max: 120 } }, { detail: 'spark' });
  assert.ok(bare.includes('<c:max val="120"/>') && bare.includes('<c:x val="0.04"/><c:y val="0.06"/><c:w val="0.92"/><c:h val="0.88"/>'));
});

test('a chart that reads right to left starts at the right, and is told so by the deck or by the template for a deck that reads right to left', () => {
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const axisOf = (xml: string, name: 'catAx' | 'valAx') => new RegExp(`<c:${name}>(?:(?!</c:${name}>).)*</c:${name}>`).exec(xml)?.[0] ?? '';
  const orientation = (part: string) => /<c:orientation val="(\w+)"\/>/.exec(part)?.[1];
  const position = (part: string) => /<c:axPos val="(\w)"\/>/.exec(part)?.[1];

  // Columns and lines: the first category on the right, the value axis on the right; it still crosses at the first category.
  for (const type of ['column', 'line'] as const) {
    const plain = buildChartXml({ ...column, type }, slot, 'rId1', 'rtl');
    const mirrored = buildChartXml({ ...column, type }, slot, 'rId1', 'rtl', { mirrored: true });
    rootElement(parseXml(mirrored));
    assert.deepEqual([orientation(axisOf(plain, 'catAx')), position(axisOf(plain, 'valAx'))], ['minMax', 'l'], type);
    assert.deepEqual([orientation(axisOf(mirrored, 'catAx')), position(axisOf(mirrored, 'catAx')), orientation(axisOf(mirrored, 'valAx')), position(axisOf(mirrored, 'valAx'))], ['maxMin', 'b', 'minMax', 'r'], type);
    assert.ok(axisOf(mirrored, 'valAx').includes('<c:crosses val="autoZero"/>'));
  }
  // Horizontal bars grow from the right, with their names on the right; the first category stays at the top, the value axis at the foot.
  const bars = buildChartXml({ ...column, type: 'bar' }, slot, 'rId1', 'rtl', { mirrored: true });
  assert.deepEqual([orientation(axisOf(bars, 'catAx')), position(axisOf(bars, 'catAx')), orientation(axisOf(bars, 'valAx')), position(axisOf(bars, 'valAx'))], ['maxMin', 'r', 'maxMin', 'b']);
  assert.ok(axisOf(bars, 'valAx').includes('<c:crosses val="max"/>'));
  // A pie is drawn as it is.
  const pie: ChartValue = { type: 'pie', categories: ['a', 'b'], series: [{ name: 's', values: [60, 40] }] };
  assert.equal(buildChartXml(pie, slot, 'rId1', 'rtl', { mirrored: true }), buildChartXml(pie, slot, 'rId1', 'rtl'));

  // A chart whose place is known keeps the room of its value axis on the right, and its legend under the middle of the plot.
  const pair: ChartValue = { ...column, series: [...column.series, { name: 'Returns', values: [4, 3, 6] }] };
  const drawn = { width: 300, height: 200, axis: { max: 100, step: 25 } };
  const plotX = (xml: string) => {
    const found = /<c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"\/><c:xMode val="edge"\/><c:yMode val="edge"\/><c:x val="([\d.]+)"\/><c:y val="[\d.]+"\/><c:w val="([\d.]+)"\/>/.exec(xml);
    assert.ok(found);
    return [Number(found[1]), Number(found[2])] as const;
  };
  const legendX = (xml: string) => Number(/<c:legend><c:legendPos val="\w"\/><c:layout><c:manualLayout><c:xMode val="edge"\/><c:yMode val="edge"\/><c:x val="([\d.]+)"\//.exec(xml)?.[1]);
  const left = buildChartXml(pair, { ...slot, drawn }, 'rId1', 'rtl');
  const right = buildChartXml(pair, { ...slot, drawn }, 'rId1', 'rtl', { mirrored: true });
  const [x, w] = plotX(left);
  const [mx, mw] = plotX(right);
  // The value axis keeps on the right the room it took on the left.
  assert.ok(Math.abs(1 - mx - mw - x) < 2e-3, `the plot is mirrored: ${x},${w} and ${mx},${mw}`);
  // PowerPoint sets a legend under a chart from the left whatever the chart reads: in one that reads
  // right to left it stands upright at the left instead, first series at the top, and the plot keeps
  // room for it there and none under it.
  assert.ok(left.includes('<c:legendPos val="b"/>') && right.includes('<c:legendPos val="l"/>'));
  assert.ok(legendX(right) < 0.05 && mx > legendX(right) + 0.05, `${legendX(right)} then ${mx}`);
  const plotHigh = (xml: string) => Number(/<c:plotArea><c:layout><c:manualLayout>(?:(?!<\/c:manualLayout>).)*<c:h val="([\d.]+)"\/>/.exec(xml)?.[1]);
  assert.ok(plotHigh(right) > plotHigh(left), 'no line is kept for a legend under the plot');

  // Rendered: the template's setting reaches a deck that reads right to left; a deck states its own; a deck that reads left to right reads so.
  const template = (chartDirection?: 'rtl' | 'ltr') => loadTemplate(pptx, { ...chartManifest(), ...(chartDirection ? { chartDirection } : {}) });
  const chartOf = (file: Uint8Array) => strFromU8(unzipSync(file)['ppt/charts/chart1.xml'] as Uint8Array);
  const sheetOf = (file: Uint8Array) => strFromU8(unzipSync(unzipSync(file)['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'] as Uint8Array)['xl/worksheets/sheet1.xml'] as Uint8Array);
  const mirroredIn = (file: Uint8Array) => orientation(axisOf(chartOf(file), 'catAx')) === 'maxMin';
  const rtlDeck = renderDeck(template('rtl'), deckOf(slide(column))).pptx;
  assert.equal(fixtureManifest.direction, 'rtl');
  assert.ok(mirroredIn(rtlDeck) && sheetOf(rtlDeck).includes('<sheetView rightToLeft="1" workbookViewId="0"/>'));
  // The view the sheet names is the workbook's own.
  assert.ok(strFromU8(unzipSync(unzipSync(rtlDeck)['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'] as Uint8Array)['xl/workbook.xml'] as Uint8Array).includes('<bookViews><workbookView/></bookViews><sheets>'));
  const plainDeck = renderDeck(template(), deckOf(slide(column))).pptx;
  assert.ok(!mirroredIn(plainDeck) && !sheetOf(plainDeck).includes('sheetView'), 'left out, charts read left to right');
  assert.ok(!mirroredIn(renderDeck(template('rtl'), { ...deckOf(slide(column)), direction: 'ltr' }).pptx), 'a deck that reads left to right');
  assert.ok(!mirroredIn(renderDeck(template('rtl'), { ...deckOf(slide(column)), chartDirection: 'ltr' }).pptx), 'the deck says left to right');
  assert.ok(mirroredIn(renderDeck(template(), { ...deckOf(slide(column)), chartDirection: 'rtl' }).pptx), 'the deck says right to left');
  assert.ok(mirroredIn(renderDeck(template(), { ...deckOf(slide(column)), direction: 'ltr', chartDirection: 'rtl' }).pptx), 'whatever its text reads');
});

test('a chart is marked with the language of its names: Hebrew names in a deck that reads right to left are Hebrew', () => {
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const hebrew: ChartValue = { type: 'column', categories: ['רבעון 1', 'רבעון 2'], series: [{ name: 'הכנסות (מיליונים)', values: [3, 4] }] };
  const marks = (xml: string) => [...xml.matchAll(/<a:(defRPr|endParaRPr)\b([^>]*)>/g)].map((found) => `${found[1]}:${/lang="([^"]+)"/.exec(found[2] ?? '')?.[1] ?? '-'}`);
  assert.deepEqual(marks(buildChartXml(hebrew, slot, 'rId1', 'rtl', { language: 'he' })), ['defRPr:he-IL', 'endParaRPr:he-IL']);
  // English names in a Hebrew deck are English, as every chart was: the end of the paragraph only.
  assert.deepEqual(marks(buildChartXml(column, slot, 'rId1', 'rtl', { language: 'he' })), ['defRPr:-', 'endParaRPr:en-US']);
  // A French deck's chart is French; a Hebrew name among English ones in a deck that reads left to right is not enough to make it Hebrew.
  assert.deepEqual(marks(buildChartXml(column, slot, 'rId1', 'ltr', { language: 'fr' })), ['defRPr:fr-FR', 'endParaRPr:fr-FR']);
  assert.deepEqual(marks(buildChartXml({ ...column, categories: ['Q1', 'Q2', 'רבעון'] }, slot, 'rId1', 'ltr')), ['defRPr:-', 'endParaRPr:en-US']);
  // The names under columns whose place is known are marked alike.
  const drawn = buildChartXml(hebrew, { ...slot, drawn: { width: 300, height: 200 } }, 'rId1', 'rtl', { language: 'he' });
  assert.ok(drawn.includes('<a:defRPr lang="he-IL"/></a:pPr><a:endParaRPr lang="he-IL"/>'));
});

test('a value that is not known is a gap in the chart, an empty cell in its data and said so to a screen reader', () => {
  const gaps: ChartValue = { type: 'line', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Rides', values: [5, null, 9] }] };
  const slot = chartManifest().patterns.at(-1)?.slots[1] as Parameters<typeof buildChartXml>[1];
  const xml = buildChartXml(gaps, slot, 'rId1');
  assert.ok(xml.includes('<c:ptCount val="3"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>9</c:v></c:pt></c:numCache>'), 'no point for the second value');
  assert.ok(xml.includes('<c:dispBlanksAs val="gap"/>'));
  const sheet = strFromU8(unzipSync(buildChartWorkbook(gaps))['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.ok(sheet.includes('<row r="3"><c r="A3" t="inlineStr"><is><t xml:space="preserve">Q2</t></is></c></row>'), 'an empty cell');

  // A deck may say so; the deck's check takes it, and refuses a chart with no value known at all.
  const template = loadTemplate(pptx, chartManifest());
  assert.deepEqual(validateDeck(template, deckOf(slide(gaps))).filter((issue) => issue.severity === 'error'), []);
  assert.match(validateDeck(template, deckOf(slide({ ...gaps, series: [{ name: 'Rides', values: [null, null, null] }] })))[0]?.message ?? '', /no value that is known/);
  assert.match(validateDeck(template, deckOf(slide({ ...gaps, series: [{ name: 'Rides', values: [5, 'many', 9] }] })))[0]?.message ?? '', /not a number/);
  const file = renderDeck(template, deckOf(slide(gaps))).pptx;
  assert.ok(strFromU8(unzipSync(file)['ppt/charts/chart1.xml'] as Uint8Array).includes('<c:pt idx="2"><c:v>9</c:v></c:pt></c:numCache>'));
  // The fixture is a Hebrew template: its decks say it in Hebrew.
  assert.match(strFromU8(unzipSync(file)['ppt/slides/slide1.xml'] as Uint8Array), /descr="תרשים קו של Rides: Q1 5, Q2 אין נתון, Q3 9"/);
  const english = renderDeck(template, { ...deckOf(slide(gaps)), direction: 'ltr', language: 'en' }).pptx;
  assert.match(strFromU8(unzipSync(english)['ppt/slides/slide1.xml'] as Uint8Array), /descr="Line chart of Rides: Q1 5, Q2 no value, Q3 9"/);
});

test('a chart says its kind and its values to a screen reader, in the language of the deck', () => {
  assert.equal(chartAlt(column), 'Column chart of Adoptions: Q1 62, Q2 74, Q3 95');
  assert.equal(chartAlt({ ...column, format: 'percent' }), 'Column chart of Adoptions: Q1 62%, Q2 74%, Q3 95%');
  const pair: ChartValue = { type: 'bar', categories: ['A', 'B'], series: [{ name: 'In', values: [1, 2] }, { name: 'Out', values: [3, null] }] };
  assert.equal(chartAlt(pair), 'Bar chart. In: A 1, B 2. Out: A 3, B no value');
  assert.equal(chartAlt({ type: 'pie', categories: ['כן', 'לא'], series: [{ name: '', values: [60, 40] }] }, 'he', 'rtl'), 'תרשים עוגה: כן 60, לא 40');
  assert.equal(chartAlt(pair, 'he-IL', 'rtl'), 'תרשים פסים. In: A 1, B 2. Out: A 3, B אין נתון');
  // A Hebrew deck turned left to right says it in English.
  assert.equal(chartAlt(column, 'he', 'ltr'), 'Column chart of Adoptions: Q1 62, Q2 74, Q3 95');
});

/** Whether the first chart of a file runs its categories from the right. */
const mirroredOf = (file: Uint8Array): boolean => /<c:catAx>(?:(?!<\/c:catAx>).)*<c:orientation val="maxMin"\/>/.test(strFromU8(unzipSync(file)["ppt/charts/chart1.xml"] as Uint8Array));

test('a studio sets the direction of charts, and clears it; a value that is no direction is reported', async () => {
  const template = loadTemplate(pptx, chartManifest());
  const set = changeManifest(template, [{ op: 'setManifest', chartDirection: 'rtl' }]);
  assert.deepEqual(set.rejected, []);
  assert.equal(set.manifest.chartDirection, 'rtl');
  assert.match(set.applied[0]?.message ?? '', /draws the charts of a deck that reads right to left from the right/);
  const cleared = changeManifest(loadTemplate(pptx, set.manifest), [{ op: 'setManifest', chartDirection: null }]);
  assert.equal('chartDirection' in cleared.manifest, false);
  assert.match(changeManifest(template, [{ op: 'setManifest', chartDirection: 'up' as 'rtl' }]).rejected[0]?.reason ?? '', /the direction of charts is "rtl" or "ltr"/);
  assert.deepEqual(
    checkManifest(loadTemplate(pptx, { ...chartManifest(), chartDirection: 'right' as 'rtl' })).filter((issue) => issue.message.includes('chartDirection')).map((issue) => issue.message),
    ['chartDirection is "right"; it is "rtl" or "ltr", and charts read left to right'],
  );
  assert.ok(!mirroredOf(renderDeck(loadTemplate(pptx, { ...chartManifest(), chartDirection: 'right' as 'rtl' }), deckOf(slide(column))).pptx));

  // An instruction in words reaches it through editManifest; "" clears it.
  const answers = [
    { note: '', changes: [{ op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'chartDirection', value: 'rtl' }] }] },
    { note: '', changes: [{ op: 'setManifest', pattern: '', shape: '', slot: '', values: [{ property: 'chartDirection', value: '' }] }] },
  ];
  const { llm, requests } = scripted(...answers);
  const asked = await editManifest(template, 'Charts in Hebrew decks should start on the right.', { llm });
  assert.equal(asked.manifest.chartDirection, 'rtl');
  assert.match(requests[0]?.prompt ?? '', /chartDirection \("rtl" draws the charts of a deck that reads right to left from the right/);
  const back = await editManifest(loadTemplate(pptx, asked.manifest), 'Draw charts left to right again.', { llm });
  assert.equal('chartDirection' in back.manifest, false);
  assert.match(requests[1]?.prompt ?? '', /the charts of a deck that reads right to left are drawn from the right/);
});

test('keptGaps: a gap stays where the chart keeps its shape and the request gives no number for it', () => {
  const before: SlideSpec = { pattern: 'chart', slots: { title: 'T', chart: { type: 'line', categories: ['a', 'b', 'c'], series: [{ name: 'x', values: [1, null, 3] }, { name: 'y', values: [null, 5, 6] }] } } };
  const after = (values: Array<Array<number | null>>, categories = ['A', 'B', 'C'], names = ['X', 'Y']): SlideSpec => ({ pattern: 'chart', slots: { title: 'T', chart: { type: 'line', categories, series: names.map((name, at) => ({ name, values: values[at] ?? [] })) } } });
  const valuesOf = (slide: SlideSpec) => (slide.slots.chart as ChartValue).series.map((series) => series.values);
  // Renamed (translated): matched by place.
  assert.deepEqual(valuesOf(keptGaps(before, after([[1, 0, 3], [9, 5, 6]]), new Set([9]))), [[1, null, 3], [9, 5, 6]]);
  // Fewer series: matched by name, and one that is new is left as it is.
  assert.deepEqual(valuesOf(keptGaps(before, after([[1, 7, 3]], ['a', 'b', 'c'], ['x']), new Set())), [[1, null, 3]]);
  assert.deepEqual(valuesOf(keptGaps(before, after([[1, 7, 3]], ['a', 'b', 'c'], ['z']), new Set())), [[1, 7, 3]]);
  // Another number of categories is another chart; nothing changed gives the same slide.
  const longer = after([[1, 2, 3, 4], [5, 6, 7, 8]], ['a', 'b', 'c', 'd']);
  assert.equal(keptGaps(before, longer, new Set()), longer);
  const same = after([[1, null, 3], [null, 5, 6]]);
  assert.equal(keptGaps(before, same, new Set()), same);
});
