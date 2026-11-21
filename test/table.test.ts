import assert from 'node:assert/strict';
import { test } from 'node:test';
import { placeholderPhoto } from '../examples/placeholder.ts';
import { DeckValidationError, PptxPackage, changeManifest, loadTemplate, planDeck, planEdit, renderDeck, validateDeck } from '../src/index.ts';
import type { DeckSpec, JsonRequest, LlmProvider, SlideSpec, TemplateManifest } from '../src/index.ts';
import { slotSources } from '../src/template/manifest.ts';
import { NS, childElement, childElements, descendants, parseXml, path, rootElement } from '../src/pptx/xml.ts';
import { fitTable, setTable, tableFrameXml } from '../src/render/table.ts';
import type { Element } from '../src/pptx/xml.ts';
import { buildFixturePptx, fixtureManifest, fixtureWithSlide, tableManifest } from './fixture.ts';

const pptx = buildFixturePptx();
const template = loadTemplate(pptx, tableManifest);

const tableSlide = (header: string[], rows: string[][]): SlideSpec => ({ pattern: 'table', slots: { title: 'Breeds', table: { header, rows } } });

function renderTable(slide: SlideSpec, deck: Partial<DeckSpec> = {}): { root: Element; tbl: Element } {
  const result = renderDeck(template, { version: 1, ...deck, slides: [slide] });
  const root = rootElement(parseXml(PptxPackage.load(result.pptx).readText('ppt/slides/slide1.xml')));
  return { root, tbl: descendants(root, NS.a, 'tbl')[0] as Element };
}

const cellText = (cell: Element) =>
  descendants(cell, NS.a, 't')
    .map((node) => node.textContent ?? '')
    .join('');
const rowsOf = (tbl: Element) => childElements(tbl, NS.a, 'tr');
const cellsOf = (row: Element) => childElements(row, NS.a, 'tc');
const fillOf = (cell: Element) => path(cell, [NS.a, 'tcPr'], [NS.a, 'solidFill'], [NS.a, 'srgbClr'])?.getAttribute('val') ?? null;

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

test('a table takes the size of its content and keeps the sample rows as its styles', () => {
  const { root, tbl } = renderTable(
    tableSlide(
      ['Breed', 'Size'],
      [
        ['Labrador', 'Large'],
        ['Bulldog', 'Small'],
        ['Collie', 'Medium'],
      ],
    ),
  );

  // Two columns share the sample table's width.
  const columns = childElements(childElement(tbl, NS.a, 'tblGrid') as Element, NS.a, 'gridCol');
  assert.deepEqual(
    columns.map((column) => column.getAttribute('w')),
    ['3000000', '3000000'],
  );

  const rows = rowsOf(tbl);
  assert.deepEqual(
    rows.map((row) => cellsOf(row).map(cellText)),
    [
      ['Breed', 'Size'],
      ['Labrador', 'Large'],
      ['Bulldog', 'Small'],
      ['Collie', 'Medium'],
    ],
  );
  // Header fill, then the two body samples in turn.
  assert.deepEqual(
    rows.map((row) => fillOf(cellsOf(row)[0] as Element)),
    ['00683A', null, 'DCDDDE', null],
  );

  assert.ok(!PptxPackageText(root).includes('SAMPLE'), 'no sample text is left');
  assert.equal(descendants(tbl, 'http://schemas.microsoft.com/office/drawing/2014/main', 'rowId').length, 0, 'copied rows carry no duplicate ids');
  const frame = tbl.parentNode?.parentNode?.parentNode as Element;
  assert.equal(path(frame, [NS.p, 'xfrm'], [NS.a, 'ext'])?.getAttribute('cy'), '1200000', 'frame height follows the rows');
});

function PptxPackageText(root: Element): string {
  return descendants(root, NS.a, 't')
    .map((node) => node.textContent ?? '')
    .join(' ');
}

test('a table can have more columns than its sample', () => {
  const { tbl } = renderTable(tableSlide(['A', 'B', 'C', 'D'], [['1', '2', '3', '4']]));
  assert.equal(childElements(childElement(tbl, NS.a, 'tblGrid') as Element, NS.a, 'gridCol').length, 4);
  assert.deepEqual(
    rowsOf(tbl).map((row) => cellsOf(row).length),
    [4, 4],
  );
  assert.equal(fillOf(cellsOf(rowsOf(tbl)[0] as Element)[3] as Element), '00683A', 'the extra header cell copies the last sample cell');
});

test('a table runs in the reading direction of the deck', () => {
  const slide = tableSlide(['A', 'B'], [['1', '2']]);
  const look = (deck: Partial<DeckSpec>) => {
    const { tbl } = renderTable(slide, deck);
    const firstParagraph = descendants(tbl, NS.a, 'pPr')[0] as Element;
    return [childElement(tbl, NS.a, 'tblPr')?.getAttribute('rtl'), firstParagraph.getAttribute('algn')];
  };
  assert.deepEqual(look({}), ['1', 'r'], 'the fixture template reads right to left');
  assert.deepEqual(look({ direction: 'ltr' }), ['0', 'l']);
});

test('table text takes the formatting the manifest gives it', () => {
  const { tbl } = renderTable(tableSlide(['Breed', 'Size'], [['Labrador', 'Large']]));
  const runProps = (row: number) => descendants(cellsOf(rowsOf(tbl)[row] as Element)[0] as Element, NS.a, 'rPr')[0] as Element;

  const header = runProps(0);
  assert.equal(header.getAttribute('sz'), '1000');
  assert.equal(header.getAttribute('b'), '1');
  // The schema fixes the order of run properties: fill first, typefaces after.
  assert.deepEqual(
    childElements(header).map((child) => child.localName),
    ['solidFill', 'latin', 'ea', 'cs', 'sym'],
  );
  assert.equal(path(header, [NS.a, 'solidFill'], [NS.a, 'srgbClr'])?.getAttribute('val'), 'FFFFFF');
  assert.equal(childElement(header, NS.a, 'latin')?.getAttribute('typeface'), 'Assistant');

  const body = runProps(1);
  assert.equal(body.getAttribute('sz'), '900');
  assert.equal(body.getAttribute('b'), null);
});

test('the rows of a placed table share its height equally, or as they are told to', () => {
  const area = { x: 100000, y: 200000, cx: 6000000, cy: 1200000 };
  const placed = (shares?: number[]) => {
    const frame = rootElement(parseXml(tableFrameXml('9', area, { header: '0E7C7B', band: 'F2F2F2', line: 'D9D9D9' })));
    setTable(frame, { header: ['Station', 'Docks'], rows: [['Tram Square', '24'], ['River Gate', '18']] }, { name: 'table', kind: 'table', shapeId: '9' }, 'ltr');
    fitTable(frame, area, shares);
    const tbl = descendants(frame, NS.a, 'tbl')[0] as Element;
    const extent = path(frame, [NS.p, 'xfrm'], [NS.a, 'ext']);
    return { rows: rowsOf(tbl).map((row) => Number(row.getAttribute('h'))), height: Number(extent?.getAttribute('cy')), above: path(cellsOf(rowsOf(tbl)[0] as Element)[0] as Element, [NS.a, 'tcPr'])?.getAttribute('marT') };
  };
  assert.deepEqual(placed(), { rows: [400000, 400000, 400000], height: 1200000, above: '40000' });
  // A row whose cells take two lines is higher than one whose cells take one: 20, 40 and 40 parts of the height.
  assert.deepEqual(placed([20, 40, 40]), { rows: [240000, 480000, 480000], height: 1200000, above: '24000' });
  // Shares that do not match the rows are not followed.
  assert.deepEqual(placed([1, 2]).rows, [400000, 400000, 400000]);
});

test('table values are validated', () => {
  const check = (slots: SlideSpec['slots']) => validateDeck(tableManifest, { version: 1, slides: [{ pattern: 'table', slots }] });
  const table = (header: string[], rows: string[][]) => ({ title: 'T', table: { header, rows } });

  assert.deepEqual(check(table(['A', 'B'], [['1', '2']])), []);
  assert.match(check(table(['A', 'B'], [['1']]))[0]?.message ?? '', /row 1 has 1 cells; the table has 2 columns/);
  assert.match(check(table([], []))[0]?.message ?? '', /no columns/);
  assert.match(check({ title: 'T', table: 'not a table' })[0]?.message ?? '', /expected a table/);
  assert.match(check({ title: { header: ['A'], rows: [] }, table: { header: ['A', 'B'], rows: [] } })[0]?.message ?? '', /expected text/);

  const big = check(table(['A', 'B', 'C', 'D'], [['1', '2', '3', 'a cell that is too long'], ['1', '2', '3', '4'], ['1', '2', '3', '4'], ['1', '2', '3', '4']]));
  assert.deepEqual(
    big.map((issue) => [issue.severity, issue.message.replace(/\d+/g, 'N')]),
    [
      ['warning', 'the table has N columns; the slot fits N'],
      ['warning', 'the table has N rows; the slot fits N'],
      ['warning', 'a cell is N characters; cells fit about N'],
    ],
  );

  assert.throws(() => renderDeck(template, { version: 1, slides: [{ pattern: 'table', slots: table(['A', 'B'], [['1']]) }] }), DeckValidationError);
});

test('the planner asks for a table as header and rows, and squares up uneven rows', async () => {
  const { llm, requests } = scripted(
    { title: 'T', slides: [{ pattern: 'table', purpose: 'Compare breeds' }] },
    { s1: { title: 'Breeds', table: { header: ['Breed', 'Size'], rows: [['Labrador', 'Large'], ['Bulldog'], ['Collie', 'Medium', 'extra']] } } },
  );
  const { deck } = await planDeck(template, { prompt: 'Compare dog breeds', language: 'en', imageDensity: 'none' }, { llm });

  const schema = (requests[1]?.schema as { properties: { s1: { properties: { table: Record<string, unknown> } } } }).properties.s1.properties.table as {
    properties: { header: { maxItems: number }; rows: { maxItems: number; items: { type: string; maxItems: number } } };
    required: string[];
  };
  assert.deepEqual(schema.required, ['header', 'rows']);
  assert.equal(schema.properties.header.maxItems, 3);
  assert.equal(schema.properties.rows.maxItems, 3);
  assert.equal(schema.properties.rows.items.type, 'array');
  assert.match(requests[0]?.prompt ?? '', /table \(table: at most 3 columns, at most 3 rows, at most 10 characters in a cell\)/);

  assert.deepEqual(deck.slides[0]?.slots.table, {
    header: ['Breed', 'Size'],
    rows: [
      ['Labrador', 'Large'],
      ['Bulldog', ''],
      ['Collie', 'Medium'],
    ],
  });
});

test('an edit shows the model the table as it is', async () => {
  const deck: DeckSpec = { version: 1, language: 'en', direction: 'ltr', slides: [tableSlide(['Breed', 'Size'], [['Labrador', 'Large']])] };
  const { llm, requests } = scripted(
    { summary: 'Added a row.', language: 'en', title: '', slides: [{ from: 1, pattern: 'table', change: 'Add a row for the collie', images: false }] },
    { s1: { title: 'Breeds', table: { header: ['Breed', 'Size'], rows: [['Labrador', 'Large'], ['Collie', 'Medium']] } } },
  );
  const result = await planEdit(template, deck, 'Add a row for the collie', { llm });

  assert.match(requests[0]?.prompt ?? '', /table: table \{"header":\["Breed","Size"\],"rows":\[\["Labrador","Large"\]\]\}/);
  assert.match(requests[1]?.prompt ?? '', /Current content: \{"title":"Breeds","table":\{"header":\["Breed","Size"\]/);
  assert.deepEqual((result.deck.slides[0]?.slots.table as { rows: string[][] }).rows.length, 2);
});

test('a picture slot can be given a larger area than its sample frame', () => {
  const withArea: TemplateManifest = {
    ...fixtureManifest,
    patterns: fixtureManifest.patterns.map((pattern) => ({
      ...pattern,
      slots: pattern.slots.map((slot) =>
        slot.name === 'photo' && slot.kind === 'image' ? { ...slot, imageKind: 'graphic' as const, box: { x: 0.1, y: 0.1, w: 0.5, h: 0.5 } } : slot,
      ),
    })),
  };
  const square = placeholderPhoto(50, 50, [0, 0, 0], [9, 9, 9]);
  const slide: SlideSpec = { pattern: 'content', slots: { title: 'T', body: 'B', note: 'N', photo: { asset: 'chart' } } };
  const result = renderDeck(loadTemplate(pptx, withArea), { version: 1, slides: [slide] }, { assets: { chart: square } });
  const root = rootElement(parseXml(PptxPackage.load(result.pptx).readText('ppt/slides/slide1.xml')));
  const pic = descendants(root, NS.p, 'pic').find((el) => path(el, [NS.p, 'nvPicPr'], [NS.p, 'cNvPr'])?.getAttribute('id') === '4') as Element;
  const xfrm = path(pic, [NS.p, 'spPr'], [NS.a, 'xfrm']) as Element;

  // The area is 4572000 x 2571750 at (914400, 514350); a square picture is centred in it.
  assert.equal(childElement(xfrm, NS.a, 'ext')?.getAttribute('cx'), '2571750');
  assert.equal(childElement(xfrm, NS.a, 'ext')?.getAttribute('cy'), '2571750');
  assert.equal(childElement(xfrm, NS.a, 'off')?.getAttribute('x'), '1914525');
  assert.equal(childElement(xfrm, NS.a, 'off')?.getAttribute('y'), '514350');

  const grouped: TemplateManifest = {
    ...fixtureManifest,
    patterns: fixtureManifest.patterns.map((pattern) => ({
      ...pattern,
      slots: pattern.slots.map((slot) => (slot.name === 'icon' && slot.kind === 'image' ? { ...slot, box: { x: 0, y: 0, w: 0.2, h: 0.2 } } : slot)),
    })),
  };
  const withIcon: SlideSpec = { pattern: 'content', slots: { ...slide.slots, icon: { asset: 'chart' } } };
  assert.throws(() => renderDeck(loadTemplate(pptx, grouped), { version: 1, slides: [withIcon] }, { assets: { chart: square } }), /not inside a group/);
});

test('each row of a grid is mirrored on its own', () => {
  const sources = slotSources(
    {
      id: 'grid',
      sourceSlide: 1,
      description: '',
      slots: ['a1', 'a2', 'a3', 'b1', 'b2', 'b3'].map((name, index) => ({ name, kind: 'text' as const, shapeId: String(index) })),
      sequences: [
        [['a1'], ['a2'], ['a3']],
        [['b1'], ['b2'], ['b3']],
      ],
    },
    true,
  );
  assert.deepEqual(Object.fromEntries(sources), { a1: 'a3', a2: 'a2', a3: 'a1', b1: 'b3', b2: 'b2', b3: 'b1' });
});

test('the columns of a placed table share its width as they are told to, single columns can end on one edge, and cells have their text in the middle', () => {
  const area = { x: 100000, y: 200000, cx: 6000000, cy: 1200000 };
  const placed = (widths?: number[], aligns?: Array<'l' | 'r' | undefined>) => {
    const frame = rootElement(parseXml(tableFrameXml('9', area, { header: '0E7C7B', band: 'F2F2F2', line: 'D9D9D9' })));
    setTable(frame, { header: ['Station', 'Docks', 'Rides'], rows: [['Tram Square', '24', '410'], ['River Gate', '18', '365']] }, { name: 'table', kind: 'table', shapeId: '9' }, 'ltr', aligns);
    fitTable(frame, area, undefined, widths);
    const tbl = descendants(frame, NS.a, 'tbl')[0] as Element;
    return {
      columns: childElements(childElement(tbl, NS.a, 'tblGrid') as Element, NS.a, 'gridCol').map((column) => Number(column.getAttribute('w'))),
      aligns: cellsOf(rowsOf(tbl)[1] as Element).map((cell) => descendants(cell, NS.a, 'pPr')[0]?.getAttribute('algn')),
      heading: cellsOf(rowsOf(tbl)[0] as Element).map((cell) => descendants(cell, NS.a, 'pPr')[0]?.getAttribute('algn')),
      anchors: [...new Set(descendants(tbl, NS.a, 'tcPr').map((props) => props.getAttribute('anchor')))],
    };
  };
  // Left to themselves the columns share the width equally, and every cell starts where reading starts.
  assert.deepEqual(placed(), { columns: [2000000, 2000000, 2000000], aligns: ['l', 'l', 'l'], heading: ['l', 'l', 'l'], anchors: ['ctr'] });
  // A column of names twice as wide as two columns of numbers, which end on one edge, their headings with them.
  const told = placed([200, 100, 100], [undefined, 'r', 'r']);
  assert.deepEqual(told, { columns: [3000000, 1500000, 1500000], aligns: ['l', 'r', 'r'], heading: ['l', 'r', 'r'], anchors: ['ctr'] });
  // The columns always take the whole width, and widths that do not match the columns are not followed.
  assert.equal(placed([1, 1, 1.0001]).columns.reduce((sum, width) => sum + width, 0), 6000000);
  assert.deepEqual(placed([2, 1]).columns, [2000000, 2000000, 2000000]);
});

test('a cell of a figure without letters reads left to right in a deck that reads left to right from a template that reads right to left', () => {
  const directions = (sample: string, deck: Partial<DeckSpec>) => {
    // The sample table's cells read right to left (rtl="1"), or state no direction and take the template's.
    const hebrew = loadTemplate(fixtureWithSlide(3, (xml) => xml.replaceAll(' rtl="0"', sample)), tableManifest);
    const file = renderDeck(hebrew, { version: 1, ...deck, slides: [tableSlide(['Station', 'Change'], [['Tram Square', '+12%'], ['River Gate', '4.6 / 5']])] }).pptx;
    const tbl = descendants(rootElement(parseXml(PptxPackage.load(file).readText('ppt/slides/slide1.xml'))), NS.a, 'tbl')[0] as Element;
    return rowsOf(tbl).flatMap((row) => cellsOf(row).map((cell) => `${cellText(cell)}:${path(cell, [NS.a, 'txBody'], [NS.a, 'p'], [NS.a, 'pPr'])?.getAttribute('rtl') ?? '-'}`));
  };
  // "+12%" in a paragraph that reads right to left shows as "12%+".
  for (const sample of [' rtl="1"', '']) {
    assert.deepEqual(directions(sample, { direction: 'ltr' }), ['Station:0', 'Change:0', 'Tram Square:0', '+12%:0', 'River Gate:0', '4.6 / 5:0'], sample);
  }
  // In a deck that reads as the template does, a figure keeps the direction the template gives it.
  assert.deepEqual(directions(' rtl="1"', {}), ['Station:0', 'Change:0', 'Tram Square:0', '+12%:1', 'River Gate:0', '4.6 / 5:1']);
  assert.deepEqual(directions('', {}), ['Station:0', 'Change:0', 'Tram Square:0', '+12%:-', 'River Gate:0', '4.6 / 5:-']);
});

test('a table that fills its frame keeps the sample\'s height with fewer rows, and its columns of numbers end on the right in a deck read from the left', () => {
  const filling: TemplateManifest = { ...tableManifest, patterns: tableManifest.patterns.map((pattern) => ({ ...pattern, slots: pattern.slots.map((slot) => (slot.kind === 'table' ? { ...slot, fill: true } : slot)) })) };
  const draw = (manifest: TemplateManifest, rows: string[][], deck: Partial<DeckSpec> = {}) => {
    const file = renderDeck(loadTemplate(pptx, manifest), { version: 1, ...deck, slides: [tableSlide(['Station', 'Docks'], rows)] }).pptx;
    const root = rootElement(parseXml(PptxPackage.load(file).readText('ppt/slides/slide1.xml')));
    const frame = descendants(root, NS.p, 'graphicFrame')[0] as Element;
    const tbl = descendants(frame, NS.a, 'tbl')[0] as Element;
    return {
      height: Number(path(frame, [NS.p, 'xfrm'], [NS.a, 'ext'])?.getAttribute('cy')),
      rows: rowsOf(tbl).map((row) => Number(row.getAttribute('h'))),
      aligns: rowsOf(tbl).map((row) => cellsOf(row).map((cell) => path(cell, [NS.a, 'txBody'], [NS.a, 'p'], [NS.a, 'pPr'])?.getAttribute('algn')).join('')),
      anchors: descendants(tbl, NS.a, 'tcPr').map((props) => props.getAttribute('anchor') ?? '-').join(''),
    };
  };
  const one = [['Tram Square', '24']];
  // The sample table is three rows of 300,000 EMU: a heading and one row end two thirds of the way down.
  assert.deepEqual(draw(tableManifest, one).rows, [300000, 300000]);
  const filled = draw(filling, one);
  assert.deepEqual([filled.height, filled.rows, filled.anchors], [900000, [450000, 450000], 'ctrctrctrctr']);
  // The fixture reads right to left: every column starts at the right already. Read from the left, the docks end on the right.
  assert.deepEqual(filled.aligns, ['rr', 'rr']);
  assert.deepEqual(draw(filling, one, { direction: 'ltr' }).aligns, ['lr', 'lr']);
  assert.deepEqual(draw(tableManifest, one, { direction: 'ltr' }).aligns, ['ll', 'll'], 'as before without it');
  // A table of as many rows as the sample, or more, is set as before; only its numbers move.
  const three = [['Tram Square', '24'], ['River Gate', '18']];
  assert.deepEqual(draw(filling, three).rows, draw(tableManifest, three).rows);
  // A studio sets it, and false takes it away.
  const template = loadTemplate(pptx, tableManifest);
  const set = changeManifest(template, [{ op: 'setSlot', pattern: 'table', slot: 'table', fill: true }]);
  assert.deepEqual(set.rejected, []);
  assert.match(set.applied[0]?.message ?? '', /filling its frame/);
  const unset = changeManifest(loadTemplate(pptx, set.manifest), [{ op: 'setSlot', pattern: 'table', slot: 'table', fill: false }]);
  assert.equal('fill' in (unset.manifest.patterns.find((pattern) => pattern.id === 'table')?.slots.find((slot) => slot.name === 'table') ?? {}), false);
  assert.match(changeManifest(template, [{ op: 'setSlot', pattern: 'table', slot: 'title', fill: true }]).rejected[0]?.reason ?? '', /fill does not apply to it/);
});
