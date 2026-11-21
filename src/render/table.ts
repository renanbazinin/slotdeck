import type { TableValue } from '../deck/spec.ts';
import { NS, childElement, childElements, descendants, escapeXmlAttr, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import type { Direction, TableSlotDef } from '../template/manifest.ts';
import { setBodyText } from './text.ts';
import type { TextOptions } from './text.ts';

/**
 * Rebuilds a table from the sample table's own rows and cells, so borders, fills and banding
 * stay as designed while the number of rows and columns follows the content.
 *
 * - The first sample row is the header; the body rows that follow are reused in turn.
 * - Columns share the table's width equally.
 * - The table runs in the deck's reading direction, with text aligned to its start.
 *
 * `aligns` sets the cells of single columns otherwise: `'r'` for a column of numbers, which are
 * read by their last digits. A column it does not name starts where reading starts.
 *
 * `writing` names the deck's language and the template's direction: in a deck that reads left to
 * right from a template that reads right to left, a cell of a figure without letters ("+12%")
 * reads left to right, as the text of a slide does (see `applyDirection` in text.ts).
 */
export function setTable(frame: Element, table: TableValue, slot: TableSlotDef, direction: Direction, aligns?: ReadonlyArray<'l' | 'r' | undefined>, writing: { language?: string; templateDirection?: Direction } = {}): void {
  const tbl = descendants(frame, NS.a, 'tbl')[0];
  if (frame.localName !== 'graphicFrame' || !tbl) throw new Error('table slots must point at a table');
  const grid = childElement(tbl, NS.a, 'tblGrid');
  const sampleRows = childElements(tbl, NS.a, 'tr');
  const sampleColumns = grid ? childElements(grid, NS.a, 'gridCol') : [];
  const headerSample = sampleRows[0];
  if (!grid || !headerSample || sampleColumns.length === 0) throw new Error('the sample table has no rows or columns to copy');
  if (descendants(tbl, NS.a, 'tc').some((cell) => ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some((name) => cell.hasAttribute(name)))) {
    throw new Error('sample tables with merged cells are not supported');
  }
  // With two or more body rows the first two carry the banding; a single one is simply repeated.
  const bodySamples = sampleRows.slice(1, 3);
  if (bodySamples.length === 0) bodySamples.push(headerSample);

  const doc = tbl.ownerDocument;
  if (!doc) throw new Error('table is detached from its document');
  const prefix = tbl.prefix ? `${tbl.prefix}:` : '';
  const columns = table.header.length;
  const align = direction === 'rtl' ? 'r' : 'l';

  // Columns: equal shares of the original width.
  const width = sampleColumns.reduce((sum, column) => sum + Number(column.getAttribute('w') ?? 0), 0);
  const share = Math.floor(width / columns);
  for (const column of sampleColumns) removeNode(column);
  for (let index = 0; index < columns; index++) {
    const column = doc.createElementNS(NS.a, `${prefix}gridCol`);
    column.setAttribute('w', String(index === columns - 1 ? width - share * (columns - 1) : share));
    grid.appendChild(column);
  }

  const buildRow = (sample: Element, cells: string[], style: TableSlotDef['headerText']): Element => {
    const row = sample.cloneNode(true) as Element;
    // Row ids must be unique, and a copied row would repeat its sample's.
    for (const extension of childElements(row, NS.a, 'extLst')) removeNode(extension);
    const sampleCells = childElements(row, NS.a, 'tc');
    const lastSample = sampleCells[sampleCells.length - 1];
    if (!lastSample) throw new Error('a sample table row has no cells');
    for (const cell of sampleCells) removeNode(cell);
    cells.forEach((text, index) => {
      const cell = (sampleCells[index] ?? lastSample).cloneNode(true) as Element;
      const body = childElement(cell, NS.a, 'txBody');
      const cellAlign = aligns?.[index] ?? align;
      const options: TextOptions = { flipAlign: false, align: cellAlign, direction, ...(writing.language ? { language: writing.language } : {}), ...(writing.templateDirection ? { templateDirection: writing.templateDirection } : {}) };
      if (body) setBodyText(body, [text], style ? { ...options, style } : options);
      row.insertBefore(cell, childElement(row, NS.a, 'extLst') ?? null);
    });
    return row;
  };

  const rows = [
    buildRow(headerSample, table.header, slot.headerText),
    ...table.rows.map((cells, index) => buildRow(bodySamples[index % bodySamples.length] as Element, cells, slot.bodyText)),
  ];
  for (const row of sampleRows) removeNode(row);
  for (const row of rows) tbl.appendChild(row);

  let props = childElement(tbl, NS.a, 'tblPr');
  if (!props) {
    props = doc.createElementNS(NS.a, `${prefix}tblPr`);
    tbl.insertBefore(props, tbl.firstChild);
  }
  props.setAttribute('rtl', direction === 'rtl' ? '1' : '0');

  // PowerPoint recomputes the frame height from the rows; other viewers use the stored value.
  const extent = descendants(frame, NS.a, 'ext').find((el) => el.parentNode && (el.parentNode as Element).localName === 'xfrm');
  const height = rows.reduce((sum, row) => sum + Number(row.getAttribute('h') ?? 0), 0);
  if (extent && height > 0) extent.setAttribute('cy', String(height));
}

/** The area a table's frame takes on its slide, in EMU; `undefined` for a frame that states none. */
export function tableArea(frame: Element): Area | undefined {
  const xfrm = childElement(frame, NS.p, 'xfrm');
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  const area = { x: Number(off?.getAttribute('x')), y: Number(off?.getAttribute('y')), cx: Number(ext?.getAttribute('cx')), cy: Number(ext?.getAttribute('cy')) };
  return Object.values(area).every(Number.isFinite) && area.cx > 0 && area.cy > 0 ? area : undefined;
}

/** A cell that is a number: a sign, a currency, then a digit. */
const NUMBER_CELL = /^[+\-−]?[€$₪£]?\s?\d/;

/**
 * How the columns of a table are aligned where its columns of numbers end on one edge: `'r'` for a
 * column whose every body cell is a number, in a deck that reads left to right. A deck that reads
 * right to left sets every column from the right already, so nothing changes there.
 */
export function numberAligns(table: TableValue, direction: Direction): Array<'l' | 'r' | undefined> | undefined {
  if (direction === 'rtl' || table.rows.length === 0) return undefined;
  const numeric = table.header.map((_, column) => table.rows.every((row) => NUMBER_CELL.test((row[column] ?? '').trim())));
  return numeric.some(Boolean) ? numeric.map((yes) => (yes ? 'r' : undefined)) : undefined;
}

/** An area of the slide in EMU. */
export interface Area {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/**
 * Puts a table that {@link setTable} has just rebuilt into an area of the slide: the columns
 * share its width and the rows its height. The space kept free inside a cell shrinks with the
 * cell, so that a small table is not all margin. PowerPoint still makes a row taller when its
 * text needs it, so the text size has to suit the row: {@link tableTextSize} says which does.
 *
 * `shares` gives the rows unequal heights, for a table whose cells take more lines in one row
 * than in another: one number for each row, the heading row first, and the rows share the
 * height in these proportions. Left out, they share it equally. `widths` does the same for the
 * columns: one number for each, and the columns share the width in these proportions, so that
 * a column of long names is wider than a column of short numbers.
 *
 * Every cell has its text in the middle of its height: rows that share a height are higher
 * than their text, and text at the top of a high row hangs. `keepMargins` leaves the space kept
 * free inside each cell as the template's table has it: for a table of a layout that fills its
 * frame (`TableSlotDef.fill`), whose cells are as the designer drew them.
 */
export function fitTable(frame: Element, area: Area, shares?: readonly number[], widths?: readonly number[], options: { keepMargins?: boolean } = {}): void {
  const tbl = descendants(frame, NS.a, 'tbl')[0];
  const grid = tbl && childElement(tbl, NS.a, 'tblGrid');
  const xfrm = childElement(frame, NS.p, 'xfrm');
  const off = xfrm && childElement(xfrm, NS.a, 'off');
  const ext = xfrm && childElement(xfrm, NS.a, 'ext');
  if (!tbl || !grid || !off || !ext) throw new Error('a table can only be placed when it has a frame and a grid');
  const columns = childElements(grid, NS.a, 'gridCol');
  const rows = childElements(tbl, NS.a, 'tr');
  if (columns.length === 0 || rows.length === 0) return;
  const width = Math.max(1, Math.round(area.cx));
  const height = Math.max(1, Math.round(area.cy));
  const columnWidth = Math.floor(width / columns.length);
  const whole = shares && shares.length === rows.length && shares.every((share) => share > 0) ? shares.reduce((sum, share) => sum + share, 0) : 0;
  const heights = rows.map((_, index) => Math.floor(whole > 0 ? (height * ((shares as readonly number[])[index] as number)) / whole : height / rows.length));
  // What is kept free above and below the text is one measure for the whole table: the lowest row's.
  const rowHeight = heights.reduce((least, value) => Math.min(least, value), Infinity);
  const across = widths && widths.length === columns.length && widths.every((share) => share > 0) ? widths.reduce((sum, share) => sum + share, 0) : 0;
  const columnWidths = columns.map((_, index) => (across > 0 ? Math.floor((width * ((widths as readonly number[])[index] as number)) / across) : columnWidth));
  columnWidths[columnWidths.length - 1] = width - columnWidths.slice(0, -1).reduce((sum, value) => sum + value, 0);
  columns.forEach((column, index) => column.setAttribute('w', String(columnWidths[index])));
  rows.forEach((row, index) => row.setAttribute('h', String(heights[index])));
  off.setAttribute('x', String(Math.round(area.x)));
  off.setAttribute('y', String(Math.round(area.y)));
  ext.setAttribute('cx', String(width));
  ext.setAttribute('cy', String(heights.reduce((sum, value) => sum + value, 0)));

  // A cell keeps 0.1 inch free at its sides and 0.05 above and below unless it says otherwise.
  const side = Math.min(91440, Math.round(columnWidth * 0.06));
  const above = Math.min(45720, Math.round(rowHeight * 0.1));
  const tight = side < 91440 || above < 45720;
  const doc = tbl.ownerDocument;
  for (const cell of descendants(tbl, NS.a, 'tc')) {
    let props = childElement(cell, NS.a, 'tcPr');
    if (!props && doc) {
      props = doc.createElementNS(NS.a, `${cell.prefix ? `${cell.prefix}:` : ''}tcPr`);
      cell.appendChild(props);
    }
    if (!props) continue;
    props.setAttribute('anchor', 'ctr');
    if (!tight || options.keepMargins) continue;
    props.setAttribute('marL', String(side));
    props.setAttribute('marR', String(side));
    props.setAttribute('marT', String(above));
    props.setAttribute('marB', String(above));
  }
}

/**
 * The text size in points at which one line of text fits a row of a table that shares `height`
 * (EMU) among `rows` rows, and no larger than `most`. Half points, never under 4.
 */
export function tableTextSize(height: number, rows: number, most: number): number {
  const rowHeight = height / Math.max(1, rows);
  const margins = 2 * Math.min(45720, Math.round(rowHeight * 0.1));
  const fits = Math.floor(((rowHeight - margins) / 12700 / 1.25) * 2) / 2;
  return Math.max(4, Math.min(most, fits));
}

/** What a table that the software draws itself looks like: colours as six hex digits. */
export interface TableLook {
  /** Fill of the heading row. */
  header: string;
  /** Fill of every second row. */
  band: string;
  /** The lines between rows. */
  line: string;
}

/**
 * A plain table for a template that has none to copy: a filled heading row, then rows with a
 * line under each and every second one lightly filled. It is a sample for {@link setTable},
 * which gives it the rows, columns and text of a deck; the text takes its look from the slot.
 */
export function tableFrameXml(shapeId: string, area: Area, look: TableLook): string {
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const hex = (value: string) => escapeXmlAttr(value.replace(/^#/, '').toUpperCase());
  const none = (side: string) => `<a:${side} w="0"><a:noFill/></a:${side}>`;
  const under = `<a:lnB w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${hex(look.line)}"/></a:solidFill><a:prstDash val="solid"/></a:lnB>`;
  const cell = (fill: string) =>
    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>` +
    `<a:tcPr anchor="ctr">${none('lnL')}${none('lnR')}${none('lnT')}${under}${fill ? `<a:solidFill><a:srgbClr val="${hex(fill)}"/></a:solidFill>` : '<a:noFill/>'}</a:tcPr></a:tc>`;
  const half = Math.floor(Math.round(area.cx) / 2);
  const rowHeight = Math.floor(Math.round(area.cy) / 3);
  const row = (fill: string) => `<a:tr h="${rowHeight}">${cell(fill)}${cell(fill)}</a:tr>`;
  return (
    `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${A}">` +
    `<p:nvGraphicFramePr><p:cNvPr id="${escapeXmlAttr(shapeId)}" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(area.x)}" y="${Math.round(area.y)}"/><a:ext cx="${Math.round(area.cx)}" cy="${rowHeight * 3}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>` +
    `<a:tblGrid><a:gridCol w="${half}"/><a:gridCol w="${Math.round(area.cx) - half}"/></a:tblGrid>` +
    `${row(look.header)}${row('')}${row(look.band)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}
