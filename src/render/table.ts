import type { TableValue } from '../deck/spec.ts';
import { NS, childElement, childElements, descendants, removeNode } from '../pptx/xml.ts';
import type { Element } from '../pptx/xml.ts';
import type { Direction, TableSlotDef } from '../template/manifest.ts';
import { setBodyText } from './text.ts';

/**
 * Rebuilds a table from the sample table's own rows and cells, so borders, fills and banding
 * stay as designed while the number of rows and columns follows the content.
 *
 * - The first sample row is the header; the body rows that follow are reused in turn.
 * - Columns share the table's width equally.
 * - The table runs in the deck's reading direction, with text aligned to its start.
 */
export function setTable(frame: Element, table: TableValue, slot: TableSlotDef, direction: Direction): void {
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
      if (body) setBodyText(body, [text], style ? { flipAlign: false, align, style } : { flipAlign: false, align });
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
