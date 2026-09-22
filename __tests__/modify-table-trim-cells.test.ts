import Automizer, { modify } from '../src/index';
import { TableData } from '../src/types/table-types';
import { expectXml } from './helpers/expect-xml';
import { DOMParser, Element as XmlElement } from '@xmldom/xmldom';
import ModifyTextHelper from '../src/helper/modify-text-helper';
import { XmlElement as SrcXmlElement } from '../src/types/xml-types';

// setTable must leave every <a:tr> with exactly as many grid positions as
// <a:tblGrid> declares. Surplus template cells are tolerated by PowerPoint,
// but LibreOffice then draws the table at the top edge of the slide.

const directChildren = (element: XmlElement, tag: string): XmlElement[] =>
  Array.from(element.childNodes).filter(
    (node) => (node as XmlElement).tagName === tag,
  ) as XmlElement[];

const readTable = async (outputFile: string, name: string) => {
  const slide = await expectXml(outputFile, 'ppt/slides/slide2.xml');
  const frame = Array.from(
    slide.doc().getElementsByTagName('p:graphicFrame'),
  ).find(
    (frame) =>
      frame.getElementsByTagName('p:cNvPr').item(0).getAttribute('name') ===
      name,
  );
  const tbl = frame.getElementsByTagName('a:tbl').item(0);
  const gridCols = tbl.getElementsByTagName('a:gridCol').length;
  const rows = directChildren(tbl, 'a:tr').map((tr) =>
    directChildren(tr, 'a:tc').map((tc) => ({
      text: Array.from(tc.getElementsByTagName('a:t'))
        .map((t) => t.textContent)
        .join(''),
      gridSpan: tc.getAttribute('gridSpan'),
      hMerge: tc.getAttribute('hMerge'),
    })),
  );
  return { gridCols, rows };
};

const body = (rows: number, cols: number): TableData => ({
  body: Array.from({ length: rows }, (_, r) => ({
    label: `r${r}`,
    values: Array.from({ length: cols }, (_, c) => `r${r}c${c}`),
  })),
});

const write = async (
  file: string,
  slideFile: string,
  shape: string,
  data: TableData,
) => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(slideFile, 'tables')
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement(shape, [modify.setTable(data)]);
    })
    .write(file);
};

test('setTable trims surplus template cells of every row to the grid', async () => {
  const file = 'modify-table-trim-cells-narrow.test.pptx';
  await write(file, 'SlideWithTables.pptx', 'TableDefault', body(3, 2));

  const table = await readTable(file, 'TableDefault');
  expect(table.gridCols).toBe(2);
  expect(table.rows.map((row) => row.length)).toEqual([2, 2, 2]);
  expect(table.rows[2].map((cell) => cell.text)).toEqual(['r2c0', 'r2c1']);
});

test('setTable widens every row to the grid', async () => {
  const file = 'modify-table-trim-cells-wide.test.pptx';
  await write(file, 'SlideWithTables.pptx', 'TableDefault', body(3, 7));

  const table = await readTable(file, 'TableDefault');
  expect(table.gridCols).toBe(7);
  expect(table.rows.map((row) => row.length)).toEqual([7, 7, 7]);
  expect(table.rows[1][6].text).toBe('r1c6');
});

test('setTable pads rows with fewer values than the widest row', async () => {
  const file = 'modify-table-trim-cells-ragged.test.pptx';
  const data: TableData = {
    body: [
      { label: 'r0', values: ['a', 'b', 'c', 'd', 'e', 'f'] },
      { label: 'r1', values: ['g'] },
    ],
  };
  await write(file, 'SlideWithTables.pptx', 'TableDefault', data);

  const table = await readTable(file, 'TableDefault');
  expect(table.gridCols).toBe(6);
  expect(table.rows.map((row) => row.length)).toEqual([6, 6]);
  expect(table.rows[1][0].text).toBe('g');
  // Template cells 1-3 keep their text; cells 4-5 are empty padding.
  expect(table.rows[1][4].text).toBe('');
  expect(table.rows[1][5].text).toBe('');
});

test('setTable clamps gridSpan of merged cells cut by the trim', async () => {
  const file = 'modify-table-trim-cells-merged.test.pptx';
  await write(file, 'NestedTables.pptx', 'NestedTable1', body(3, 3));

  const table = await readTable(file, 'NestedTable1');
  expect(table.gridCols).toBe(3);
  expect(table.rows.map((row) => row.length)).toEqual([3, 3, 3]);
  // B1-D1 spanned 3 columns from position 1; only 2 remain.
  expect(table.rows[0][1].gridSpan).toBe('2');
  expect(table.rows[0][2].hMerge).toBe('1');
  // B2-C2 spans 2 from position 1 and still fits.
  expect(table.rows[1][1].gridSpan).toBe('2');

  const file2 = 'modify-table-trim-cells-merged-2.test.pptx';
  await write(file2, 'NestedTables.pptx', 'NestedTable1', body(3, 2));
  const table2 = await readTable(file2, 'NestedTable1');
  expect(table2.gridCols).toBe(2);
  expect(table2.rows.map((row) => row.length)).toEqual([2, 2, 2]);
  expect(table2.rows[0][1].gridSpan).toBeNull();
  expect(table2.rows[1][1].gridSpan).toBeNull();
});

// null and '' clear a cell; undefined keeps the template text (used e.g. to
// keep labels of merged header cells, see modify-nested-table.test.ts).
test('setTable clears cells for null and keeps them for undefined', async () => {
  const file = 'modify-table-trim-cells-empty.test.pptx';
  const data = {
    body: [
      { label: 'r0', values: [null, undefined, 'x', ''] },
      { label: 'r1', values: ['a', 'b', 'c', 'd'] },
      { label: 'r2', values: ['a', 'b', 'c', 'd'] },
    ],
  } as unknown as TableData;
  await write(file, 'SlideWithTables.pptx', 'TableDefault', data);

  const table = await readTable(file, 'TableDefault');
  expect(table.rows[0].map((cell) => cell.text)).toEqual(['', 'cell', 'x', '']);
});

test('ModifyTextHelper.content fills an <a:t/> without a text node', () => {
  const doc = new DOMParser().parseFromString(
    '<a:r xmlns:a="a"><a:t/></a:r>',
    'text/xml',
  );
  const t = doc.getElementsByTagName('a:t').item(0);
  ModifyTextHelper.content('filled')(t as unknown as SrcXmlElement);
  expect(t.textContent).toBe('filled');
});
