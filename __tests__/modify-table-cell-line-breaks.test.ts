import fs from 'fs';
import JSZip from 'jszip';
import {
  DOMParser,
  Document as XmlDocument,
  Element as XmlElement,
  XMLSerializer,
} from '@xmldom/xmldom';
import Automizer, { modify } from '../src/index';
import { ModifyTableParams, TableData } from '../src/types/table-types';
import { expectXml } from './helpers/expect-xml';

// A "\n" (or U+000B) in a table cell value is a soft line break: the cell's
// run is repeated per line, separated by an <a:br/> carrying the run's
// <a:rPr>, all in the same paragraph.

const directChildren = (element: XmlElement, tag?: string): XmlElement[] =>
  Array.from(element.childNodes).filter(
    (node) => node.nodeType === 1 && (!tag || node.nodeName === tag),
  ) as XmlElement[];

const serialize = (element: XmlElement) =>
  new XMLSerializer().serializeToString(element);

const tableRows = (doc: XmlDocument, name: string): XmlElement[][] => {
  const frame = Array.from(doc.getElementsByTagName('p:graphicFrame')).find(
    (frame) =>
      frame.getElementsByTagName('p:cNvPr').item(0).getAttribute('name') ===
      name,
  );
  const tbl = frame.getElementsByTagName('a:tbl').item(0);
  return directChildren(tbl, 'a:tr').map((tr) => directChildren(tr, 'a:tc'));
};

const readTable = async (file: string) => {
  const slide = await expectXml(file, 'ppt/slides/slide2.xml');
  return tableRows(slide.doc(), 'TableDefault');
};

// Children of the (single) paragraph of a cell, e.g. ['a:pPr', 'a:r', ...]
const paragraphOf = (tc: XmlElement) => {
  const paragraphs = tc.getElementsByTagName('a:p');
  expect(paragraphs.length).toBe(1);
  return paragraphs.item(0);
};
const sequence = (tc: XmlElement) =>
  directChildren(paragraphOf(tc)).map((child) => child.nodeName);
const runTexts = (tc: XmlElement) =>
  directChildren(paragraphOf(tc), 'a:r').map(
    (r) => r.getElementsByTagName('a:t').item(0).textContent,
  );

const write = async (
  file: string,
  data: TableData,
  params?: ModifyTableParams,
) => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load('SlideWithTables.pptx', 'tables')
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableDefault', [modify.setTable(data, params)]);
    })
    .write(file);
};

const style = {
  color: { type: 'srgbClr' as const, value: 'CCAA4F' },
  size: 1800,
  isBold: true,
};

test('a line break in a cell value becomes an <a:br/> between two styled runs', async () => {
  const file = 'modify-table-cell-line-breaks.test.pptx';
  await write(file, {
    body: [
      {
        values: ['Stick\n(solid)', 'a\u000Bb', 'plain', 'x'],
        styles: [style],
      },
      { values: ['a\n\nb', '\nx', 'y\n', 'z\r\nw'] },
      { values: ['1', '2', '3', '4'] },
    ],
  });
  const rows = await readTable(file);

  const stick = rows[0][0];
  expect(sequence(stick)).toEqual([
    'a:pPr',
    'a:r',
    'a:br',
    'a:r',
    'a:endParaRPr',
  ]);
  expect(runTexts(stick)).toEqual(['Stick', '(solid)']);

  // Both runs and the break carry the cell's run properties incl. the style
  const rPrs = Array.from(stick.getElementsByTagName('a:rPr'));
  expect(rPrs).toHaveLength(3);
  expect(rPrs[1].parentNode.nodeName).toBe('a:br');
  rPrs.forEach((rPr) => {
    expect(rPr.getAttribute('sz')).toBe('1800');
    expect(rPr.getAttribute('b')).toBe('1');
    expect(rPr.getAttribute('lang')).toBe('de-DE');
    const fill = directChildren(rPr, 'a:solidFill')[0];
    expect(
      fill.getElementsByTagName('a:srgbClr').item(0).getAttribute('val'),
    ).toBe('CCAA4F');
    expect(serialize(rPr)).toBe(serialize(rPrs[0]));
  });

  // U+000B is the same break
  expect(sequence(rows[0][1])).toEqual([
    'a:pPr',
    'a:r',
    'a:br',
    'a:r',
    'a:endParaRPr',
  ]);
  expect(runTexts(rows[0][1])).toEqual(['a', 'b']);

  // Without a break, a cell keeps its single run
  expect(sequence(rows[0][2])).toEqual(['a:pPr', 'a:r', 'a:endParaRPr']);
  expect(runTexts(rows[0][2])).toEqual(['plain']);

  // An empty line is the break alone
  expect(sequence(rows[1][0])).toEqual([
    'a:pPr',
    'a:r',
    'a:br',
    'a:br',
    'a:r',
    'a:endParaRPr',
  ]);
  expect(runTexts(rows[1][0])).toEqual(['a', 'b']);
  expect(sequence(rows[1][1])).toEqual([
    'a:pPr',
    'a:br',
    'a:r',
    'a:endParaRPr',
  ]);
  expect(runTexts(rows[1][1])).toEqual(['x']);
  expect(sequence(rows[1][2])).toEqual([
    'a:pPr',
    'a:r',
    'a:br',
    'a:endParaRPr',
  ]);
  expect(runTexts(rows[1][2])).toEqual(['y']);
  expect(runTexts(rows[1][3])).toEqual(['z', 'w']);
  // No literal break characters left in any <a:t>
  rows.flat().forEach((tc) =>
    Array.from(tc.getElementsByTagName('a:t')).forEach((t) =>
      // eslint-disable-next-line no-control-regex -- U+000B is what we look for
      expect(t.textContent).not.toMatch(/[\r\n\u000B]/),
    ),
  );
});

test('a cell value without a break writes the same XML as before', async () => {
  const file = 'modify-table-cell-line-breaks-plain.test.pptx';
  const values = [
    ['a', 'b', 'c', 'd'],
    ['e', 12, 'g h', 'i'],
    ['j', 'k', 'l', 'm'],
  ];
  await write(file, { body: values.map((row) => ({ values: row })) });
  const rows = await readTable(file);

  // Expected: the template cell with nothing but its text replaced
  const template = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-templates/SlideWithTables.pptx`),
  );
  const templateDoc = new DOMParser().parseFromString(
    await template.file('ppt/slides/slide1.xml').async('text'),
    'application/xml',
  );
  const templateRows = tableRows(templateDoc, 'TableDefault');

  values.forEach((row, r) =>
    row.forEach((value, c) => {
      const expected = templateRows[r][c];
      const texts = expected.getElementsByTagName('a:t');
      expect(texts).toHaveLength(1);
      texts.item(0).firstChild.textContent = String(value);
      expect(serialize(rows[r][c])).toBe(serialize(expected));
    }),
  );
});

test('a cell cloned from a broken neighbour does not inherit its break', async () => {
  const file = 'modify-table-cell-line-breaks-expand.test.pptx';
  // Any `expand` makes missing cells clone their (already written)
  // predecessor; the tag matches nothing, so no rows/columns are added.
  await write(
    file,
    {
      body: [
        { values: ['a', 'b', 'c', 'Stick\n(solid)', 'plain', 'p2'] },
        { values: ['a', 'b', 'c', 'd', 'e', 'f'] },
        { values: ['a', 'b', 'c', 'd', 'e', 'f'] },
      ],
    },
    { expand: [{ mode: 'row', tag: '{{no-such-tag}}', count: 1 }] },
  );
  const rows = await readTable(file);

  expect(rows[0]).toHaveLength(6);
  expect(runTexts(rows[0][3])).toEqual(['Stick', '(solid)']);
  [4, 5].forEach((c) => {
    expect(sequence(rows[0][c])).toEqual(['a:pPr', 'a:r', 'a:endParaRPr']);
    expect(rows[0][c].getElementsByTagName('a:br')).toHaveLength(0);
  });
  expect(runTexts(rows[0][4])).toEqual(['plain']);
  expect(runTexts(rows[0][5])).toEqual(['p2']);
});

test('a padding cell cloned from a broken cell is cleared incl. its breaks', async () => {
  const file = 'modify-table-cell-line-breaks-pad.test.pptx';
  await write(file, {
    body: [
      { values: ['a', 'b', 'c', 'd', 'e', 'f'] },
      { values: ['a', 'b', 'c', 'Stick\n(solid)'] },
    ],
  });
  const rows = await readTable(file);

  expect(rows.map((row) => row.length)).toEqual([6, 6]);
  expect(runTexts(rows[1][3])).toEqual(['Stick', '(solid)']);
  [4, 5].forEach((c) => {
    const pad = rows[1][c];
    expect(pad.getElementsByTagName('a:br')).toHaveLength(0);
    Array.from(pad.getElementsByTagName('a:t')).forEach((t) =>
      expect(t.textContent).toBe(''),
    );
  });
});
