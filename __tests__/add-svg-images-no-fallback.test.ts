import JSZip from 'jszip';
import fs from 'fs';
import Automizer from '../src/automizer';

const templateDir = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

/**
 * An SVG picture written without a raster fallback is a bare <a:blip> with no
 * r:embed at all; the only relationship is on the <asvg:svgBlip> inside its
 * extLst. ContentTracker used to push a target for the attribute-less blip,
 * which resolved to no relationship and crashed the cleanup collector.
 */
test('copies a slide with a fallback-less svg picture under cleanup', async () => {
  const automizer = new Automizer({
    templateDir,
    outputDir,
    cleanup: true,
  });

  const pres = automizer
    .loadRoot(`SVGImagesNoFallback.pptx`)
    .load(`SVGImagesNoFallback.pptx`, 'svg');

  pres.addSlide('svg', 1);

  const result = await pres.write(`add-svg-images-no-fallback.test.pptx`);

  // The three svg files are imported once each and none is dropped by cleanup.
  expect(result.images).toBe(3);

  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${outputDir}/add-svg-images-no-fallback.test.pptx`),
  );

  const rels = await archive
    .file('ppt/slides/_rels/slide2.xml.rels')
    .async('string');
  const svgTargets = (rels.match(/Target="\.\.\/media\/[^"]+\.svg"/g) || [])
    .length;

  expect(svgTargets).toBe(3);
});
