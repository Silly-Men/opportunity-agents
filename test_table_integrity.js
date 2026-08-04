// Table-integrity smoke test (cleanup batch, 3 Aug).
// Run: node test_table_integrity.js
// Proves: authored pipe tables emit as REAL <w:tbl> elements, and the guard
// fails loudly when the output disagrees with the authored count.

const { Document, Packer, Paragraph } = require('docx');
const { parseBlocks, bodyWithTables, assertTableIntegrity } = require('./server.js');

const SAMPLE = [
  'Runtime note. Values here are the source of derived fields.',
  'Field | Value |',
  'Business name | TallowGlow | [STATED]',
  'Price point | $39.99 per kit | [STATED]',
  'Category | Cosmetics | [STATED]',
  '',
  'Prose between tables stays prose.',
  '',
  '| Month | Spend |',
  '| --- | --- |',
  '| M1 | $924 |',
  '| M2 | $635 |',
].join('\n');

(async () => {
  const blocks = parseBlocks(SAMPLE);
  const tables = blocks.filter(b => b.type === 'table');
  if (tables.length !== 2) throw new Error(`expected 2 parsed tables, got ${tables.length}`);
  if (tables[1].rows.length !== 3) throw new Error('separator row must not become a data row');

  const counter = { tables: 0 };
  const children = bodyWithTables(SAMPLE, counter);
  if (counter.tables !== 2) throw new Error(`expected 2 authored tables, counted ${counter.tables}`);

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  const emitted = await assertTableIntegrity(buffer, counter.tables, 'smoke');
  console.log(`PASS: ${emitted} real <w:tbl> element(s) in the output, matching authored count`);

  // The negative case: a claimed table the output does not carry must FAIL.
  const flat = new Document({ sections: [{ children: [new Paragraph('no tables here')] }] });
  const flatBuffer = await Packer.toBuffer(flat);
  try {
    await assertTableIntegrity(flatBuffer, 1, 'smoke-negative');
    throw new Error('the guard passed a flattened document — it must not');
  } catch (err) {
    if (!String(err.message).includes('table integrity check failed')) throw err;
    console.log('PASS: the guard refuses a flattened-table document, naming the mismatch');
  }
})().catch(err => { console.error('FAIL:', err.message); process.exit(1); });
