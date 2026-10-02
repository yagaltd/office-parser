// Golden round-trip corpus: every fixture through MorphEditor's BlockModel.
// Usage: node tools/golden_roundtrip.mjs <deck.md> <marks.md> <table.md> [columns.md] [mermaid.md]
// Requires the MorphEditor checkout at ../MorphFam/MorphEditor.
import { readFileSync } from 'fs';
import { parseBlocks } from '/home/aurel/Documents/current/MorphFam/MorphEditor/src/modules/editor/dom-editor/BlockModel.js';
import { parseSlidesFrontmatter } from '/home/aurel/Documents/current/MorphFam/MorphEditor/src/modules/publish/frontmatter.js';

const [, , deckPath, marksPath, tablePath, columnsPath, mermaidPath] = process.argv;
let failures = 0;
const check = (name, cond, detail) => {
  if (cond) console.log(`PASS ${name}`);
  else { failures++; console.log(`FAIL ${name}: ${detail}`); }
};

// 1. Deck: slides ratio + hr sections + note fence
{
  const md = readFileSync(deckPath, 'utf8');
  const fm = parseSlidesFrontmatter(md);
  check('deck frontmatter', !!fm && fm.ratio === '16:9', JSON.stringify(fm && fm.ratio));
  const blocks = parseBlocks(fm.rest).blocks;
  const h1 = blocks.filter(b => b.type === 'heading-1');
  check('deck sections', h1.length === 2 && blocks.filter(b => b.type === 'horizontal-rule').length === 1, `${h1.length} headings`);
  const note = blocks.find(b => b.type === 'note');
  check('note fence', !!note && note.text.includes('Q4'), JSON.stringify(note && note.text));
  const esc = blocks.find(b => b.type === 'paragraph' && b.text.includes('looks like a bullet'));
  check('sigil escape', !!esc && esc.type === 'paragraph', 're-detected as ' + (esc && esc.type));
}

// 2. Marks: bold/italic/strike/link byte-exact
{
  const blocks = parseBlocks(readFileSync(marksPath, 'utf8')).blocks;
  const p = blocks.find(b => b.type === 'paragraph' && b.text.startsWith('Plain '));
  const m = (p && p.marks) || [];
  check('marks', JSON.stringify(m) === JSON.stringify([
    { type: 'bold', from: 6, to: 10 },
    { type: 'italic', from: 15, to: 21 },
    { type: 'strikethrough', from: 21, to: 26 },
  ]), JSON.stringify(m));
  const lp = blocks.find(b => b.type === 'paragraph' && b.text.startsWith('Visit '));
  const link = ((lp && lp.marks) || []).find(x => x.type === 'link');
  check('link mark', !!link && link.from === 6 && link.to === 13 && link.href === 'https://example.com', JSON.stringify(link));
}

// 3. Table wrapper: widths survive
{
  const doc = parseBlocks(readFileSync(tablePath, 'utf8'));
  const t = doc.blocks.find(b => b.type === 'table');
  check('table widths', !!t && (t.widths || []).join(',') === '3000,1000', JSON.stringify(t && t.widths));
}

// 4. Columns: column-scoped groups survive
if (columnsPath) {
  const doc = parseBlocks(readFileSync(columnsPath, 'utf8'));
  const cols = doc.blocks.filter(b => b.colId);
  check('columns groups', cols.length === 2 && new Set(cols.map(b => b.colId)).size === 2,
    JSON.stringify(cols.map(b => b.colId)));
}

// 5. Mermaid hardening: arrow directions + group flattening
if (mermaidPath) {
  const blocks = parseBlocks(readFileSync(mermaidPath, 'utf8')).blocks;
  const diagram = blocks.find(b => (b.text || '').includes('flowchart'));
  const mm = (diagram && diagram.text) || '';
  check('mermaid both-arrow', mm.includes('n2 <--> n3'), mm.slice(0, 200));
  check('mermaid reverse-swap', mm.includes('n2 --> n3'), mm.slice(0, 200));
  check('mermaid group member', mm.includes('label: \"Grouped A\"'), mm.slice(0, 200));
}

if (failures) { console.log(`${failures} FAILURES`); process.exit(1); }
console.log('GOLDEN ROUND-TRIP OK');
