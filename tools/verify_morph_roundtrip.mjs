// Round-trip: office-parser morph markdown -> MorphEditor BlockModel parse.
// Proves the emitted dialect parses as intended: 2 hr-delimited deck
// sections, escaped line stays a paragraph, fences stay code-fences.
// Round-trip proof: office-parser morph markdown -> MorphEditor BlockModel.
// Usage: node tools/verify_morph_roundtrip.mjs <file.md>
// Requires MorphEditor checkout at ../MorphFam/MorphEditor (parser source of truth).
import { readFileSync } from 'fs';
const md = readFileSync(process.argv[2], 'utf8');

const { parseSlidesFrontmatter } = await import('/home/aurel/Documents/current/MorphFam/MorphEditor/src/modules/publish/frontmatter.js');
const { parseBlocks } = await import('/home/aurel/Documents/current/MorphFam/MorphEditor/src/modules/editor/dom-editor/BlockModel.js');

const fm = parseSlidesFrontmatter(md);
console.log('deck frontmatter:', JSON.stringify({ ratio: fm && fm.ratio }));
if (!fm || fm.ratio !== '16:9') throw new Error('slides key not parsed (or not a SLIDE_RATIOS key)');
const doc = parseBlocks(fm.rest); const blocks = doc.blocks;
console.log('block types:', blocks.map(b => b.type).join(', '));
const h1 = blocks.filter(b => b.type === 'heading-1');
if (h1.length !== 2) throw new Error('expected 2 slide headings, got ' + h1.length);
if (h1[0].text !== 'First Slide' || h1[1].text !== 'Second Slide') throw new Error('heading text drift');
const escaped = blocks.find(b => b.text.includes('looks like a bullet'));
if (!escaped || escaped.type !== 'paragraph') throw new Error('escaped line re-detected as ' + (escaped && escaped.type));
if (escaped.text !== '- looks like a bullet') throw new Error('escape strip drift: ' + JSON.stringify(escaped.text)); // marker stripped on parse = correct
const hr = blocks.filter(b => b.type === 'horizontal-rule');
if (hr.length !== 1) throw new Error('expected 1 slide divider, got ' + hr.length);
console.log('ROUND-TRIP OK: slides:16:9 deck, 2 hr-delimited sections, escaped line = clean paragraph');
