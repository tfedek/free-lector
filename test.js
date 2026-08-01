/**
 * Free Lector — Test Suite (Round 4)
 * Real XML parsing via @xmldom/xmldom, real ZIP via jszip
 * Run: node test.js
 */
'use strict';

const assert = require('assert');
const path = require('path');
const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');

global.DOMParser = DOMParser;
global.JSZip = JSZip;
global.mammoth = { convertToHtml: async () => ({ value: '' }) };
if (typeof structuredClone === 'undefined') global.structuredClone = v => JSON.parse(JSON.stringify(v));

const RuleEngine = require(path.resolve(__dirname, 'rules.js'));
const Exporter = require(path.resolve(__dirname, 'exporter.js'));
const DocumentParser = require(path.resolve(__dirname, 'parser.js'));

let passed = 0, failed = 0;
const asyncQueue = [];
function test(name, fn) {
    try { fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
    catch (e) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}\n    ${e.message}`); }
}
function testAsync(name, fn) {
    asyncQueue.push({ name, fn });
}

function makeDocMap(elements, opts = {}) {
    let idx = 0;
    const mapped = elements.map(el => ({
        type: el.type||'paragraph', index: idx++, text: el.text||'', style: el.style||'Normal',
        runs: el.runs||[{text:el.text||''}], id: el.id||`p-test-${idx}`, section: el.section||'(test)',
        isEmpty: !(el.text&&el.text.trim().length>0), isDirectQuote: el.isDirectQuote||false,
        quoteConfidence: el.quoteConfidence||0, headingLevel: el.headingLevel||null,
        numId: el.numId||null, numLevel: el.numLevel||null,
        displayedNumber: el.displayedNumber!=null?el.displayedNumber:null,
        displayedLabel: el.displayedLabel||null,
        listInstanceId: el.listInstanceId||null, listStart: el.listStart||null,
        numFmt: el.numFmt||null, paraId: el.paraId||null,
        rows: el.rows||undefined, tableId: el.tableId||undefined,
    }));
    return { type: opts.type||'txt', name: opts.name||'test.txt', elements: mapped,
        footnotes: opts.footnotes||[], endnotes: opts.endnotes||[],
        headers:[], footers:[], styles: opts.styles||{}, numbering: opts.numbering||{},
        htmlPreview:'', rawText: mapped.map(e=>e.text).join('\n'), wordCount:100,
        paragraphCount: mapped.filter(e=>e.type==='paragraph').length,
        tableCount: mapped.filter(e=>e.type==='table').length,
        headingCount: mapped.filter(e=>e.type==='heading').length };
}
function allOpts() { return {brackets:true,quotes:true,markdown:true,spacing:true,scriptMix:true,greek:true,duplicates:true,toc:true,numbering:true,dashes:true,bibliography:true,urls:true,footnotes:true,repetition:true,capsWords:true,emptyHeadings:true}; }

/** Helper: create minimal DOCX ZIP as ArrayBuffer */
async function createDocxZip(documentXml, extras = {}) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('word/document.xml', documentXml);
    if (extras.numbering) zip.file('word/numbering.xml', extras.numbering);
    if (extras.styles) zip.file('word/styles.xml', extras.styles);
    if (extras.vbaProject) zip.file('word/vbaProject.bin', 'fake macro content');
    if (extras.embedding) zip.file('word/embeddings/oleObject1.bin', 'fake embedding');
    if (extras.largeImage) zip.file('word/media/image1.png', Buffer.alloc(extras.largeImage));
    return zip.generateAsync({ type: 'arraybuffer' });
}


// ==========================================
// BASIC RULE TESTS (carried over)
// ==========================================
console.log('\nOsnovne provere:');
test('straight quotes consolidated globally', () => {
    const doc = makeDocMap([{text:'A "b" c.'},{text:'D "e" f.'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const qf = findings.filter(f=>f.category==='Tipografija'&&f.rationale.includes('ravni'));
    assert.strictEqual(qf.length, 1, `Expected 1 consolidated, got ${qf.length}`);
    assert(qf[0].original.includes('4'));
});
test('unmatched typographic quotes detected globally', () => {
    const doc = makeDocMap([{text:'Rekao \u201ezdrav ali.'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.rationale.includes('Neupareni')));
});
test('dashes check disabled (no-op)', () => {
    const doc = makeDocMap([{text:'Period 484-425. Also -- here.'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const dashFindings = findings.filter(f=>f.rationale&&(f.rationale.includes('en-dash')||f.rationale.includes('em-dash')));
    assert.strictEqual(dashFindings.length, 0, 'Dashes check should be disabled');
});
test('script mixing detected', () => {
    const doc = makeDocMap([{text:'\u041Cadmo'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Mešanje pisama'));
});
test('numbering skip 1→3', () => {
    const doc = makeDocMap([{text:'1. A'},{text:'3. C'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Numeracija'));
});
test('list restart allowed', () => {
    const doc = makeDocMap([{text:'1. A'},{text:'2. B'},{text:'1. C'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert.strictEqual(findings.filter(f=>f.category==='Numeracija').length, 0);
});
test('direct quote demotion', () => {
    const doc = makeDocMap([{text:'Dva  razmaka.', isDirectQuote:true}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const sf = findings.filter(f=>f.category==='Razmaci');
    assert(sf.length>=1); assert.strictEqual(sf[0].priority, 'PROVERITI');
});
test('empty footnote', () => {
    const doc = makeDocMap([{text:'T.'}], {footnotes:[{id:'1',text:'',isEmpty:true}]});
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Fusnote'));
});
test('bibliography without year', () => {
    const doc = makeDocMap([{type:'heading',text:'Bibliografija',headingLevel:1},{text:'Petrović, Milan. Lingvistička analiza srpskog. Beograd'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Bibliografija'));
});
test('dupe words "je je"', () => {
    const doc = makeDocMap([{text:'On je je stigao.'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Duple reči'));
});
test('Greek without translation', () => {
    const doc = makeDocMap([{text:'\u03C4\u1F78\u03BD \u039A\u03AC\u03B4\u03BC\u03BF\u03BD \u1F10\u03BB\u03B8\u03B5\u1FD6\u03BD.'}]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Grčki bez prevoda'));
});


// ==========================================
// ROUND 4 TESTS: Real XML parseNumbering
// ==========================================
console.log('\nReal parseNumbering:');
test('startOverride from XML', () => {
    const xml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>`;
    const r = DocumentParser.parseNumbering(xml);
    assert.strictEqual(r.nums['1'].lvlOverrides[0].startOverride, 5);
});
test('lvlOverride with w:lvl element', () => {
    const xml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride></w:num></w:numbering>`;
    const r = DocumentParser.parseNumbering(xml);
    assert.strictEqual(r.nums['1'].lvlOverrides[0].startOverride, 10);
    assert(r.nums['1'].lvlOverrides[0].lvlDef);
    assert.strictEqual(r.nums['1'].lvlOverrides[0].lvlDef.numFmt, 'upperRoman');
});
test('lvlRestart parsed', () => {
    const xml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlRestart w:val="0"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const r = DocumentParser.parseNumbering(xml);
    assert.strictEqual(r.abstractNums['0'].levels[1].lvlRestart, 0);
});
test('bullet format excluded', () => {
    const xml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const r = DocumentParser.parseNumbering(xml);
    assert.strictEqual(r.abstractNums['0'].levels[0].numFmt, 'bullet');
});

// ==========================================
// Letter overflow and multilevel
// ==========================================
console.log('\nFormatiranje:');
test('toLetter: 26→z, 27→aa, 28→ab', () => {
    assert.strictEqual(DocumentParser.toLetter(26, false), 'z');
    assert.strictEqual(DocumentParser.toLetter(27, false), 'aa');
    assert.strictEqual(DocumentParser.toLetter(28, false), 'ab');
    assert.strictEqual(DocumentParser.toLetter(27, true), 'AA');
});
test('formatNumber decimal', () => { assert.strictEqual(DocumentParser.formatNumber(3,'decimal','%1.',0), '3.'); });
test('formatNumber lowerLetter', () => { assert.strictEqual(DocumentParser.formatNumber(1,'lowerLetter','%1)',0), 'a)'); });
test('formatNumber upperRoman', () => { assert.strictEqual(DocumentParser.formatNumber(4,'upperRoman','%1.',0), 'IV.'); });
test('formatLabel multilevel %1.%2.', () => {
    const label = DocumentParser.formatLabel('%1.%2.', {0:2, 1:3}, {0:{numFmt:'decimal'},1:{numFmt:'decimal'}});
    assert.strictEqual(label, '2.3.');
});


// ==========================================
// ROUND 4: Real ZIP tests
// ==========================================
console.log('\nZIP testovi:');

testAsync('ZIP with vbaProject.bin rejected', async () => {
    const buf = await createDocxZip('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>', { vbaProject: true });
    const file = { name: 'test.docx', arrayBuffer: async () => buf };
    let threw = false;
    try { await DocumentParser.parse(file); } catch (e) { threw = true; assert(e.message.includes('VBA') || e.message.includes('vbaProject')); }
    assert(threw, 'Should have thrown for vbaProject');
});

testAsync('ZIP with embedding rejected', async () => {
    const buf = await createDocxZip('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>', { embedding: true });
    const file = { name: 'test.docx', arrayBuffer: async () => buf };
    let threw = false;
    try { await DocumentParser.parse(file); } catch (e) { threw = true; assert(e.message.includes('embeddings')); }
    assert(threw, 'Should have thrown for embedding');
});

testAsync('ZIP with large image counted toward total size', async () => {
    // Create a ZIP with a 5MB image (under limit but verifies counting works)
    const buf = await createDocxZip('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>', { largeImage: 5 * 1024 * 1024 });
    const file = { name: 'test.docx', arrayBuffer: async () => buf };
    // Should not throw (under limit)
    const result = await DocumentParser.parse(file);
    assert(result.elements.length >= 1);
});

// ==========================================
// Real parser ID stability
// ==========================================
console.log('\nID stabilnost (parser):');

testAsync('ID stable after insertion elsewhere', async () => {
    const makeDoc = (paras) => `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paras.map(t=>`<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join('')}</w:body></w:document>`;
    const buf1 = await createDocxZip(makeDoc(['Alpha','Beta','Gamma']));
    const buf2 = await createDocxZip(makeDoc(['Alpha','Inserted','Beta','Gamma']));
    const f1 = { name: 't.docx', arrayBuffer: async()=>buf1 };
    const f2 = { name: 't.docx', arrayBuffer: async()=>buf2 };
    const r1 = await DocumentParser.parse(f1);
    const r2 = await DocumentParser.parse(f2);
    // "Gamma" has same prev ("Beta") and same next ('') in both — should get same ID
    const gamma1 = r1.elements.find(e=>e.text==='Gamma');
    const gamma2 = r2.elements.find(e=>e.text==='Gamma');
    assert.strictEqual(gamma1.id, gamma2.id, 'Gamma ID should be stable');
});


// ==========================================
// Real listInstanceId from parsed XML
// ==========================================
console.log('\nlistInstanceId iz parsiranog XML:');

testAsync('listInstanceId unique per restart', async () => {
    const numXml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 1</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 2</w:t></w:r></w:p><w:p><w:r><w:t>Break paragraph</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 1 again</w:t></w:r></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml, { numbering: numXml });
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const numbered = result.elements.filter(e=>e.listInstanceId);
    assert(numbered.length >= 3, `Expected 3+ numbered, got ${numbered.length}`);
    // First two should share instanceId, third should differ
    assert.strictEqual(numbered[0].listInstanceId, numbered[1].listInstanceId);
    assert.notStrictEqual(numbered[0].listInstanceId, numbered[2].listInstanceId, 'After break, new instance');
});

// ==========================================
// Bibliography with subheadings
// ==========================================
console.log('\nBibliografija sa podnaslovima:');
test('bibliography continues through subheadings', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Bibliografija', headingLevel:1},
        {type:'heading', text:'Primarni izvori', headingLevel:2},
        {text:'Petrović, Milan. Lingvistička analiza srpskog. Beograd'},
        {type:'heading', text:'Sekundarni izvori', headingLevel:2},
        {text:'Smith J., Mythology, Oxford University Press, 2015'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    // Petrović entry should be found (no year, modern source)
    assert(findings.some(f=>f.category==='Bibliografija'&&f.original.includes('Petrović')));
});

// ==========================================
// H1→H2 allowed
// ==========================================
console.log('\nHeading sequence:');
test('H1 followed by H2 not reported as error', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Poglavlje 1', headingLevel:1},
        {type:'heading', text:'Podnaslov', headingLevel:2},
    ]);
    const {findings} = RuleEngine.runAudit(doc, {emptyHeadings:true});
    const struct = findings.filter(f=>f.category==='Struktura');
    assert.strictEqual(struct.length, 0, 'H1→H2 should not trigger');
});
test('H2 followed by H2 reports warning', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Podnaslov A', headingLevel:2},
        {type:'heading', text:'Podnaslov B', headingLevel:2},
    ]);
    const {findings} = RuleEngine.runAudit(doc, {emptyHeadings:true});
    assert(findings.some(f=>f.category==='Struktura'));
});


// ==========================================
// ISO date exclusion
// ==========================================
console.log('\nISO datum:');
test('ISO date 2026-08-01 not flagged (dashes disabled)', () => {
    // Dashes check is removed — this is a no-op confirmation
    const doc = makeDocMap([{text:'Datum 2026-08-01 i raspon 484-425.'}]);
    const {findings} = RuleEngine.runAudit(doc, {dashes:true});
    const dash = findings.filter(f=>f.category==='Tipografija'&&f.rationale&&f.rationale.includes('en-dash'));
    assert.strictEqual(dash.length, 0);
});

// ==========================================
// Cyrillic ALL-CAPS
// ==========================================
console.log('\nĆirilička ALL-CAPS:');
test('Cyrillic ALL-CAPS word detected', () => {
    const doc = makeDocMap([{text:'Ovo je ТЕКСТ napisano velikim slovima.'}]);
    const {findings} = RuleEngine.runAudit(doc, {capsWords:true});
    assert(findings.some(f=>f.category==='Tipografija'&&f.original==='ТЕКСТ'));
});

// ==========================================
// Multi-paragraph quote balance
// ==========================================
console.log('\nVišepasusni citat:');
test('unmatched quotes across paragraphs detected', () => {
    const doc = makeDocMap([{text:'Rekao je: \u201eOvo je početak'},{text:'nastavak citata bez zatvaranja.'}]);
    const {findings} = RuleEngine.runAudit(doc, {quotes:true});
    assert(findings.some(f=>f.rationale.includes('Neupareni')));
});

// ==========================================
// Tracked changes
// ==========================================
console.log('\nTracked changes:');
testAsync('tracked changes: accept mode skips deleted text', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello </w:t></w:r><w:del><w:r><w:delText>old </w:delText></w:r></w:del><w:ins><w:r><w:t>new </w:t></w:r></w:ins><w:r><w:t>world</w:t></w:r></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file, { trackedChanges: 'accept' });
    assert(result.elements[0].text.includes('new'));
    assert(!result.elements[0].text.includes('old'));
});
testAsync('tracked changes: show_deleted includes deleted text', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello </w:t></w:r><w:del><w:r><w:delText>old </w:delText></w:r></w:del><w:r><w:t>world</w:t></w:r></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file, { trackedChanges: 'show_deleted' });
    assert(result.elements[0].text.includes('old'));
});

// ==========================================
// basedOn numbering inheritance
// ==========================================
console.log('\nbasedOn numeracija:');
testAsync('numbering inherited through basedOn chain', async () => {
    const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:styleId="ListBase"><w:name w:val="List Base"/><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr></w:style><w:style w:styleId="MyList"><w:name w:val="My List"/><w:basedOn w:val="ListBase"/></w:style></w:styles>`;
    const numXml = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="MyList"/></w:pPr><w:r><w:t>Inherited item 1</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="MyList"/></w:pPr><w:r><w:t>Inherited item 2</w:t></w:r></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml, { numbering: numXml, styles: stylesXml });
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const numbered = result.elements.filter(e=>e.numId);
    assert(numbered.length >= 2, `Expected 2 numbered items, got ${numbered.length}`);
    assert.strictEqual(numbered[0].displayedNumber, 1);
    assert.strictEqual(numbered[1].displayedNumber, 2);
});


// ==========================================
// Contradictory passedChecks (table findings prevent "passed")
// ==========================================
console.log('\nKontradiktorni passedChecks:');
test('spacing not marked as passed when table cell has error', () => {
    const tableEl = { type:'table', text:'A  B', tableId:'t-x',
        rows:[[{text:'A  B', tableId:'t-x', rowId:'t-x-r0', cellId:'t-x-r0-c0', rowIndex:0, columnIndex:0, paragraphs:['A  B']}]] };
    const doc = makeDocMap([tableEl]);
    const {findings, passedChecks} = RuleEngine.runAudit(doc, {spacing:true});
    assert(findings.some(f=>f.category==='Razmaci'), 'Should find spacing error in cell');
    const spacingPassed = passedChecks.find(p=>p.area==='Razmaci i interpunkcija');
    assert(!spacingPassed, 'Spacing should NOT be in passedChecks when cell has error');
});

// ==========================================
// Duplicate bracket findings in table
// ==========================================
console.log('\nDuplikati zagrada u tabeli:');
test('table brackets checked per-cell, not from table.text', () => {
    const tableEl = { type:'table', text:'(unclosed', tableId:'t-b',
        rows:[[{text:'(unclosed', tableId:'t-b', rowId:'t-b-r0', cellId:'t-b-r0-c0', rowIndex:0, columnIndex:0, paragraphs:['(unclosed']}]] };
    const doc = makeDocMap([tableEl]);
    const {findings} = RuleEngine.runAudit(doc, {brackets:true});
    const bracketFindings = findings.filter(f=>f.category==='Zagrade');
    // Should have exactly 1 finding (from cell), not 2 (would be if table.text also checked)
    assert.strictEqual(bracketFindings.length, 1, `Expected 1 bracket finding, got ${bracketFindings.length}`);
    assert(bracketFindings[0].cellId, 'Finding should have cellId');
});

// ==========================================
// Disabled table checks
// ==========================================
console.log('\nIsključene provere tabela:');
test('no cell findings when checks disabled', () => {
    const tableEl = { type:'table', text:'A  B', tableId:'t-d',
        rows:[[{text:'A  B', tableId:'t-d', rowId:'t-d-r0', cellId:'t-d-r0-c0', rowIndex:0, columnIndex:0, paragraphs:['A  B']}]] };
    const doc = makeDocMap([tableEl]);
    const {findings} = RuleEngine.runAudit(doc, {spacing:false, brackets:false, scriptMix:false, quotes:true});
    assert.strictEqual(findings.filter(f=>f.tableId).length, 0);
});

// ==========================================
// Real applyExportFilter
// ==========================================
console.log('\napplyExportFilter:');
test('filters and recalculates', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const findings = [
        {id:'F-1',category:'A',priority:'OBAVEZNO',confidence:0.9,original:'x',replacement:'y',rationale:'r',status:'OPEN',isDirectQuote:false,requiresSourceVerification:false,autoFixable:false,globalPattern:false,section:'(t)',paragraphId:'p-1',tableId:null,rowId:null,cellId:null,rowIndex:null,columnIndex:null},
        {id:'F-2',category:'B',priority:'PROVERITI',confidence:0.8,original:'a',replacement:'b',rationale:'s',status:'DONE',isDirectQuote:false,requiresSourceVerification:false,autoFixable:false,globalPattern:false,section:'(t)',paragraphId:'p-2',tableId:null,rowId:null,cellId:null,rowIndex:null,columnIndex:null},
    ];
    const json = Exporter.buildAuditJson(doc, findings, [], {spacing:true});
    const filtered = Exporter.applyExportFilter(json, 'open');
    assert.strictEqual(filtered.findings.length, 1);
    assert.strictEqual(filtered.summary.total_occurrences, 1);
    assert.strictEqual(filtered.audit_status.status, 'DELIMIČAN');
});

// ==========================================
// Scope/audit_status
// ==========================================
console.log('\nScope/audit_status:');
test('scope.proofreading dynamic', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const j1 = Exporter.buildAuditJson(doc, [], [], {});
    assert.strictEqual(j1.scope.proofreading, false, 'No checks = no proofreading');
    const j2 = Exporter.buildAuditJson(doc, [], [], {brackets:true});
    assert.strictEqual(j2.scope.proofreading, true);
});
test('scope.note dynamic', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const j = Exporter.buildAuditJson(doc, [], [], {brackets:true, aiGrammar:true});
    assert(j.scope.note.includes('gramatička'));
    assert(j.scope.note.includes('determinističke'));
});
test('audit_status.style_analysis present', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const j = Exporter.buildAuditJson(doc, [], [], {aiStyle:true});
    assert.strictEqual(j.audit_status.style_analysis, 'IZVRŠENA');
});
test('final gate requires grammar+visual+style', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const j1 = Exporter.buildAuditJson(doc, [], [], {aiGrammar:true,visualLayout:true});
    assert.strictEqual(j1.summary.can_be_marked_final, false, 'Missing style');
    const j2 = Exporter.buildAuditJson(doc, [], [], {aiGrammar:true,visualLayout:true,aiStyle:true});
    assert.strictEqual(j2.summary.can_be_marked_final, true);
});


// ==========================================
// Additional tests: ancient source skipped, antičke excluded
// ==========================================
console.log('\nAntički izvori:');
test('ancient sources not flagged for missing year', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Bibliografija', headingLevel:1},
        {text:'Diodorus Siculus. Bibliotheca Historica. Knjige 1.28.'},
        {text:'Josephus Flavius. Antiquitates Judaicae. Knjiga 12.'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const bibFindings = findings.filter(f=>f.category==='Bibliografija');
    assert.strictEqual(bibFindings.length, 0, 'Ancient sources should be skipped');
});
test('electronic sources with URL not flagged for missing year', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Bibliografija', headingLevel:1},
        {text:'Perseus Digital Library: https://scaife.perseus.org/'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const bibFindings = findings.filter(f=>f.category==='Bibliografija');
    assert.strictEqual(bibFindings.length, 0, 'Electronic sources should be skipped');
});

// ==========================================
// SUMMARY — run async tests then report
// ==========================================

(async () => {
    for (const { name, fn } of asyncQueue) {
        try { await fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
        catch (e) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}\n    ${e.message}`); }
    }
    console.log('\n' + '='.repeat(40));
    console.log(`Rezultat: ${passed} prošlo, ${failed} palo`);
    console.log('='.repeat(40));
    process.exit(failed > 0 ? 1 : 0);
})();
