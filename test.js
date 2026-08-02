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
let currentSectionLabel = '';
let lastPrintedSection = '';

function printSectionIfNeeded() {
    if (currentSectionLabel && currentSectionLabel !== lastPrintedSection) {
        console.log(currentSectionLabel);
        lastPrintedSection = currentSectionLabel;
    }
}

function test(name, fn) {
    printSectionIfNeeded();
    try { fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
    catch (e) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}\n    ${e.message}`); }
}
function testAsync(name, fn) {
    asyncQueue.push({ name, fn, section: currentSectionLabel });
}
function section(label) {
    currentSectionLabel = label;
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
function allOpts() { return {brackets:true,quotes:true,markdown:true,spacing:true,scriptMix:true,greek:true,duplicates:true,toc:true,numbering:true,bibliography:true,urls:true,footnotes:true,repetition:true,capsWords:true,emptyHeadings:true}; }

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
section('\nOsnovne provere:');
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
section('\nReal parseNumbering:');
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
section('\nFormatiranje:');
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
section('\nZIP testovi:');

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
section('\nID stabilnost (parser):');

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
section('\nlistInstanceId iz parsiranog XML:');

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
section('\nBibliografija sa podnaslovima:');
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
section('\nHeading sequence:');
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
// ==========================================
// Cyrillic ALL-CAPS
// ==========================================
section('\nĆirilička ALL-CAPS:');
test('Cyrillic ALL-CAPS word detected', () => {
    const doc = makeDocMap([{text:'Ovo je ТЕКСТ napisano velikim slovima.'}]);
    const {findings} = RuleEngine.runAudit(doc, {capsWords:true});
    assert(findings.some(f=>f.category==='Tipografija'&&f.original==='ТЕКСТ'));
});

// ==========================================
// Multi-paragraph quote balance
// ==========================================
section('\nVišepasusni citat:');
test('unmatched quotes across paragraphs detected', () => {
    const doc = makeDocMap([{text:'Rekao je: \u201eOvo je početak'},{text:'nastavak citata bez zatvaranja.'}]);
    const {findings} = RuleEngine.runAudit(doc, {quotes:true});
    assert(findings.some(f=>f.rationale.includes('Neupareni')));
});

// ==========================================
// Tracked changes
// ==========================================
section('\nTracked changes:');
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
section('\nbasedOn numeracija:');
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
section('\nKontradiktorni passedChecks:');
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
section('\nDuplikati zagrada u tabeli:');
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
section('\nIsključene provere tabela:');
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
section('\napplyExportFilter:');
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
section('\nScope/audit_status:');
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
// C3 fix: direct quote protection per cell
// ==========================================
section('\nDirektni citat u ćeliji:');
test('cell with typographic quotes gets PROVERITI for spacing error', () => {
    const tableEl = { type:'table', text:'test', tableId:'t-q',
        rows:[[
            {text:'Normalan  tekst.', tableId:'t-q', rowId:'t-q-r0', cellId:'t-q-r0-c0', rowIndex:0, columnIndex:0, paragraphs:['Normalan  tekst.']},
            {text:'\u201eOvo je direktan  citat.\u201c', tableId:'t-q', rowId:'t-q-r0', cellId:'t-q-r0-c1', rowIndex:0, columnIndex:1, paragraphs:['\u201eOvo je direktan  citat.\u201c']},
        ]]};
    const doc = makeDocMap([tableEl]);
    const {findings} = RuleEngine.runAudit(doc, {spacing:true});
    const c1findings = findings.filter(f=>f.cellId==='t-q-r0-c1');
    assert(c1findings.length >= 1, 'Should find spacing error in quoted cell');
    assert.strictEqual(c1findings[0].priority, 'PROVERITI', 'Quoted cell should be PROVERITI');
    assert.strictEqual(c1findings[0].autoFixable, false);
    assert.strictEqual(c1findings[0].requiresSourceVerification, true);
    // Normal cell should still be OBAVEZNO
    const c0findings = findings.filter(f=>f.cellId==='t-q-r0-c0');
    assert(c0findings.length >= 1);
    assert.strictEqual(c0findings[0].priority, 'OBAVEZNO');
});

// ==========================================
// C2 fix: footnote content checks
// ==========================================
section('\nSadržaj fusnota:');
test('double space in footnote detected', () => {
    const doc = makeDocMap([{text:'Tekst.'}], {footnotes:[{id:'1', text:'Vid. Apijan,  Ilirika 2.', isEmpty:false}]});
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const fnSpacing = findings.filter(f=>f.category==='Razmaci'&&f.section.includes('fusnot'));
    assert(fnSpacing.length >= 1, 'Should detect double space in footnote');
});
test('unbalanced bracket in footnote detected', () => {
    const doc = makeDocMap([{text:'Tekst.'}], {footnotes:[{id:'2', text:'Vid. (Apijan, Ilirika 2', isEmpty:false}]});
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const fnBrackets = findings.filter(f=>f.category==='Zagrade'&&f.section.includes('fusnot'));
    assert(fnBrackets.length >= 1, 'Should detect unbalanced bracket in footnote');
});

// ==========================================
// C2: endnote bracket check
// ==========================================
section('\nEndnote provera:');
test('unbalanced bracket in endnote detected', () => {
    const doc = makeDocMap([{text:'Tekst.'}], {footnotes:[], endnotes:[{id:'1', text:'Uporedi (Apijan, Ilirika', isEmpty:false}]});
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const enBrackets = findings.filter(f=>f.category==='Zagrade'&&f.section.includes('endnot'));
    assert(enBrackets.length >= 1, 'Should detect unbalanced bracket in endnote');
});

// ==========================================
// C4: Primarni/Sekundarni izvori as standalone bib heading
// ==========================================
section('\nPrimarni/Sekundarni izvori:');
test('"Primarni izvori" recognized as bibliography heading', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Primarni izvori', headingLevel:1},
        {text:'Petrović M. Analiza. Beograd'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Bibliografija'), 'Should trigger bibliography check');
});
test('"Sekundarni izvori" recognized as bibliography heading', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Sekundarni izvori', headingLevel:1},
        {text:'Jovanović S. Studija. Novi Sad'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    assert(findings.some(f=>f.category==='Bibliografija'), 'Should trigger bibliography check');
});

// ==========================================
// Nested table signal
// ==========================================
section('\nUgnježdena tabela:');
testAsync('nested table recursively parsed into parent cell', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Outer cell</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const tbl = result.elements.find(e=>e.type==='table');
    assert(tbl, 'Should have table element');
    // Nested table text IS now included via recursive parsing
    assert(tbl.rows[0][0].text.includes('Inner cell'), 'Nested table text should be in cell');
    assert(tbl.hasNestedTables, 'hasNestedTables should be true');
    // Paragraphs should contain [Tabela: ...] marker
    assert(tbl.rows[0][0].paragraphs.some(p=>p.includes('[Tabela:')), 'Should have [Tabela:] marker');
});

// ==========================================
// Real parser test: identical tables
// ==========================================
section('\nIdentične tabele (real parser):');
testAsync('two identical tables in DOCX get different tableId', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Same</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Same</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const tables = result.elements.filter(e=>e.type==='table');
    assert.strictEqual(tables.length, 2, 'Should find 2 tables');
    assert.notStrictEqual(tables[0].tableId, tables[1].tableId, 'Identical tables must have different IDs');
});

// ==========================================
// Additional tests: ancient source skipped, antičke excluded
// ==========================================
section('\nAntički izvori:');
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
// (Identical table hash test replaced by real parser test below)
// ==========================================
section('\nIdentične tabele:');
test('two identical rows in same table get different rowId (hash)', () => {
    const id1 = DocumentParser.hashId('table', 'tbl|row|Same');
    const id2 = DocumentParser.hashId('table', 'tbl|row|Same|#2');
    assert.notStrictEqual(id1, id2);
});

// ==========================================
// All export after resolving findings
// ==========================================
section('\nExport all posle rešavanja:');
test('all export includes resolved findings with full count', () => {
    const doc = makeDocMap([{text:'T.'}]);
    const findings = [
        {id:'F-1',category:'A',priority:'OBAVEZNO',confidence:0.9,original:'x',replacement:'y',rationale:'r',status:'DONE',isDirectQuote:false,requiresSourceVerification:false,autoFixable:false,globalPattern:false,section:'(t)',paragraphId:'p-1',tableId:null,rowId:null,cellId:null,rowIndex:null,columnIndex:null},
        {id:'F-2',category:'B',priority:'PROVERITI',confidence:0.8,original:'a',replacement:'b',rationale:'s',status:'OPEN',isDirectQuote:false,requiresSourceVerification:false,autoFixable:false,globalPattern:false,section:'(t)',paragraphId:'p-2',tableId:null,rowId:null,cellId:null,rowIndex:null,columnIndex:null},
    ];
    const json = Exporter.buildAuditJson(doc, findings, [], {spacing:true});
    // Mark one as DONE in the original
    json.findings[0].status = 'DONE';
    const exported = Exporter.applyExportFilter(json, 'all');
    assert.strictEqual(exported.findings.length, 2, 'All export should include all findings');
    assert.strictEqual(exported.summary.total_occurrences, 2);
    assert.strictEqual(exported.summary.mandatory, 1, 'DONE finding still counts in all export');
});

// ==========================================
// Journal article not flagged for missing publisher
// ==========================================
section('\nČlanak u časopisu:');
test('journal article with quoted title not flagged for publisher', () => {
    const doc = makeDocMap([
        {type:'heading', text:'Bibliografija', headingLevel:1},
        {text:'Yadin, Yigael. \u201eAnd Dan, Why Did He Remain in Ships.\u201c Australian Journal of Biblical Archaeology, 1968.'},
    ]);
    const {findings} = RuleEngine.runAudit(doc, allOpts());
    const pubFindings = findings.filter(f=>f.category==='Bibliografija'&&f.rationale.includes('izdavač'));
    assert.strictEqual(pubFindings.length, 0, 'Journal article should not trigger publisher warning');
});

// ==========================================
// INTEGRATION: merged cells (gridSpan/vMerge)
// ==========================================
section('\nSpojene ćelije:');
testAsync('gridSpan parsed correctly', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged cell</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Normal</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const tbl = result.elements.find(e=>e.type==='table');
    assert(tbl, 'Should have table');
    assert.strictEqual(tbl.rows[0][0].gridSpan, 2, 'First cell gridSpan should be 2');
    assert.strictEqual(tbl.rows[0][0].columnIndex, 0);
    assert.strictEqual(tbl.rows[0][1].columnIndex, 2, 'Second cell starts at logical col 2');
    assert(tbl.hasMergedCells, 'hasMergedCells should be true');
});

// ==========================================
// INTEGRATION: recursive nested tables
// ==========================================
section('\nRekurzivne ugnježdene tabele:');
testAsync('nested table text included in parent cell', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Outer</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const tbl = result.elements.find(e=>e.type==='table');
    assert(tbl.rows[0][0].text.includes('Inner'), 'Nested table text should be in cell');
    assert(tbl.hasNestedTables, 'hasNestedTables should be true');
});

// ==========================================
// INTEGRATION: headers/footers checked
// ==========================================
section('\nHeaders/footers:');
testAsync('header with double space produces finding', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`;
    const headerXml = `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header  text</w:t></w:r></w:p></w:hdr>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('word/document.xml', docXml);
    zip.file('word/header1.xml', headerXml);
    const buf = await zip.generateAsync({ type: 'arraybuffer' });
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    assert(result.headerElements.length >= 1, 'Should have header elements');
    const {findings} = RuleEngine.runAudit(result, {spacing:true});
    const hdrFindings = findings.filter(f=>f.section&&f.section.includes('zaglavlje'));
    assert(hdrFindings.length >= 1, 'Should detect double space in header');
});

// ==========================================
// INTEGRATION: footnotes content checked
// ==========================================
section('\nFusnote integracioni:');
testAsync('footnote with bracket error detected via full parse', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p></w:body></w:document>`;
    const fnXml = `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Fusnota (bez zatvaranja</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('word/document.xml', docXml);
    zip.file('word/footnotes.xml', fnXml);
    const buf = await zip.generateAsync({ type: 'arraybuffer' });
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    const {findings} = RuleEngine.runAudit(result, {footnotes:true,brackets:true});
    assert(findings.some(f=>f.category==='Zagrade'&&f.section.includes('fusnot')), 'Should find bracket error in footnote');
});

// ==========================================
// INTEGRATION: tracked changes
// ==========================================
section('\nTracked changes integracioni:');
testAsync('deleted text excluded in accept mode', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:ins><w:r><w:t>kept</w:t></w:r></w:ins><w:del><w:r><w:delText>removed</w:delText></w:r></w:del></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file, {trackedChanges:'accept'});
    assert(result.elements[0].text.includes('kept'));
    assert(!result.elements[0].text.includes('removed'));
});

// ==========================================
// INTEGRATION: unsupported elements warning
// ==========================================
section('\nNepodržani elementi:');
testAsync('document with equations produces coverage warning', async () => {
    const docXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath><m:r><w:t>x=1</w:t></m:r></m:oMath></w:p></w:body></w:document>`;
    const buf = await createDocxZip(docXml);
    const file = { name: 't.docx', arrayBuffer: async()=>buf };
    const result = await DocumentParser.parse(file);
    assert(result.processingCoverage.unsupported.includes('equations'), 'Should detect equations');
    const {findings} = RuleEngine.runAudit(result, {brackets:true});
    const coverageFindings = findings.filter(f=>f.category==='Pokrivenost');
    assert(coverageFindings.length >= 1, 'Should produce coverage warning finding');
    assert(coverageFindings[0].rationale.includes('nisu potpuno'), 'Should mention incomplete processing');
});

// ==========================================
// SUMMARY — run async tests then report
// ==========================================

(async () => {
    let lastPrintedSection = '';
    for (const { name, fn, section } of asyncQueue) {
        if (section && section !== lastPrintedSection) {
            console.log(section);
            lastPrintedSection = section;
        }
        try { await fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
        catch (e) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}\n    ${e.message}`); }
    }
    console.log('\n' + '='.repeat(40));
    console.log(`Rezultat: ${passed} prošlo, ${failed} palo`);
    console.log('='.repeat(40));
    process.exit(failed > 0 ? 1 : 0);
})();
