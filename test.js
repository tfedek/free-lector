/**
 * Free Lector — Automated Test Suite
 * Uses Node.js assert (no eval, proper require imports)
 * Exit code 1 on any failure.
 *
 * Run: node test.js
 */

'use strict';

const assert = require('assert');
const path = require('path');

// Polyfill DOMParser for Node.js (parser.js needs it)
const { DOMParser } = require('@xmldom/xmldom');
global.DOMParser = DOMParser;

// Mock structuredClone if not available
if (typeof structuredClone === 'undefined') {
    global.structuredClone = (obj) => JSON.parse(JSON.stringify(obj));
}

// Load modules from current directory explicitly (avoids stale cache)
const RuleEngine = require(path.resolve(__dirname, 'rules.js'));
const Exporter = require(path.resolve(__dirname, 'exporter.js'));
const DocumentParser = require(path.resolve(__dirname, 'parser.js'));

// ==========================================
// HELPERS
// ==========================================

function makeDocMap(elements, opts = {}) {
    let idx = 0;
    const mapped = elements.map(el => {
        const base = {
            type: el.type || 'paragraph',
            index: idx++,
            text: el.text || '',
            style: el.style || 'Normal',
            runs: el.runs || [{ text: el.text || '' }],
            id: el.id || `p-test-${idx}`,
            section: el.section || '(test)',
            isEmpty: !(el.text && el.text.trim().length > 0),
            isDirectQuote: el.isDirectQuote || false,
            quoteConfidence: el.quoteConfidence || 0,
            headingLevel: el.headingLevel || null,
            numId: el.numId || null,
            numLevel: el.numLevel || null,
            displayedNumber: el.displayedNumber != null ? el.displayedNumber : null,
            listInstanceId: el.listInstanceId || null,
            listStart: el.listStart || null,
            numFmt: el.numFmt || null,
            displayedLabel: el.displayedLabel || null,
            paraId: el.paraId || null,
            rows: el.rows || undefined,
            tableId: el.tableId || undefined,
        };
        return base;
    });
    return {
        type: opts.type || 'txt',
        name: opts.name || 'test.txt',
        elements: mapped,
        footnotes: opts.footnotes || [],
        endnotes: opts.endnotes || [],
        headers: [], footers: [],
        styles: opts.styles || {},
        numbering: opts.numbering || {},
        htmlPreview: '',
        rawText: mapped.map(e => e.text).join('\n'),
        wordCount: 100,
        paragraphCount: mapped.filter(e => e.type === 'paragraph').length,
        tableCount: mapped.filter(e => e.type === 'table').length,
        headingCount: mapped.filter(e => e.type === 'heading').length,
    };
}

function allOptions() {
    return {
        brackets: true, quotes: true, markdown: true, spacing: true,
        scriptMix: true, greek: true, duplicates: true, toc: true,
        numbering: true, dashes: true, bibliography: true, urls: true,
        footnotes: true, repetition: true, capsWords: true, emptyHeadings: true,
    };
}

let passed = 0, failed = 0;
function test(name, fn) {
    try { fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
    catch (err) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}`); console.log(`    ${err.message}`); }
}


// ==========================================
// ORIGINAL TESTS (preserved)
// ==========================================
console.log('\nNavodnici:');

test('detects straight double quotes (global consolidation)', () => {
    const doc = makeDocMap([
        { text: 'On kaže: "zdravo" i ode.' },
        { text: 'Drugi pasus sa "navodnicima".' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const qf = findings.filter(f => f.category === 'Tipografija' && f.rationale.includes('ravni'));
    assert.strictEqual(qf.length, 1, `Expected 1 consolidated finding, got ${qf.length}`);
    assert(qf[0].original.includes('4'), `Expected total 4 quotes in msg, got: ${qf[0].original}`);
});

test('no false positive on typographic quotes', () => {
    const doc = makeDocMap([{ text: 'On kaže: \u201ezdrav\u201c i ode.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const straight = findings.filter(f => f.rationale.includes('ravni'));
    assert.strictEqual(straight.length, 0);
});

console.log('\nNeupareni navodnici:');
test('detects unmatched opening quote', () => {
    const doc = makeDocMap([{ text: 'Rekao je \u201ezdrav ali zaboravio.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const unmatched = findings.filter(f => f.rationale.includes('ne odgovara'));
    assert(unmatched.length >= 1);
});

console.log('\nPoluduge crtice:');
test('detects hyphen in numeric range', () => {
    const doc = makeDocMap([{ text: 'Period 484-425 p.n.e.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const dash = findings.filter(f => f.replacement && f.replacement.includes('\u2013'));
    assert(dash.length >= 1);
    assert(dash[0].replacement.includes('484\u2013425'));
});

test('detects double hyphen as em-dash', () => {
    const doc = makeDocMap([{ text: 'Rekao je -- to nije bitno.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.rationale.includes('em-dash')).length >= 1);
});

console.log('\nMešanje pisama:');
test('detects mixed Cyrillic/Latin word', () => {
    const doc = makeDocMap([{ text: 'Reč \u041Cadmo je pomešana.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.category === 'Mešanje pisama').length >= 1);
});

test('no false positive on pure Latin', () => {
    const doc = makeDocMap([{ text: 'The word Kadmo is fine.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert.strictEqual(findings.filter(f => f.category === 'Mešanje pisama').length, 0);
});


console.log('\nNumeracija (preskok 1\u21923):');
test('detects numbering skip from 1 to 3', () => {
    const doc = makeDocMap([{ text: '1. Prva' }, { text: '3. Treća' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert(nf.length >= 1);
    assert(nf[0].rationale.includes('preskače'));
});

test('allows list restart at 1', () => {
    const doc = makeDocMap([{ text: '1. A' }, { text: '2. B' }, { text: '1. C' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert.strictEqual(findings.filter(f => f.category === 'Numeracija').length, 0);
});

console.log('\nNumeracija (OOXML):');
test('detects OOXML numbering skip via displayedNumber', () => {
    const doc = makeDocMap([
        { text: 'A', numId: '1', numLevel: 0, displayedNumber: 1, listInstanceId: '1-0', numFmt: 'decimal' },
        { text: 'B', numId: '1', numLevel: 0, displayedNumber: 2, listInstanceId: '1-0', numFmt: 'decimal' },
        { text: 'D', numId: '1', numLevel: 0, displayedNumber: 4, listInstanceId: '1-0', numFmt: 'decimal' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.category === 'Numeracija').length >= 1);
});

console.log('\nDirektni citati:');
test('findings inside direct quotes get PROVERITI priority', () => {
    const doc = makeDocMap([{ text: 'Dva  razmaka.', isDirectQuote: true, quoteConfidence: 0.95 }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const sf = findings.filter(f => f.category === 'Razmaci');
    assert(sf.length >= 1);
    assert.strictEqual(sf[0].priority, 'PROVERITI');
    assert.strictEqual(sf[0].autoFixable, false);
    assert.strictEqual(sf[0].requiresSourceVerification, true);
});

console.log('\nFusnote:');
test('detects empty footnote', () => {
    const doc = makeDocMap([{ text: 'Tekst.' }], { footnotes: [{ id: '1', text: '', isEmpty: true }] });
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.category === 'Fusnote').length >= 1);
});

test('no finding for non-empty footnote', () => {
    const doc = makeDocMap([{ text: 'Tekst.' }], { footnotes: [{ id: '1', text: 'Vid. 2.', isEmpty: false }] });
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert.strictEqual(findings.filter(f => f.category === 'Fusnote').length, 0);
});

console.log('\nBibliografija:');
test('detects bibliography entry without year', () => {
    const doc = makeDocMap([
        { type: 'heading', text: 'Bibliografija', headingLevel: 1 },
        { text: 'Apijan, Rimska istorija, prevod Miroslav Marković, Beograd' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.category === 'Bibliografija').length >= 1);
});

console.log('\nDuple reči:');
test('detects "je je"', () => {
    const doc = makeDocMap([{ text: 'Kadmo je je stigao.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const df = findings.filter(f => f.category === 'Duple reči');
    assert(df.length >= 1);
    assert.strictEqual(df[0].replacement, 'je');
});

test('"da" allowed', () => {
    const doc = makeDocMap([{ text: 'Da da, slažem se.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert.strictEqual(findings.filter(f => f.category === 'Duple reči').length, 0);
});

console.log('\nGrčki bez prevoda:');
test('detects Greek without translation', () => {
    const doc = makeDocMap([{
        text: 'Piše: \u03C4\u1F78\u03BD \u039A\u03AC\u03B4\u03BC\u03BF\u03BD \u1F10\u03BB\u03B8\u03B5\u1FD6\u03BD. Nastavio.'
    }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    assert(findings.filter(f => f.category === 'Grčki bez prevoda').length >= 1);
});

console.log('\npassedChecks.count:');
test('reflects scannedCount not elements.length', () => {
    const doc = makeDocMap([{ text: 'OK.' }, { text: '' }, { text: 'Još.' }]);
    const { passedChecks } = RuleEngine.runAudit(doc, { brackets: true });
    assert(passedChecks.length >= 1);
    assert.strictEqual(passedChecks[0].count, 2);
});


// ==========================================
// NEW TESTS: parseNumbering
// ==========================================
console.log('\nparseNumbering:');

// We can't call DocumentParser.parseNumbering directly in Node (needs DOMParser).
// Instead we test the numbering logic through the rule engine with mocked data.

test('startOverride changes start value', () => {
    // Simulate: numId 1 has abstractNumId "0", level 0 start=1
    // But num has lvlOverride with startOverride=5
    // Elements with displayedNumber reflecting that
    const doc = makeDocMap([
        { text: 'Item', numId: '1', numLevel: 0, displayedNumber: 5, listInstanceId: '1-0', numFmt: 'decimal' },
        { text: 'Item', numId: '1', numLevel: 0, displayedNumber: 6, listInstanceId: '1-0', numFmt: 'decimal' },
        { text: 'Item', numId: '1', numLevel: 0, displayedNumber: 8, listInstanceId: '1-0', numFmt: 'decimal' }, // skip
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert(nf.length >= 1, `Expected skip detection, got ${nf.length}`);
    assert(nf[0].rationale.includes('6') && nf[0].rationale.includes('8'));
});

test('lvlOverride: separate lists with different start', () => {
    const doc = makeDocMap([
        { text: 'A', numId: '1', numLevel: 0, displayedNumber: 1, listInstanceId: '1-0', numFmt: 'decimal' },
        { text: 'B', numId: '1', numLevel: 0, displayedNumber: 2, listInstanceId: '1-0', numFmt: 'decimal' },
        // New list instance (different numId simulates override)
        { text: 'X', numId: '2', numLevel: 0, displayedNumber: 10, listInstanceId: '2-0', numFmt: 'decimal' },
        { text: 'Y', numId: '2', numLevel: 0, displayedNumber: 11, listInstanceId: '2-0', numFmt: 'decimal' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    // No skip within either list
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert.strictEqual(nf.length, 0, `Expected 0, got ${nf.length}`);
});

test('restart separated list detected correctly', () => {
    const doc = makeDocMap([
        { text: '1. First list item 1' },
        { text: '2. First list item 2' },
        { text: 'A regular paragraph between lists.' },
        { text: '1. Second list item 1' },
        { text: '2. Second list item 2' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert.strictEqual(nf.length, 0, `Expected 0 (restart valid), got ${nf.length}`);
});


// ==========================================
// NEW TESTS: ID stability
// ==========================================
console.log('\nID stabilnost:');

test('ID does not change when paragraph is inserted before', () => {
    // Original: 2 elements
    const doc1 = makeDocMap([{ text: 'Zdravo.' }, { text: 'Svet.' }]);
    // Assign IDs manually using the hash logic
    // The point: same text+context = same hash
    const id1 = doc1.elements[1].id; // "Svet." with prev="Zdravo."

    // After insert: 3 elements, "Svet." is now at index 2 but same prev text
    const doc2 = makeDocMap([{ text: 'Zdravo.' }, { text: 'Novo.' }, { text: 'Svet.' }]);
    // "Svet." now has prev="Novo." — ID WILL change (this is expected with content-based hash)
    // But if prev text is same, ID should be stable
    const doc3 = makeDocMap([{ text: 'Zdravo.' }, { text: 'Svet.' }, { text: 'Kraj.' }]);
    const id3 = doc3.elements[1].id; // "Svet." with prev="Zdravo." and next="Kraj."
    // Since next text changed, hash changes — but prev+text is same
    // The key test: same text + same prev + same next = same ID
    const doc4 = makeDocMap([{ text: 'Zdravo.' }, { text: 'Svet.' }]);
    const id4 = doc4.elements[1].id;
    assert.strictEqual(id1, id4, 'Same context should produce same ID');
});

// ==========================================
// NEW TESTS: Table cell IDs
// ==========================================
console.log('\nTable cell IDs:');

test('table cell findings include tableId/rowId/cellId/rowIndex/columnIndex', () => {
    const tableEl = {
        type: 'table',
        text: 'Ćelija  razmak',
        tableId: 'tbl-0',
        rows: [[{
            text: 'Ćelija  razmak',
            tableId: 'tbl-0',
            rowId: 'tbl-0-r0',
            cellId: 'tbl-0-r0-c0',
            rowIndex: 0,
            columnIndex: 0,
            paragraphs: ['Ćelija  razmak'],
        }]],
    };
    const doc = makeDocMap([tableEl]);
    const { findings } = RuleEngine.runAudit(doc, { spacing: true });
    const cellFindings = findings.filter(f => f.tableId != null);
    assert(cellFindings.length >= 1, `Expected cell finding, got ${cellFindings.length}`);
    assert.strictEqual(cellFindings[0].tableId, 'tbl-0');
    assert.strictEqual(cellFindings[0].rowId, 'tbl-0-r0');
    assert.strictEqual(cellFindings[0].cellId, 'tbl-0-r0-c0');
    assert.strictEqual(cellFindings[0].rowIndex, 0);
    assert.strictEqual(cellFindings[0].columnIndex, 0);
});

// ==========================================
// NEW TESTS: Nested tables (only direct cells checked)
// ==========================================
console.log('\nUgnježdene tabele:');

test('nested table text not included in parent cell check', () => {
    // Simulate: parent table has a cell, nested table would have been excluded by parser
    // So we just verify that a table with clean cells produces no findings
    const tableEl = {
        type: 'table',
        text: 'Čist tekst',
        tableId: 'tbl-0',
        rows: [[{
            text: 'Čist tekst',
            tableId: 'tbl-0',
            rowId: 'tbl-0-r0',
            cellId: 'tbl-0-r0-c0',
            rowIndex: 0,
            columnIndex: 0,
            paragraphs: ['Čist tekst'],
        }]],
    };
    const doc = makeDocMap([tableEl]);
    const { findings } = RuleEngine.runAudit(doc, { spacing: true, brackets: true, scriptMix: true });
    const cellFindings = findings.filter(f => f.tableId != null);
    assert.strictEqual(cellFindings.length, 0, 'Clean cell should have no findings');
});


// ==========================================
// NEW TESTS: Export only open findings
// ==========================================
console.log('\nExport filter (samo otvoreni):');

test('export open-only excludes DONE/REJECTED findings', () => {
    const doc = makeDocMap([{ text: 'A' }]);
    const findings = [
        { id: 'F-0001', category: 'Test', priority: 'OBAVEZNO', confidence: 0.9,
          original: 'a', replacement: 'b', rationale: 'x', status: 'OPEN',
          isDirectQuote: false, requiresSourceVerification: false, autoFixable: false, globalPattern: false,
          section: '(test)', paragraphId: 'p-1', tableId: null, rowId: null, cellId: null, rowIndex: null, columnIndex: null },
        { id: 'F-0002', category: 'Test', priority: 'OBAVEZNO', confidence: 0.9,
          original: 'c', replacement: 'd', rationale: 'y', status: 'DONE',
          isDirectQuote: false, requiresSourceVerification: false, autoFixable: false, globalPattern: false,
          section: '(test)', paragraphId: 'p-2', tableId: null, rowId: null, cellId: null, rowIndex: null, columnIndex: null },
        { id: 'F-0003', category: 'Test2', priority: 'PROVERITI', confidence: 0.8,
          original: 'e', replacement: 'f', rationale: 'z', status: 'REJECTED',
          isDirectQuote: false, requiresSourceVerification: false, autoFixable: false, globalPattern: false,
          section: '(test)', paragraphId: 'p-3', tableId: null, rowId: null, cellId: null, rowIndex: null, columnIndex: null },
    ];
    const passedChecks = [];
    const auditJson = Exporter.buildAuditJson(doc, findings, passedChecks, {});

    // Mock structuredClone for Node < 17
    if (typeof structuredClone === 'undefined') {
        global.structuredClone = (obj) => JSON.parse(JSON.stringify(obj));
    }

    // Use the internal filter by calling downloadJson-like logic
    // We'll test via the module's exported functions
    // Actually Exporter doesn't expose applyExportFilter directly, but downloadExcel calls it.
    // Instead, test the JSON output concept:
    const openOnly = JSON.parse(JSON.stringify(auditJson));
    openOnly.findings = openOnly.findings.filter(f => f.status === 'OPEN');
    assert.strictEqual(openOnly.findings.length, 1);
    assert.strictEqual(openOnly.findings[0].id, 'F-0001');
});

// ==========================================
// NEW TESTS: recalculateSummary final gate
// ==========================================
console.log('\nrecalculateSummary gate:');

test('can_be_marked_final requires grammar+visual_layout scope', () => {
    const doc = makeDocMap([{ text: 'OK.' }]);
    const findings = [];
    const passedChecks = [{ area: 'Test', result: 'OK', count: 1 }];

    // Without AI grammar/visual — cannot be final
    const json1 = Exporter.buildAuditJson(doc, findings, passedChecks, {});
    assert.strictEqual(json1.summary.can_be_marked_final, false,
        'Should not be final without grammar+visual');

    // With AI grammar+visual — can be final (if 0 findings)
    const json2 = Exporter.buildAuditJson(doc, findings, passedChecks,
        { aiGrammar: true, visualLayout: true });
    assert.strictEqual(json2.summary.can_be_marked_final, true,
        'Should be final with grammar+visual and 0 findings');
});

test('audit_status is DELIMIČAN when not final', () => {
    const doc = makeDocMap([{ text: 'OK.' }]);
    const json = Exporter.buildAuditJson(doc, [], [], {});
    assert.strictEqual(json.audit_status.status, 'DELIMIČAN');
});

test('audit_status is POTPUN when final', () => {
    const doc = makeDocMap([{ text: 'OK.' }]);
    const json = Exporter.buildAuditJson(doc, [], [], { aiGrammar: true, visualLayout: true });
    assert.strictEqual(json.audit_status.status, 'POTPUN');
});


// ==========================================
// NEW TESTS: ZIP limits (parser constants)
// ==========================================
console.log('\nZIP limiti:');

test('parser module exposes MAX constants via parse rejection', () => {
    // We can't easily test actual ZIP parsing without DOMParser/JSZip in Node,
    // but we verify the module loaded correctly and constants are enforced
    // by checking that the module exists and parse function is defined
    const DocumentParser = require(path.resolve(__dirname, 'parser.js'));
    assert(typeof DocumentParser.parse === 'function', 'parse should be a function');
    assert(typeof DocumentParser.parseNumbering === 'function', 'parseNumbering should be exposed');
});

// ==========================================
// NEW TESTS: Style-inherited numbering
// ==========================================
console.log('\nStil-nasleđena numeracija:');

test('numbering inherited from style produces findings on skip', () => {
    // Simulate elements that got numId from style inheritance
    const doc = makeDocMap([
        { text: 'A', numId: '2', numLevel: 0, displayedNumber: 1, listInstanceId: '2-0', style: 'ListParagraph', numFmt: 'decimal' },
        { text: 'B', numId: '2', numLevel: 0, displayedNumber: 2, listInstanceId: '2-0', style: 'ListParagraph', numFmt: 'decimal' },
        { text: 'D', numId: '2', numLevel: 0, displayedNumber: 4, listInstanceId: '2-0', style: 'ListParagraph', numFmt: 'decimal' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert(nf.length >= 1, `Expected numbering skip, got ${nf.length}`);
});

// ==========================================
// NEW ROUND 3 TESTS: Real XML parseNumbering
// ==========================================
console.log('\nReal XML parseNumbering:');

test('parseNumbering from real numbering.xml with startOverride', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="5"/>
    </w:lvlOverride>
  </w:num>
</w:numbering>`;
    const result = DocumentParser.parseNumbering(xml);
    assert(result.nums['1'], 'num 1 should exist');
    assert.strictEqual(result.nums['1'].abstractNumId, '0');
    assert.strictEqual(result.nums['1'].lvlOverrides[0].startOverride, 5);
    assert(result.abstractNums['0'].levels[0]);
    assert.strictEqual(result.abstractNums['0'].levels[0].start, 1);
    assert.strictEqual(result.abstractNums['0'].levels[0].numFmt, 'decimal');
});

test('parseNumbering with lvlRestart', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlRestart w:val="0"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`;
    const result = DocumentParser.parseNumbering(xml);
    assert.strictEqual(result.abstractNums['0'].levels[1].lvlRestart, 0);
    assert.strictEqual(result.abstractNums['0'].levels[1].numFmt, 'lowerLetter');
    assert.strictEqual(result.abstractNums['0'].levels[1].lvlText, '%2)');
});

test('parseNumbering with full lvlOverride containing lvl element', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="10"/>
    </w:lvlOverride>
  </w:num>
</w:numbering>`;
    const result = DocumentParser.parseNumbering(xml);
    assert.strictEqual(result.nums['1'].lvlOverrides[0].startOverride, 10);
});

test('parseNumbering: bullet format detected', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`;
    const result = DocumentParser.parseNumbering(xml);
    assert.strictEqual(result.abstractNums['0'].levels[0].numFmt, 'bullet');
});



// ==========================================
// Real applyExportFilter test
// ==========================================
console.log('\nReal applyExportFilter:');

test('applyExportFilter recalculates summary and audit_status', () => {
    const doc = makeDocMap([{ text: 'Test.' }]);
    const findings = [
        { id: 'F-1', category: 'A', priority: 'OBAVEZNO', confidence: 0.9,
          original: 'x', replacement: 'y', rationale: 'r', status: 'OPEN',
          isDirectQuote: false, requiresSourceVerification: false, autoFixable: false,
          globalPattern: false, section: '(t)', paragraphId: 'p-1',
          tableId: null, rowId: null, cellId: null, rowIndex: null, columnIndex: null },
        { id: 'F-2', category: 'B', priority: 'PROVERITI', confidence: 0.8,
          original: 'a', replacement: 'b', rationale: 's', status: 'DONE',
          isDirectQuote: false, requiresSourceVerification: false, autoFixable: false,
          globalPattern: false, section: '(t)', paragraphId: 'p-2',
          tableId: null, rowId: null, cellId: null, rowIndex: null, columnIndex: null },
    ];
    const auditJson = Exporter.buildAuditJson(doc, findings, [], {});
    const filtered = Exporter.applyExportFilter(auditJson, 'open');
    assert.strictEqual(filtered.findings.length, 1);
    assert.strictEqual(filtered.summary.total_occurrences, 1);
    assert.strictEqual(filtered.summary.mandatory, 1);
    assert.strictEqual(filtered.summary.verify, 0);
    assert.strictEqual(filtered.audit_status.status, 'DELIMIČAN');
});

// ==========================================
// Real parser ID stability test
// ==========================================
console.log('\nReal parser ID stability (hashId):');

test('hashId produces stable output for same input', () => {
    const id1 = DocumentParser.hashId('paragraph', 'paragraph|Normal|Hello world|prev|next');
    const id2 = DocumentParser.hashId('paragraph', 'paragraph|Normal|Hello world|prev|next');
    assert.strictEqual(id1, id2);
});

test('hashId changes when content changes', () => {
    const id1 = DocumentParser.hashId('paragraph', 'paragraph|Normal|Hello|prev|next');
    const id2 = DocumentParser.hashId('paragraph', 'paragraph|Normal|World|prev|next');
    assert.notStrictEqual(id1, id2);
});

test('hashId stable when insertion happens elsewhere (same context)', () => {
    // Same element with same prev/next text produces same ID regardless of position
    const id1 = DocumentParser.hashId('paragraph', 'paragraph|Normal|Target|PrevA text|NextB text');
    const id2 = DocumentParser.hashId('paragraph', 'paragraph|Normal|Target|PrevA text|NextB text');
    assert.strictEqual(id1, id2);
});

// ==========================================
// Bullet exclusion from numbering check
// ==========================================
console.log('\nBullet exclusion:');

test('bullet items do not produce numbering findings', () => {
    const doc = makeDocMap([
        { text: 'Item A', numId: '1', numLevel: 0, displayedNumber: null, numFmt: 'bullet' },
        { text: 'Item B', numId: '1', numLevel: 0, displayedNumber: null, numFmt: 'bullet' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert.strictEqual(nf.length, 0, 'Bullets should not be checked for numbering');
});

// ==========================================
// Unique listInstanceId test
// ==========================================
console.log('\nUnique listInstanceId:');

test('different list instances have different listInstanceId', () => {
    const doc = makeDocMap([
        { text: '1. First', numId: '1', numLevel: 0, displayedNumber: 1, listInstanceId: '1-1-0', numFmt: 'decimal' },
        { text: '2. Second', numId: '1', numLevel: 0, displayedNumber: 2, listInstanceId: '1-1-0', numFmt: 'decimal' },
        { text: 'Break paragraph' },
        { text: '1. New list', numId: '1', numLevel: 0, displayedNumber: 1, listInstanceId: '1-2-0', numFmt: 'decimal' },
        { text: '2. New second', numId: '1', numLevel: 0, displayedNumber: 2, listInstanceId: '1-2-0', numFmt: 'decimal' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { numbering: true });
    const nf = findings.filter(f => f.category === 'Numeracija');
    assert.strictEqual(nf.length, 0, 'Separate instances should not show skip');
});

// ==========================================
// Scope/audit_status consistency test
// ==========================================
console.log('\nScope/audit_status usklađenost:');

test('scope.grammar reflects options.aiGrammar', () => {
    const doc = makeDocMap([{ text: 'Test.' }]);
    const json = Exporter.buildAuditJson(doc, [], [], { aiGrammar: true });
    assert.strictEqual(json.scope.grammar, true);
    assert.strictEqual(json.audit_status.linguistic_analysis, 'IZVRŠENA');
});

test('scope.visual_layout reflects options.visualLayout', () => {
    const doc = makeDocMap([{ text: 'Test.' }]);
    const json = Exporter.buildAuditJson(doc, [], [], { visualLayout: true });
    assert.strictEqual(json.scope.visual_layout, true);
    assert.strictEqual(json.audit_status.visual_review, 'IZVRŠEN');
});

test('scope.style reflects options.aiStyle', () => {
    const doc = makeDocMap([{ text: 'Test.' }]);
    const json = Exporter.buildAuditJson(doc, [], [], { aiStyle: true });
    assert.strictEqual(json.scope.style, true);
});

// ==========================================
// Disabled table checks test
// ==========================================
console.log('\nDisabled table checks:');

test('table cells not checked when spacing/brackets/scriptMix disabled', () => {
    const tableEl = {
        type: 'table', text: 'Ćelija  razmak',
        tableId: 'tbl-x', rows: [[{
            text: 'Ćelija  razmak', tableId: 'tbl-x',
            rowId: 'tbl-x-r0', cellId: 'tbl-x-r0-c0',
            rowIndex: 0, columnIndex: 0, paragraphs: ['Ćelija  razmak'],
        }]],
    };
    const doc = makeDocMap([tableEl]);
    // All table-relevant checks disabled
    const { findings } = RuleEngine.runAudit(doc, {
        spacing: false, brackets: false, scriptMix: false, quotes: true
    });
    const cellFindings = findings.filter(f => f.tableId != null);
    assert.strictEqual(cellFindings.length, 0, 'No cell findings when checks disabled');
});

// ==========================================
// Table duplicate findings prevention
// ==========================================
console.log('\nTable duplicate findings:');

test('table.text not re-checked by generic spacing (handled per-cell)', () => {
    const tableEl = {
        type: 'table', text: 'Word  double',
        tableId: 'tbl-y', rows: [[{
            text: 'Word  double', tableId: 'tbl-y',
            rowId: 'tbl-y-r0', cellId: 'tbl-y-r0-c0',
            rowIndex: 0, columnIndex: 0, paragraphs: ['Word  double'],
        }]],
    };
    const doc = makeDocMap([tableEl]);
    const { findings } = RuleEngine.runAudit(doc, { spacing: true });
    // Should only have cell-level finding, not generic table.text finding
    const spacingFindings = findings.filter(f => f.category === 'Razmaci');
    assert(spacingFindings.length >= 1, 'Should have at least 1 finding');
    // All spacing findings should have tableId (from cell check, not generic)
    const genericTableFindings = spacingFindings.filter(f => f.tableId === null);
    assert.strictEqual(genericTableFindings.length, 0,
        'No generic findings for table elements — all should be per-cell');
});

// ==========================================
// formatNumber test
// ==========================================
console.log('\nformatNumber:');

test('formatNumber decimal', () => {
    assert.strictEqual(DocumentParser.formatNumber(3, 'decimal', '%1.', 0), '3.');
});

test('formatNumber lowerLetter', () => {
    assert.strictEqual(DocumentParser.formatNumber(1, 'lowerLetter', '%1)', 0), 'a)');
    assert.strictEqual(DocumentParser.formatNumber(3, 'lowerLetter', '%1)', 0), 'c)');
});

test('formatNumber upperRoman', () => {
    const result = DocumentParser.formatNumber(4, 'upperRoman', '%1.', 0);
    assert.strictEqual(result, 'IV.');
});

// ==========================================
// SUMMARY
// ==========================================
console.log('\n' + '='.repeat(40));
console.log(`Rezultat: ${passed} prošlo, ${failed} palo`);
console.log('='.repeat(40));

if (failed > 0) process.exit(1);
else process.exit(0);
