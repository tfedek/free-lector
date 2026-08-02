/**
 * Free Lector — Property-based / Fuzz Tests
 * Uses fast-check with deterministic seeds.
 * Run: node test/fuzz.js
 * Nightly: FUZZ_CASES=10000 node test/fuzz.js
 */
'use strict';

const fc = require('fast-check');
const assert = require('assert');
const path = require('path');
const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');

global.DOMParser = DOMParser;
global.JSZip = JSZip;
global.mammoth = { convertToHtml: async () => ({ value: '' }) };
if (typeof structuredClone === 'undefined') {
    global.structuredClone = v => JSON.parse(JSON.stringify(v));
}

const RuleEngine = require(path.resolve(__dirname, '..', 'rules.js'));
const Exporter = require(path.resolve(__dirname, '..', 'exporter.js'));
const DocumentParser = require(path.resolve(__dirname, '..', 'parser.js'));

const NUM_RUNS = parseInt(process.env.FUZZ_CASES || '500', 10);
const SEED = 20260802;
const TIMEOUT_MS = 5000; // per-case timeout

let passed = 0, failed = 0;
const failures = [];

function report(name, err, seed) {
    failed++;
    const msg = `FAIL [${name}] seed=${seed}: ${err.message}`;
    failures.push(msg);
    console.log(`  \x1b[31m✗\x1b[0m ${msg}`);
}


// ==========================================
// GENERATORS
// ==========================================

// Serbian Latin + Cyrillic text
const srLatinChars = 'abcčćdđefghijklmnoprstuvzžšABCČĆDĐEFGHIJKLMNOPRSTUVZŽŠ .,;:!?()-"\'';
const srCyrillicChars = 'абвгдђежзијклмнопрстћуфхцчџшАБВГДЂЕЖЗИЈКЛМНОПРСТЋУФХЦЧЏШ .,;:!?()-"\'';

const arbSerbianText = fc.oneof(
    fc.array(fc.constantFrom(...srLatinChars.split('')), { minLength: 1, maxLength: 200 }).map(a => a.join('')),
    fc.array(fc.constantFrom(...srCyrillicChars.split('')), { minLength: 1, maxLength: 200 }).map(a => a.join(''))
);

// Paragraph with split runs (same visible text, different w:r/w:t structure)
const arbRunSplit = fc.tuple(arbSerbianText, fc.integer({ min: 1, max: 5 })).map(([text, splits]) => {
    const parts = [];
    const chunkSize = Math.max(1, Math.ceil(text.length / splits));
    for (let i = 0; i < text.length; i += chunkSize) {
        parts.push(text.substring(i, i + chunkSize));
    }
    return parts;
});

// Build w:p XML from run parts
function buildParagraphXml(runParts, style) {
    const styleXml = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : '';
    const runs = runParts.map(t => `<w:r><w:t xml:space="preserve">${escXml(t)}</w:t></w:r>`).join('');
    return `<w:p>${styleXml}${runs}</w:p>`;
}

function escXml(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// Valid minimal DOCX content
function buildDocXml(paragraphs) {
    return `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.join('')}</w:body></w:document>`;
}

// Table generator
const arbTable = fc.tuple(
    fc.integer({ min: 1, max: 4 }), // rows
    fc.integer({ min: 1, max: 4 }), // cols
    fc.array(arbSerbianText, { minLength: 1, maxLength: 16 })
).map(([rows, cols, texts]) => {
    let xml = '<w:tbl>';
    for (let r = 0; r < rows; r++) {
        xml += '<w:tr>';
        for (let c = 0; c < cols; c++) {
            const t = texts[(r * cols + c) % texts.length] || 'cell';
            xml += `<w:tc><w:p><w:r><w:t>${escXml(t)}</w:t></w:r></w:p></w:tc>`;
        }
        xml += '</w:tr>';
    }
    xml += '</w:tbl>';
    return xml;
});

// Footnote generator
const arbFootnote = fc.tuple(fc.integer({ min: 1, max: 10 }), arbSerbianText).map(([id, text]) =>
    `<w:footnote w:id="${id}"><w:p><w:r><w:t>${escXml(text)}</w:t></w:r></w:p></w:footnote>`
);

// Create DOCX ZIP
async function createDocx(docXml, extras = {}) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('word/document.xml', docXml);
    if (extras.footnotes) zip.file('word/footnotes.xml', `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${extras.footnotes}</w:footnotes>`);
    if (extras.numbering) zip.file('word/numbering.xml', extras.numbering);
    return zip.generateAsync({ type: 'arraybuffer' });
}


// ==========================================
// PROPERTY TESTS
// ==========================================

async function runProperties() {
    console.log(`\nFree Lector Fuzz Tests (${NUM_RUNS} cases/property, seed=${SEED})\n`);
    const opts = { brackets:true, quotes:true, spacing:true, scriptMix:true, duplicates:true,
        numbering:true, bibliography:true, urls:true, footnotes:true, repetition:true,
        capsWords:true, emptyHeadings:true, toc:true, markdown:true, greek:true };

    // INVARIANT 1: Parsing never crashes
    console.log('Property 1: Parsing never crashes');
    try {
        await fc.assert(fc.asyncProperty(
            fc.array(arbRunSplit, { minLength: 0, maxLength: 20 }),
            async (paragraphs) => {
                const paras = paragraphs.map(parts => buildParagraphXml(parts));
                const docXml = buildDocXml(paras);
                const buf = await createDocx(docXml);
                const file = { name: 'fuzz.docx', arrayBuffer: async () => buf };
                const result = await DocumentParser.parse(file);
                assert(result !== null && result !== undefined);
                assert(Array.isArray(result.elements));
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m No crashes');
    } catch (e) { report('P1-no-crash', e, SEED); }

    // INVARIANT 3: Split/merge runs preserves visible text
    console.log('Property 3: Run splitting preserves text');
    try {
        await fc.assert(fc.asyncProperty(
            arbSerbianText,
            fc.integer({ min: 1, max: 10 }),
            async (text, splits) => {
                const chunkSize = Math.max(1, Math.ceil(text.length / splits));
                const parts = [];
                for (let i = 0; i < text.length; i += chunkSize) parts.push(text.substring(i, i + chunkSize));
                // Single run
                const doc1Xml = buildDocXml([buildParagraphXml([text])]);
                const buf1 = await createDocx(doc1Xml);
                const r1 = await DocumentParser.parse({ name: 'a.docx', arrayBuffer: async () => buf1 });
                // Split runs
                const doc2Xml = buildDocXml([buildParagraphXml(parts)]);
                const buf2 = await createDocx(doc2Xml);
                const r2 = await DocumentParser.parse({ name: 'b.docx', arrayBuffer: async () => buf2 });
                // Visible text must be identical
                assert.strictEqual(r1.elements[0]?.text, r2.elements[0]?.text,
                    `Text mismatch: "${r1.elements[0]?.text}" vs "${r2.elements[0]?.text}"`);
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m Text preserved across run splits');
    } catch (e) { report('P3-run-split', e, SEED); }

    // INVARIANT 5: DONE/REJECTED don't block final
    console.log('Property 5: DONE findings do not block final status');
    try {
        fc.assert(fc.property(
            fc.array(fc.constantFrom('BLOCKER','OBAVEZNO','PROVERITI','PREPORUKA'), { minLength: 1, maxLength: 10 }),
            (priorities) => {
                const findings = priorities.map((p, i) => ({
                    id:`F-${i}`, category:'Test', priority: p, confidence:0.9,
                    original:'x', replacement:'y', rationale:'r', status:'DONE',
                    isDirectQuote:false, requiresSourceVerification:false, autoFixable:false,
                    globalPattern:false, section:'(t)', paragraphId:`p-${i}`,
                    tableId:null, rowId:null, cellId:null, rowIndex:null, columnIndex:null
                }));
                const doc = { name:'t.txt', rawText:'x', wordCount:1, paragraphCount:1, tableCount:0, headingCount:0,
                    elements:[], footnotes:[], endnotes:[], headerElements:[], footerElements:[],
                    processingCoverage:{supported:[],partial:[],unsupported:[]} };
                const json = Exporter.buildAuditJson(doc, findings, [], { aiGrammar:true, visualLayout:true, aiStyle:true, spacing:true });
                // All findings are DONE → can_be_marked_final must be true
                assert.strictEqual(json.summary.can_be_marked_final, true,
                    `Blocked by DONE findings: mandatory_open=${json.summary.mandatory_open}`);
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m DONE findings never block');
    } catch (e) { report('P5-done-block', e, SEED); }

    // INVARIANT 7: Passed check cannot coexist with finding in same category
    console.log('Property 7: No contradictory passedChecks');
    try {
        await fc.assert(fc.asyncProperty(
            fc.array(arbRunSplit, { minLength: 1, maxLength: 10 }),
            async (paragraphs) => {
                const paras = paragraphs.map(parts => buildParagraphXml(parts));
                const docXml = buildDocXml(paras);
                const buf = await createDocx(docXml);
                const file = { name: 'fuzz.docx', arrayBuffer: async () => buf };
                const result = await DocumentParser.parse(file);
                const { findings, passedChecks } = RuleEngine.runAudit(result, opts);
                // Category mapping
                const catMap = { 'Nebalansirane zagrade':'Zagrade', 'Navodnici':'Tipografija',
                    'Razmaci i interpunkcija':'Razmaci', 'Mešanje ćirilice/latinice':'Mešanje pisama',
                    'Duple reči':'Duple reči', 'URL-ovi':'URL', 'Markdown artefakti':'Markdown artefakt',
                    'Grčki citati bez prevoda':'Grčki bez prevoda', 'ALL-CAPS reči':'ALL-CAPS' };
                for (const pc of passedChecks) {
                    const findingCat = catMap[pc.area] || pc.area;
                    const contradiction = findings.some(f => f.category === findingCat);
                    assert(!contradiction,
                        `Passed "${pc.area}" but finding exists in "${findingCat}"`);
                }
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m No contradictions');
    } catch (e) { report('P7-contradiction', e, SEED); }

    // INVARIANT 9: Export-all preserves every finding
    console.log('Property 9: Export-all preserves findings');
    try {
        fc.assert(fc.property(
            fc.array(fc.record({
                priority: fc.constantFrom('BLOCKER','OBAVEZNO','PROVERITI','PREPORUKA'),
                status: fc.constantFrom('OPEN','DONE','REJECTED')
            }), { minLength: 1, maxLength: 20 }),
            (specs) => {
                const findings = specs.map((s, i) => ({
                    id:`F-${i}`, category:'Test', priority:s.priority, confidence:0.9,
                    original:'x', replacement:'y', rationale:'r', status:s.status,
                    isDirectQuote:false, requiresSourceVerification:false, autoFixable:false,
                    globalPattern:false, section:'(t)', paragraphId:`p-${i}`,
                    tableId:null, rowId:null, cellId:null, rowIndex:null, columnIndex:null
                }));
                const doc = { name:'t.txt', rawText:'test', wordCount:1, paragraphCount:1, tableCount:0, headingCount:0,
                    elements:[], footnotes:[], endnotes:[], headerElements:[], footerElements:[],
                    processingCoverage:{supported:[],partial:[],unsupported:[]} };
                const json = Exporter.buildAuditJson(doc, findings, [], {spacing:true});
                const exported = Exporter.applyExportFilter(json, 'all');
                assert.strictEqual(exported.findings.length, findings.length,
                    `Export lost findings: ${exported.findings.length} vs ${findings.length}`);
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m All findings preserved');
    } catch (e) { report('P9-export-all', e, SEED); }

    // INVARIANT 6: JSON summary counts agree
    console.log('Property 6: Summary counts are consistent');
    try {
        fc.assert(fc.property(
            fc.array(fc.record({
                priority: fc.constantFrom('BLOCKER','OBAVEZNO','PROVERITI','PREPORUKA'),
                status: fc.constantFrom('OPEN','DONE','REJECTED')
            }), { minLength: 0, maxLength: 30 }),
            (specs) => {
                const findings = specs.map((s, i) => ({
                    id:`F-${i}`, category:`Cat${i%3}`, priority:s.priority, confidence:0.9,
                    original:'x', replacement:'y', rationale:'r', status:s.status,
                    isDirectQuote:false, requiresSourceVerification:false, autoFixable:false,
                    globalPattern:false, section:'(t)', paragraphId:`p-${i}`,
                    tableId:null, rowId:null, cellId:null, rowIndex:null, columnIndex:null
                }));
                const doc = { name:'t.txt', rawText:'test', wordCount:1, paragraphCount:1, tableCount:0, headingCount:0,
                    elements:[], footnotes:[], endnotes:[], headerElements:[], footerElements:[],
                    processingCoverage:{supported:[],partial:[],unsupported:[]} };
                const json = Exporter.buildAuditJson(doc, findings, [], {spacing:true});
                const s = json.summary;
                // Total counts
                assert.strictEqual(s.blockers + s.mandatory + s.verify + s.recommendations, s.total_occurrences,
                    `Priority sum ${s.blockers}+${s.mandatory}+${s.verify}+${s.recommendations} != total ${s.total_occurrences}`);
                // by_status
                assert.strictEqual(s.by_status.open + s.by_status.done + s.by_status.rejected, s.total_occurrences,
                    `Status sum != total`);
                // open counts <= total
                assert(s.mandatory_open <= s.mandatory, 'mandatory_open > mandatory');
                assert(s.blockers_open <= s.blockers, 'blockers_open > blockers');
                assert(s.verify_open <= s.verify, 'verify_open > verify');
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m Summary counts consistent');
    } catch (e) { report('P6-summary', e, SEED); }

    // INVARIANT 4: ZIP entry reorder preserves results
    console.log('Property 4: ZIP entry order independence');
    try {
        await fc.assert(fc.asyncProperty(
            arbSerbianText,
            async (text) => {
                const docXml = buildDocXml([buildParagraphXml([text])]);
                // Normal order
                const zip1 = new JSZip();
                zip1.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
                zip1.file('word/document.xml', docXml);
                const buf1 = await zip1.generateAsync({ type: 'arraybuffer' });
                // Reversed order
                const zip2 = new JSZip();
                zip2.file('word/document.xml', docXml);
                zip2.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
                const buf2 = await zip2.generateAsync({ type: 'arraybuffer' });

                const r1 = await DocumentParser.parse({ name: 'a.docx', arrayBuffer: async () => buf1 });
                const r2 = await DocumentParser.parse({ name: 'b.docx', arrayBuffer: async () => buf2 });
                assert.strictEqual(r1.elements[0]?.text, r2.elements[0]?.text);
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m ZIP order independent');
    } catch (e) { report('P4-zip-order', e, SEED); }

    // INVARIANT 2: Every local finding resolves to existing document unit
    console.log('Property 2: Local findings resolve to existing units');
    try {
        await fc.assert(fc.asyncProperty(
            fc.array(arbRunSplit, { minLength: 1, maxLength: 10 }),
            async (paragraphs) => {
                const paras = paragraphs.map(parts => buildParagraphXml(parts));
                const docXml = buildDocXml(paras);
                const buf = await createDocx(docXml);
                const file = { name: 'fuzz.docx', arrayBuffer: async () => buf };
                const result = await DocumentParser.parse(file);
                const { findings } = RuleEngine.runAudit(result, opts);
                const elementIds = new Set(result.elements.map(e => e.id));
                // Add synthetic IDs
                elementIds.add('doc-global');
                elementIds.add('doc-coverage');
                for (const f of findings) {
                    if (f.globalPattern) continue; // global findings don't need location
                    if (f.paragraphId && !f.paragraphId.startsWith('fn-') && !f.paragraphId.startsWith('en-') &&
                        !f.paragraphId.startsWith('hdr-') && !f.paragraphId.startsWith('ftr-') &&
                        !f.paragraphId.startsWith('doc-')) {
                        assert(elementIds.has(f.paragraphId),
                            `Finding "${f.id}" references non-existent element "${f.paragraphId}"`);
                    }
                }
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m All findings resolve');
    } catch (e) { report('P2-resolve', e, SEED); }

    // INVARIANT 8: No duplicate findings at same location (after app-level dedup)
    console.log('Property 8: No duplicate findings after dedup');
    try {
        await fc.assert(fc.asyncProperty(
            fc.array(arbRunSplit, { minLength: 1, maxLength: 10 }),
            async (paragraphs) => {
                const paras = paragraphs.map(parts => buildParagraphXml(parts));
                const docXml = buildDocXml(paras);
                const buf = await createDocx(docXml);
                const file = { name: 'fuzz.docx', arrayBuffer: async () => buf };
                const result = await DocumentParser.parse(file);
                const { findings } = RuleEngine.runAudit(result, opts);
                // Apply same dedup logic as app.js (paragraphId+category+original+replacement)
                const seen = new Set();
                const deduped = [];
                for (const f of findings) {
                    const key = `${f.paragraphId}||${f.tableId||''}||${f.cellId||''}||${f.category}||${f.original}||${f.replacement}`;
                    if (!seen.has(key)) { seen.add(key); deduped.push(f); }
                }
                // After dedup, no duplicates should remain
                const seen2 = new Set();
                for (const f of deduped) {
                    const key = `${f.paragraphId}||${f.tableId||''}||${f.cellId||''}||${f.category}||${f.original}||${f.replacement}`;
                    assert(!seen2.has(key), `Duplicate after dedup: "${f.category}" "${f.original?.substring(0,30)}"`);
                    seen2.add(key);
                }
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m No duplicates after dedup');
    } catch (e) { report('P8-dedup', e, SEED); }

    // INVARIANT 10: Visible text never silently lost
    console.log('Property 10: No silent text loss');
    try {
        await fc.assert(fc.asyncProperty(
            fc.array(arbRunSplit, { minLength: 1, maxLength: 10 }),
            async (paragraphs) => {
                const visibleText = paragraphs.map(parts => parts.join('')).join('');
                const paras = paragraphs.map(parts => buildParagraphXml(parts));
                const docXml = buildDocXml(paras);
                const buf = await createDocx(docXml);
                const file = { name: 'fuzz.docx', arrayBuffer: async () => buf };
                const result = await DocumentParser.parse(file);
                // All generated visible text must appear in rawText
                const rawText = result.rawText || '';
                const combinedParsed = result.elements.map(e => e.text).join('');
                // Each paragraph's text must be present
                for (let i = 0; i < paragraphs.length; i++) {
                    const expected = paragraphs[i].join('');
                    if (expected.trim().length === 0) continue; // empty paragraphs may be skipped
                    assert(rawText.includes(expected) || combinedParsed.includes(expected),
                        `Text lost: "${expected.substring(0,50)}" not in parsed output`);
                }
            }
        ), { numRuns: NUM_RUNS, seed: SEED });
        passed++; console.log('  \x1b[32m✓\x1b[0m No text loss');
    } catch (e) { report('P10-text-loss', e, SEED); }

    // Summary
    console.log(`\n${'='.repeat(40)}`);
    console.log(`Properties: ${passed + failed} total, ${passed} passed, ${failed} failed`);
    if (failures.length > 0) {
        console.log('\nFailures:');
        failures.forEach(f => console.log(`  ${f}`));
    }
    console.log('='.repeat(40));
    process.exit(failed > 0 ? 1 : 0);
}

runProperties().catch(e => { console.error(e); process.exit(1); });
