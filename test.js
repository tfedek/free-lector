/**
 * Free Lector — Automated Test Suite
 * Uses Node.js assert (no eval, proper module imports)
 * Exit code 1 on any failure.
 *
 * Run: node test.js
 */

'use strict';

const assert = require('assert');
const RuleEngine = require('./rules.js');

// ==========================================
// HELPERS: build minimal docMap for testing
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
        };
        return base;
    });
    return {
        type: opts.type || 'txt',
        name: opts.name || 'test.txt',
        elements: mapped,
        footnotes: opts.footnotes || [],
        endnotes: opts.endnotes || [],
        headers: [],
        footers: [],
        styles: {},
        numbering: opts.numbering || {},
        htmlPreview: '',
        rawText: mapped.map(e => e.text).join('\n'),
        wordCount: 100,
        paragraphCount: mapped.filter(e => e.type === 'paragraph').length,
        tableCount: 0,
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

let passed = 0;
let failed = 0;

function test(name, fn) {
    try {
        fn();
        passed++;
        console.log(`  \x1b[32m✓\x1b[0m ${name}`);
    } catch (err) {
        failed++;
        console.log(`  \x1b[31m✗\x1b[0m ${name}`);
        console.log(`    ${err.message}`);
    }
}


// ==========================================
// TEST: Straight quotes detection
// ==========================================
console.log('\nNavodnici:');

test('detects straight double quotes', () => {
    const doc = makeDocMap([{ text: 'On kaže: "zdravo" i ode.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const quoteFindings = findings.filter(f => f.category === 'Tipografija' && f.original.includes('ravni'));
    assert(quoteFindings.length >= 1, `Expected quote finding, got ${quoteFindings.length}`);
});

test('no false positive on typographic quotes', () => {
    const doc = makeDocMap([{ text: 'On kaže: \u201ezdravо\u201c i ode.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const straight = findings.filter(f => f.category === 'Tipografija' && f.rationale.includes('ravni'));
    assert.strictEqual(straight.length, 0, `Expected 0 straight quote findings, got ${straight.length}`);
});

// ==========================================
// TEST: Unmatched typographic quotes
// ==========================================
console.log('\nNeupareni navodnici:');

test('detects unmatched opening quote', () => {
    const doc = makeDocMap([{ text: 'Rekao je \u201ezdravо ali zaboravio da zatvori.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const unmatched = findings.filter(f => f.rationale.includes('ne odgovara'));
    assert(unmatched.length >= 1, `Expected unmatched quote finding, got ${unmatched.length}`);
});

// ==========================================
// TEST: En-dash for numeric ranges
// ==========================================
console.log('\nPoluduge crtice:');

test('detects hyphen in numeric range 484-425', () => {
    const doc = makeDocMap([{ text: 'Period je bio 484-425 p.n.e.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const dash = findings.filter(f => f.category === 'Tipografija' && f.replacement && f.replacement.includes('\u2013'));
    assert(dash.length >= 1, `Expected en-dash finding, got ${dash.length}`);
    assert(dash[0].replacement.includes('484\u2013425'), `Expected 484–425, got ${dash[0].replacement}`);
});

test('detects double hyphen as em-dash', () => {
    const doc = makeDocMap([{ text: 'Rekao je -- to nije bitno.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const emDash = findings.filter(f => f.rationale.includes('em-dash'));
    assert(emDash.length >= 1, `Expected em-dash finding, got ${emDash.length}`);
});

// ==========================================
// TEST: Script mixing (Cyrillic/Latin)
// ==========================================
console.log('\nMešanje pisama:');

test('detects mixed Cyrillic/Latin word', () => {
    // К is Cyrillic, admo is Latin
    const doc = makeDocMap([{ text: 'Reč \u041Cadmo je pomešana.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const mix = findings.filter(f => f.category === 'Mešanje pisama');
    assert(mix.length >= 1, `Expected script mixing finding, got ${mix.length}`);
});

test('no false positive on pure Latin word', () => {
    const doc = makeDocMap([{ text: 'The word Kadmo is fine.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const mix = findings.filter(f => f.category === 'Mešanje pisama');
    assert.strictEqual(mix.length, 0, `Expected 0, got ${mix.length}`);
});


// ==========================================
// TEST: Numbering skip 1 → 3
// ==========================================
console.log('\nNumeracija (preskok 1→3):');

test('detects numbering skip from 1 to 3', () => {
    const doc = makeDocMap([
        { text: '1. Prva stavka' },
        { text: '3. Treća stavka' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const numFindings = findings.filter(f => f.category === 'Numeracija');
    assert(numFindings.length >= 1, `Expected numbering gap, got ${numFindings.length}`);
    assert(numFindings[0].rationale.includes('preskače'), `Expected "preskače" in rationale`);
});

test('allows list restart at 1', () => {
    const doc = makeDocMap([
        { text: '1. Prva lista stavka 1' },
        { text: '2. Prva lista stavka 2' },
        { text: '1. Druga lista stavka 1' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const numFindings = findings.filter(f => f.category === 'Numeracija');
    assert.strictEqual(numFindings.length, 0, `Expected 0, restart is valid, got ${numFindings.length}`);
});

// ==========================================
// TEST: OOXML numbering (displayedNumber)
// ==========================================
console.log('\nNumeracija (OOXML):');

test('detects OOXML numbering skip via displayedNumber', () => {
    const doc = makeDocMap([
        { text: 'Stavka A', numId: '1', numLevel: 0, displayedNumber: 1, listInstanceId: '1-0', listStart: 1 },
        { text: 'Stavka B', numId: '1', numLevel: 0, displayedNumber: 2, listInstanceId: '1-0', listStart: 1 },
        { text: 'Stavka D', numId: '1', numLevel: 0, displayedNumber: 4, listInstanceId: '1-0', listStart: 1 },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const numFindings = findings.filter(f => f.category === 'Numeracija');
    assert(numFindings.length >= 1, `Expected OOXML numbering gap, got ${numFindings.length}`);
});

// ==========================================
// TEST: Direct quotes handling
// ==========================================
console.log('\nDirektni citati:');

test('findings inside direct quotes get PROVERITI priority', () => {
    const doc = makeDocMap([
        { text: 'On kaže  dva razmaka unutar citata.', isDirectQuote: true, quoteConfidence: 0.95 },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const spacingInQuote = findings.filter(f => f.category === 'Razmaci');
    assert(spacingInQuote.length >= 1, `Expected spacing finding in quote`);
    assert.strictEqual(spacingInQuote[0].priority, 'PROVERITI',
        `Expected PROVERITI for quote, got ${spacingInQuote[0].priority}`);
    assert.strictEqual(spacingInQuote[0].autoFixable, false);
    assert.strictEqual(spacingInQuote[0].requiresSourceVerification, true);
});

// ==========================================
// TEST: Tables (table elements pass through)
// ==========================================
console.log('\nTabele:');

test('table elements are included in docMap and scanned', () => {
    const doc = makeDocMap([
        { type: 'table', text: 'Kolona 1 | Kolona 2\nRed 1 | Red  2' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, { ...allOptions() });
    // Should detect double space in "Red  2"
    const spacing = findings.filter(f => f.category === 'Razmaci');
    assert(spacing.length >= 1, `Expected spacing finding in table text, got ${spacing.length}`);
});


// ==========================================
// TEST: Footnotes (empty)
// ==========================================
console.log('\nFusnote:');

test('detects empty footnote', () => {
    const doc = makeDocMap(
        [{ text: 'Tekst sa fusnotom.' }],
        { footnotes: [{ id: '1', text: '', isEmpty: true }] }
    );
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const fnFindings = findings.filter(f => f.category === 'Fusnote');
    assert(fnFindings.length >= 1, `Expected empty footnote finding, got ${fnFindings.length}`);
});

test('no finding for non-empty footnote', () => {
    const doc = makeDocMap(
        [{ text: 'Tekst sa fusnotom.' }],
        { footnotes: [{ id: '1', text: 'Vid. Apijan, Ilirika 2.', isEmpty: false }] }
    );
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const fnFindings = findings.filter(f => f.category === 'Fusnote');
    assert.strictEqual(fnFindings.length, 0, `Expected 0, got ${fnFindings.length}`);
});

// ==========================================
// TEST: Bibliography
// ==========================================
console.log('\nBibliografija:');

test('detects bibliography entry without year', () => {
    const doc = makeDocMap([
        { type: 'heading', text: 'Bibliografija', headingLevel: 1 },
        { text: 'Apijan, Rimska istorija, prevod Miroslav Marković, Beograd' },
    ]);
    const opts = allOptions();
    const { findings } = RuleEngine.runAudit(doc, opts);
    const bibFindings = findings.filter(f => f.category === 'Bibliografija');
    assert(bibFindings.length >= 1, `Expected bib finding, got ${bibFindings.length}. All: ${findings.map(f=>f.category).join(',')}`);
});

test('no year warning for entry with year', () => {
    const doc = makeDocMap([
        { type: 'heading', text: 'Bibliografija', headingLevel: 1 },
        { text: 'Smith J., Greek Mythology, Oxford University Press, 2015' },
    ]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const yearFindings = findings.filter(f => f.category === 'Bibliografija' && f.rationale.includes('godina'));
    assert.strictEqual(yearFindings.length, 0, `Expected 0 year findings, got ${yearFindings.length}`);
});

// ==========================================
// TEST: Duplicate words
// ==========================================
console.log('\nDuple reči:');

test('detects "je je" as duplicate', () => {
    const doc = makeDocMap([{ text: 'Kadmo je je stigao u Iliriju.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const dupes = findings.filter(f => f.category === 'Duple reči');
    assert(dupes.length >= 1, `Expected dupe finding, got ${dupes.length}`);
    assert.strictEqual(dupes[0].replacement, 'je');
});

test('does not flag "da" as duplicate (allowed)', () => {
    const doc = makeDocMap([{ text: 'Da da, slažem se.' }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const dupes = findings.filter(f => f.category === 'Duple reči');
    assert.strictEqual(dupes.length, 0, `Expected 0 (da is allowed), got ${dupes.length}`);
});

// ==========================================
// TEST: Greek without translation
// ==========================================
console.log('\nGrčki bez prevoda:');

test('detects Greek text without nearby translation', () => {
    const doc = makeDocMap([{
        text: 'Apijan piše: \u03C4\u1F78\u03BD \u039A\u03AC\u03B4\u03BC\u03BF\u03BD \u1F10\u03BB\u03B8\u03B5\u1FD6\u03BD \u03B5\u1F30\u03C2 \u1F38\u03BB\u03BB\u03C5\u03C1\u03B9\u03BF\u03CD\u03C2. Nastavio je putovanje.'
    }]);
    const { findings } = RuleEngine.runAudit(doc, allOptions());
    const greek = findings.filter(f => f.category === 'Grčki bez prevoda');
    assert(greek.length >= 1, `Expected Greek finding, got ${greek.length}`);
});


// ==========================================
// TEST: passedChecks uses scannedCount not elements.length
// ==========================================
console.log('\npassedChecks.count:');

test('passedChecks.count reflects scannedCount, not elements.length', () => {
    // 3 elements but one is empty → scannedCount should be 2
    const doc = makeDocMap([
        { text: 'Tekst bez problema.' },
        { text: '' },
        { text: 'Još teksta.' },
    ]);
    const opts = { brackets: true };
    const { passedChecks } = RuleEngine.runAudit(doc, opts);
    assert(passedChecks.length >= 1, 'Expected brackets to pass');
    // scannedCount should be 2 (skipped the empty one)
    assert.strictEqual(passedChecks[0].count, 2,
        `Expected scannedCount=2, got ${passedChecks[0].count}`);
});

// ==========================================
// SUMMARY
// ==========================================
console.log('\n' + '='.repeat(40));
console.log(`Rezultat: ${passed} prošlo, ${failed} palo`);
console.log('='.repeat(40));

if (failed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
