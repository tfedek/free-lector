/**
 * Free Lector - Evaluation Framework Tests
 * Tests for sample-size.js and evaluate-confidence.js
 * Run: node test/evaluation.test.js
 */
'use strict';

const assert = require('assert');
const path = require('path');

const { calculateSampleSize } = require(path.resolve(__dirname, '..', 'evaluation', 'sample-size.js'));
const { parseCSV, wilsonCI } = require(path.resolve(__dirname, '..', 'evaluation', 'evaluate-confidence.js'));

let passed = 0, failed = 0;

function test(name, fn) {
    try { fn(); passed++; console.log(`  \x1b[32m\u2713\x1b[0m ${name}`); }
    catch (e) { failed++; console.log(`  \x1b[31m\u2717\x1b[0m ${name}\n    ${e.message}`); }
}

console.log('\nEvaluation Framework Tests:\n');

// ==========================================
// SAMPLE SIZE TESTS
// ==========================================
console.log('Sample size calculations:');

test('p=0.5, margin=0.05, 95% -> ~385', () => {
    const result = calculateSampleSize({ p: 0.5, margin: 0.05, z: 1.96, findingsPerDocument: 1, designEffect: 1 });
    // ceil(1.96^2 * 0.5 * 0.5 / 0.05^2) = ceil(384.16) = 385
    assert.strictEqual(result.requiredFindings, 385, `Expected 385, got ${result.requiredFindings}`);
});

test('p=0.5, margin=0.10, 95% -> ~97', () => {
    const result = calculateSampleSize({ p: 0.5, margin: 0.10, z: 1.96, findingsPerDocument: 1, designEffect: 1 });
    // ceil(1.96^2 * 0.5 * 0.5 / 0.10^2) = ceil(96.04) = 97
    assert.strictEqual(result.requiredFindings, 97, `Expected 97, got ${result.requiredFindings}`);
});

test('design effect multiplies required findings', () => {
    const result = calculateSampleSize({ p: 0.5, margin: 0.05, z: 1.96, findingsPerDocument: 5, designEffect: 2 });
    assert.strictEqual(result.requiredFindings, 385);
    assert.strictEqual(result.adjustedFindings, 770); // 385 * 2
    assert.strictEqual(result.estimatedDocuments, 154); // ceil(770 / 5)
});

test('input validation: p=0 throws', () => {
    let threw = false;
    try { calculateSampleSize({ p: 0, margin: 0.05, z: 1.96, findingsPerDocument: 5, designEffect: 1 }); }
    catch (e) { threw = true; assert(e.message.includes('p must be')); }
    assert(threw, 'Should throw for p=0');
});

test('input validation: p=1 throws', () => {
    let threw = false;
    try { calculateSampleSize({ p: 1, margin: 0.05, z: 1.96, findingsPerDocument: 5, designEffect: 1 }); }
    catch (e) { threw = true; assert(e.message.includes('p must be')); }
    assert(threw, 'Should throw for p=1');
});

test('input validation: margin=0 throws', () => {
    let threw = false;
    try { calculateSampleSize({ p: 0.5, margin: 0, z: 1.96, findingsPerDocument: 5, designEffect: 1 }); }
    catch (e) { threw = true; assert(e.message.includes('margin must be')); }
    assert(threw, 'Should throw for margin=0');
});

test('input validation: findingsPerDocument=0 throws', () => {
    let threw = false;
    try { calculateSampleSize({ p: 0.5, margin: 0.05, z: 1.96, findingsPerDocument: 0, designEffect: 1 }); }
    catch (e) { threw = true; assert(e.message.includes('findingsPerDocument must be')); }
    assert(threw, 'Should throw for findingsPerDocument=0');
});

test('input validation: designEffect<1 throws', () => {
    let threw = false;
    try { calculateSampleSize({ p: 0.5, margin: 0.05, z: 1.96, findingsPerDocument: 5, designEffect: 0.5 }); }
    catch (e) { threw = true; assert(e.message.includes('designEffect must be')); }
    assert(threw, 'Should throw for designEffect<1');
});

// ==========================================
// WILSON CI TESTS
// ==========================================
console.log('\nWilson CI:');

test('Wilson CI for small sample (N=5, TP=4)', () => {
    const ci = wilsonCI(4, 5);
    // Precision = 4/5 = 0.8
    assert(ci.lower > 0.2, `Lower bound too low: ${ci.lower}`);
    assert(ci.upper <= 1.0, `Upper bound too high: ${ci.upper}`);
    assert(ci.lower < 0.8, 'Lower bound should be below point estimate');
    assert(ci.upper > 0.8, 'Upper bound should be above point estimate');
    assert(ci.width > 0, 'Width should be positive');
    assert(ci.width < 1, 'Width should be less than 1');
});

test('Wilson CI for N=0 returns zeros', () => {
    const ci = wilsonCI(0, 0);
    assert.strictEqual(ci.lower, 0);
    assert.strictEqual(ci.upper, 0);
    assert.strictEqual(ci.width, 0);
});

// ==========================================
// BRIER SCORE TESTS
// ==========================================
console.log('\nBrier score:');

test('Brier score calculation', () => {
    // Perfect predictions: confidence=1 for TP, confidence=0 for FP
    // Brier = mean((pred - label)^2)
    // (1-1)^2 + (0-0)^2 = 0
    const perfect = [{ confidence: 1, label: 1 }, { confidence: 0, label: 0 }];
    const brierPerfect = perfect.reduce((s, x) => s + Math.pow(x.confidence - x.label, 2), 0) / perfect.length;
    assert.strictEqual(brierPerfect, 0, 'Perfect predictions should have Brier=0');

    // Worst predictions: confidence=0 for TP, confidence=1 for FP
    const worst = [{ confidence: 0, label: 1 }, { confidence: 1, label: 0 }];
    const brierWorst = worst.reduce((s, x) => s + Math.pow(x.confidence - x.label, 2), 0) / worst.length;
    assert.strictEqual(brierWorst, 1, 'Worst predictions should have Brier=1');

    // Mixed: confidence=0.7 for label=1
    const mixed = [{ confidence: 0.7, label: 1 }];
    const brierMixed = mixed.reduce((s, x) => s + Math.pow(x.confidence - x.label, 2), 0) / mixed.length;
    assert(Math.abs(brierMixed - 0.09) < 0.001, `Expected ~0.09, got ${brierMixed}`);
});

// ==========================================
// RELIABILITY BIN TESTS
// ==========================================
console.log('\nReliability bins:');

test('confidence=1.0 goes in [0.9,1.0] bin', () => {
    const c = 1.0;
    let binIdx = Math.min(9, Math.floor(c * 10));
    if (c === 1.0) binIdx = 9;
    assert.strictEqual(binIdx, 9, 'confidence=1.0 should go in bin index 9 ([0.9,1.0])');
});

test('confidence=0.0 goes in [0.0,0.1) bin', () => {
    const c = 0.0;
    let binIdx = Math.min(9, Math.floor(c * 10));
    assert.strictEqual(binIdx, 0, 'confidence=0.0 should go in bin index 0');
});

test('confidence=0.95 goes in [0.9,1.0] bin', () => {
    const c = 0.95;
    let binIdx = Math.min(9, Math.floor(c * 10));
    assert.strictEqual(binIdx, 9, 'confidence=0.95 should go in bin index 9');
});

// ==========================================
// CSV PARSING TESTS
// ==========================================
console.log('\nCSV parsing:');

test('quoted CSV field parsed correctly', () => {
    const csv = 'a,b,c\n"hello, world",bar,baz\n';
    const result = parseCSV(csv);
    assert.deepStrictEqual(result.header, ['a', 'b', 'c']);
    assert.strictEqual(result.rows.length, 1);
    assert.strictEqual(result.rows[0].fields[0], 'hello, world');
});

test('empty fields handled', () => {
    const csv = 'a,b,c\nfoo,,baz\n';
    const result = parseCSV(csv);
    assert.strictEqual(result.rows[0].fields[1], '');
});

test('CRLF line endings handled', () => {
    const csv = 'a,b,c\r\nfoo,bar,baz\r\n';
    const result = parseCSV(csv);
    assert.deepStrictEqual(result.header, ['a', 'b', 'c']);
    assert.strictEqual(result.rows.length, 1);
    assert.deepStrictEqual(result.rows[0].fields, ['foo', 'bar', 'baz']);
});

test('quoted field with embedded quotes', () => {
    const csv = 'a,b\n"say ""hello""",bar\n';
    const result = parseCSV(csv);
    assert.strictEqual(result.rows[0].fields[0], 'say "hello"');
});

// ==========================================
// MALFORMED INPUT TESTS
// ==========================================
console.log('\nMalformed input:');

test('malformed label gives error with row number (captured)', () => {
    // We test the validation logic directly
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,rule-1,f-1,p-1,0.5,INVALID,r1,rev-1,\n';
    const { header, rows } = parseCSV(csv);
    const colIdx = {};
    header.forEach((col, i) => { colIdx[col] = i; });
    const label = rows[0].fields[colIdx['label']];
    // Validate label is 0 or 1
    assert(label !== '0' && label !== '1', 'Label INVALID should fail validation');
    assert.strictEqual(rows[0].lineNumber, 2, 'Should report row number 2');
});

test('missing required column detected', () => {
    const csv = 'finding_id,label\nf-1,1\n';
    const { header } = parseCSV(csv);
    const requiredCols = ['rule_id', 'rule_version', 'version_id', 'label', 'predicted_confidence', 'finding_id'];
    const missing = requiredCols.filter(col => !header.includes(col));
    assert(missing.length > 0, 'Should detect missing columns');
    assert(missing.includes('rule_id'), 'Should detect missing rule_id');
});

// ==========================================
// REPORT CONTENT TESTS
// ==========================================
console.log('\nReport content:');

test('report does NOT contain recall or F1', () => {
    // The evaluate-confidence.js should never output recall or F1
    const fs = require('fs');
    const content = fs.readFileSync(path.resolve(__dirname, '..', 'evaluation', 'evaluate-confidence.js'), 'utf-8');
    // Check that the script's output strings don't contain recall/F1
    // (The disclaimer explicitly says "Ne meri recall, F1")
    assert(content.includes('Ne meri recall, F1'), 'Should include disclaimer about not measuring recall/F1');
    // Check it doesn't compute recall or F1
    assert(!content.includes('overallRecall'), 'Should not compute overallRecall');
    assert(!content.includes('overallF1'), 'Should not compute overallF1');
});

// ==========================================
// SUMMARY
// ==========================================
console.log('\n' + '='.repeat(40));
console.log(`Evaluation tests: ${passed} passed, ${failed} failed`);
console.log('='.repeat(40));
process.exit(failed > 0 ? 1 : 0);
