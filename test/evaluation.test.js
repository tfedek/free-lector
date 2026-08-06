/**
 * Free Lector - Evaluation Framework Tests
 * Tests for sample-size.js and evaluate-confidence.js
 * Run: node test/evaluation.test.js
 */
'use strict';

const assert = require('assert');
const path = require('path');
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');

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
    assert.strictEqual(result.requiredFindings, 385, `Expected 385, got ${result.requiredFindings}`);
});

test('p=0.5, margin=0.10, 95% -> ~97', () => {
    const result = calculateSampleSize({ p: 0.5, margin: 0.10, z: 1.96, findingsPerDocument: 1, designEffect: 1 });
    assert.strictEqual(result.requiredFindings, 97, `Expected 97, got ${result.requiredFindings}`);
});

test('design effect multiplies required findings', () => {
    const result = calculateSampleSize({ p: 0.5, margin: 0.05, z: 1.96, findingsPerDocument: 5, designEffect: 2 });
    assert.strictEqual(result.requiredFindings, 385);
    assert.strictEqual(result.adjustedFindings, 770);
    assert.strictEqual(result.estimatedDocuments, 154);
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
    const perfect = [{ confidence: 1, label: 1 }, { confidence: 0, label: 0 }];
    const brierPerfect = perfect.reduce((s, x) => s + Math.pow(x.confidence - x.label, 2), 0) / perfect.length;
    assert.strictEqual(brierPerfect, 0, 'Perfect predictions should have Brier=0');

    const worst = [{ confidence: 0, label: 1 }, { confidence: 1, label: 0 }];
    const brierWorst = worst.reduce((s, x) => s + Math.pow(x.confidence - x.label, 2), 0) / worst.length;
    assert.strictEqual(brierWorst, 1, 'Worst predictions should have Brier=1');

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
// CLI TESTS - evaluate-confidence.js
// ==========================================
console.log('\nCLI tests (evaluate-confidence.js):');

const evalScript = path.resolve(__dirname, '..', 'evaluation', 'evaluate-confidence.js');
const sampleScript = path.resolve(__dirname, '..', 'evaluation', 'sample-size.js');

function writeTempCSV(content) {
    const tmpFile = path.join(os.tmpdir(), `fl-test-${Date.now()}-${Math.random().toString(36).slice(2)}.csv`);
    fs.writeFileSync(tmpFile, content);
    return tmpFile;
}

test('malformed label exits with code 1 and shows row number', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,rule-1,f-1,p-1,0.5,INVALID,r1,rev-1,\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
        assert(result.stderr.includes('red 2'), `stderr should mention row number: ${result.stderr}`);
    } finally { fs.unlinkSync(tmpFile); }
});

test('0.9abc confidence rejected', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,rule-1,f-1,p-1,0.9abc,1,r1,rev-1,\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
        assert(result.stderr.includes('0.9abc'), `stderr should mention the bad value: ${result.stderr}`);
    } finally { fs.unlinkSync(tmpFile); }
});

test('unclosed quotes rejected', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,rule-1,f-1,p-1,0.9,1,r1,rev-1,"unclosed\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
        assert(result.stderr.includes('nezatvorene'), `stderr should mention unclosed quotes: ${result.stderr}`);
    } finally { fs.unlinkSync(tmpFile); }
});

test('wrong column count rejected', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,rule-1,f-1,p-1,0.9,1,r1\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
        assert(result.stderr.includes('kolona'), `stderr should mention column count: ${result.stderr}`);
    } finally { fs.unlinkSync(tmpFile); }
});

test('empty required field rejected', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,,f-1,p-1,0.9,1,r1,rev-1,\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
        assert(result.stderr.includes('prazna') || result.stderr.includes('prazan'), `stderr should mention empty field: ${result.stderr}`);
    } finally { fs.unlinkSync(tmpFile); }
});

test('valid CSV generates report', () => {
    const csv = 'document_id,version_id,rule_id,finding_id,location_key,predicted_confidence,label,rule_version,reviewer,note\ndoc-1,v-1,spacing_body,f-1,p-1,0.95,1,r1,rev-1,ok\n';
    const tmpFile = writeTempCSV(csv);
    try {
        const result = spawnSync('node', [evalScript, tmpFile], { encoding: 'utf-8' });
        assert.strictEqual(result.status, 0, `Expected exit code 0, got ${result.status}. stderr: ${result.stderr}`);
        // Check report was generated
        const reportDir = path.resolve(__dirname, '..', 'evaluation', 'report');
        assert(fs.existsSync(path.join(reportDir, 'confidence-report.json')), 'JSON report should exist');
        assert(fs.existsSync(path.join(reportDir, 'confidence-report.md')), 'MD report should exist');
    } finally { fs.unlinkSync(tmpFile); }
});

// ==========================================
// CLI TESTS - sample-size.js
// ==========================================
console.log('\nCLI tests (sample-size.js):');

test('simultaneous --confidence and --z rejected', () => {
    const result = spawnSync('node', [sampleScript, '--confidence', '0.95', '--z', '1.96'], { encoding: 'utf-8' });
    assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
    assert(result.stderr.includes('istovremeno'), `stderr should mention simultaneous use: ${result.stderr}`);
});

test('argument without value rejected', () => {
    const result = spawnSync('node', [sampleScript, '--p'], { encoding: 'utf-8' });
    assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
    assert(result.stderr.includes('vrednost'), `stderr should mention missing value: ${result.stderr}`);
});

test('NaN input rejected (--p abc)', () => {
    const result = spawnSync('node', [sampleScript, '--p', 'abc'], { encoding: 'utf-8' });
    assert.strictEqual(result.status, 1, `Expected exit code 1, got ${result.status}`);
});

test('valid sample-size invocation succeeds', () => {
    const result = spawnSync('node', [sampleScript, '--p', '0.5', '--margin', '0.05'], { encoding: 'utf-8' });
    assert.strictEqual(result.status, 0, `Expected exit code 0, got ${result.status}. stderr: ${result.stderr}`);
    assert(result.stdout.includes('385'), `Output should contain 385: ${result.stdout}`);
});

// ==========================================
// REPORT CONTENT TESTS
// ==========================================
console.log('\nReport content:');

test('report does NOT contain recall or F1', () => {
    const content = fs.readFileSync(path.resolve(__dirname, '..', 'evaluation', 'evaluate-confidence.js'), 'utf-8');
    assert(content.includes('Ne meri recall, F1'), 'Should include disclaimer about not measuring recall/F1');
    assert(!content.includes('overallRecall'), 'Should not compute overallRecall');
    assert(!content.includes('overallF1'), 'Should not compute overallF1');
});

test('report includes bootstrap note', () => {
    const content = fs.readFileSync(path.resolve(__dirname, '..', 'evaluation', 'evaluate-confidence.js'), 'utf-8');
    assert(content.includes('bootstrap na nivou celog dokumenta'), 'Should include bootstrap future note');
});

// ==========================================
// McNEMAR, FDR, CLIFF'S DELTA TESTS
// ==========================================

console.log('\nMcNemar test:');
test('McNemar detects significant improvement', () => {
    const { mcNemar } = require('../evaluation/evaluate-confidence.js');
    // 20 cases where before was wrong and after is right, 2 vice versa
    const paired = [];
    for (let i = 0; i < 20; i++) paired.push({before: 0, after: 1});
    for (let i = 0; i < 2; i++) paired.push({before: 1, after: 0});
    for (let i = 0; i < 78; i++) paired.push({before: 1, after: 1});
    const result = mcNemar(paired);
    assert(result.significant, 'Should be significant: 20 improvements vs 2 regressions');
    assert(result.c === 20, 'c (improvements) should be 20');
    assert(result.b === 2, 'b (regressions) should be 2');
});

test('McNemar not significant when no change', () => {
    const { mcNemar } = require('../evaluation/evaluate-confidence.js');
    const paired = [];
    for (let i = 0; i < 5; i++) paired.push({before: 0, after: 1});
    for (let i = 0; i < 5; i++) paired.push({before: 1, after: 0});
    for (let i = 0; i < 90; i++) paired.push({before: 1, after: 1});
    const result = mcNemar(paired);
    assert(!result.significant, 'Should NOT be significant: equal improvements and regressions');
});

console.log('\nFDR (Benjamini-Hochberg):');
test('FDR controls false discoveries', () => {
    const { benjaminiHochberg } = require('../evaluation/evaluate-confidence.js');
    const pValues = [
        {rule_id: 'a', pValue: 0.001},
        {rule_id: 'b', pValue: 0.01},
        {rule_id: 'c', pValue: 0.04},
        {rule_id: 'd', pValue: 0.06},
        {rule_id: 'e', pValue: 0.5},
    ];
    const results = benjaminiHochberg(pValues, 0.05);
    const sig = results.filter(r => r.significant);
    assert(sig.length >= 2, 'At least a and b should be significant');
    assert(!results.find(r => r.rule_id === 'e').significant, 'e (p=0.5) should not be significant');
});

console.log('\nCliff\'s delta:');
test('Cliff delta: clear superiority', () => {
    const { cliffsDelta } = require('../evaluation/evaluate-confidence.js');
    const group1 = [10, 11, 12, 13, 14]; // all higher
    const group2 = [1, 2, 3, 4, 5];      // all lower
    const result = cliffsDelta(group1, group2);
    assert.strictEqual(result.delta, 1.0, 'Delta should be 1.0 (complete superiority)');
    assert.strictEqual(result.magnitude, 'large');
});

test('Cliff delta: no difference', () => {
    const { cliffsDelta } = require('../evaluation/evaluate-confidence.js');
    const group1 = [1, 2, 3, 4, 5];
    const group2 = [1, 2, 3, 4, 5];
    const result = cliffsDelta(group1, group2);
    assert.strictEqual(result.delta, 0, 'Delta should be 0 (identical)');
    assert.strictEqual(result.magnitude, 'negligible');
});

console.log('\nBootstrap BCa:');
test('Bootstrap BCa returns valid interval', () => {
    const { bootstrapBCa } = require('../evaluation/evaluate-confidence.js');
    const data = [1,1,1,1,1,1,1,1,0,0]; // 80% success rate
    const result = bootstrapBCa(data, d => d.filter(x=>x===1).length / d.length, 2000, 0.05);
    assert(result.lower >= 0.1 && result.lower <= 0.8, 'Lower bound should be reasonable: ' + result.lower);
    assert(result.upper >= 0.8 && result.upper <= 1.0, 'Upper bound should be reasonable: ' + result.upper);
    assert(result.lower < result.upper, 'Lower must be less than upper');
});

// ==========================================
// SUMMARY
// ==========================================
console.log('\n' + '='.repeat(40));
console.log(`Evaluation tests: ${passed} passed, ${failed} failed`);
console.log('='.repeat(40));
process.exit(failed > 0 ? 1 : 0);
