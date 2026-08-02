#!/usr/bin/env node
'use strict';

/**
 * Evaluate confidence metrics from labeled samples.
 * Reads labels.csv and generates precision/recall/F1 per category.
 *
 * Usage: node evaluation/evaluate-confidence.js [path-to-labels.csv]
 */

const fs = require('fs');
const path = require('path');

const labelsPath = process.argv[2] || path.join(__dirname, 'labels.csv');
const reportDir = path.join(__dirname, 'report');

if (!fs.existsSync(labelsPath)) {
    console.error(`Error: Labels file not found: ${labelsPath}`);
    process.exit(1);
}

// Parse CSV
const raw = fs.readFileSync(labelsPath, 'utf-8').trim();
const lines = raw.split('\n');
const header = lines[0].split(',');

const requiredCols = ['finding_id', 'category', 'human_label'];
for (const col of requiredCols) {
    if (!header.includes(col)) {
        console.error(`Error: Missing required column "${col}" in CSV header.`);
        process.exit(1);
    }
}

const records = [];
for (let i = 1; i < lines.length; i++) {
    const parts = lines[i].split(',');
    if (parts.length < 3) continue;
    const record = {};
    header.forEach((col, idx) => { record[col.trim()] = (parts[idx] || '').trim(); });
    if (!record.finding_id || !record.human_label) continue;
    records.push(record);
}

if (records.length === 0) {
    console.error('Error: No valid records found in labels.csv');
    process.exit(1);
}

// Compute metrics per category
const categories = [...new Set(records.map(r => r.category))].sort();
const results = {};
let totalTP = 0, totalFP = 0, totalFN = 0;

for (const cat of categories) {
    const catRecords = records.filter(r => r.category === cat);
    const tp = catRecords.filter(r => r.human_label === 'TP').length;
    const fp = catRecords.filter(r => r.human_label === 'FP').length;
    const fn = catRecords.filter(r => r.human_label === 'FN').length;
    const precision = tp + fp > 0 ? tp / (tp + fp) : 0;
    const recall = tp + fn > 0 ? tp / (tp + fn) : 0;
    const f1 = precision + recall > 0 ? 2 * precision * recall / (precision + recall) : 0;

    results[cat] = { tp, fp, fn, precision, recall, f1, total: catRecords.length };
    totalTP += tp;
    totalFP += fp;
    totalFN += fn;
}

// Overall metrics
const overallPrecision = totalTP + totalFP > 0 ? totalTP / (totalTP + totalFP) : 0;
const overallRecall = totalTP + totalFN > 0 ? totalTP / (totalTP + totalFN) : 0;
const overallF1 = overallPrecision + overallRecall > 0
    ? 2 * overallPrecision * overallRecall / (overallPrecision + overallRecall) : 0;

// Print report
console.log('\n=== Free Lector Confidence Evaluation ===\n');
console.log(`Total labeled samples: ${records.length}`);
console.log(`Overall: Precision=${(overallPrecision * 100).toFixed(1)}% Recall=${(overallRecall * 100).toFixed(1)}% F1=${(overallF1 * 100).toFixed(1)}%\n`);
console.log('Per category:');
console.log('-'.repeat(70));
console.log(`${'Category'.padEnd(25)} ${'TP'.padStart(4)} ${'FP'.padStart(4)} ${'FN'.padStart(4)} ${'Prec'.padStart(7)} ${'Rec'.padStart(7)} ${'F1'.padStart(7)}`);
console.log('-'.repeat(70));

for (const cat of categories) {
    const r = results[cat];
    console.log(`${cat.padEnd(25)} ${String(r.tp).padStart(4)} ${String(r.fp).padStart(4)} ${String(r.fn).padStart(4)} ${(r.precision * 100).toFixed(1).padStart(6)}% ${(r.recall * 100).toFixed(1).padStart(6)}% ${(r.f1 * 100).toFixed(1).padStart(6)}%`);
}
console.log('-'.repeat(70));

// Write report to file
if (!fs.existsSync(reportDir)) fs.mkdirSync(reportDir, { recursive: true });
const reportPath = path.join(reportDir, `report-${new Date().toISOString().slice(0, 10)}.json`);
const report = {
    date: new Date().toISOString(),
    total_samples: records.length,
    overall: { precision: overallPrecision, recall: overallRecall, f1: overallF1 },
    per_category: results,
};
fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
console.log(`\nReport saved: ${reportPath}`);
