#!/usr/bin/env node
'use strict';

/**
 * Calculate required sample size for confidence evaluation.
 *
 * Usage: node evaluation/sample-size.js [--population N] [--margin 0.05] [--confidence 0.95]
 */

const args = process.argv.slice(2);

function getArg(name, defaultVal) {
    const idx = args.indexOf(`--${name}`);
    if (idx !== -1 && args[idx + 1]) return parseFloat(args[idx + 1]);
    return defaultVal;
}

const population = getArg('population', 1000);
const margin = getArg('margin', 0.05);
const confidence = getArg('confidence', 0.95);

// Z-scores for common confidence levels
const zScores = { 0.90: 1.645, 0.95: 1.96, 0.99: 2.576 };
const z = zScores[confidence] || 1.96;

// Conservative estimate (p = 0.5 for maximum sample size)
const p = 0.5;

// Cochran's formula for finite population
const n0 = Math.ceil((z * z * p * (1 - p)) / (margin * margin));
const n = Math.ceil(n0 / (1 + (n0 - 1) / population));

console.log('\n=== Sample Size Calculator ===\n');
console.log(`Population size:    ${population}`);
console.log(`Confidence level:   ${(confidence * 100).toFixed(0)}%`);
console.log(`Margin of error:    +/- ${(margin * 100).toFixed(1)}%`);
console.log(`Z-score:            ${z}`);
console.log('');
console.log(`Cochran's n0:       ${n0} (infinite population)`);
console.log(`Adjusted n:         ${n} (finite population correction)`);
console.log('');
console.log(`Required sample size: ${n}`);
console.log(`Sampling rate:        ${((n / population) * 100).toFixed(1)}%`);
