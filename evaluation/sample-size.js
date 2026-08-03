#!/usr/bin/env node
'use strict';

/**
 * Sample size calculator for confidence evaluation.
 *
 * Formula:
 *   requiredFindings = ceil(z^2 * p * (1-p) / e^2)
 *   adjustedFindings = ceil(requiredFindings * designEffect)
 *   estimatedDocuments = ceil(adjustedFindings / findingsPerDocument)
 *
 * Usage:
 *   node evaluation/sample-size.js --p 0.5 --margin 0.05 --confidence 0.95 --findings-per-document 5 --design-effect 1
 *   node evaluation/sample-size.js --p 0.5 --margin 0.05 --z 1.96
 */

const args = process.argv.slice(2);

function getArg(name, defaultVal) {
    const idx = args.indexOf(`--${name}`);
    if (idx !== -1 && args[idx + 1] !== undefined) return args[idx + 1];
    if (defaultVal === null || defaultVal === undefined) return null;
    return String(defaultVal);
}

function calculateSampleSize({ p, margin, z, findingsPerDocument, designEffect }) {
    // Validation
    if (typeof p !== 'number' || p <= 0 || p >= 1) {
        throw new Error('p must be between 0 and 1 (exclusive)');
    }
    if (typeof margin !== 'number' || margin <= 0 || margin >= 1) {
        throw new Error('margin must be between 0 and 1 (exclusive)');
    }
    if (typeof z !== 'number' || z <= 0) {
        throw new Error('z must be a positive number');
    }
    if (typeof findingsPerDocument !== 'number' || findingsPerDocument <= 0) {
        throw new Error('findingsPerDocument must be greater than 0');
    }
    if (typeof designEffect !== 'number' || designEffect < 1) {
        throw new Error('designEffect must be >= 1');
    }

    const requiredFindings = Math.ceil(z * z * p * (1 - p) / (margin * margin));
    const adjustedFindings = Math.ceil(requiredFindings * designEffect);
    const estimatedDocuments = Math.ceil(adjustedFindings / findingsPerDocument);

    return { requiredFindings, adjustedFindings, estimatedDocuments };
}

// Z-scores for common confidence levels
const zScores = { '0.90': 1.645, '0.95': 1.96, '0.99': 2.576 };

function resolveZ(confidenceStr, explicitZ) {
    if (explicitZ !== null) return parseFloat(explicitZ);
    const z = zScores[confidenceStr];
    if (z) return z;
    throw new Error(`Unknown confidence level: ${confidenceStr}. Use --z for explicit override.`);
}

function main() {
    const pStr = getArg('p', '0.5');
    const marginStr = getArg('margin', '0.05');
    const confidenceStr = getArg('confidence', '0.95');
    const findingsPerDocStr = getArg('findings-per-document', '5');
    const designEffectStr = getArg('design-effect', '1');
    const explicitZStr = getArg('z', null);

    const p = parseFloat(pStr);
    const margin = parseFloat(marginStr);
    const z = resolveZ(confidenceStr, explicitZStr);
    const findingsPerDocument = parseFloat(findingsPerDocStr);
    const designEffect = parseFloat(designEffectStr);

    const result = calculateSampleSize({ p, margin, z, findingsPerDocument, designEffect });

    console.log('\n=== Sample Size Calculator ===\n');
    console.log(`p (expected precision):     ${p}`);
    console.log(`Margin of error:            +/- ${(margin * 100).toFixed(1)}%`);
    console.log(`z-score:                    ${z}`);
    console.log(`Findings per document:      ${findingsPerDocument}`);
    console.log(`Design effect:              ${designEffect}`);
    console.log('');
    console.log(`Required findings:          ${result.requiredFindings}`);
    console.log(`Adjusted findings (DE):     ${result.adjustedFindings}`);
    console.log(`Estimated documents:        ${result.estimatedDocuments}`);

    if (designEffect === 1) {
        console.log('\nOvo je optimistična donja procena koja pretpostavlja nezavisne nalaze.');
    }
}

// Export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { calculateSampleSize, resolveZ };
}

// Run if called directly
if (require.main === module) {
    main();
}
