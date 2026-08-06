#!/usr/bin/env node
'use strict';

/**
 * Evaluate confidence metrics from labeled samples.
 * Reads labels.csv and generates precision, Wilson CI, and Brier score per rule_id.
 *
 * Ova evaluacija meri preciznost prijavljenih nalaza.
 * Ne meri recall, F1 niti greške koje alat nije pronašao.
 *
 * Nalazi iz istog dokumenta nisu statistički nezavisni.
 * Wilson intervali po pojedinačnim nalazima mogu potceniti stvarnu neizvesnost.
 *
 * Usage: node evaluation/evaluate-confidence.js [path-to-labels.csv]
 */

const fs = require('fs');
const path = require('path');

const labelsPath = process.argv[2] || path.join(__dirname, 'labels.csv');
const reportDir = path.join(__dirname, 'report');

// ==========================================
// CSV PARSING (handles quoted fields, CRLF/LF, commas in quotes)
// ==========================================
function parseCSV(text) {
    const lines = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
            inQuotes = !inQuotes;
            current += ch;
        } else if ((ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) && !inQuotes) {
            lines.push(current);
            current = '';
            if (ch === '\r') i++; // skip \n after \r
        } else if (ch === '\r' && !inQuotes) {
            lines.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    if (current.length > 0) lines.push(current);

    if (lines.length === 0) return { header: [], rows: [] };

    const header = splitCSVLine(lines[0], 1);
    if (!header) { process.exit(1); return { header: [], rows: [] }; }
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim().length === 0) continue;
        const fields = splitCSVLine(lines[i], i + 1);
        if (!fields) { process.exit(1); return { header, rows }; }
        rows.push({ lineNumber: i + 1, fields });
    }
    return { header, rows };
}

function splitCSVLine(line, rowNum) {
    const fields = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            if (inQuotes && line[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            fields.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    if (inQuotes) {
        console.error(`Greška red ${rowNum}: nezatvoreni navodnici.`);
        return null;
    }
    fields.push(current);
    return fields;
}

// ==========================================
// WILSON CONFIDENCE INTERVAL
// ==========================================
function wilsonCI(successes, total, z = 1.96) {
    if (total === 0) return { lower: 0, upper: 0, center: 0, width: 0 };
    const p = successes / total;
    const n = total;
    const denom = 1 + z * z / n;
    const center = (p + z * z / (2 * n)) / denom;
    const margin = (z / denom) * Math.sqrt(p * (1 - p) / n + z * z / (4 * n * n));
    const lower = Math.max(0, center - margin);
    const upper = Math.min(1, center + margin);
    return { lower, upper, center, width: upper - lower };
}

// ==========================================
// MAIN
// ==========================================
function run() {
    if (!fs.existsSync(labelsPath)) {
        console.error(`Error: Labels file not found: ${labelsPath}`);
        process.exit(1);
    }

    const raw = fs.readFileSync(labelsPath, 'utf-8');

    // Detect unclosed quotes
    let inQuotes = false;
    for (let i = 0; i < raw.length; i++) {
        if (raw[i] === '"') inQuotes = !inQuotes;
    }
    if (inQuotes) {
        console.error('Greška: CSV sadrži nezatvorene navodnike.');
        process.exit(1);
    }

    const { header, rows } = parseCSV(raw);

    // Require exactly 10 columns in header
    if (header.length !== 10) {
        console.error(`Greška: zaglavlje mora imati tačno 10 kolona, pronađeno ${header.length}.`);
        process.exit(1);
    }

    // Validate required columns
    const requiredCols = ['document_id', 'version_id', 'rule_id', 'finding_id', 'location_key', 'predicted_confidence', 'label', 'rule_version', 'reviewer', 'note'];
    for (const col of requiredCols) {
        if (!header.includes(col)) {
            console.error(`Greška: nedostaje obavezna kolona "${col}" u zaglavlju CSV-a.`);
            process.exit(1);
        }
    }

    const colIdx = {};
    header.forEach((col, i) => { colIdx[col] = i; });

    const records = [];
    let hasErrors = false;

    for (const { lineNumber, fields } of rows) {
        // Equal column count per row
        if (fields.length !== header.length) {
            console.error(`Greška (red ${lineNumber}): očekivano ${header.length} kolona, pronađeno ${fields.length}.`);
            hasErrors = true;
            continue;
        }

        // All fields except 'note' must be non-empty
        const noteIdx = colIdx['note'];
        for (let i = 0; i < fields.length; i++) {
            if (i === noteIdx) continue;
            if (!fields[i] || fields[i].trim().length === 0) {
                console.error(`Greška (red ${lineNumber}): kolona "${header[i]}" ne sme biti prazna.`);
                hasErrors = true;
                break;
            }
        }
        if (hasErrors && records.length === 0) continue;

        const label = fields[colIdx['label']];
        const confidence = fields[colIdx['predicted_confidence']];
        const ruleId = fields[colIdx['rule_id']];
        const ruleVersion = fields[colIdx['rule_version']];
        const versionId = fields[colIdx['version_id']];

        // Validate label is 0 or 1
        if (label !== '0' && label !== '1') {
            console.error(`Greška (red ${lineNumber}): label mora biti 0 ili 1, dobijeno "${label}".`);
            hasErrors = true;
            continue;
        }

        // Validate predicted_confidence is finite 0-1 (use Number not parseFloat)
        const conf = Number(confidence);
        if (!Number.isFinite(conf) || conf < 0 || conf > 1) {
            console.error(`Greška (red ${lineNumber}): predicted_confidence mora biti konačan broj 0-1, dobijeno "${confidence}".`);
            hasErrors = true;
            continue;
        }

        // Validate rule_id, rule_version, version_id not empty
        if (!ruleId || ruleId.trim().length === 0) {
            console.error(`Greška (red ${lineNumber}): rule_id ne sme biti prazan.`);
            hasErrors = true;
            continue;
        }
        if (!ruleVersion || ruleVersion.trim().length === 0) {
            console.error(`Greška (red ${lineNumber}): rule_version ne sme biti prazan.`);
            hasErrors = true;
            continue;
        }
        if (!versionId || versionId.trim().length === 0) {
            console.error(`Greška (red ${lineNumber}): version_id ne sme biti prazan.`);
            hasErrors = true;
            continue;
        }

        records.push({
            rule_id: ruleId.trim(),
            rule_version: ruleVersion.trim(),
            label: parseInt(label, 10),
            predicted_confidence: conf,
            finding_id: (fields[colIdx['finding_id']] || '').trim(),
        });
    }

    // If ANY invalid row, exit without generating report
    if (hasErrors) {
        console.error('Evaluacija prekinuta zbog grešaka u ulaznim podacima.');
        process.exit(1);
    }

    if (records.length === 0) {
        console.log('Nema obeleženih uzoraka. Dodajte redove u labels.csv za pokretanje evaluacije.');
        process.exit(0);
    }

    // Group by rule_id
    const byRule = {};
    for (const r of records) {
        if (!byRule[r.rule_id]) byRule[r.rule_id] = [];
        byRule[r.rule_id].push(r);
    }

    const ruleResults = {};
    const allLabels = [];
    const allConfidences = [];

    for (const [ruleId, recs] of Object.entries(byRule)) {
        const tp = recs.filter(r => r.label === 1).length;
        const fp = recs.filter(r => r.label === 0).length;
        const n = tp + fp;
        const precision = n > 0 ? tp / n : 0;
        const wilson = wilsonCI(tp, n);

        // Brier score
        const brier = recs.reduce((sum, r) => sum + Math.pow(r.predicted_confidence - r.label, 2), 0) / n;

        // Small sample warning
        let sampleWarning = null;
        if (n < 10) sampleWarning = 'Veoma mali uzorak - rezultati nepouzdani';
        else if (n < 30) sampleWarning = 'Mali uzorak - širok interval poverenja';

        // Collect rule_versions
        const rule_versions = [...new Set(recs.map(r => r.rule_version))];

        ruleResults[ruleId] = { tp, fp, n, precision, wilson, brier, sampleWarning, rule_versions };
        allLabels.push(...recs.map(r => r.label));
        allConfidences.push(...recs.map(r => r.predicted_confidence));
    }

    // Overall Brier score
    const overallBrier = allLabels.reduce((sum, label, i) =>
        sum + Math.pow(allConfidences[i] - label, 2), 0) / allLabels.length;

    // Reliability bins: [0,0.1), [0.1,0.2), ..., [0.9,1.0]
    const bins = [];
    for (let i = 0; i < 10; i++) {
        bins.push({ lower: i * 0.1, upper: (i + 1) * 0.1, items: [] });
    }
    for (let i = 0; i < allConfidences.length; i++) {
        const c = allConfidences[i];
        // 1.0 goes in last bin [0.9, 1.0]
        let binIdx = Math.min(9, Math.floor(c * 10));
        if (c === 1.0) binIdx = 9;
        bins[binIdx].items.push({ confidence: c, label: allLabels[i] });
    }

    const reliabilityBins = bins
        .filter(b => b.items.length > 0)
        .map(b => {
            const n = b.items.length;
            const meanPredicted = b.items.reduce((s, x) => s + x.confidence, 0) / n;
            const observedCorrect = b.items.filter(x => x.label === 1).length / n;
            const calibrationDiff = observedCorrect - meanPredicted;
            const brierContribution = b.items.reduce((s, x) =>
                s + Math.pow(x.confidence - x.label, 2), 0) / allLabels.length;
            return {
                range: `[${b.lower.toFixed(1)},${b.upper === 1.0 ? '1.0]' : b.upper.toFixed(1) + ')'}`,
                n, meanPredicted, observedCorrect, calibrationDiff, brierContribution
            };
        });

    // Generate reports
    if (!fs.existsSync(reportDir)) fs.mkdirSync(reportDir, { recursive: true });

    const jsonReport = {
        generated: new Date().toISOString(),
        disclaimer: 'Ova evaluacija meri preciznost prijavljenih nalaza. Ne meri recall, F1 niti greške koje alat nije pronašao.',
        correlation_warning: 'Nalazi iz istog dokumenta nisu statistički nezavisni. Wilson intervali po pojedinačnim nalazima mogu potceniti stvarnu neizvesnost.',
        future_note: 'Buduća evaluacija može koristiti bootstrap na nivou celog dokumenta.',
        total_samples: records.length,
        overall_brier: overallBrier,
        per_rule: ruleResults,
        reliability_bins: reliabilityBins,
    };

    const jsonPath = path.join(reportDir, 'confidence-report.json');
    fs.writeFileSync(jsonPath, JSON.stringify(jsonReport, null, 2));

    // Markdown report
    const md = [];
    md.push('# Confidence Evaluation Report');
    md.push('');
    md.push('> Ova evaluacija meri preciznost prijavljenih nalaza. Ne meri recall, F1 niti greške koje alat nije pronašao.');
    md.push('');
    md.push('> Nalazi iz istog dokumenta nisu statistički nezavisni. Wilson intervali po pojedinačnim nalazima mogu potceniti stvarnu neizvesnost.');
    md.push('');
    md.push('> Buduća evaluacija može koristiti bootstrap na nivou celog dokumenta.');
    md.push('');
    md.push(`**Ukupno obeleženih uzoraka:** ${records.length}`);
    md.push(`**Ukupni Brier score:** ${overallBrier.toFixed(4)}`);
    md.push('');
    md.push('## Per-rule results');
    md.push('');
    md.push('| Rule ID | TP | FP | N | Precision | Wilson CI | Width | Brier | Versions | Warning |');
    md.push('|---------|----|----|---|-----------|-----------|-------|-------|----------|---------|');
    for (const [ruleId, r] of Object.entries(ruleResults)) {
        md.push(`| ${ruleId} | ${r.tp} | ${r.fp} | ${r.n} | ${(r.precision * 100).toFixed(1)}% | [${r.wilson.lower.toFixed(3)}, ${r.wilson.upper.toFixed(3)}] | ${r.wilson.width.toFixed(3)} | ${r.brier.toFixed(4)} | ${r.rule_versions.join(', ')} | ${r.sampleWarning || '-'} |`);
    }
    md.push('');
    md.push('## Reliability diagram');
    md.push('');
    md.push('| Bin | N | Mean predicted | Observed correct | Calibration diff | Brier contribution |');
    md.push('|-----|---|----------------|-----------------|-----------------|-------------------|');
    for (const b of reliabilityBins) {
        md.push(`| ${b.range} | ${b.n} | ${b.meanPredicted.toFixed(3)} | ${b.observedCorrect.toFixed(3)} | ${b.calibrationDiff.toFixed(3)} | ${b.brierContribution.toFixed(4)} |`);
    }
    md.push('');

    const mdPath = path.join(reportDir, 'confidence-report.md');
    fs.writeFileSync(mdPath, md.join('\n'));

    // Console output
    console.log('\n=== Confidence Evaluation ===\n');
    console.log('Ova evaluacija meri preciznost prijavljenih nalaza. Ne meri recall, F1 niti greške koje alat nije pronašao.');
    console.log('Nalazi iz istog dokumenta nisu statistički nezavisni. Wilson intervali po pojedinačnim nalazima mogu potceniti stvarnu neizvesnost.');
    console.log('Buduća evaluacija može koristiti bootstrap na nivou celog dokumenta.\n');
    console.log(`Total samples: ${records.length}`);
    console.log(`Overall Brier: ${overallBrier.toFixed(4)}\n`);
    for (const [ruleId, r] of Object.entries(ruleResults)) {
        console.log(`  ${ruleId}: precision=${(r.precision * 100).toFixed(1)}% (${r.tp}/${r.n}) Wilson=[${r.wilson.lower.toFixed(3)},${r.wilson.upper.toFixed(3)}] Brier=${r.brier.toFixed(4)} versions=[${r.rule_versions.join(',')}]${r.sampleWarning ? ` [${r.sampleWarning}]` : ''}`);
    }
    console.log(`\nReports saved: ${jsonPath}, ${mdPath}`);
}

// ==========================================
// McNEMAR TEST - paired comparison of two rule versions
// ==========================================
function mcNemar(pairedResults) {
    // pairedResults: array of {before: 0|1, after: 0|1}
    // b = before correct, after wrong; c = before wrong, after correct
    let b = 0, c = 0;
    for (const p of pairedResults) {
        if (p.before === 1 && p.after === 0) b++;
        if (p.before === 0 && p.after === 1) c++;
    }
    const n = b + c;
    if (n === 0) return { chi2: 0, pValue: 1, b, c, n: pairedResults.length, significant: false };
    // McNemar chi-squared with continuity correction
    const chi2 = Math.pow(Math.abs(b - c) - 1, 2) / (b + c);
    // Approximate p-value from chi2 with 1 df using normal approximation
    const z = Math.sqrt(chi2);
    const pValue = 2 * (1 - normalCDF(z));
    return { chi2, pValue, b, c, n: pairedResults.length, significant: pValue < 0.05 };
}

function normalCDF(x) {
    const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429;
    const p = 0.3275911;
    const sign = x < 0 ? -1 : 1;
    x = Math.abs(x) / Math.sqrt(2);
    const t = 1.0 / (1.0 + p * x);
    const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
    return 0.5 * (1.0 + sign * y);
}

// ==========================================
// FDR (Benjamini-Hochberg) - controls false discovery rate
// ==========================================
function benjaminiHochberg(pValues, alpha) {
    // pValues: array of {rule_id, pValue}
    // Returns array with {rule_id, pValue, rank, critical, significant}
    alpha = alpha || 0.05;
    const sorted = pValues.slice().sort((a, b) => a.pValue - b.pValue);
    const m = sorted.length;
    const results = [];
    let lastSignificant = -1;
    for (let i = 0; i < m; i++) {
        const rank = i + 1;
        const critical = (rank / m) * alpha;
        if (sorted[i].pValue <= critical) lastSignificant = i;
        results.push({ ...sorted[i], rank, critical, significant: false });
    }
    // All up to lastSignificant are significant
    for (let i = 0; i <= lastSignificant; i++) {
        results[i].significant = true;
    }
    return results;
}

// ==========================================
// CLIFF'S DELTA - non-parametric effect size
// ==========================================
function cliffsDelta(group1, group2) {
    // group1, group2: arrays of numeric values
    // Returns delta in [-1, 1] and magnitude
    let more = 0, less = 0;
    for (const x of group1) {
        for (const y of group2) {
            if (x > y) more++;
            else if (x < y) less++;
        }
    }
    const n = group1.length * group2.length;
    if (n === 0) return { delta: 0, magnitude: 'negligible' };
    const delta = (more - less) / n;
    const absDelta = Math.abs(delta);
    let magnitude;
    if (absDelta < 0.147) magnitude = 'negligible';
    else if (absDelta < 0.33) magnitude = 'small';
    else if (absDelta < 0.474) magnitude = 'medium';
    else magnitude = 'large';
    return { delta, magnitude };
}

// ==========================================
// BOOTSTRAP BCa INTERVAL
// ==========================================
function bootstrapBCa(data, statFn, B, alpha) {
    B = B || 10000;
    alpha = alpha || 0.05;
    const n = data.length;
    if (n === 0) return { lower: 0, upper: 0, center: 0 };
    const observed = statFn(data);
    // Bootstrap resamples
    const bootStats = [];
    for (let i = 0; i < B; i++) {
        const sample = [];
        for (let j = 0; j < n; j++) sample.push(data[Math.floor(Math.random() * n)]);
        bootStats.push(statFn(sample));
    }
    bootStats.sort((a, b) => a - b);
    // Bias correction (z0)
    const below = bootStats.filter(s => s < observed).length;
    const z0 = normalInv(below / B);
    // Acceleration (a) via jackknife
    const jackValues = [];
    for (let i = 0; i < n; i++) {
        const jack = data.filter((_, j) => j !== i);
        jackValues.push(statFn(jack));
    }
    const jackMean = jackValues.reduce((s, v) => s + v, 0) / n;
    const num = jackValues.reduce((s, v) => s + Math.pow(jackMean - v, 3), 0);
    const den = jackValues.reduce((s, v) => s + Math.pow(jackMean - v, 2), 0);
    const a = den === 0 ? 0 : num / (6 * Math.pow(den, 1.5));
    // Adjusted percentiles
    const zAlpha = normalInv(alpha / 2);
    const zUpper = normalInv(1 - alpha / 2);
    const adjLower = normalCDF(z0 + (z0 + zAlpha) / (1 - a * (z0 + zAlpha)));
    const adjUpper = normalCDF(z0 + (z0 + zUpper) / (1 - a * (z0 + zUpper)));
    const lower = bootStats[Math.max(0, Math.floor(adjLower * B))] || bootStats[0];
    const upper = bootStats[Math.min(B - 1, Math.floor(adjUpper * B))] || bootStats[B - 1];
    return { lower, upper, center: observed };
}

function normalInv(p) {
    // Rational approximation for inverse normal CDF
    if (p <= 0) return -Infinity;
    if (p >= 1) return Infinity;
    if (p === 0.5) return 0;
    const a = [
        -3.969683028665376e+01, 2.209460984245205e+02,
        -2.759285104469687e+02, 1.383577518672690e+02,
        -3.066479806614716e+01, 2.506628277459239e+00
    ];
    const b = [
        -5.447609879822406e+01, 1.615858368580409e+02,
        -1.556989798598866e+02, 6.680131188771972e+01,
        -1.328068155288572e+01
    ];
    const c = [
        -7.784894002430293e-03, -3.223964580411365e-01,
        -2.400758277161838e+00, -2.549732539343734e+00,
        4.374664141464968e+00, 2.938163982698783e+00
    ];
    const d = [
        7.784695709041462e-03, 3.224671290700398e-01,
        2.445134137142996e+00, 3.754408661907416e+00
    ];
    const pLow = 0.02425, pHigh = 1 - pLow;
    let q, r;
    if (p < pLow) {
        q = Math.sqrt(-2 * Math.log(p));
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    } else if (p <= pHigh) {
        q = p - 0.5; r = q * q;
        return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1);
    } else {
        q = Math.sqrt(-2 * Math.log(1 - p));
        return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
    }
}

// Export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { parseCSV, splitCSVLine, wilsonCI, mcNemar, benjaminiHochberg, cliffsDelta, bootstrapBCa, run };
}

// Run if called directly
if (require.main === module) {
    run();
}
