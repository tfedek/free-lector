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

    const header = splitCSVLine(lines[0]);
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim().length === 0) continue;
        rows.push({ lineNumber: i + 1, fields: splitCSVLine(lines[i]) });
    }
    return { header, rows };
}

function splitCSVLine(line) {
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
    const requiredCols = ['rule_id', 'rule_version', 'version_id', 'label', 'predicted_confidence', 'finding_id'];
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
        if (n < 10) sampleWarning = 'Veoma mali uzorak — rezultati nepouzdani';
        else if (n < 30) sampleWarning = 'Mali uzorak — širok interval poverenja';

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

// Export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { parseCSV, splitCSVLine, wilsonCI, run };
}

// Run if called directly
if (require.main === module) {
    run();
}
