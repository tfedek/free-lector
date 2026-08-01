/**
 * Exporter Module
 * Generates Excel (XLSX), Markdown, and JSON reports
 */

const Exporter = (() => {
    'use strict';

    /**
     * Build the full audit JSON object
     */
    function buildAuditJson(docMap, findings, passedChecks, options) {
        const blockers = findings.filter(f => f.priority === 'BLOCKER').length;
        const mandatory = findings.filter(f => f.priority === 'OBAVEZNO').length;
        const verify = findings.filter(f => f.priority === 'PROVERITI').length;
        const recommendations = findings.filter(f => f.priority === 'PREPORUKA').length;

        const requiredChecksComplete =
            options.aiGrammar === true &&
            options.visualLayout === true;

        const canBeMarkedFinal =
            requiredChecksComplete &&
            blockers === 0 &&
            mandatory === 0 &&
            verify === 0;

        let finalAssessment;
        if (canBeMarkedFinal) {
            finalAssessment = 'Audit je završen. Sve provere su prošle bez nalaza.';
        } else if (blockers === 0 && mandatory === 0 && verify === 0) {
            finalAssessment = 'Determinističke provere su završene. Gramatika, stil i vizuelni prelom nisu provereni.';
        } else {
            finalAssessment = `Dokument ima ${mandatory} obaveznih ispravki i ${blockers} blokirajućih problema. Nije spreman za objavljivanje.`;
        }

        return {
            document: {
                name: docMap.name,
                document_id: generateDocId(docMap.name, docMap.rawText),
                version_id: new Date().toISOString(),
                language: options.houseStyle || 'sr-Latn',
                word_count: docMap.wordCount,
                paragraph_count: docMap.paragraphCount,
                table_count: docMap.tableCount,
                heading_count: docMap.headingCount,
                audit_mode: options.auditMode || 'FULL_AUDIT',
                render_engine: null,
            },
            scope: {
                proofreading: true,
                grammar: options.aiGrammar === true,
                style: options.aiStyle === true,
                bibliography: options.bibliography || false,
                visual_layout: options.visualLayout === true,
                fact_checking: false,
                web_research: false,
                note: 'Samo determinističke provere (rule-based). Jezička analiza zahteva AI modul.',
            },
            audit_status: {
                status: canBeMarkedFinal ? 'POTPUN' : 'DELIMIČAN',
                linguistic_analysis: options.aiGrammar === true ? 'IZVRŠENA' : 'NIJE IZVRŠENA',
                visual_review: options.visualLayout === true ? 'IZVRŠEN' : 'NIJE IZVRŠEN',
            },
            summary: {
                total_finding_categories: new Set(findings.map(f => f.category)).size,
                total_occurrences: findings.length,
                blockers,
                mandatory,
                verify,
                recommendations,
                passed_checks: passedChecks.length,
                can_be_marked_final: canBeMarkedFinal,
                final_assessment: finalAssessment,
            },
            findings,
            passed_checks: passedChecks,
            global_patterns: extractGlobalPatterns(findings),
        };
    }

    /**
     * Generate Excel workbook with 3 sheets
     */
    function generateExcel(auditJson) {
        const wb = XLSX.utils.book_new();

        // SHEET 1: Summary
        const summaryData = [
            ['LEKTORSKI AUDIT — SAŽETAK'],
            [],
            ['Naziv dokumenta', auditJson.document.name],
            ['ID verzije', auditJson.document.version_id],
            ['Datum audita', new Date().toLocaleDateString('sr-Latn-RS')],
            ['Režim', auditJson.document.audit_mode],
            ['Broj reči', auditJson.document.word_count],
            ['Broj pasusa', auditJson.document.paragraph_count],
            ['Broj tabela', auditJson.document.table_count],
            ['Broj naslova', auditJson.document.heading_count],
            [],
            ['STATUS AUDITA'],
            ['Status audita', auditJson.audit_status.status],
            ['Jezička analiza', auditJson.audit_status.linguistic_analysis],
            ['Vizuelni pregled', auditJson.audit_status.visual_review],
            [],
            ['REZULTATI'],
            ['Ukupno nalaza', auditJson.summary.total_occurrences],
            ['Blocker', auditJson.summary.blockers],
            ['Obavezno', auditJson.summary.mandatory],
            ['Proveriti', auditJson.summary.verify],
            ['Preporuke', auditJson.summary.recommendations],
            ['Provere bez grešaka', auditJson.summary.passed_checks],
            [],
            ['Završna procena', auditJson.summary.final_assessment],
            [],
            ['NAPOMENA'],
            ['Ovaj audit je izvršen isključivo determinističkim (rule-based) proverama.'],
            ['Jezička, stilska i sadržajna analiza nisu uključene (zahtevaju AI modul).'],
        ];
        const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
        ws1['!cols'] = [{ wch: 25 }, { wch: 60 }];
        XLSX.utils.book_append_sheet(wb, ws1, 'Sažetak');

        // SHEET 2: Findings
        const findingsHeader = [
            'Br.', 'ID', 'Odeljak', 'Pasus/ID', 'Kategorija', 'Prioritet',
            'Pouzdanost', 'Original', 'Predložena ispravka', 'Obrazloženje',
            'Direktan citat', 'Provera izvora', 'Automatski primenljivo',
            'Broj pojavljivanja', 'Status', 'Napomena korisnika'
        ];
        const findingsRows = auditJson.findings.map((f, i) => [
            i + 1,
            f.id,
            f.section,
            f.paragraphId,
            f.category,
            f.priority,
            f.confidence,
            f.original,
            f.replacement,
            f.rationale,
            f.isDirectQuote ? 'Da' : 'Ne',
            f.requiresSourceVerification ? 'Da' : 'Ne',
            f.autoFixable ? 'Da' : 'Ne',
            f.globalPattern ? 'Globalno' : '1',
            f.status,
            '',
        ]);

        const ws2 = XLSX.utils.aoa_to_sheet([findingsHeader, ...findingsRows]);
        ws2['!cols'] = [
            { wch: 4 }, { wch: 8 }, { wch: 20 }, { wch: 10 }, { wch: 15 }, { wch: 12 },
            { wch: 8 }, { wch: 40 }, { wch: 40 }, { wch: 35 },
            { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 20 },
        ];
        XLSX.utils.book_append_sheet(wb, ws2, 'Nalazi');

        // SHEET 3: Passed checks
        const passedHeader = ['Oblast provere', 'Rezultat', 'Proveravanih elemenata', 'Napomena'];
        const passedRows = auditJson.passed_checks.map(p => [
            p.area, p.result, p.count, ''
        ]);
        const ws3 = XLSX.utils.aoa_to_sheet([passedHeader, ...passedRows]);
        ws3['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 20 }, { wch: 30 }];
        XLSX.utils.book_append_sheet(wb, ws3, 'Provere bez grešaka');

        return wb;
    }

    /**
     * Download Excel file
     */
    function downloadExcel(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        const wb = generateExcel(filtered);
        const fileName = `${filtered.document.name.replace(/\.[^.]+$/, '')}_audit.xlsx`;
        XLSX.writeFile(wb, fileName);
    }

    /**
     * Generate Markdown report
     */
    function generateMarkdown(auditJson) {
        const lines = [];
        const s = auditJson.summary;
        const d = auditJson.document;
        const st = auditJson.audit_status;

        lines.push(`# Lektorsko-korektorski audit — ${d.name}`);
        lines.push('');
        lines.push(`**Izvor:** ${d.name}  `);
        lines.push(`**Verzija:** ${d.version_id}  `);
        lines.push(`**Režim:** ${d.audit_mode}  `);
        lines.push(`**Datum:** ${new Date().toLocaleDateString('sr-Latn-RS')}`);
        lines.push('');
        lines.push('## Status audita');
        lines.push('');
        lines.push(`- Status audita: **${st.status}**`);
        lines.push(`- Jezička analiza: **${st.linguistic_analysis}**`);
        lines.push(`- Vizuelni pregled: **${st.visual_review}**`);
        lines.push('');

        lines.push('## Opseg');
        lines.push('');
        lines.push('Izvršene su determinističke (rule-based) provere:');
        lines.push('zagrade, navodnici, razmaci, interpunkcija, mešanje pisama, grčki citati, duple reči, numeracija, crtice, bibliografija, URL-ovi, fusnote, ponavljanja, ALL-CAPS, struktura naslova.');
        lines.push('');
        lines.push('**Nije proveravano:** jezička analiza (gramatika, stil, sintaksa), vizuelni prelom, faktografska tačnost.');
        lines.push('');

        lines.push('## Sažetak');
        lines.push('');
        lines.push(`- Ukupno nalaza: **${s.total_occurrences}**`);
        lines.push(`- Obavezne ispravke: **${s.mandatory}**`);
        lines.push(`- Za proveru: **${s.verify}**`);
        lines.push(`- Preporuke: **${s.recommendations}**`);
        lines.push(`- Blocker: **${s.blockers}**`);
        lines.push(`- Provere bez grešaka: **${s.passed_checks}**`);
        lines.push(`- Dokument ${s.can_be_marked_final ? 'MOŽE' : 'NE MOŽE'} biti označen kao finalan`);
        lines.push('');

        // Critical interventions
        const critical = auditJson.findings.filter(f => f.priority === 'BLOCKER' || f.priority === 'OBAVEZNO');
        if (critical.length > 0) {
            lines.push('## Kritične intervencije');
            lines.push('');
            critical.slice(0, 20).forEach((f, i) => {
                lines.push(`${i + 1}. **[${f.category}]** ${f.section} — ${f.rationale}`);
            });
            if (critical.length > 20) {
                lines.push(`\n... i još ${critical.length - 20} obaveznih stavki (videti Excel).`);
            }
            lines.push('');
        }

        // Detailed findings
        lines.push('## Detaljni nalazi');
        lines.push('');
        auditJson.findings.forEach((f, i) => {
            lines.push(`### ${i + 1}. [${f.section}] — ${f.category} (${f.priority})`);
            lines.push('');
            lines.push(`**Original/problem:**  `);
            lines.push(`\`${f.original}\``);
            lines.push('');
            lines.push(`**Predložena ispravka:**  `);
            lines.push(`\`${f.replacement}\``);
            lines.push('');
            lines.push(`**Obrazloženje:** ${f.rationale}  `);
            lines.push(`**Pouzdanost:** ${f.confidence}  `);
            lines.push(`**Automatski primenljivo:** ${f.autoFixable ? 'Da' : 'Ne'}  `);
            lines.push(`**Provera izvora potrebna:** ${f.requiresSourceVerification ? 'Da' : 'Ne'}`);
            lines.push('');
        });

        // Passed checks
        lines.push('## Provere bez pronađenih grešaka');
        lines.push('');
        auditJson.passed_checks.forEach(p => {
            lines.push(`- ${p.area}: ${p.result} (${p.count} elemenata provereno)`);
        });
        lines.push('');

        // Final assessment
        lines.push('## Završna procena');
        lines.push('');
        lines.push(s.final_assessment);
        lines.push('');
        lines.push('---');
        lines.push('*Generisano alatom Free Lector (rule-based, bez AI). Jezička analiza nije izvršena.*');

        return lines.join('\n');
    }

    /**
     * Download Markdown file
     */
    function downloadMarkdown(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        const md = generateMarkdown(filtered);
        const fileName = `${filtered.document.name.replace(/\.[^.]+$/, '')}_audit.md`;
        downloadTextFile(md, fileName, 'text/markdown');
    }

    /**
     * Download JSON file
     */
    function downloadJson(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        const json = JSON.stringify(filtered, null, 2);
        const fileName = `${filtered.document.name.replace(/\.[^.]+$/, '')}_audit.json`;
        downloadTextFile(json, fileName, 'application/json');
    }

    /**
     * Apply export filter (all findings or only open)
     * Creates a deep copy and recalculates summary/global_patterns
     */
    function applyExportFilter(auditJson, exportFilter) {
        if (!exportFilter || exportFilter === 'all') return auditJson;

        const filtered = structuredClone(auditJson);
        filtered.findings = filtered.findings.filter(f => f.status === 'OPEN');

        // Recalculate summary
        const openFindings = filtered.findings;
        const s = filtered.summary;
        s.blockers = openFindings.filter(f => f.priority === 'BLOCKER').length;
        s.mandatory = openFindings.filter(f => f.priority === 'OBAVEZNO').length;
        s.verify = openFindings.filter(f => f.priority === 'PROVERITI').length;
        s.recommendations = openFindings.filter(f => f.priority === 'PREPORUKA').length;
        s.total_occurrences = openFindings.length;
        s.total_finding_categories = new Set(openFindings.map(f => f.category)).size;

        // Recalculate final gate
        const scope = filtered.scope;
        const requiredComplete = scope.grammar === true && scope.visual_layout === true;
        s.can_be_marked_final = requiredComplete &&
            s.blockers === 0 && s.mandatory === 0 && s.verify === 0;

        if (s.can_be_marked_final) {
            s.final_assessment = 'Audit je završen. Sve provere su prošle bez nalaza.';
        } else if (s.blockers === 0 && s.mandatory === 0 && s.verify === 0) {
            s.final_assessment = 'Determinističke provere su završene. Gramatika, stil i vizuelni prelom nisu provereni.';
        } else {
            s.final_assessment = `Dokument ima ${s.mandatory} obaveznih ispravki i ${s.blockers} blokirajućih problema. Nije spreman za objavljivanje.`;
        }

        // Update audit_status.status
        filtered.audit_status.status = s.can_be_marked_final ? 'POTPUN' : 'DELIMIČAN';

        // Recalculate global patterns
        filtered.global_patterns = extractGlobalPatterns(openFindings);

        return filtered;
    }

    /**
     * Helper: trigger text file download
     */
    function downloadTextFile(content, fileName, mimeType) {
        const blob = new Blob([content], { type: mimeType + ';charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /**
     * Extract global patterns from findings
     */
    function extractGlobalPatterns(findings) {
        const categoryCount = {};
        for (const f of findings) {
            const key = `${f.category}::${f.rationale}`;
            if (!categoryCount[key]) {
                categoryCount[key] = { category: f.category, rationale: f.rationale, count: 0 };
            }
            categoryCount[key].count++;
        }
        return Object.values(categoryCount).filter(p => p.count >= 3)
            .sort((a, b) => b.count - a.count);
    }

    /**
     * Generate document ID from full content using FNV-1a hash.
     * Hashes entire rawText for uniqueness (not truncated).
     */
    function generateDocId(name, rawText) {
        const input = (name || '') + '::' + (rawText || '');
        // FNV-1a 64-bit emulated with two 32-bit halves for better distribution
        let h1 = 0x811c9dc5;
        let h2 = 0x01000193;
        for (let i = 0; i < input.length; i++) {
            const c = input.charCodeAt(i);
            h1 ^= c;
            h1 = Math.imul(h1, 0x01000193);
            h2 ^= (c ^ 0x5f);
            h2 = Math.imul(h2, 0x01000193);
        }
        const part1 = (h1 >>> 0).toString(36);
        const part2 = (h2 >>> 0).toString(36);
        return 'doc-' + part1 + part2;
    }

    // Public API
    return {
        buildAuditJson,
        downloadExcel,
        downloadMarkdown,
        downloadJson,
        generateMarkdown,
        applyExportFilter,
    };
})();

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = Exporter;
}
