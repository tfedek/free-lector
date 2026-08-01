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

        const canBeMarkedFinal = blockers === 0 && mandatory === 0;

        let finalAssessment;
        if (canBeMarkedFinal && verify === 0) {
            finalAssessment = 'Tekstualni audit je završen. Nema otvorenih obaveznih stavki.';
        } else if (canBeMarkedFinal) {
            finalAssessment = `Tekstualni audit je završen. Nema obaveznih ispravki, ali ${verify} stavki zahteva ručnu proveru.`;
        } else {
            finalAssessment = `Dokument ima ${mandatory} obaveznih ispravki i ${blockers} blokirajućih problema. Nije spreman za objavljivanje.`;
        }

        return {
            document: {
                name: docMap.name,
                document_id: generateDocId(docMap.name),
                version_id: new Date().toISOString(),
                language: options.houseStyle || 'sr-Latn',
                word_count: docMap.wordCount,
                paragraph_count: docMap.paragraphCount,
                table_count: docMap.tableCount,
                heading_count: docMap.headingCount,
                audit_mode: options.auditMode || 'FULL_AUDIT',
                render_engine: 'browser_deterministic',
            },
            scope: {
                proofreading: true,
                grammar: false, // No AI grammar check
                style: false,
                bibliography: options.bibliography || false,
                visual_layout: false,
                fact_checking: false,
                web_research: false,
                note: 'Samo determinističke provere (rule-based). Jezička analiza zahteva AI modul.',
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
    function downloadExcel(auditJson) {
        const wb = generateExcel(auditJson);
        const fileName = `${auditJson.document.name.replace(/\.[^.]+$/, '')}_audit.xlsx`;
        XLSX.writeFile(wb, fileName);
    }

    /**
     * Generate Markdown report
     */
    function generateMarkdown(auditJson) {
        const lines = [];
        const s = auditJson.summary;
        const d = auditJson.document;

        lines.push(`# Lektorsko-korektorski audit — ${d.name}`);
        lines.push('');
        lines.push(`**Izvor:** ${d.name}  `);
        lines.push(`**Verzija:** ${d.version_id}  `);
        lines.push(`**Režim:** ${d.audit_mode}  `);
        lines.push(`**Datum:** ${new Date().toLocaleDateString('sr-Latn-RS')}`);
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
    function downloadMarkdown(auditJson) {
        const md = generateMarkdown(auditJson);
        const fileName = `${auditJson.document.name.replace(/\.[^.]+$/, '')}_audit.md`;
        downloadTextFile(md, fileName, 'text/markdown');
    }

    /**
     * Download JSON file
     */
    function downloadJson(auditJson) {
        const json = JSON.stringify(auditJson, null, 2);
        const fileName = `${auditJson.document.name.replace(/\.[^.]+$/, '')}_audit.json`;
        downloadTextFile(json, fileName, 'application/json');
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
     * Generate a simple document ID from filename
     */
    function generateDocId(name) {
        return 'doc-' + name.replace(/[^a-zA-Z0-9]/g, '').substring(0, 16).toLowerCase();
    }

    // Public API
    return {
        buildAuditJson,
        downloadExcel,
        downloadMarkdown,
        downloadJson,
        generateMarkdown,
    };
})();
