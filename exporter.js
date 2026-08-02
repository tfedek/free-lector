/**
 * Exporter Module — Round 4
 * Dynamic scope, configurable requiredCapabilities, FNV-based content ID,
 * structuredClone fallback, table cell columns in Excel/Markdown
 */

const Exporter = (() => {
    'use strict';

    // Production fallback for structuredClone
    function deepClone(value) {
        if (typeof structuredClone === 'function') return structuredClone(value);
        return JSON.parse(JSON.stringify(value));
    }

    // Configurable required capabilities for final gate
    const DEFAULT_REQUIRED_CAPABILITIES = ['grammar', 'visual_layout', 'style'];

    /**
     * Build the full audit JSON object
     */
    function buildAuditJson(docMap, findings, passedChecks, options) {
        const blockers = findings.filter(f => f.priority === 'BLOCKER').length;
        const mandatory = findings.filter(f => f.priority === 'OBAVEZNO').length;
        const verify = findings.filter(f => f.priority === 'PROVERITI').length;
        const recommendations = findings.filter(f => f.priority === 'PREPORUKA').length;

        // Dynamic scope from actual options
        const scope = {
            proofreading: hasAnyProofreadingCheck(options),
            grammar: options.aiGrammar === true,
            style: options.aiStyle === true,
            bibliography: options.bibliography === true,
            visual_layout: options.visualLayout === true,
            fact_checking: false,
            web_research: false,
        };

        // Dynamic scope.note
        const noteParts = [];
        if (scope.proofreading) noteParts.push('determinističke provere');
        if (scope.grammar) noteParts.push('gramatička analiza');
        if (scope.style) noteParts.push('stilska analiza');
        if (scope.visual_layout) noteParts.push('vizuelni pregled');
        const missingParts = [];
        if (!scope.grammar) missingParts.push('gramatika');
        if (!scope.style) missingParts.push('stil');
        if (!scope.visual_layout) missingParts.push('vizuelni prelom');
        scope.note = `Izvršeno: ${noteParts.join(', ') || 'ništa'}.` +
            (missingParts.length > 0 ? ` Nije izvršeno: ${missingParts.join(', ')}.` : '');

        // Configurable final gate
        const requiredCaps = options.requiredCapabilities || DEFAULT_REQUIRED_CAPABILITIES;
        const requiredComplete = requiredCaps.every(cap => scope[cap] === true);

        const blockersOpen = findings.filter(f => f.priority === 'BLOCKER' && f.status === 'OPEN').length;
        const mandatoryOpen = findings.filter(f => f.priority === 'OBAVEZNO' && f.status === 'OPEN').length;
        const verifyOpen = findings.filter(f => f.priority === 'PROVERITI' && f.status === 'OPEN').length;

        const canBeMarkedFinal = requiredComplete &&
            blockersOpen === 0 && mandatoryOpen === 0 && verifyOpen === 0;

        let finalAssessment;
        if (canBeMarkedFinal) {
            finalAssessment = 'Audit je završen. Sve provere su prošle bez nalaza.';
        } else if (blockersOpen === 0 && mandatoryOpen === 0 && verifyOpen === 0) {
            finalAssessment = `Determinističke provere su završene. Nedostaje: ${missingParts.join(', ') || 'ništa'}.`;
        } else {
            finalAssessment = `Dokument ima ${mandatoryOpen} obaveznih, ${blockersOpen} blokirajućih i ${verifyOpen} za proveru. Nije spreman za objavljivanje.`;
        }

        // FNV-based content ID (content only, no filename)
        const documentId = generateDocId(docMap.rawText || '');
        // Content + time version_id
        const versionId = generateVersionId(docMap.rawText || '');


        return {
            document: {
                name: docMap.name,
                document_id: documentId,
                version_id: versionId,
                language: options.houseStyle || 'sr-Latn',
                word_count: docMap.wordCount,
                paragraph_count: docMap.paragraphCount,
                table_count: docMap.tableCount,
                heading_count: docMap.headingCount,
                audit_mode: options.auditMode || 'FULL_AUDIT',
                render_engine: null,
            },
            scope,
            audit_status: {
                status: canBeMarkedFinal ? 'POTPUN' : 'DELIMIČAN',
                linguistic_analysis: scope.grammar ? 'IZVRŠENA' : 'NIJE IZVRŠENA',
                style_analysis: scope.style ? 'IZVRŠENA' : 'NIJE IZVRŠENA',
                visual_review: scope.visual_layout ? 'IZVRŠEN' : 'NIJE IZVRŠEN',
            },
            summary: {
                total_finding_categories: new Set(findings.map(f => f.category)).size,
                total_occurrences: findings.length,
                blockers, mandatory, verify, recommendations,
                passed_checks: passedChecks.length,
                by_status: {
                    open: findings.filter(f => f.status === 'OPEN').length,
                    done: findings.filter(f => f.status === 'DONE').length,
                    rejected: findings.filter(f => f.status === 'REJECTED').length,
                },
                mandatory_open: findings.filter(f => f.priority === 'OBAVEZNO' && f.status === 'OPEN').length,
                mandatory_total: mandatory,
                blockers_open: findings.filter(f => f.priority === 'BLOCKER' && f.status === 'OPEN').length,
                blockers_total: blockers,
                verify_open: findings.filter(f => f.priority === 'PROVERITI' && f.status === 'OPEN').length,
                verify_total: verify,
                can_be_marked_final: canBeMarkedFinal,
                final_assessment: finalAssessment,
            },
            findings,
            passed_checks: passedChecks,
            global_patterns: extractGlobalPatterns(findings),
            required_capabilities: requiredCaps,
            processing_coverage: docMap.processingCoverage || { supported: [], partial: [], unsupported: [] },
        };
    }

    /**
     * Determine if any proofreading check is active
     */
    function hasAnyProofreadingCheck(options) {
        const proofKeys = ['brackets','quotes','markdown','spacing','scriptMix','greek',
            'duplicates','toc','numbering','bibliography','urls','footnotes',
            'repetition','capsWords','emptyHeadings'];
        return proofKeys.some(k => options[k] === true);
    }


    /**
     * Generate Excel workbook — includes table cell columns
     */
    function generateExcel(auditJson) {
        const wb = XLSX.utils.book_new();

        // SHEET 1: Summary
        const summaryData = [
            ['LEKTORSKI AUDIT — SAŽETAK'],
            [],
            ['Naziv dokumenta', auditJson.document.name],
            ['Document ID', auditJson.document.document_id],
            ['Version ID', auditJson.document.version_id],
            ['Datum audita', new Date().toLocaleDateString('sr-Latn-RS')],
            ['Režim', auditJson.document.audit_mode],
            ['Broj reči', auditJson.document.word_count],
            ['Broj pasusa', auditJson.document.paragraph_count],
            ['Broj tabela', auditJson.document.table_count],
            ['Broj naslova', auditJson.document.heading_count],
            [],
            ['STATUS AUDITA'],
            ['Status', auditJson.audit_status.status],
            ['Jezička analiza', auditJson.audit_status.linguistic_analysis],
            ['Stilska analiza', auditJson.audit_status.style_analysis],
            ['Vizuelni pregled', auditJson.audit_status.visual_review],
            [],
            ['OPSEG'],
            ['Napomena', auditJson.scope.note],
            [],
            ['REZULTATI'],
            ['Ukupno nalaza', auditJson.summary.total_occurrences],
            ['Blocker', auditJson.summary.blockers],
            ['Obavezno', auditJson.summary.mandatory],
            ['Proveriti', auditJson.summary.verify],
            ['Preporuke', auditJson.summary.recommendations],
            ['Provere bez grešaka', auditJson.summary.passed_checks],
            [],
            ['STATUS PO NALAZIMA'],
            ['Otvoreno', auditJson.summary.by_status ? auditJson.summary.by_status.open : ''],
            ['Rešeno', auditJson.summary.by_status ? auditJson.summary.by_status.done : ''],
            ['Nije greška', auditJson.summary.by_status ? auditJson.summary.by_status.rejected : ''],
            ['Obavezno otvoreno', auditJson.summary.mandatory_open != null ? auditJson.summary.mandatory_open : ''],
            ['Obavezno ukupno', auditJson.summary.mandatory_total != null ? auditJson.summary.mandatory_total : ''],
            ['Blocker otvoreno', auditJson.summary.blockers_open != null ? auditJson.summary.blockers_open : ''],
            ['Blocker ukupno', auditJson.summary.blockers_total != null ? auditJson.summary.blockers_total : ''],
            ['Proveriti otvoreno', auditJson.summary.verify_open != null ? auditJson.summary.verify_open : ''],
            ['Proveriti ukupno', auditJson.summary.verify_total != null ? auditJson.summary.verify_total : ''],
            [],
            ['Završna procena', auditJson.summary.final_assessment],
        ];
        const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
        ws1['!cols'] = [{ wch: 25 }, { wch: 70 }];
        XLSX.utils.book_append_sheet(wb, ws1, 'Sažetak');

        // SHEET 2: Findings — with Table ID, Row ID, Cell ID, Red, Kolona columns
        const findingsHeader = [
            'Br.', 'ID', 'Odeljak', 'Pasus/ID', 'Kategorija', 'Prioritet',
            'Pouzdanost', 'Original', 'Predložena ispravka', 'Obrazloženje',
            'Direktan citat', 'Provera izvora', 'Auto-fix',
            'Globalno', 'Status',
            'Table ID', 'Row ID', 'Cell ID', 'Red', 'Kolona',
            'Napomena'
        ];
        const findingsRows = auditJson.findings.map((f, i) => [
            i + 1, f.id, f.section, f.paragraphId, f.category, f.priority,
            f.confidence, f.original, f.replacement, f.rationale,
            f.isDirectQuote ? 'Da' : 'Ne',
            f.requiresSourceVerification ? 'Da' : 'Ne',
            f.autoFixable ? 'Da' : 'Ne',
            f.globalPattern ? 'Globalno' : '1',
            f.status,
            f.tableId || '', f.rowId || '', f.cellId || '',
            f.rowIndex != null ? f.rowIndex : '',
            f.columnIndex != null ? f.columnIndex : '',
            '',
        ]);
        const ws2 = XLSX.utils.aoa_to_sheet([findingsHeader, ...findingsRows]);
        ws2['!cols'] = [
            {wch:4},{wch:8},{wch:20},{wch:10},{wch:15},{wch:12},{wch:6},{wch:35},{wch:35},{wch:30},
            {wch:6},{wch:6},{wch:6},{wch:6},{wch:10},{wch:12},{wch:12},{wch:12},{wch:4},{wch:4},{wch:20},
        ];
        XLSX.utils.book_append_sheet(wb, ws2, 'Nalazi');

        // SHEET 3: Passed checks
        const passedHeader = ['Oblast provere', 'Rezultat', 'Proveravanih elemenata'];
        const passedRows = auditJson.passed_checks.map(p => [p.area, p.result, p.count]);
        const ws3 = XLSX.utils.aoa_to_sheet([passedHeader, ...passedRows]);
        ws3['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 20 }];
        XLSX.utils.book_append_sheet(wb, ws3, 'Provere bez grešaka');

        return wb;
    }


    function downloadExcel(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        const wb = generateExcel(filtered);
        const fileName = `${filtered.document.name.replace(/\.[^.]+$/, '')}_audit.xlsx`;
        XLSX.writeFile(wb, fileName);
    }

    /**
     * Markdown report — includes table cell info for cell-level findings
     */
    function generateMarkdown(auditJson) {
        const lines = [];
        const s = auditJson.summary;
        const d = auditJson.document;
        const st = auditJson.audit_status;

        lines.push(`# Lektorsko-korektorski audit — ${d.name}`);
        lines.push('');
        lines.push(`**Document ID:** ${d.document_id}  `);
        lines.push(`**Version ID:** ${d.version_id}  `);
        lines.push(`**Režim:** ${d.audit_mode}  `);
        lines.push(`**Datum:** ${new Date().toLocaleDateString('sr-Latn-RS')}`);
        lines.push('');
        lines.push('## Status audita');
        lines.push('');
        lines.push(`- Status: **${st.status}**`);
        lines.push(`- Jezička analiza: **${st.linguistic_analysis}**`);
        lines.push(`- Stilska analiza: **${st.style_analysis}**`);
        lines.push(`- Vizuelni pregled: **${st.visual_review}**`);
        lines.push('');
        lines.push(`> ${auditJson.scope.note}`);
        lines.push('');
        lines.push('## Sažetak');
        lines.push('');
        lines.push(`- Ukupno nalaza: **${s.total_occurrences}**`);
        lines.push(`- Obavezne: **${s.mandatory}** | Proveriti: **${s.verify}** | Preporuke: **${s.recommendations}** | Blocker: **${s.blockers}**`);
        lines.push(`- Provere bez grešaka: **${s.passed_checks}**`);
        if (s.by_status) lines.push(`- Otvoreno: **${s.by_status.open}** | Rešeno: **${s.by_status.done}** | Nije greška: **${s.by_status.rejected}**`);
        if (s.mandatory_open != null) lines.push(`- Obavezno: **${s.mandatory_open}**/${s.mandatory_total} | Blocker: **${s.blockers_open != null ? s.blockers_open : s.blockers}**/${s.blockers_total != null ? s.blockers_total : s.blockers} | Proveriti: **${s.verify_open != null ? s.verify_open : s.verify}**/${s.verify_total != null ? s.verify_total : s.verify} *(otvoreno/ukupno)*`);
        lines.push(`- Finalan: ${s.can_be_marked_final ? '**DA**' : '**NE**'}`);
        lines.push('');

        lines.push('## Nalazi');
        lines.push('');
        auditJson.findings.forEach((f, i) => {
            lines.push(`### ${i+1}. [${f.section}] ${f.category} (${f.priority})`);
            if (f.tableId) {
                lines.push(`> Tabela: ${f.tableId} | Red: ${f.rowIndex} | Kolona: ${f.columnIndex} | Cell ID: ${f.cellId}`);
            }
            lines.push(`- **Original:** \`${f.original}\``);
            lines.push(`- **Ispravka:** \`${f.replacement}\``);
            lines.push(`- ${f.rationale} (pouzdanost: ${Math.round(f.confidence*100)}%)`);
            lines.push('');
        });

        lines.push('## Provere bez grešaka');
        lines.push('');
        auditJson.passed_checks.forEach(p => lines.push(`- ${p.area} (${p.count})`));
        lines.push('');

        // Processing coverage
        if (auditJson.processing_coverage) {
            const pc = auditJson.processing_coverage;
            lines.push('## Pokrivenost obrade');
            lines.push('');
            if (pc.supported.length > 0) lines.push(`- **Obrađeno:** ${pc.supported.join(', ')}`);
            if (pc.partial.length > 0) lines.push(`- **Delimično obrađeno:** ${pc.partial.join(', ')}`);
            if (pc.unsupported.length > 0) lines.push(`- **Nije obrađeno:** ${pc.unsupported.join(', ')}`);
            lines.push('');
        }

        lines.push('## Završna procena');
        lines.push('');
        lines.push(s.final_assessment);
        lines.push('');
        lines.push('---');
        lines.push(`*Free Lector (rule-based). ${auditJson.scope.note}*`);
        return lines.join('\n');
    }

    function downloadMarkdown(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        const md = generateMarkdown(filtered);
        downloadTextFile(md, `${filtered.document.name.replace(/\.[^.]+$/,'')}_audit.md`, 'text/markdown');
    }

    function downloadJson(auditJson, exportFilter) {
        const filtered = applyExportFilter(auditJson, exportFilter);
        downloadTextFile(JSON.stringify(filtered,null,2), `${filtered.document.name.replace(/\.[^.]+$/,'')}_audit.json`, 'application/json');
    }


    /**
     * Apply export filter with full recalculation
     */
    function applyExportFilter(auditJson, exportFilter) {
        // Always deep clone to avoid exporting mutated in-memory state
        const filtered = deepClone(auditJson);

        if (exportFilter === 'open') {
            filtered.findings = filtered.findings.filter(f => f.status === 'OPEN');
        }
        // else 'all': keep all findings as-is

        // Recalculate summary from the exported findings set
        const targetFindings = filtered.findings;
        const s = filtered.summary;
        s.blockers = targetFindings.filter(f => f.priority === 'BLOCKER').length;
        s.mandatory = targetFindings.filter(f => f.priority === 'OBAVEZNO').length;
        s.verify = targetFindings.filter(f => f.priority === 'PROVERITI').length;
        s.recommendations = targetFindings.filter(f => f.priority === 'PREPORUKA').length;
        s.total_occurrences = targetFindings.length;
        s.total_finding_categories = new Set(targetFindings.map(f => f.category)).size;
        s.by_status = {
            open: targetFindings.filter(f => f.status === 'OPEN').length,
            done: targetFindings.filter(f => f.status === 'DONE').length,
            rejected: targetFindings.filter(f => f.status === 'REJECTED').length,
        };
        s.mandatory_open = targetFindings.filter(f => f.priority === 'OBAVEZNO' && f.status === 'OPEN').length;
        s.mandatory_total = targetFindings.filter(f => f.priority === 'OBAVEZNO').length;
        s.blockers_open = targetFindings.filter(f => f.priority === 'BLOCKER' && f.status === 'OPEN').length;
        s.blockers_total = targetFindings.filter(f => f.priority === 'BLOCKER').length;
        s.verify_open = targetFindings.filter(f => f.priority === 'PROVERITI' && f.status === 'OPEN').length;
        s.verify_total = targetFindings.filter(f => f.priority === 'PROVERITI').length;

        const scope = filtered.scope;
        const reqCaps = filtered.required_capabilities || DEFAULT_REQUIRED_CAPABILITIES;
        const reqComplete = reqCaps.every(cap => scope[cap] === true);

        // Final status always computed from OPEN findings only
        const openFindings = targetFindings.filter(f => f.status === 'OPEN');
        const openBlockers = openFindings.filter(f => f.priority === 'BLOCKER').length;
        const openMandatory = openFindings.filter(f => f.priority === 'OBAVEZNO').length;
        const openVerify = openFindings.filter(f => f.priority === 'PROVERITI').length;
        s.can_be_marked_final = reqComplete && openBlockers === 0 && openMandatory === 0 && openVerify === 0;

        if (s.can_be_marked_final) s.final_assessment = 'Audit završen. Sve provere prošle.';
        else if (openBlockers === 0 && openMandatory === 0 && openVerify === 0) s.final_assessment = 'Determinističke provere završene. Nedostaju obavezne sposobnosti.';
        else s.final_assessment = `${openMandatory} obaveznih, ${openBlockers} blokirajućih i ${openVerify} za proveru. Nije spreman.`;

        filtered.audit_status.status = s.can_be_marked_final ? 'POTPUN' : 'DELIMIČAN';
        filtered.global_patterns = extractGlobalPatterns(targetFindings);
        return filtered;
    }

    function downloadTextFile(content, fileName, mimeType) {
        const blob = new Blob([content], { type: mimeType + ';charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = fileName;
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }

    function extractGlobalPatterns(findings) {
        const cc = {};
        for (const f of findings) {
            const k = `${f.category}::${f.rationale}`;
            if (!cc[k]) cc[k] = { category: f.category, rationale: f.rationale, count: 0 };
            cc[k].count++;
        }
        return Object.values(cc).filter(p => p.count >= 3).sort((a,b) => b.count - a.count);
    }

    /**
     * FNV-based content hash for document ID (normalized content, no filename).
     */
    function generateDocId(rawText) {
        const input = (rawText || '').replace(/\s+/g, ' ').trim();
        let h1 = 0x811c9dc5, h2 = 0x01000193, h3 = 0x6c62272e, h4 = 0x3b6ef4be;
        for (let i = 0; i < input.length; i++) {
            const c = input.charCodeAt(i);
            h1 ^= c; h1 = Math.imul(h1, 0x01000193);
            h2 ^= (c ^ 0x5f); h2 = Math.imul(h2, 0x01000193);
            h3 ^= (c ^ 0xa3); h3 = Math.imul(h3, 0x01000193);
            h4 ^= (c ^ 0xf7); h4 = Math.imul(h4, 0x01000193);
        }
        return 'doc-' + [(h1>>>0),(h2>>>0),(h3>>>0),(h4>>>0)].map(h => h.toString(36)).join('');
    }

    /**
     * Version ID from content hash + processing time
     */
    function generateVersionId(rawText) {
        const contentHash = generateDocId(rawText).substring(4, 14);
        const timestamp = Date.now().toString(36);
        return `v-${contentHash}-${timestamp}`;
    }

    // Public API
    return { buildAuditJson, downloadExcel, downloadMarkdown, downloadJson, generateMarkdown, applyExportFilter };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = Exporter; }
