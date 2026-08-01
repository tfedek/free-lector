/**
 * Free Lector — Main Application Controller
 * Wires together: upload → parse → check → display → export
 */

(function () {
    'use strict';

    // State
    let currentFile = null;
    let currentDocMap = null;
    let currentAuditJson = null;
    let auditInProgress = false;

    // DOM refs
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileBtn = document.getElementById('file-btn');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const removeFileBtn = document.getElementById('remove-file');
    const optionsPanel = document.getElementById('options-panel');
    const runBtn = document.getElementById('run-btn');

    const progressSection = document.getElementById('progress-section');
    const progressFill = document.getElementById('progress-fill');
    const progressPhase = document.getElementById('progress-phase');
    const progressText = document.getElementById('progress-text');

    const resultsSection = document.getElementById('results-section');
    const statsDiv = document.getElementById('stats');
    const finalAssessmentEl = document.getElementById('final-assessment');
    const resultsBody = document.getElementById('results-body');
    const noResults = document.getElementById('no-results');
    const passedList = document.getElementById('passed-list');

    const filterPriority = document.getElementById('filter-priority');
    const filterCategory = document.getElementById('filter-category');
    const filterSection = document.getElementById('filter-section');
    const filterAutofix = document.getElementById('filter-autofixable');

    const downloadXlsx = document.getElementById('download-xlsx');
    const downloadMd = document.getElementById('download-md');
    const downloadJsonBtn = document.getElementById('download-json');
    const resetBtn = document.getElementById('reset-btn');
    const exportFilter = document.getElementById('export-filter');


    // ==========================================
    // FILE UPLOAD HANDLING
    // ==========================================
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    fileBtn.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('click', (e) => {
        if (e.target !== fileBtn) fileInput.click();
    });
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) handleFile(fileInput.files[0]);
    });
    removeFileBtn.addEventListener('click', () => resetState());

    function handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        const allowed = ['docx', 'md', 'txt', 'text'];
        if (!allowed.includes(ext)) {
            alert(`Nepodržan format: .${ext}\nPodržani formati: .docx, .md, .txt`);
            return;
        }
        if (file.size > 20 * 1024 * 1024) {
            alert('Fajl je prevelik (maksimum 20MB).');
            return;
        }
        if (ext === 'docm') {
            alert('Makro-omogućeni dokumenti (.docm) nisu podržani.');
            return;
        }
        currentFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.classList.remove('hidden');
        dropZone.classList.add('hidden');
        optionsPanel.classList.remove('hidden');
    }


    // ==========================================
    // RUN AUDIT (blocks concurrent runs)
    // ==========================================
    runBtn.addEventListener('click', async () => {
        if (!currentFile || auditInProgress) return;
        auditInProgress = true;
        runBtn.disabled = true;
        runBtn.textContent = 'Audit u toku...';

        const options = {};
        document.querySelectorAll('[data-check]').forEach(cb => {
            options[cb.dataset.check] = cb.checked;
        });
        options.auditMode = document.querySelector('input[name="audit-mode"]:checked').value;
        options.houseStyle = document.getElementById('house-style').value;

        document.getElementById('upload-section').classList.add('hidden');
        progressSection.classList.remove('hidden');
        resultsSection.classList.add('hidden');

        try {
            updateProgress(10, 'Parsiranje dokumenta', 'Čitanje strukture fajla...');
            currentDocMap = await DocumentParser.parse(currentFile);

            updateProgress(30, 'Analiza strukture', `${currentDocMap.elements.length} elemenata pronađeno`);
            assignSections(currentDocMap);

            updateProgress(50, 'Determinističke provere', 'Pokretanje pravila...');
            await sleep(50);
            const { findings, passedChecks } = RuleEngine.runAudit(currentDocMap, options);

            updateProgress(75, 'Filtriranje nalaza', `${findings.length} kandidata pronađeno`);
            const verified = filterAndDeduplicateFindings(findings, currentDocMap);

            updateProgress(90, 'Generisanje izveštaja', 'Priprema izvoza...');
            currentAuditJson = Exporter.buildAuditJson(currentDocMap, verified, passedChecks, options);

            updateProgress(100, 'Gotovo', `${verified.length} nalaza`);
            await sleep(300);
            displayResults(currentAuditJson);

        } catch (err) {
            console.error(err);
            progressPhase.textContent = 'Greška';
            progressText.textContent = err.message;
            progressFill.style.width = '0%';
            progressFill.style.background = 'var(--danger)';
            setTimeout(() => {
                if (confirm(`Greška: ${err.message}\n\nPokušati ponovo?`)) resetState();
            }, 500);
        } finally {
            auditInProgress = false;
            runBtn.disabled = false;
            runBtn.textContent = 'Pokreni audit';
        }
    });

    function updateProgress(percent, phase, detail) {
        progressFill.style.width = percent + '%';
        progressPhase.textContent = phase;
        progressText.textContent = detail;
    }

    function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }


    // ==========================================
    // ASSIGN SECTIONS
    // ==========================================
    function assignSections(docMap) {
        let currentSection = '(početak dokumenta)';
        for (const el of docMap.elements) {
            if (el.type === 'heading' && el.text.trim()) {
                currentSection = el.text.trim();
            }
            el.section = currentSection;
        }
    }

    // ==========================================
    // FILTER AND DEDUPLICATE FINDINGS
    // Validates original exists in element text
    // ==========================================
    function filterAndDeduplicateFindings(findings, docMap) {
        const verified = [];
        const seen = new Set();

        for (const f of findings) {
            // Skip if original and replacement are identical
            if (f.original === f.replacement) continue;

            // Skip duplicates (same original + category + element)
            const key = `${f.paragraphId}::${f.category}::${f.original}`;
            if (seen.has(key)) continue;
            seen.add(key);

            // Skip low confidence
            if (f.confidence < 0.60) continue;

            // Validate: original should exist in element text
            // (skip validation for synthetic/consolidated findings)
            if (f.original && !f.globalPattern && f.paragraphId) {
                const el = docMap.elements.find(e => e.id === f.paragraphId);
                if (el && el.text && !f.original.startsWith('[') &&
                    !f.original.startsWith('TOC:') &&
                    !f.original.startsWith('Stavka') &&
                    !f.original.startsWith('Lista') &&
                    !f.original.startsWith('Izvor')) {
                    // Check if original text actually appears in element
                    const cleanOriginal = f.original.replace(/^\.\.\./, '').replace(/\.\.\.$/, '');
                    if (cleanOriginal.length > 3 && !el.text.includes(cleanOriginal)) {
                        continue; // Original not found in element — skip
                    }
                }
            }

            verified.push(f);
        }
        return verified;
    }


    // ==========================================
    // DISPLAY RESULTS
    // ==========================================
    function displayResults(auditJson) {
        progressSection.classList.add('hidden');
        resultsSection.classList.remove('hidden');
        updateStatsDisplay(auditJson);
        populateFilters(auditJson.findings);
        renderTable(auditJson.findings);
        renderPassedChecks(auditJson.passed_checks);
    }

    function updateStatsDisplay(auditJson) {
        const s = auditJson.summary;
        statsDiv.innerHTML = '';
        if (s.blockers > 0) addBadge('Blocker: ' + s.blockers, 'stat-blocker');
        if (s.mandatory > 0) addBadge('Obavezno: ' + s.mandatory, 'stat-mandatory');
        if (s.verify > 0) addBadge('Proveriti: ' + s.verify, 'stat-verify');
        if (s.recommendations > 0) addBadge('Preporuke: ' + s.recommendations, 'stat-recommendation');
        if (s.passed_checks > 0) addBadge('Prošlo: ' + s.passed_checks, 'stat-passed');
        finalAssessmentEl.textContent = s.final_assessment;
    }

    function addBadge(text, cls) {
        const span = document.createElement('span');
        span.className = 'stat-badge ' + cls;
        span.textContent = text;
        statsDiv.appendChild(span);
    }

    function populateFilters(findings) {
        const categories = [...new Set(findings.map(f => f.category))].sort();
        filterCategory.innerHTML = '<option value="all">Sve</option>';
        categories.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c; opt.textContent = c;
            filterCategory.appendChild(opt);
        });
        const sections = [...new Set(findings.map(f => f.section))].sort();
        filterSection.innerHTML = '<option value="all">Svi</option>';
        sections.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s.length > 40 ? s.substring(0, 40) + '...' : s;
            filterSection.appendChild(opt);
        });
    }


    function renderTable(findings) {
        const priority = filterPriority.value;
        const category = filterCategory.value;
        const section = filterSection.value;
        const autoOnly = filterAutofix.checked;

        let filtered = findings;
        if (priority !== 'all') filtered = filtered.filter(f => f.priority === priority);
        if (category !== 'all') filtered = filtered.filter(f => f.category === category);
        if (section !== 'all') filtered = filtered.filter(f => f.section === section);
        if (autoOnly) filtered = filtered.filter(f => f.autoFixable);

        resultsBody.innerHTML = '';
        if (filtered.length === 0) { noResults.classList.remove('hidden'); return; }
        noResults.classList.add('hidden');

        filtered.forEach((f, i) => {
            const tr = document.createElement('tr');
            if (f.status === 'DONE') tr.style.opacity = '0.4';
            if (f.status === 'REJECTED') tr.style.opacity = '0.3';

            tr.innerHTML = `
                <td>${i + 1}</td>
                <td title="${escHtml(f.section)}">${escHtml(truncate(f.section, 30))}</td>
                <td>${escHtml(f.category)}</td>
                <td><span class="priority-badge pri-${f.priority}">${f.priority}</span></td>
                <td><span class="cell-orig">${escHtml(truncate(f.original, 80))}</span></td>
                <td><span class="cell-fix">${escHtml(truncate(f.replacement, 80))}</span></td>
                <td><span class="cell-reason">${escHtml(f.rationale)}</span></td>
                <td>${renderConfidence(f.confidence)}</td>
                <td class="action-btns">
                    <button class="btn btn-sm btn-accept" data-id="${f.id}" title="Označi kao rešeno">&#10003;</button>
                    <button class="btn btn-sm btn-reject" data-id="${f.id}" title="Označi kao nije greška">&#10007;</button>
                </td>
            `;
            resultsBody.appendChild(tr);
        });

        // Action handlers with stats recalculation
        resultsBody.querySelectorAll('.btn-accept').forEach(btn => {
            btn.addEventListener('click', () => markFinding(btn.dataset.id, 'DONE'));
        });
        resultsBody.querySelectorAll('.btn-reject').forEach(btn => {
            btn.addEventListener('click', () => markFinding(btn.dataset.id, 'REJECTED'));
        });
    }

    /**
     * Mark a finding and recalculate summary stats
     */
    function markFinding(findingId, newStatus) {
        const finding = currentAuditJson.findings.find(f => f.id === findingId);
        if (!finding) return;
        finding.status = newStatus;
        recalculateSummary();
        renderTable(currentAuditJson.findings);
    }

    /**
     * Recalculate summary after status changes
     */
    function recalculateSummary() {
        const open = currentAuditJson.findings.filter(f => f.status === 'OPEN');
        const s = currentAuditJson.summary;
        s.mandatory = open.filter(f => f.priority === 'OBAVEZNO').length;
        s.verify = open.filter(f => f.priority === 'PROVERITI').length;
        s.recommendations = open.filter(f => f.priority === 'PREPORUKA').length;
        s.blockers = open.filter(f => f.priority === 'BLOCKER').length;
        s.can_be_marked_final = s.blockers === 0 && s.mandatory === 0 && s.verify === 0;

        if (s.blockers === 0 && s.mandatory === 0 && s.verify === 0) {
            s.final_assessment = 'Determinističke provere su završene. Gramatika, stil i vizuelni prelom nisu provereni.';
        } else {
            s.final_assessment = `Dokument ima ${s.mandatory} obaveznih ispravki i ${s.blockers} blokirajućih problema. Nije spreman za objavljivanje.`;
        }
        updateStatsDisplay(currentAuditJson);
    }


    function renderConfidence(conf) {
        const pct = Math.round(conf * 100);
        const cls = conf >= 0.9 ? 'conf-high' : conf >= 0.75 ? 'conf-mid' : 'conf-low';
        return `<div class="confidence-bar"><div class="confidence-fill ${cls}" style="width:${pct}%"></div></div><small>${pct}%</small>`;
    }

    function renderPassedChecks(checks) {
        passedList.innerHTML = '';
        checks.forEach(c => {
            const div = document.createElement('div');
            div.className = 'passed-item';
            div.textContent = `${c.area} (${c.count})`;
            passedList.appendChild(div);
        });
    }

    // ==========================================
    // FILTER HANDLERS
    // ==========================================
    filterPriority.addEventListener('change', () => renderTable(currentAuditJson.findings));
    filterCategory.addEventListener('change', () => renderTable(currentAuditJson.findings));
    filterSection.addEventListener('change', () => renderTable(currentAuditJson.findings));
    filterAutofix.addEventListener('change', () => renderTable(currentAuditJson.findings));

    // ==========================================
    // EXPORT HANDLERS (with export filter)
    // ==========================================
    downloadXlsx.addEventListener('click', () => {
        if (currentAuditJson) {
            const filter = exportFilter ? exportFilter.value : 'all';
            Exporter.downloadExcel(currentAuditJson, filter);
        }
    });
    downloadMd.addEventListener('click', () => {
        if (currentAuditJson) {
            const filter = exportFilter ? exportFilter.value : 'all';
            Exporter.downloadMarkdown(currentAuditJson, filter);
        }
    });
    downloadJsonBtn.addEventListener('click', () => {
        if (currentAuditJson) {
            const filter = exportFilter ? exportFilter.value : 'all';
            Exporter.downloadJson(currentAuditJson, filter);
        }
    });

    // ==========================================
    // RESET
    // ==========================================
    resetBtn.addEventListener('click', resetState);

    function resetState() {
        currentFile = null;
        currentDocMap = null;
        currentAuditJson = null;
        auditInProgress = false;
        runBtn.disabled = false;
        runBtn.textContent = 'Pokreni audit';

        fileInput.value = '';
        fileInfo.classList.add('hidden');
        dropZone.classList.remove('hidden');
        optionsPanel.classList.add('hidden');
        progressSection.classList.add('hidden');
        resultsSection.classList.add('hidden');
        document.getElementById('upload-section').classList.remove('hidden');
        progressFill.style.width = '0%';
        progressFill.style.background = '';
    }

    // ==========================================
    // HELPERS
    // ==========================================
    function escHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
    function truncate(str, maxLen) {
        if (!str) return '';
        return str.length > maxLen ? str.substring(0, maxLen) + '...' : str;
    }
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

})();
