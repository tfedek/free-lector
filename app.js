/**
 * Free Lector
 */

(function () {
    'use strict';

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
    const filterStatus = document.getElementById('filter-status');
    const downloadXlsx = document.getElementById('download-xlsx');
    const downloadMd = document.getElementById('download-md');
    const downloadJsonBtn = document.getElementById('download-json');
    const resetBtn = document.getElementById('reset-btn');
    const exportFilter = document.getElementById('export-filter');

    // ==========================================
    // PRESET SYSTEM
    // ==========================================
    let activePreset = 'basic';
    let applyingPreset = false;

    // Use global presets from presets.js (loaded via <script> before app.js)
    const BASIC_PRESET = window.BASIC_PRESET;
    const FULL_PRESET = window.FULL_PRESET;

    function applyPreset(name) {
        applyingPreset = true;
        const preset = name === 'full' ? FULL_PRESET : BASIC_PRESET;
        document.querySelectorAll('[data-check]').forEach(cb => {
            const key = cb.dataset.check;
            if (key in preset) { cb.checked = preset[key]; }
        });
        const radio = document.querySelector(`input[name="audit-mode"][value="${name === 'full' ? 'FULL_AUDIT' : 'PROOFREADING'}"]`);
        if (radio) radio.checked = true;
        // Hide custom label when a preset is applied
        const customLabel = document.getElementById('custom-preset-label');
        if (customLabel) customLabel.classList.add('hidden');
        activePreset = name;
        applyingPreset = false;
    }

    const modeRadios = document.querySelectorAll('input[name="audit-mode"]');
    modeRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            if (radio.value === 'FULL_AUDIT') applyPreset('full');
            else applyPreset('basic');
        });
    });

    document.querySelectorAll('[data-check]').forEach(cb => {
        cb.addEventListener('change', () => {
            if (!applyingPreset) {
                activePreset = 'custom';
                // Uncheck both preset radios to indicate custom mode
                modeRadios.forEach(r => r.checked = false);
                // Show 'Prilagođeno' label if available
                const customLabel = document.getElementById('custom-preset-label');
                if (customLabel) customLabel.classList.remove('hidden');
            }
        });
    });

    applyPreset('basic');


    // ==========================================
    // FILE UPLOAD
    // ==========================================
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]); });
    fileBtn.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('click', (e) => { if (e.target !== fileBtn) fileInput.click(); });
    dropZone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
    fileInput.addEventListener('change', () => { if (fileInput.files.length > 0) handleFile(fileInput.files[0]); });
    removeFileBtn.addEventListener('click', () => resetState());

    function handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext === 'docm') { alert('Makro-omogućeni dokumenti (.docm) nisu podržani.'); return; }
        const allowed = ['docx', 'md', 'txt', 'text'];
        if (!allowed.includes(ext)) { alert(`Nepodržan format: .${ext}\nPodržani: .docx, .md, .txt`); return; }
        if (file.size > 20 * 1024 * 1024) { alert('Fajl prevelik (maks 20MB).'); return; }
        currentFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.classList.remove('hidden');
        dropZone.classList.add('hidden');
        optionsPanel.classList.remove('hidden');
    }

    // ==========================================
    // RUN AUDIT
    // ==========================================
    runBtn.addEventListener('click', async () => {
        if (!currentFile || auditInProgress) return;
        auditInProgress = true; runBtn.disabled = true; runBtn.textContent = 'Audit u toku...';

        const options = {};
        document.querySelectorAll('[data-check]').forEach(cb => { options[cb.dataset.check] = cb.checked; });
        options.auditMode = document.querySelector('input[name="audit-mode"]:checked')?.value ?? 'CUSTOM';
        options.preset = activePreset;

        document.getElementById('upload-section').classList.add('hidden');
        progressSection.classList.remove('hidden');
        resultsSection.classList.add('hidden');

        try {
            updateProgress(10, 'Parsiranje', 'Čitanje strukture...');
            currentDocMap = await DocumentParser.parse(currentFile);
            updateProgress(30, 'Struktura', `${currentDocMap.elements.length} elemenata`);
            assignSections(currentDocMap);
            updateProgress(50, 'Provere', 'Pokretanje pravila...');
            await sleep(50);
            const { findings, passedChecks } = RuleEngine.runAudit(currentDocMap, options);
            updateProgress(75, 'Filtriranje', `${findings.length} kandidata`);
            const verified = filterAndDeduplicateFindings(findings, currentDocMap);
            updateProgress(90, 'Izveštaj', 'Priprema...');
            currentAuditJson = Exporter.buildAuditJson(currentDocMap, verified, passedChecks, options);
            updateProgress(100, 'Gotovo', `${verified.length} nalaza`);
            await sleep(300);
            displayResults(currentAuditJson);
        } catch (err) {
            console.error(err);
            progressPhase.textContent = 'Greška'; progressText.textContent = err.message;
            progressFill.style.width = '0%'; progressFill.style.background = 'var(--danger)';
            setTimeout(() => { if (confirm(`Greška: ${err.message}\n\nPonovo?`)) resetState(); }, 500);
        } finally {
            auditInProgress = false; runBtn.disabled = false; runBtn.textContent = 'Pokreni audit';
        }
    });

    function updateProgress(pct, phase, detail) { progressFill.style.width = pct+'%'; progressPhase.textContent = phase; progressText.textContent = detail; }
    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }


    // ==========================================
    // SECTIONS & DEDUP
    // ==========================================
    function assignSections(docMap) {
        let sec = '(početak dokumenta)';
        for (const el of docMap.elements) {
            if (el.type === 'heading' && el.text.trim()) sec = el.text.trim();
            el.section = sec;
        }
    }

    function filterAndDeduplicateFindings(findings, docMap) {
        const verified = []; const seen = new Set();
        for (const f of findings) {
            if (f.original === f.replacement) continue;
            // Dedup key includes table cell
            const key = `${f.paragraphId}::${f.tableId||''}::${f.cellId||''}::${f.category}::${f.original}::${f.replacement}`;
            if (seen.has(key)) continue; seen.add(key);
            if (f.confidence < 0.60) continue;

            // Validate original exists - skip synthetic/global IDs
            const isSynthetic = f.paragraphId && (f.paragraphId.startsWith('fn-') || f.paragraphId.startsWith('en-') || f.paragraphId.startsWith('hdr-') || f.paragraphId.startsWith('ftr-') || f.paragraphId.startsWith('doc-'));
            if (f.original && !f.globalPattern && f.paragraphId && !isSynthetic && !f.original.startsWith('[') &&
                !f.original.startsWith('TOC:') && !f.original.startsWith('Stavka') &&
                !f.original.startsWith('Lista') && !f.original.startsWith('Izvor') &&
                !f.original.startsWith('Citiran')) {
                if (f.cellId) {
                    const tbl = findTableById(docMap, f.tableId);
                    if (!tbl) continue; // Location doesn't exist - skip
                    if (tbl.rows) {
                        const cell = tbl.rows.flat().find(c => c.cellId === f.cellId);
                        if (cell) {
                            const clean = f.original.replace(/^\.\.\./, '').replace(/\.\.\.$/, '');
                            if (clean.length > 3 && !cell.text.includes(clean)) continue;
                        } else continue; // Cell not found - skip
                    }
                } else {
                    const el = docMap.elements.find(e => e.id === f.paragraphId);
                    if (!el) continue; // Location doesn't exist - skip
                    if (el.text) {
                        const clean = f.original.replace(/^\.\.\./, '').replace(/\.\.\.$/, '');
                        if (clean.length > 3 && !el.text.includes(clean)) continue;
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
        renderCoverage(auditJson.processing_coverage);
    }

    function updateStatsDisplay(auditJson) {
        const s = auditJson.summary;
        statsDiv.innerHTML = '';
        const bOpen = s.blockers_open !== undefined ? s.blockers_open : s.blockers;
        const mOpen = s.mandatory_open !== undefined ? s.mandatory_open : s.mandatory;
        const vOpen = s.verify_open !== undefined ? s.verify_open : s.verify;
        if (bOpen > 0) addBadge('Blocker: '+bOpen, 'stat-blocker');
        if (mOpen > 0) addBadge('Obavezno: '+mOpen, 'stat-mandatory');
        if (vOpen > 0) addBadge('Proveriti: '+vOpen, 'stat-verify');
        const rOpen = s.recommendations_open !== undefined ? s.recommendations_open : s.recommendations;
        if (rOpen > 0) addBadge('Preporuke: '+rOpen, 'stat-recommendation');
        if (s.passed_checks > 0) addBadge('Prošlo: '+s.passed_checks, 'stat-passed');
        // Dynamic final message from scope
        finalAssessmentEl.textContent = s.final_assessment;
    }

    function addBadge(text, cls) { const s = document.createElement('span'); s.className = 'stat-badge '+cls; s.textContent = text; statsDiv.appendChild(s); }

    function populateFilters(findings) {
        const cats = [...new Set(findings.map(f => f.category))].sort();
        filterCategory.innerHTML = '<option value="all">Sve</option>';
        cats.forEach(c => { const o = document.createElement('option'); o.value = c; o.textContent = c; filterCategory.appendChild(o); });
        const secs = [...new Set(findings.map(f => f.section))].sort();
        filterSection.innerHTML = '<option value="all">Svi</option>';
        secs.forEach(s => { const o = document.createElement('option'); o.value = s; o.textContent = s.length>40?s.substring(0,40)+'...':s; filterSection.appendChild(o); });
    }


    // ==========================================
    // RENDER TABLE - with status filter, hide actions on resolved, "Vrati u otvoreno"
    // ==========================================
    function renderTable(findings) {
        const priority = filterPriority.value;
        const category = filterCategory.value;
        const section = filterSection.value;
        const autoOnly = filterAutofix ? filterAutofix.checked : false;
        const statusFilter = filterStatus ? filterStatus.value : 'all';

        let filtered = findings;
        if (priority !== 'all') filtered = filtered.filter(f => f.priority === priority);
        if (category !== 'all') filtered = filtered.filter(f => f.category === category);
        if (section !== 'all') filtered = filtered.filter(f => f.section === section);
        if (autoOnly) filtered = filtered.filter(f => f.autoFixable);
        if (statusFilter !== 'all') {
            const statusMap = { 'open': 'OPEN', 'done': 'DONE', 'rejected': 'REJECTED' };
            filtered = filtered.filter(f => f.status === statusMap[statusFilter]);
        }

        resultsBody.innerHTML = '';
        if (filtered.length === 0) { noResults.classList.remove('hidden'); return; }
        noResults.classList.add('hidden');

        filtered.forEach((f, i) => {
            const tr = document.createElement('tr');
            if (f.status === 'DONE') tr.style.opacity = '0.4';
            if (f.status === 'REJECTED') tr.style.opacity = '0.3';

            // Actions: hide mark buttons on resolved, show "Vrati" instead
            let actionsHtml;
            if (f.status === 'OPEN') {
                actionsHtml = `<button class="btn btn-sm btn-accept" data-id="${f.id}" title="Označi kao rešeno">&#10003;</button><button class="btn btn-sm btn-reject" data-id="${f.id}" title="Označi kao nije greška">&#10007;</button>`;
            } else {
                actionsHtml = `<button class="btn btn-sm btn-outline" data-id="${f.id}" data-action="reopen" title="Vrati u otvoreno">&#8634;</button>`;
            }

            tr.innerHTML = `
                <td>${i+1}</td>
                <td title="${escHtml(f.section)}">${escHtml(truncate(f.section,30))}</td>
                <td>${escHtml(f.category)}</td>
                <td><span class="priority-badge pri-${f.priority}">${f.priority}</span></td>
                <td><span class="cell-orig">${escHtml(truncate(f.original,80))}</span></td>
                <td><span class="cell-fix">${escHtml(truncate(f.replacement,80))}</span></td>
                <td><span class="cell-reason">${escHtml(f.rationale)}</span></td>
                <td>${renderConfidence(f.confidence)}</td>
                <td class="action-btns">${actionsHtml}</td>
            `;
            resultsBody.appendChild(tr);
        });

        // Event delegation for actions
        resultsBody.querySelectorAll('.btn-accept').forEach(b => b.addEventListener('click', () => markFinding(b.dataset.id, 'DONE')));
        resultsBody.querySelectorAll('.btn-reject').forEach(b => b.addEventListener('click', () => markFinding(b.dataset.id, 'REJECTED')));
        resultsBody.querySelectorAll('[data-action="reopen"]').forEach(b => b.addEventListener('click', () => markFinding(b.dataset.id, 'OPEN')));
    }

    function markFinding(id, status) {
        const f = currentAuditJson.findings.find(x => x.id === id);
        if (!f) return;
        f.status = status;
        recalculateSummary();
        renderTable(currentAuditJson.findings);
    }

    function recalculateSummary() {
        const all = currentAuditJson.findings;
        const open = all.filter(f => f.status === 'OPEN');
        const s = currentAuditJson.summary;

        // Update open-specific counts (do NOT overwrite totals)
        s.mandatory_open = open.filter(f => f.priority === 'OBAVEZNO').length;
        s.blockers_open = open.filter(f => f.priority === 'BLOCKER').length;
        s.verify_open = open.filter(f => f.priority === 'PROVERITI').length;

        // Update by_status
        s.by_status = {
            open: open.length,
            done: all.filter(f => f.status === 'DONE').length,
            rejected: all.filter(f => f.status === 'REJECTED').length,
        };

        // total_occurrences stays as total, not open
        // mandatory/blockers/verify/recommendations stay as totals (set at build time)

        const scope = currentAuditJson.scope;
        const reqCaps = currentAuditJson.required_capabilities || ['grammar', 'visual_layout', 'style'];
        const reqComplete = reqCaps.every(cap => scope[cap] === true);
        s.can_be_marked_final = reqComplete && s.blockers_open === 0 && s.mandatory_open === 0 && s.verify_open === 0;

        const recommendationsOpen = all.filter(f => f.status === 'OPEN' && f.priority === 'PREPORUKA').length;
        s.recommendations_open = recommendationsOpen;

        if (s.can_be_marked_final && recommendationsOpen === 0) {
            const done = s.by_status ? s.by_status.done : 0;
            const rejected = s.by_status ? s.by_status.rejected : 0;
            if (done + rejected > 0) {
                s.final_assessment = `Nema otvorenih nalaza. Rešeno: ${done}. Odbačeno: ${rejected}.`;
            } else {
                s.final_assessment = 'Audit završen. Sve provere prošle bez nalaza.';
            }
        } else if (s.can_be_marked_final) {
            s.final_assessment = `Nema otvorenih blokirajućih, obaveznih ni nalaza za proveru. Postoji ${recommendationsOpen} preporuka.`;
        } else if (s.blockers_open === 0 && s.mandatory_open === 0 && s.verify_open === 0) {
            const missing = [];
            if (!scope.grammar) missing.push('gramatika');
            if (!scope.style) missing.push('stil');
            if (!scope.visual_layout) missing.push('vizuelni prelom');
            s.final_assessment = missing.length > 0
                ? `Determinističke provere završene. Nedostaje: ${missing.join(', ')}.`
                : 'Determinističke provere završene.';
        } else {
            s.final_assessment = `${s.mandatory_open} obaveznih, ${s.blockers_open} blokirajućih i ${s.verify_open} za proveru. Nije spreman.`;
        }
        currentAuditJson.audit_status.status = s.can_be_marked_final ? 'POTPUN' : 'DELIMIČAN';
        updateStatsDisplay(currentAuditJson);
    }

    function renderConfidence(conf) {
        const pct = Math.round(conf*100);
        const cls = conf >= 0.9 ? 'conf-high' : conf >= 0.75 ? 'conf-mid' : 'conf-low';
        return `<div class="confidence-bar"><div class="confidence-fill ${cls}" style="width:${pct}%"></div></div><small>${pct}%</small>`;
    }

    function renderPassedChecks(checks) {
        passedList.innerHTML = '';
        checks.forEach(c => { const d = document.createElement('div'); d.className = 'passed-item'; d.textContent = `${c.area} (${c.count})`; passedList.appendChild(d); });
    }

    function renderCoverage(coverage) {
        const section = document.getElementById('coverage-section');
        if (!section || !coverage) return;
        let html = '';
        if (coverage.supported.length > 0) {
            html += `<p><strong>Obrađeno:</strong> ${coverage.supported.join(', ')}</p>`;
        }
        if (coverage.partial.length > 0) {
            html += `<p class="coverage-partial"><strong>Delimično obrađeno:</strong> ${coverage.partial.join(', ')}</p>`;
        }
        if (coverage.unsupported.length > 0) {
            html += `<p class="coverage-unsupported"><strong>Nije obrađeno:</strong> ${coverage.unsupported.join(', ')}</p>`;
        }
        if (!html) html = '<p>Svi elementi dokumenta su potpuno obrađeni.</p>';
        section.innerHTML = `<h3>Pokrivenost obrade</h3>${html}`;
    }


    // ==========================================
    // FILTERS
    // ==========================================
    filterPriority.addEventListener('change', () => renderTable(currentAuditJson.findings));
    filterCategory.addEventListener('change', () => renderTable(currentAuditJson.findings));
    filterSection.addEventListener('change', () => renderTable(currentAuditJson.findings));
    if (filterAutofix) filterAutofix.addEventListener('change', () => renderTable(currentAuditJson.findings));
    if (filterStatus) filterStatus.addEventListener('change', () => renderTable(currentAuditJson.findings));

    // ==========================================
    // EXPORT
    // ==========================================
    downloadXlsx.addEventListener('click', () => { if (currentAuditJson) Exporter.downloadExcel(currentAuditJson, exportFilter ? exportFilter.value : 'all'); });
    downloadMd.addEventListener('click', () => { if (currentAuditJson) Exporter.downloadMarkdown(currentAuditJson, exportFilter ? exportFilter.value : 'all'); });
    downloadJsonBtn.addEventListener('click', () => { if (currentAuditJson) Exporter.downloadJson(currentAuditJson, exportFilter ? exportFilter.value : 'all'); });

    // ==========================================
    // RESET
    // ==========================================
    resetBtn.addEventListener('click', resetState);
    function resetState() {
        currentFile = null; currentDocMap = null; currentAuditJson = null;
        auditInProgress = false; runBtn.disabled = false; runBtn.textContent = 'Pokreni audit';
        fileInput.value = ''; fileInfo.classList.add('hidden');
        dropZone.classList.remove('hidden'); optionsPanel.classList.add('hidden');
        progressSection.classList.add('hidden'); resultsSection.classList.add('hidden');
        document.getElementById('upload-section').classList.remove('hidden');
        progressFill.style.width = '0%'; progressFill.style.background = '';
    }

    // ==========================================
    // HELPERS
    // ==========================================
    function escHtml(str) { const d = document.createElement('div'); d.textContent = str; return d.innerHTML; }
    function truncate(str, max) { if (!str) return ''; return str.length > max ? str.substring(0,max)+'...' : str; }
    function formatFileSize(b) { if (b<1024) return b+' B'; if (b<1048576) return (b/1024).toFixed(1)+' KB'; return (b/1048576).toFixed(1)+' MB'; }

    function findTableById(docMap, tableId) {
        // Search top-level tables
        const topLevel = docMap.elements.find(e => e.tableId === tableId);
        if (topLevel) return topLevel;
        // Search nested tables recursively
        for (const el of docMap.elements) {
            if (el.type !== 'table' || !el.rows) continue;
            const found = findNestedTable(el, tableId);
            if (found) return found;
        }
        return null;
    }

    function findNestedTable(tbl, tableId) {
        for (const row of tbl.rows) {
            for (const cell of row) {
                if (!cell.nestedTables) continue;
                for (const nt of cell.nestedTables) {
                    if (nt.tableId === tableId) return nt;
                    const deeper = findNestedTable(nt, tableId);
                    if (deeper) return deeper;
                }
            }
        }
        return null;
    }

})();
