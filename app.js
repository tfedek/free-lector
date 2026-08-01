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
    const finalAssessment = document.getElementById('final-assessment');
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



    // ==========================================
    // FILE UPLOAD HANDLING
    // ==========================================

    // Drag and drop
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
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // Click to upload
    fileBtn.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('click', (e) => {
        if (e.target !== fileBtn) fileInput.click();
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            handleFile(fileInput.files[0]);
        }
    });

    // Remove file
    removeFileBtn.addEventListener('click', () => {
        resetState();
    });

    function handleFile(file) {
        // Validate extension
        const ext = file.name.split('.').pop().toLowerCase();
        const allowed = ['docx', 'md', 'txt', 'text'];
        if (!allowed.includes(ext)) {
            alert(`Nepodržan format: .${ext}\nPodržani formati: .docx, .md, .txt`);
            return;
        }

        // Validate size (max 20MB)
        if (file.size > 20 * 1024 * 1024) {
            alert('Fajl je prevelik (maksimum 20MB).');
            return;
        }

        // Block .docm disguised as .docx
        if (ext === 'docm') {
            alert('Makro-omogućeni dokumenti (.docm) nisu podržani iz bezbednosnih razloga.');
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
    // RUN AUDIT
    // ==========================================

    runBtn.addEventListener('click', async () => {
        if (!currentFile) return;

        // Gather options
        const options = {};
        document.querySelectorAll('[data-check]').forEach(cb => {
            options[cb.dataset.check] = cb.checked;
        });
        options.auditMode = document.querySelector('input[name="audit-mode"]:checked').value;
        options.houseStyle = document.getElementById('house-style').value;

        // Show progress
        document.getElementById('upload-section').classList.add('hidden');
        progressSection.classList.remove('hidden');
        resultsSection.classList.add('hidden');

        try {
            // Phase 1: Parse
            updateProgress(10, 'Parsiranje dokumenta', 'Čitanje strukture fajla...');
            currentDocMap = await DocumentParser.parse(currentFile);

            // Phase 2: Assign sections to elements
            updateProgress(30, 'Analiza strukture', `${currentDocMap.elements.length} elemenata pronađeno`);
            assignSections(currentDocMap);

            // Phase 3: Run deterministic checks
            updateProgress(50, 'Determinističke provere', 'Pokretanje pravila...');
            await sleep(50); // Allow UI to update
            const { findings, passedChecks } = RuleEngine.runAudit(currentDocMap, options);

            // Phase 4: Verify findings (dedup, validate)
            updateProgress(75, 'Verifikacija nalaza', `${findings.length} kandidata pronađeno`);
            const verifiedFindings = verifyFindings(findings, currentDocMap);

            // Phase 5: Build audit JSON
            updateProgress(90, 'Generisanje izveštaja', 'Priprema izvoza...');
            currentAuditJson = Exporter.buildAuditJson(currentDocMap, verifiedFindings, passedChecks, options);

            // Phase 6: Display results
            updateProgress(100, 'Gotovo', `${verifiedFindings.length} nalaza`);
            await sleep(300);

            displayResults(currentAuditJson);

        } catch (err) {
            console.error(err);
            progressPhase.textContent = 'Greška';
            progressText.textContent = err.message;
            progressFill.style.width = '0%';
            progressFill.style.background = 'var(--danger)';

            // Show reset option
            setTimeout(() => {
                if (confirm(`Greška pri obradi: ${err.message}\n\nPokušati ponovo?`)) {
                    resetState();
                }
            }, 500);
        }
    });

    function updateProgress(percent, phase, detail) {
        progressFill.style.width = percent + '%';
        progressPhase.textContent = phase;
        progressText.textContent = detail;
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }



    // ==========================================
    // ASSIGN SECTIONS (nearest heading) TO ELEMENTS
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
    // VERIFY FINDINGS (dedup, sanity check)
    // ==========================================
    function verifyFindings(findings, docMap) {
        const verified = [];
        const seen = new Set();

        for (const f of findings) {
            // Skip if original and replacement are identical
            if (f.original === f.replacement) continue;

            // Skip duplicates (same original + same category in same element)
            const key = `${f.paragraphId}::${f.category}::${f.original}`;
            if (seen.has(key)) continue;
            seen.add(key);

            // Skip if confidence is too low
            if (f.confidence < 0.60) continue;

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

        const s = auditJson.summary;

        // Stats badges
        statsDiv.innerHTML = '';
        if (s.blockers > 0) addBadge('Blocker: ' + s.blockers, 'stat-blocker');
        if (s.mandatory > 0) addBadge('Obavezno: ' + s.mandatory, 'stat-mandatory');
        if (s.verify > 0) addBadge('Proveriti: ' + s.verify, 'stat-verify');
        if (s.recommendations > 0) addBadge('Preporuke: ' + s.recommendations, 'stat-recommendation');
        if (s.passed_checks > 0) addBadge('Prošlo: ' + s.passed_checks, 'stat-passed');

        finalAssessment.textContent = s.final_assessment;

        // Populate filter dropdowns
        populateFilters(auditJson.findings);

        // Render table
        renderTable(auditJson.findings);

        // Render passed checks
        renderPassedChecks(auditJson.passed_checks);
    }

    function addBadge(text, cls) {
        const span = document.createElement('span');
        span.className = 'stat-badge ' + cls;
        span.textContent = text;
        statsDiv.appendChild(span);
    }

    function populateFilters(findings) {
        // Categories
        const categories = [...new Set(findings.map(f => f.category))].sort();
        filterCategory.innerHTML = '<option value="all">Sve</option>';
        categories.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c;
            opt.textContent = c;
            filterCategory.appendChild(opt);
        });

        // Sections
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

        if (filtered.length === 0) {
            noResults.classList.remove('hidden');
            return;
        }
        noResults.classList.add('hidden');

        filtered.forEach((f, i) => {
            const tr = document.createElement('tr');
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
                    <button class="btn btn-sm btn-accept" data-id="${f.id}" title="Prihvati">&#10003;</button>
                    <button class="btn btn-sm btn-reject" data-id="${f.id}" title="Odbij">&#10007;</button>
                </td>
            `;
            resultsBody.appendChild(tr);
        });

        // Action button handlers
        resultsBody.querySelectorAll('.btn-accept').forEach(btn => {
            btn.addEventListener('click', () => {
                const finding = currentAuditJson.findings.find(f => f.id === btn.dataset.id);
                if (finding) { finding.status = 'DONE'; btn.closest('tr').style.opacity = '0.4'; }
            });
        });
        resultsBody.querySelectorAll('.btn-reject').forEach(btn => {
            btn.addEventListener('click', () => {
                const finding = currentAuditJson.findings.find(f => f.id === btn.dataset.id);
                if (finding) { finding.status = 'REJECTED'; btn.closest('tr').style.opacity = '0.3'; }
            });
        });
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
            div.textContent = c.area;
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
    // EXPORT HANDLERS
    // ==========================================
    downloadXlsx.addEventListener('click', () => {
        if (currentAuditJson) Exporter.downloadExcel(currentAuditJson);
    });

    downloadMd.addEventListener('click', () => {
        if (currentAuditJson) Exporter.downloadMarkdown(currentAuditJson);
    });

    downloadJsonBtn.addEventListener('click', () => {
        if (currentAuditJson) Exporter.downloadJson(currentAuditJson);
    });

    // ==========================================
    // RESET
    // ==========================================
    resetBtn.addEventListener('click', resetState);

    function resetState() {
        currentFile = null;
        currentDocMap = null;
        currentAuditJson = null;

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
