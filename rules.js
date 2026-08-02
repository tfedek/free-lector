/**
 * Rule Engine Module — Round 4
 * Deterministic, rule-based checks for document auditing
 */

const RuleEngine = (() => {
    'use strict';

    let findingCounter = 0;
    function resetCounter() { findingCounter = 0; }
    function nextId() { findingCounter++; return `F-${String(findingCounter).padStart(4, '0')}`; }

    /**
     * Run all enabled checks. Table cell checks run BEFORE passedChecks are formed.
     */
    function runAudit(docMap, options) {
        resetCounter();
        const findings = [];
        const passedChecks = [];

        const checks = [
            { key: 'brackets', name: 'Nebalansirane zagrade', fn: checkBrackets },
            { key: 'quotes', name: 'Navodnici', fn: checkQuotes },
            { key: 'markdown', name: 'Markdown artefakti', fn: checkMarkdownArtifacts },
            { key: 'spacing', name: 'Razmaci i interpunkcija', fn: checkSpacing },
            { key: 'scriptMix', name: 'Mešanje ćirilice/latinice', fn: checkScriptMixing },
            { key: 'greek', name: 'Grčki citati bez prevoda', fn: checkGreekWithoutTranslation },
            { key: 'duplicates', name: 'Duple reči', fn: checkDuplicateWords },
            { key: 'toc', name: 'TOC vs naslovi', fn: checkTocVsHeadings },
            { key: 'numbering', name: 'Numeracija lista', fn: checkNumbering },
            { key: 'bibliography', name: 'Bibliografija', fn: checkBibliography },
            { key: 'urls', name: 'URL-ovi', fn: checkUrls },
            { key: 'footnotes', name: 'Fusnote', fn: (dm) => checkFootnotes(dm, options) },
            { key: 'repetition', name: 'Ponovljeni pasusi', fn: checkRepetition },
            { key: 'capsWords', name: 'ALL-CAPS reči', fn: checkAllCaps },
            { key: 'emptyHeadings', name: 'Prazni naslovi', fn: checkEmptyHeadings },
        ];

        // Run per-element checks
        const checkResults = {};
        for (const check of checks) {
            if (!options[check.key]) continue;
            const result = check.fn(docMap);
            for (const f of result.findings) {
                if (f._element && f._element.isDirectQuote) {
                    f.priority = 'PROVERITI'; f.autoFixable = false;
                    f.requiresSourceVerification = true; f.isDirectQuote = true;
                }
                delete f._element;
            }
            checkResults[check.key] = result;
            findings.push(...result.findings);
        }

        // Run table cell checks BEFORE forming passedChecks
        const cellFindings = checkTableCells(docMap, options);
        findings.push(...cellFindings);

        // Run checks on headers/footers
        const hfFindings = checkHeadersFooters(docMap, options);
        findings.push(...hfFindings);

        // Report unsupported elements as warning
        if (docMap.processingCoverage) {
            const pc = docMap.processingCoverage;
            if (pc.unsupported.length > 0 || pc.partial.length > 0) {
                const parts = [];
                if (pc.partial.length > 0) parts.push(`Delimično: ${pc.partial.join(', ')}`);
                if (pc.unsupported.length > 0) parts.push(`Nije obrađeno: ${pc.unsupported.join(', ')}`);
                findings.push(makeFinding({
                    element: { id: 'doc-coverage', type: 'document', text: '', section: '(pokrivenost)', isDirectQuote: false },
                    category: 'Pokrivenost',
                    priority: 'PROVERITI',
                    confidence: 1.0,
                    original: parts.join('. '),
                    replacement: '[proveriti navedene delove dokumenta ručno]',
                    rationale: 'Dokument sadrži elemente koji nisu potpuno obrađeni. Rezultat audita može biti nepotpun.',
                }));
            }
        }

        // Now form passedChecks — considers element-level, cell-level, header/footer, and footnote cross-category findings
        const allExtraFindings = [...cellFindings, ...hfFindings];
        // Footnote findings are already in the check result for 'footnotes', but they produce
        // findings in other categories (Razmaci, Zagrade, etc.) — collect them
        const footnoteResult = checkResults['footnotes'];
        const footnoteCrossFindings = footnoteResult ? footnoteResult.findings.filter(f => f.category !== 'Fusnote') : [];

        for (const check of checks) {
            if (!options[check.key]) continue;
            const result = checkResults[check.key];
            if (!result) continue;

            // Check if any extra source (cells, headers/footers, footnotes) produced findings in this category
            const categoryMap = { spacing: 'Razmaci', brackets: 'Zagrade', scriptMix: 'Mešanje pisama', quotes: 'Tipografija', duplicates: 'Duple reči', urls: 'URL', markdown: 'Markdown artefakt', greek: 'Grčki bez prevoda', capsWords: 'ALL-CAPS' };
            const cat = categoryMap[check.key];
            const hasExtraFindings = cat ? (
                allExtraFindings.some(f => f.category === cat) ||
                footnoteCrossFindings.some(f => f.category === cat)
            ) : false;

            if (result.findings.length === 0 && !hasExtraFindings) {
                passedChecks.push({ area: check.name, result: 'Bez grešaka', count: result.scannedCount });
            }
        }

        return { findings, passedChecks };
    }


    // ==========================================
    // TABLE CELL CHECKS — full spacing, all brackets {}, direct quote protection
    // Returns findings array (not modifying external)
    // ==========================================
    function checkTableCells(docMap, options) {
        const findings = [];
        const doSpacing = options.spacing === true;
        const doBrackets = options.brackets === true;
        const doScriptMix = options.scriptMix === true;
        const doDuplicates = options.duplicates === true;
        const doUrls = options.urls === true;
        const doGreek = options.greek === true;
        const doMarkdown = options.markdown === true;
        if (!doSpacing && !doBrackets && !doScriptMix && !doDuplicates && !doUrls && !doGreek && !doMarkdown) return findings;

        for (const el of docMap.elements) {
            if (el.type !== 'table' || !el.rows) continue;
            for (const row of el.rows) {
                for (const cell of row) {
                    if (!cell.text || cell.text.trim().length === 0) continue;
                    const text = cell.directText !== undefined ? cell.directText : cell.text;
                    if (!text || text.trim().length === 0) continue;
                    const cm = { tableId: cell.tableId, rowId: cell.rowId, cellId: cell.cellId, rowIndex: cell.rowIndex, columnIndex: cell.columnIndex };

                    if (doSpacing) {
                        // All spacing checks (not just double spaces)
                        let m;
                        const dbl = /  +/g;
                        while ((m = dbl.exec(text)) !== null) {
                            const ctx = getContext(text, m.index, 20);
                            findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.99, original: ctx, replacement: ctx.replace(/  +/g, ' '), rationale: 'Višestruki razmak u ćeliji.', autoFixable: true, ...cm }));
                        }
                        const sbp = / +([,.:;!?])/g;
                        while ((m = sbp.exec(text)) !== null) {
                            const ctx = getContext(text, m.index, 20);
                            findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.97, original: ctx, replacement: ctx.replace(/ +([,.:;!?])/, '$1'), rationale: `Razmak pre \u201e${m[1]}\u201c u ćeliji.`, autoFixable: true, ...cm }));
                        }
                        const nsa = /([,;:])([^\s\d"'\u201C\u201D\u201E\u2019)\]])/g;
                        while ((m = nsa.exec(text)) !== null) {
                            if (text.substring(Math.max(0, m.index-10), m.index+10).match(/https?:|:\\/)) continue;
                            const ctx = getContext(text, m.index, 20);
                            findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.85, original: ctx, replacement: ctx.replace(/([,;:])(\S)/, '$1 $2'), rationale: `Nedostaje razmak posle \u201e${m[1]}\u201c u ćeliji.`, autoFixable: true, ...cm }));
                        }
                    }

                    if (doBrackets) {
                        for (const [open, close] of [['(',')'],['[',']'],['{','}']]) {
                            let depth = 0;
                            let prematureClose = false;
                            for (let i = 0; i < text.length; i++) {
                                if (text[i] === open) depth++;
                                else if (text[i] === close) {
                                    depth--;
                                    if (depth < 0) { prematureClose = true; depth = 0; }
                                }
                            }
                            if (depth !== 0 || prematureClose) {
                                const rationale = prematureClose
                                    ? `Prerano zatvorena zagrada ${close} pre otvaranja ${open} u ćeliji.`
                                    : `Neuparene zagrade u ćeliji (${cell.cellId}).`;
                                findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.95, original: text.substring(0, 50), replacement: `[neuparene zagrade ${open}${close} u ćeliji]`, rationale, ...cm }));
                            }
                        }
                    }

                    if (doScriptMix) {
                        const words = text.match(/[\p{L}\p{M}]+/gu) || [];
                        for (const word of words) {
                            if (word.length < 2) continue;
                            if (/[\u0400-\u04FF]/.test(word) && /[a-zA-Z\u00C0-\u024F]/.test(word)) {
                                findings.push(makeFinding({ element: el, category: 'Mešanje pisama', priority: 'OBAVEZNO', confidence: 0.96, original: word, replacement: '[prebaciti na jedno pismo]', rationale: `Pomešana slova u ćeliji: \u201e${word}\u201c.`, ...cm }));
                            }
                        }
                    }

                    // Per-cell: duplicate words
                    if (options.duplicates) {
                        const allowedDups = new Set(['ha','da','ne','vrlo','još','baš','sve']);
                        const dupRe = /(?<=\s|^)(\p{L}+)\s+\1(?=\s|[,.:;!?)\]\}]|$)/giu; let dm;
                        while ((dm = dupRe.exec(text)) !== null) {
                            if (allowedDups.has(dm[1].toLowerCase()) || dm[1].length < 2) continue;
                            findings.push(makeFinding({ element: el, category: 'Duple reči', priority: 'OBAVEZNO', confidence: 0.95, original: dm[0], replacement: dm[1], rationale: `Ponovljena reč \u201e${dm[1]}\u201c u ćeliji.`, autoFixable: true, ...cm }));
                        }
                    }

                    // Per-cell: URLs
                    if (options.urls) {
                        const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g; let um;
                        while ((um = urlRe.exec(text)) !== null) {
                            const url = um[0];
                            let trailing = url.match(/[,.;]+$/);
                            if (trailing) { findings.push(makeFinding({ element: el, category: 'URL', priority: 'PROVERITI', confidence: 0.80, original: url, replacement: url.replace(/[,.;]+$/, ''), rationale: 'URL završen interpunkcijom u ćeliji.', ...cm })); }
                        }
                    }

                    // Per-cell: Greek without translation
                    if (options.greek) {
                        const gRe = /[\u0370-\u03FF\u1F00-\u1FFF][\u0370-\u03FF\u1F00-\u1FFF\s,;\u00B7.'"\u0300-\u036F\u1DC0-\u1DFF]{2,}/g; let gm;
                        while ((gm = gRe.exec(text)) !== null) {
                            const after = text.substring(gm.index+gm[0].length, gm.index+gm[0].length+200);
                            const trNear = /[(\u201E\u201A""][^)"\u201C"]*[\u0400-\u04FFa-zA-Z\u00C0-\u024F]{5,}[^)"\u201C"]*[)"\u201C""]/;
                            if (trNear.test(after.substring(0,150))) continue;
                            const snip = gm[0].length > 40 ? gm[0].substring(0,40)+'...' : gm[0];
                            findings.push(makeFinding({ element: el, category: 'Grčki bez prevoda', priority: 'PROVERITI', confidence: 0.75, original: snip, replacement: '[dodati prevod]', rationale: 'Grčki bez prevoda u ćeliji.', requiresSourceVerification: true, ...cm }));
                        }
                    }

                    // Per-cell: Markdown artifacts
                    if (options.markdown && docMap.type !== 'markdown') {
                        const mdPatterns = [
                            { re: /(?:^|\s)\*\*[^*]+\*\*(?:\s|$)/gm, desc: '**bold**' },
                            { re: /(?:^|\s)\*[^*]+\*(?:\s|$)/gm, desc: '*italic*' },
                            { re: /\[([^\]]+)\]\([^)]+\)/g, desc: '[link](url)' },
                            { re: /```/g, desc: '```' },
                        ];
                        for (const pat of mdPatterns) { pat.re.lastIndex = 0; let mm; while ((mm = pat.re.exec(text)) !== null) { findings.push(makeFinding({ element: el, category: 'Markdown artefakt', priority: 'OBAVEZNO', confidence: 0.92, original: mm[0].trim(), replacement: '[ukloniti markdown]', rationale: `Ostatak ${pat.desc} u ćeliji.`, ...cm })); } }
                    }

                    // Apply direct quote protection PER CELL (not per table)
                    const cellIsQuote = /^\u201E[\s\S]*\u201C$/.test(text.trim()) ||
                        /^\u00AB[\s\S]*\u00BB$/.test(text.trim()) ||
                        /^[\u2014\u2015]\s/.test(text.trim());
                    if (cellIsQuote) {
                        for (const f of findings) {
                            if (f.cellId === cell.cellId && f.priority !== 'PROVERITI') {
                                f.priority = 'PROVERITI'; f.autoFixable = false; f.requiresSourceVerification = true; f.isDirectQuote = true;
                            }
                        }
                    }
                }
            }
        }

        // Recursively check nested tables — inherit parent's section
        for (const el of docMap.elements) {
            if (el.type !== 'table' || !el.rows) continue;
            for (const row of el.rows) {
                for (const cell of row) {
                    if (!cell.nestedTables) continue;
                    for (const nestedTbl of cell.nestedTables) {
                        // Inherit section from parent table element
                        nestedTbl.section = el.section || '(tabela)';
                        nestedTbl.id = nestedTbl.tableId || el.id;
                        const nestedDoc = { elements: [nestedTbl], type: 'docx', footnotes: [], endnotes: [], headerElements: [], footerElements: [] };
                        const nf = checkTableCells(nestedDoc, options);
                        findings.push(...nf);
                    }
                }
            }
        }

        return findings;
    }


    // ==========================================
    // HEADERS/FOOTERS — run basic checks on header/footer text
    // ==========================================
    function checkHeadersFooters(docMap, options) {
        const findings = [];
        const hfElements = [
            ...(docMap.headerElements || []),
            ...(docMap.footerElements || []),
        ];
        if (hfElements.length === 0) return findings;

        for (const el of hfElements) {
            if (!el.text || !el.text.trim()) continue;
            const text = el.text;

            if (options.spacing) {
                const dbl = /  +/g; let m;
                while ((m = dbl.exec(text)) !== null) {
                    const ctx = getContext(text, m.index, 20);
                    findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.99, original: ctx, replacement: ctx.replace(/  +/g, ' '), rationale: `Višestruki razmak u ${el.type === 'header' ? 'zaglavlju' : 'podnožju'}.`, autoFixable: true }));
                }
            }

            if (options.brackets) {
                for (const [o,c] of [['(',')'],['[',']'],['{','}']]) {
                    let depth = 0; let prematureClose = false;
                    for (let i = 0; i < text.length; i++) {
                        if (text[i]===o) depth++;
                        else if (text[i]===c) { depth--; if (depth < 0) { prematureClose = true; depth = 0; } }
                    }
                    if (depth !== 0 || prematureClose) {
                        const rationale = prematureClose
                            ? `Prerano zatvorena zagrada ${c} u ${el.type === 'header' ? 'zaglavlju' : 'podnožju'}.`
                            : `Neuparene zagrade u ${el.type === 'header' ? 'zaglavlju' : 'podnožju'}.`;
                        findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.95, original: text.substring(0,50), replacement: `[neuparene zagrade ${o}${c}]`, rationale }));
                    }
                }
            }

            if (options.scriptMix) {
                const words = text.match(/[\p{L}\p{M}]+/gu) || [];
                for (const word of words) {
                    if (word.length < 2) continue;
                    if (/[\u0400-\u04FF]/.test(word) && /[a-zA-Z\u00C0-\u024F]/.test(word)) {
                        findings.push(makeFinding({ element: el, category: 'Mešanje pisama', priority: 'OBAVEZNO', confidence: 0.96, original: word, replacement: '[prebaciti na jedno pismo]', rationale: `Pomešana slova u ${el.type === 'header' ? 'zaglavlju' : 'podnožju'}.` }));
                    }
                }
            }

            const loc = el.type === 'header' ? 'zaglavlju' : 'podnožju';

            if (options.quotes) {
                const sc = (text.match(/"/g)||[]).length;
                if (sc > 0) findings.push(makeFinding({ element: el, category: 'Tipografija', priority: 'OBAVEZNO', confidence: 0.95, original: `[${sc} ravnih navodnika u ${loc}]`, replacement: '[zameniti tipografskim]', rationale: `${sc} ravnih navodnika u ${loc}.`, autoFixable: true }));
            }

            if (options.duplicates) {
                const dupeRe = /(?<=\s|^)(\p{L}+)\s+\1(?=\s|[,.:;!?)\]\}]|$)/giu; let dm;
                const allowed = new Set(['ha','da','ne','vrlo','još','baš','sve']);
                while ((dm = dupeRe.exec(text)) !== null) {
                    if (allowed.has(dm[1].toLowerCase()) || dm[1].length < 2) continue;
                    findings.push(makeFinding({ element: el, category: 'Duple reči', priority: 'OBAVEZNO', confidence: 0.95, original: dm[0], replacement: dm[1], rationale: `Ponovljena reč u ${loc}.`, autoFixable: true }));
                }
            }

            if (options.urls) {
                const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g; let um;
                while ((um = urlRe.exec(text)) !== null) {
                    if (um[0].match(/[,.;]+$/)) findings.push(makeFinding({ element: el, category: 'URL', priority: 'PROVERITI', confidence: 0.80, original: um[0], replacement: um[0].replace(/[,.;]+$/,''), rationale: `URL u ${loc} završen interpunkcijom.` }));
                }
            }

            if (options.markdown && docMap.type !== 'markdown') {
                const mdPats = [/(?:^|\s)\*\*[^*]+\*\*(?:\s|$)/gm, /(?:^|\s)\*[^*]+\*(?:\s|$)/gm, /\[([^\]]+)\]\([^)]+\)/g, /```/g];
                for (const re of mdPats) { re.lastIndex=0; let mm; while ((mm=re.exec(text))!==null) { findings.push(makeFinding({ element: el, category: 'Markdown artefakt', priority: 'OBAVEZNO', confidence: 0.92, original: mm[0].trim(), replacement: '[ukloniti markdown]', rationale: `Markdown u ${loc}.` })); } }
            }

            if (options.greek) {
                const gRe = /[\u0370-\u03FF\u1F00-\u1FFF][\u0370-\u03FF\u1F00-\u1FFF\s,;\u00B7.'"\u0300-\u036F]{2,}/g; let gm;
                while ((gm = gRe.exec(text)) !== null) {
                    const snip = gm[0].length > 40 ? gm[0].substring(0,40)+'...' : gm[0];
                    findings.push(makeFinding({ element: el, category: 'Grčki bez prevoda', priority: 'PROVERITI', confidence: 0.75, original: snip, replacement: '[dodati prevod]', rationale: `Grčki tekst u ${loc} bez prevoda.`, requiresSourceVerification: true }));
                }
            }
        }
        return findings;
    }

    // ==========================================
    // BRACKETS — skip table elements (checked per-cell)
    // ==========================================
    function checkBrackets(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const openChars = '([{'; const closeChars = ')]}';
        const names = {'(':'obla','[':'uglasta','{':'vitičasta'};
        for (const el of docMap.elements) {
            if (!el.text || !el.text.trim()) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            // Stack-based: detects interleaving ([)] and premature closing
            const stack = [];
            for (let i = 0; i < el.text.length; i++) {
                const ch = el.text[i];
                const openIdx = openChars.indexOf(ch);
                const closeIdx = closeChars.indexOf(ch);
                if (openIdx !== -1) {
                    stack.push({ char: ch, pos: i, type: openIdx });
                } else if (closeIdx !== -1) {
                    if (stack.length === 0) {
                        // Premature close — nothing to match
                        findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.98, original: getContext(el.text, i, 40), replacement: `[ukloniti višak \u201e${ch}\u201c]`, rationale: `Zatvorena zagrada ${ch} bez odgovarajuće otvorene.` }));
                    } else if (stack[stack.length-1].type !== closeIdx) {
                        // Mismatched: e.g. ( then ]
                        const top = stack.pop();
                        const expected = closeChars[top.type];
                        findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.95, original: getContext(el.text, i, 40), replacement: `[pogrešan redosled: otvoreno \u201e${top.char}\u201c ali zatvoreno \u201e${ch}\u201c]`, rationale: `Ukrštene zagrade: ${top.char}...${ch} umesto ${top.char}...${expected}.` }));
                    } else {
                        stack.pop(); // Correct match
                    }
                }
            }
            // Remaining unclosed
            for (const item of stack) {
                const expected = closeChars[item.type];
                findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.98, original: getContext(el.text, item.pos, 40), replacement: `[dodati \u201e${expected}\u201c]`, rationale: `Otvorena ${names[item.char]||''} zagrada bez zatvorene.` }));
            }
            // ; where ) should be
            const sp = /\([^)]*;(?=[^(]*$)/g; let m;
            while ((m = sp.exec(el.text)) !== null) {
                const frag = el.text.substring(m.index);
                if (frag.indexOf(')') === -1 || frag.indexOf(';') < frag.indexOf(')')) {
                    findings.push(makeFinding({ element: el, category: 'Zagrade', priority: 'PROVERITI', confidence: 0.75, original: getContext(el.text, m.index, 50), replacement: '[proveriti da li \u201e;\u201c treba da bude \u201e)\u201c]', rationale: 'Tačka-zarez unutar nezatvorene zagrade.' }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // QUOTES — global consolidation, multi-paragraph balance
    // ==========================================
    function checkQuotes(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        let totalStraight = 0;
        // Multi-paragraph typographic quote balance
        let globalOpen = 0, globalClose = 0;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            totalStraight += (el.text.match(/"/g) || []).length;
            globalOpen += (el.text.match(/\u201E/g) || []).length;
            globalClose += (el.text.match(/\u201C/g) || []).length;
        }

        if (totalStraight > 0) {
            findings.push(makeFinding({ element: { id: 'doc-global', type: 'document', text: '', section: '(ceo dokument)', isDirectQuote: false }, category: 'Tipografija', priority: 'OBAVEZNO', confidence: 0.95, original: `[${totalStraight} ravnih navodnika u dokumentu]`, replacement: '[zameniti tipografskim: \u201e...\u201c]', rationale: `Ukupno ${totalStraight} ravnih navodnika. Standard: \u201e...\u201c.`, autoFixable: true, globalPattern: true }));
        }

        // Multi-paragraph balance check
        if (globalOpen !== globalClose) {
            findings.push(makeFinding({ element: { id: 'doc-global', type: 'document', text: '', section: '(ceo dokument)', isDirectQuote: false }, category: 'Tipografija', priority: 'OBAVEZNO', confidence: 0.88, original: `[${globalOpen}\u00D7 \u201e vs ${globalClose}\u00D7 \u201c u celom dokumentu]`, replacement: '[upariti navodnike]', rationale: 'Neupareni tipografski navodnici kroz ceo dokument (moguć višepasusni citat).' }));
        }

        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // MARKDOWN ARTIFACTS
    // ==========================================
    function checkMarkdownArtifacts(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        if (docMap.type === 'markdown') return { findings, scannedCount, skippedCount };
        const patterns = [
            { re: /(?:^|\s)\*\*[^*]+\*\*(?:\s|$)/gm, desc: '**bold**' },
            { re: /(?:^|\s)\*[^*]+\*(?:\s|$)/gm, desc: '*italic*' },
            { re: /(?:^|\s)__[^_]+__(?:\s|$)/gm, desc: '__bold__' },
            { re: /(?:^|\s)_[^_]+_(?:\s|$)/gm, desc: '_italic_' },
            { re: /^#{1,6}\s/gm, desc: '# naslov' },
            { re: /^\s*[-*+]\s/gm, desc: '- lista' },
            { re: /^\s*>\s/gm, desc: '> blockquote' },
            { re: /\[([^\]]+)\]\([^)]+\)/g, desc: '[link](url)' },
            { re: /```/g, desc: '```' },
            { re: /(?:^|\s)`[^`]+`(?:\s|$)/g, desc: '`code`' },
        ];
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            for (const pat of patterns) { pat.re.lastIndex = 0; let m; while ((m = pat.re.exec(el.text)) !== null) { findings.push(makeFinding({ element: el, category: 'Markdown artefakt', priority: 'OBAVEZNO', confidence: 0.92, original: m[0].trim(), replacement: '[ukloniti markdown]', rationale: `Ostatak ${pat.desc}.` })); } }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // SPACING — skips tables (checked per-cell)
    // ==========================================
    function checkSpacing(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            const text = el.text; let m;
            const dbl = /  +/g;
            while ((m = dbl.exec(text)) !== null) { const ctx = getContext(text, m.index, 25); findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.99, original: ctx, replacement: ctx.replace(/  +/g, ' '), rationale: 'Višestruki razmak.', autoFixable: true })); }
            const sbp = / +([,.:;!?])/g;
            while ((m = sbp.exec(text)) !== null) { const b = text.substring(Math.max(0,m.index-5),m.index); if (b.match(/https?:$/) || b.match(/\d$/)) continue; const ctx = getContext(text, m.index, 20); findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.97, original: ctx, replacement: ctx.replace(/ +([,.:;!?])/, '$1'), rationale: `Razmak pre \u201e${m[1]}\u201c.`, autoFixable: true })); }
            const sbc = / +([)\]}>»\u201C])/g;
            while ((m = sbc.exec(text)) !== null) { const ctx = getContext(text, m.index, 20); findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'PREPORUKA', confidence: 0.90, original: ctx, replacement: ctx.replace(/ +([)\]}>»\u201C])/, '$1'), rationale: 'Razmak pre zatvorene zagrade.', autoFixable: true })); }
            const sao = /([(\[{<«\u201E]) +/g;
            while ((m = sao.exec(text)) !== null) { const ctx = getContext(text, m.index, 20); findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'PREPORUKA', confidence: 0.90, original: ctx, replacement: ctx.replace(/([(\[{<«\u201E]) +/, '$1'), rationale: 'Razmak posle otvorene zagrade.', autoFixable: true })); }
            const nsa = /([,;:])([^\s\d"'\u201C\u201D\u201E\u2019)\]])/g;
            while ((m = nsa.exec(text)) !== null) { const c5 = text.substring(Math.max(0,m.index-10),m.index+10); if (c5.match(/https?:/) || c5.match(/\w:\\/)) continue; const ctx = getContext(text, m.index, 20); findings.push(makeFinding({ element: el, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.85, original: ctx, replacement: ctx.replace(/([,;:])(\S)/, '$1 $2'), rationale: `Nedostaje razmak posle \u201e${m[1]}\u201c.`, autoFixable: true })); }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // SCRIPT MIXING — skips tables, Unicode tokenization
    // ==========================================
    function checkScriptMixing(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            const words = el.text.match(/[\p{L}\p{M}]+/gu) || [];
            for (const word of words) {
                if (word.length < 2) continue;
                if (/[\u0400-\u04FF]/.test(word) && /[a-zA-Z\u00C0-\u024F]/.test(word)) {
                    const cc = (word.match(/[\u0400-\u04FF]/g)||[]).length;
                    const lc = (word.match(/[a-zA-Z\u00C0-\u024F]/g)||[]).length;
                    findings.push(makeFinding({ element: el, category: 'Mešanje pisama', priority: 'OBAVEZNO', confidence: 0.96, original: word, replacement: `[prebaciti na ${cc>lc?'ćirilica':'latinica'}]`, rationale: `Reč \u201e${word}\u201c: ${cc} ćir. + ${lc} lat.` }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // GREEK WITHOUT TRANSLATION
    // ==========================================
    function checkGreekWithoutTranslation(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const gRe = /[\u0370-\u03FF\u1F00-\u1FFF][\u0370-\u03FF\u1F00-\u1FFF\s,;\u00B7.'"\u0300-\u036F\u1DC0-\u1DFF]{2,}/g;
        const trNear = /[(\u201E\u201A""][^)"\u201C"]*[\u0400-\u04FFa-zA-Z\u00C0-\u024F]{5,}[^)"\u201C"]*[)"\u201C""]/;
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++; gRe.lastIndex = 0; let m;
            while ((m = gRe.exec(el.text)) !== null) {
                const after = el.text.substring(m.index+m[0].length, m.index+m[0].length+200);
                const before = el.text.substring(Math.max(0,m.index-100), m.index);
                if (trNear.test(after.substring(0,150))) continue;
                if (before.match(/[(\u201E\u201A].*[\u0400-\u04FFa-zA-Z]{5,}.*[)\u201C"]/)) continue;
                if (before.match(/\(\s*$/) && after.match(/^\s*\)/)) continue;
                const snip = m[0].length > 40 ? m[0].substring(0,40)+'...' : m[0];
                findings.push(makeFinding({ element: el, category: 'Grčki bez prevoda', priority: 'PROVERITI', confidence: 0.75, original: snip, replacement: '[dodati prevod]', rationale: 'Grčki bez prevoda u blizini.', requiresSourceVerification: true }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // DUPLICATE WORDS
    // ==========================================
    function checkDuplicateWords(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const allowed = new Set(['ha','da','ne','vrlo','još','baš','sve']);
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            // Check per-paragraph (split on newlines) to avoid false positives across paragraphs
            const lines = el.text.split('\n');
            for (const line of lines) {
                const re = /(?<=\s|^)(\p{L}+)\s+\1(?=\s|[,.:;!?)\]\}]|$)/giu; let m;
                while ((m = re.exec(line)) !== null) {
                    if (allowed.has(m[1].toLowerCase()) || m[1].length < 2) continue;
                    findings.push(makeFinding({ element: el, category: 'Duple reči', priority: 'OBAVEZNO', confidence: 0.95, original: m[0], replacement: m[1], rationale: `Ponovljena reč \u201e${m[1]}\u201c.`, autoFixable: true }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // TOC VS HEADINGS
    // ==========================================
    function checkTocVsHeadings(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const headings = docMap.elements.filter(e => e.type === 'heading' && e.text.trim()).map(e => ({ text: e.text.trim(), level: e.headingLevel, id: e.id }));
        // Match TOC entries: dots OR tab before page number
        const tocPat = /^(.+?)(?:\.{3,}|\t)\s*\d+\s*$/;
        const tocEntries = [];
        for (const el of docMap.elements) { if (el.type !== 'paragraph') continue; const m2 = el.text.match(tocPat); if (m2) tocEntries.push({ text: m2[1].trim(), element: el }); }
        scannedCount = tocEntries.length;
        for (const te of tocEntries) {
            const norm = te.text.replace(/\s+/g,' ').toLowerCase();
            if (headings.find(h => h.text.replace(/\s+/g,' ').toLowerCase() === norm)) continue;
            const close = headings.find(h => levenshteinDistance(h.text.toLowerCase(), norm) <= 3);
            if (close) { findings.push(makeFinding({ element: te.element, category: 'TOC/naslovi', priority: 'OBAVEZNO', confidence: 0.88, original: `TOC: \u201e${te.text}\u201c`, replacement: `Uskladiti: \u201e${close.text}\u201c`, rationale: 'Razlika TOC i naslova.' })); }
            else { findings.push(makeFinding({ element: te.element, category: 'TOC/naslovi', priority: 'PROVERITI', confidence: 0.70, original: `TOC: \u201e${te.text}\u201c`, replacement: '[proveriti]', rationale: 'TOC stavka bez odgovarajućeg naslova.' })); }
        }

        // Reverse check: headings not in TOC (only if TOC exists)
        if (tocEntries.length > 0) {
            for (const h of headings) {
                const normH = h.text.replace(/\s+/g,' ').toLowerCase();
                const inToc = tocEntries.some(te => te.text.replace(/\s+/g,' ').toLowerCase() === normH || levenshteinDistance(te.text.toLowerCase(), normH) <= 3);
                if (!inToc) {
                    const hEl = docMap.elements.find(e => e.id === h.id) || { id: h.id, type: 'heading', text: h.text, section: h.text, isDirectQuote: false };
                    findings.push(makeFinding({ element: hEl, category: 'TOC/naslovi', priority: 'PROVERITI', confidence: 0.70, original: `Naslov: \u201e${h.text}\u201c`, replacement: '[dodati u sadržaj]', rationale: 'Naslov postoji u dokumentu ali nije u sadržaju.' }));
                }
            }
        }

        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // NUMBERING — numeric formats only, grouped by consecutive runs,
    // proper expectedLabel via formatter
    // ==========================================
    function checkNumbering(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;

        // Strategy 1: OOXML data (numeric formats only)
        const ooxmlNumbered = docMap.elements.filter(el =>
            el.numId && el.displayedNumber != null &&
            el.numFmt && el.numFmt !== 'bullet' && el.numFmt !== 'none');

        if (ooxmlNumbered.length > 1) {
            scannedCount = ooxmlNumbered.length;
            const byList = {};
            for (const el of ooxmlNumbered) {
                const key = el.listInstanceId || `${el.numId}-${el.numLevel||0}`;
                if (!byList[key]) byList[key] = [];
                byList[key].push(el);
            }
            for (const [, items] of Object.entries(byList)) {
                for (let i = 1; i < items.length; i++) {
                    const prev = items[i-1], curr = items[i];
                    const expected = prev.displayedNumber + 1;
                    if (curr.displayedNumber !== expected) {
                        // Use stored lvlTextPattern for proper expected label
                        const lvlText = curr.lvlTextPattern || '%1.';
                        const expLabel = lvlText.replace(/%\d/, String(expected));
                        findings.push(makeFinding({ element: curr, category: 'Numeracija', priority: 'OBAVEZNO', confidence: 0.92, original: `Stavka ${curr.displayedLabel||curr.displayedNumber} (prethodno: ${prev.displayedLabel||prev.displayedNumber})`, replacement: `Očekivano: ${expLabel}`, rationale: `Numeracija preskače sa ${prev.displayedNumber} na ${curr.displayedNumber}.` }));
                    }
                }
            }
            // Don't return early — also run text-based strategy below for manual lists
        }

        // Strategy 2: Text-based, grouped by CONSECUTIVE runs
        // A plain paragraph, heading, or table breaks a list
        const lists = [];
        let currentList = null;
        for (const el of docMap.elements) {
            if (!el.text) continue;
            const m = el.text.match(/^\s*(\d+)[.)]\s/);
            if (m) {
                if (!currentList) currentList = [];
                currentList.push({ num: parseInt(m[1], 10), element: el });
            } else {
                if (currentList) { lists.push(currentList); currentList = null; }
            }
        }
        if (currentList) lists.push(currentList);

        for (const list of lists) {
            scannedCount += list.length;
            for (let i = 1; i < list.length; i++) {
                const prev = list[i-1], curr = list[i];
                if (curr.num === 1) continue;
                const expected = prev.num + 1;
                if (curr.num !== expected) {
                    findings.push(makeFinding({ element: curr.element, category: 'Numeracija', priority: 'OBAVEZNO', confidence: 0.85, original: `Stavka ${curr.num} (prethodno: ${prev.num})`, replacement: `Očekivano: ${expected}`, rationale: `Numeracija preskače sa ${prev.num} na ${curr.num}.` }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }




    // ==========================================
    // BIBLIOGRAPHY — continues through subheadings until same/higher level heading
    // ==========================================
    function checkBibliography(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const bibPattern = /^(bibliografija|literatura|izvori i literatura|references|works cited|bibliography|primarni izvori|sekundarni izvori|arhivski izvori|korišćena literatura|spisak literature|библиографија|литература|примарни извори|секундарни извори|архивски извори)$/i;

        // Find ALL bibliography headings
        const bibIndices = [];
        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type === 'heading' && el.text.trim().match(bibPattern)) {
                bibIndices.push(i);
            }
        }
        if (bibIndices.length === 0) return { findings, scannedCount, skippedCount };

        const processedElementIndices = new Set();
        for (const bibIdx of bibIndices) {

        const bibHeading = docMap.elements[bibIdx];
        const bibLevel = bibHeading.headingLevel || 1;

        // Collect entries: continue through lower-level headings, stop at same/higher
        const bibEntries = [];
        for (let i = bibIdx + 1; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type === 'heading') {
                const hl = el.headingLevel || 1;
                if (hl <= bibLevel) break; // Same or higher level → stop
                continue; // Lower level subheading → skip but continue
            }
            if (el.type === 'paragraph' && el.text.trim().length > 10 && !processedElementIndices.has(i)) {
                bibEntries.push(el);
                processedElementIndices.add(i);
            }
        }
        scannedCount += bibEntries.length;
        if (bibEntries.length === 0) continue;

        for (const entry of bibEntries) {
            const text = entry.text.trim();

            // Skip ancient/classical sources (no modern publication year expected)
            const isAncient = text.match(/\b(Herodot(us)?|Thucydides|Plutarch|Diodorus|Strabo|Pausanias|Appian|Apijan|Josephus|Flavius|Pseudo-|Homer|Hesiod|Plato|Platon|Aristotle|Aristot|Cicero|Tacitus|Livius|Plinius|Ptolemy|Euripides|Sophocles|Aeschylus|Xenophon|Polybius|Apollodorus|Bibliotheca|Historiae|Antiquitates|Annales)\b/i);
            // Skip electronic/web sources (URLs present)
            const isElectronic = text.match(/https?:\/\//);
            // Skip religious texts
            const isReligious = text.match(/\b(Biblija|Bible|Sveto Pismo|Quran|Talmud|Torah)\b/i);

            if (!isAncient && !isElectronic && !isReligious) {
                if (!text.match(/\b(1[5-9]\d{2}|20\d{2})\b/)) {
                    findings.push(makeFinding({ element: entry, category: 'Bibliografija', priority: 'PROVERITI', confidence: 0.75, original: text.substring(0,60)+(text.length>60?'...':''), replacement: '[dodati godinu]', rationale: 'Bez godine izdanja.', requiresSourceVerification: true }));
                }
                if (!text.match(/:\s*[A-Z\u0400-\u04FF]/) && !text.match(/University|Press|Verlag|izdava/i) &&
                    // Skip journal articles (have quoted title, journal name, or vol/issue indicators)
                    !text.match(/[„"\u201E\u201C\u00AB\u00BB]/) && !text.match(/\b(Journal|Review|Bulletin|Proceedings|Annals|Quarterly|Vol\.|Issue|pp\.|str\.)\b/i)) {
                    findings.push(makeFinding({ element: entry, category: 'Bibliografija', priority: 'PROVERITI', confidence: 0.60, original: text.substring(0,60)+(text.length>60?'...':''), replacement: '[proveriti izdavača]', rationale: 'Moguć nedostatak izdavača.', requiresSourceVerification: true }));
                }
            }
        }

        // Cross-reference: use body text before FIRST bib section, check against UNION of all bib entries
        const allBibEntries = [];
        const allBibText = processedElementIndices; // reuse the set of processed indices
        for (const bi of bibIndices) {
            const bh = docMap.elements[bi];
            const bl = bh.headingLevel || 1;
            for (let i = bi + 1; i < docMap.elements.length; i++) {
                const el = docMap.elements[i];
                if (el.type === 'heading' && (el.headingLevel || 1) <= bl) break;
                if (el.type === 'paragraph' && el.text.trim().length > 10) allBibEntries.push(el);
            }
        }
        const firstBibIdx = bibIndices[0];
        const excludeWords = new Set(['Beograd','Zagreb','Sarajevo','London','Oxford','Cambridge','Berlin','Paris','Roma','Athens','Chicago','Moscow','Moskva','Atina','Review','Journal','Press','University','Institut','Academy','Society','Edition','Volume','Chapter']);
        const citPats = [/\((\p{Lu}\p{Ll}{2,}),?\s*\d{4}\)/gu, /(\p{Lu}\p{Ll}{2,})\s*\(\d{4}\)/gu];
        const cited = new Set();
        for (const el of docMap.elements.slice(0, firstBibIdx)) {
            if (!el.text) continue;
            for (const p of citPats) { p.lastIndex = 0; let m2; while ((m2 = p.exec(el.text)) !== null) {
                if (!excludeWords.has(m2[1])) cited.add(m2[1]);
            }}
        }
        for (const author of cited) {
            if (!allBibEntries.some(e => e.text.includes(author))) {
                findings.push(makeFinding({ element: docMap.elements[firstBibIdx], category: 'Bibliografija', priority: 'PROVERITI', confidence: 0.70, original: `Citiran: \u201e${author}\u201c`, replacement: '[dodati u bibliografiju]', rationale: `\u201e${author}\u201c citiran ali nije u bibliografiji.`, requiresSourceVerification: true }));
            }
        }
        } // end for (const bibIdx of bibIndices)
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // URLs
    // ==========================================
    function checkUrls(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g;
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; }
            scannedCount++; urlRe.lastIndex = 0; let m;
            while ((m = urlRe.exec(el.text)) !== null) {
                const url = m[0];
                const openP = (url.match(/\(/g)||[]).length, closeP = (url.match(/\)/g)||[]).length;
                let trailing = url.match(/[,.;]+$/);
                if (trailing) { findings.push(makeFinding({ element: el, category: 'URL', priority: 'PROVERITI', confidence: 0.80, original: url, replacement: url.replace(/[,.;]+$/, ''), rationale: 'URL završen interpunkcijom.' })); }
                else if (url.match(/\)$/) && closeP > openP) { findings.push(makeFinding({ element: el, category: 'URL', priority: 'PROVERITI', confidence: 0.70, original: url, replacement: url.replace(/\)$/, ''), rationale: 'Zatvorena zagrada verovatno nije deo URL-a.' })); }
                if (url.length > 100) { findings.push(makeFinding({ element: el, category: 'URL', priority: 'PREPORUKA', confidence: 0.65, original: url.substring(0,50)+'...', replacement: '[proveriti]', rationale: 'Veoma dugačak URL.' })); }
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // FOOTNOTES
    // ==========================================
    function checkFootnotes(docMap, options) {
        const findings = [];
        const allNotes = [...(docMap.footnotes||[]), ...(docMap.endnotes||[])];
        const scannedCount = allNotes.length;

        for (const note of allNotes) {
            const isFootnote = (docMap.footnotes||[]).includes(note);
            const noteType = isFootnote ? 'Fusnota' : 'Endnota';
            const noteEl = { id: `${isFootnote?'fn':'en'}-${note.id}`, type: isFootnote?'footnote':'endnote', text: note.text, section: `(${noteType.toLowerCase()} ${note.id})`, isDirectQuote: false };

            if (note.isEmpty) {
                findings.push(makeFinding({ element: noteEl, category: 'Fusnote', priority: 'OBAVEZNO', confidence: 0.99, original: `[${noteType} ${note.id} \u2014 prazna]`, replacement: '[dodati sadržaj]', rationale: `Prazna ${noteType.toLowerCase()}.` }));
                continue;
            }

            const text = note.text;

            // Detect if footnote content is a direct quote
            const fnIsQuote = /^\u201E[\s\S]*\u201C$/.test(text.trim()) ||
                /^\u00AB[\s\S]*\u00BB$/.test(text.trim());

            // Only run sub-checks that are enabled in options
            if (options.spacing) {
                let m;
                const dbl = /  +/g;
                while ((m = dbl.exec(text)) !== null) {
                    const ctx = getContext(text, m.index, 20);
                    findings.push(makeFinding({ element: noteEl, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.99, original: ctx, replacement: ctx.replace(/  +/g, ' '), rationale: `Višestruki razmak u ${noteType.toLowerCase()}.`, autoFixable: true }));
                }
                const sbp = / +([,.:;!?])/g;
                while ((m = sbp.exec(text)) !== null) {
                    const b = text.substring(Math.max(0,m.index-5),m.index);
                    if (b.match(/https?:$/) || b.match(/\d$/)) continue;
                    const ctx = getContext(text, m.index, 20);
                    findings.push(makeFinding({ element: noteEl, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.97, original: ctx, replacement: ctx.replace(/ +([,.:;!?])/, '$1'), rationale: `Razmak pre \u201e${m[1]}\u201c u ${noteType.toLowerCase()}.`, autoFixable: true }));
                }
                const sbc = / +([)\]}>»\u201C])/g;
                while ((m = sbc.exec(text)) !== null) {
                    const ctx = getContext(text, m.index, 20);
                    findings.push(makeFinding({ element: noteEl, category: 'Razmaci', priority: 'PREPORUKA', confidence: 0.90, original: ctx, replacement: ctx.replace(/ +([)\]}>»\u201C])/, '$1'), rationale: `Razmak pre zatvorene zagrade u ${noteType.toLowerCase()}.`, autoFixable: true }));
                }
                const nsa = /([,;:])([^\s\d"'\u201C\u201D\u201E\u2019)\]])/g;
                while ((m = nsa.exec(text)) !== null) {
                    if (text.substring(Math.max(0,m.index-10),m.index+10).match(/https?:|:\\/)) continue;
                    const ctx = getContext(text, m.index, 20);
                    findings.push(makeFinding({ element: noteEl, category: 'Razmaci', priority: 'OBAVEZNO', confidence: 0.85, original: ctx, replacement: ctx.replace(/([,;:])(\S)/, '$1 $2'), rationale: `Nedostaje razmak posle \u201e${m[1]}\u201c u ${noteType.toLowerCase()}.`, autoFixable: true }));
                }
            }

            if (options.brackets) {
                for (const [o,c] of [['(',')'],['[',']'],['{','}']]) {
                    let depth = 0; let prematureClose = false;
                    for (let i = 0; i < text.length; i++) {
                        if (text[i]===o) depth++;
                        else if (text[i]===c) { depth--; if (depth < 0) { prematureClose = true; depth = 0; } }
                    }
                    if (depth !== 0 || prematureClose) {
                        const rationale = prematureClose
                            ? `Prerano zatvorena zagrada ${c} u ${noteType.toLowerCase()} ${note.id}.`
                            : `Neuparene zagrade u ${noteType.toLowerCase()} ${note.id}.`;
                        findings.push(makeFinding({ element: noteEl, category: 'Zagrade', priority: 'OBAVEZNO', confidence: 0.95, original: text.substring(0, 50), replacement: `[neuparene zagrade ${o}${c}]`, rationale }));
                    }
                }
            }

            if (options.scriptMix) {
                const words = text.match(/[\p{L}\p{M}]+/gu) || [];
                for (const word of words) {
                    if (word.length < 2) continue;
                    if (/[\u0400-\u04FF]/.test(word) && /[a-zA-Z\u00C0-\u024F]/.test(word)) {
                        findings.push(makeFinding({ element: noteEl, category: 'Mešanje pisama', priority: 'OBAVEZNO', confidence: 0.96, original: word, replacement: '[prebaciti na jedno pismo]', rationale: `Pomešana slova u ${noteType.toLowerCase()}: \u201e${word}\u201c.` }));
                    }
                }
            }

            if (options.quotes) {
                const straightCount = (text.match(/"/g) || []).length;
                if (straightCount > 0) {
                    findings.push(makeFinding({ element: noteEl, category: 'Tipografija', priority: 'OBAVEZNO', confidence: 0.95, original: `[${straightCount} ravnih navodnika u ${noteType.toLowerCase()}]`, replacement: '[zameniti tipografskim]', rationale: `${straightCount} ravnih navodnika u ${noteType.toLowerCase()} ${note.id}.`, autoFixable: true }));
                }
            }

            if (options.duplicates) {
                const dupeRe = /(?<=\s|^)(\p{L}+)\s+\1(?=\s|[,.:;!?)\]\}]|$)/giu;
                const allowedDupes = new Set(['ha','da','ne','vrlo','još','baš','sve']);
                let md;
                while ((md = dupeRe.exec(text)) !== null) {
                    if (allowedDupes.has(md[1].toLowerCase()) || md[1].length < 2) continue;
                    findings.push(makeFinding({ element: noteEl, category: 'Duple reči', priority: 'OBAVEZNO', confidence: 0.95, original: md[0], replacement: md[1], rationale: `Ponovljena reč u ${noteType.toLowerCase()}: \u201e${md[1]}\u201c.`, autoFixable: true }));
                }
            }

            if (options.urls) {
                const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g;
                let mu;
                while ((mu = urlRe.exec(text)) !== null) {
                    const url = mu[0];
                    if (url.match(/[,.;]+$/)) {
                        findings.push(makeFinding({ element: noteEl, category: 'URL', priority: 'PROVERITI', confidence: 0.80, original: url, replacement: url.replace(/[,.;]+$/, ''), rationale: `URL u ${noteType.toLowerCase()} završen interpunkcijom.` }));
                    }
                }
            }

            // Apply direct quote protection to footnote findings
            if (fnIsQuote) {
                for (const f of findings) {
                    if (f.paragraphId === noteEl.id && f.priority !== 'PROVERITI') {
                        f.priority = 'PROVERITI'; f.autoFixable = false;
                        f.requiresSourceVerification = true; f.isDirectQuote = true;
                    }
                }
            }
        }
        return { findings, scannedCount, skippedCount: 0 };
    }

    // ==========================================
    // REPEATED PARAGRAPHS
    // ==========================================
    function checkRepetition(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const seen = new Map();
        for (const el of docMap.elements) {
            if (!el.text || el.text.trim().length < 50 || el.type === 'heading' || el.type === 'table') { skippedCount++; continue; }
            scannedCount++;
            const norm = el.text.trim().replace(/\s+/g,' ').toLowerCase();
            if (seen.has(norm)) { findings.push(makeFinding({ element: el, category: 'Ponavljanje', priority: 'OBAVEZNO', confidence: 0.98, original: el.text.substring(0,80)+'...', replacement: '[ukloniti]', rationale: `Ponovljen pasus (${seen.get(norm).id}).` })); }
            else seen.set(norm, el);
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // ALL-CAPS — Unicode tokenization (no \b for Cyrillic)
    // ==========================================
    function checkAllCaps(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        const known = new Set(['UNESCO','NATO','EU','SAD','SSSR','DNA','RNA','URL','HTML','CSS','JS','PDF','DOCX','ISBN','ISSN','DOI','NB','PS','AD','BC','PhD','USA','UK','ID','OK','IT','PR','HR','TV','CD','DVD','USB','PC','OS','AI','NSP']);

        // Collect words from headings — these are likely intentional title words
        const headingWords = new Set();
        for (const el of docMap.elements) {
            if (el.type === 'heading' && el.text) {
                const hw = el.text.match(/[\p{L}\p{M}]+/gu) || [];
                hw.forEach(w => { if (w === w.toUpperCase() && w.length >= 3) headingWords.add(w); });
            }
        }

        // Collect ALL-CAPS words from table cells (likely table headers)
        const tableCapsWords = new Set();
        for (const el of docMap.elements) {
            if (el.type === 'table' && el.rows) {
                for (const row of el.rows) {
                    for (const cell of row) {
                        if (!cell.text) continue;
                        const cw = cell.text.match(/[\p{L}\p{M}]+/gu) || [];
                        cw.forEach(w => { if (w === w.toUpperCase() && w.length >= 3) tableCapsWords.add(w); });
                    }
                }
            }
        }

        // Find first heading index — skip everything before it (title page)
        const firstHeadingIdx = docMap.elements.findIndex(e => e.type === 'heading');

        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (!el.text || el.type === 'heading' || el.type === 'table') { skippedCount++; continue; }
            // Skip title page (before first heading)
            if (firstHeadingIdx >= 0 && i < firstHeadingIdx) { skippedCount++; continue; }
            scannedCount++;
            const words = el.text.match(/[\p{L}\p{M}]+/gu) || [];
            for (const word of words) {
                if (word.length < 3) continue;
                if (word === word.toUpperCase() && word !== word.toLowerCase()) {
                    if (known.has(word)) continue;
                    if (/^[IVXLCDM]+$/.test(word)) continue;
                    if (headingWords.has(word)) continue;
                    if (tableCapsWords.has(word)) continue;
                    findings.push(makeFinding({ element: el, category: 'ALL-CAPS', priority: 'PREPORUKA', confidence: 0.70, original: word, replacement: '[normalizovati ako nije skraćenica]', rationale: `\u201e${word}\u201c ALL-CAPS u telu teksta.` }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // EMPTY HEADINGS — allow H1→H2, H2→H3 (normal nesting)
    // ==========================================
    function checkEmptyHeadings(docMap) {
        const findings = []; let scannedCount = 0, skippedCount = 0;
        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type !== 'heading') { skippedCount++; continue; }
            scannedCount++;
            if (!el.text || !el.text.trim()) {
                findings.push(makeFinding({ element: el, category: 'Struktura', priority: 'OBAVEZNO', confidence: 0.99, original: `[Prazan naslov nivoa ${el.headingLevel}]`, replacement: '[dodati tekst]', rationale: 'Naslov bez sadržaja.' }));
            }
            // Check heading followed by another heading — only report if same or HIGHER level (not normal nesting)
            if (i < docMap.elements.length - 1) {
                const next = docMap.elements[i + 1];
                if (next.type === 'heading') {
                    const currLvl = el.headingLevel || 1;
                    const nextLvl = next.headingLevel || 1;
                    // H1→H2 is fine (nesting). H2→H1 or H2→H2 without content is suspect.
                    if (nextLvl <= currLvl) {
                        findings.push(makeFinding({ element: el, category: 'Struktura', priority: 'PREPORUKA', confidence: 0.75, original: el.text || '[prazan]', replacement: '[dodati sadržaj ispod]', rationale: 'Naslov bez sadržaja pre naslova istog/višeg nivoa.' }));
                    }
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // HELPERS
    // ==========================================
    function makeFinding({ element, category, priority, confidence, original,
        replacement, rationale, autoFixable = false, globalPattern = false,
        requiresSourceVerification = false,
        tableId = null, rowId = null, cellId = null, rowIndex = null, columnIndex = null }) {
        return {
            id: nextId(), section: getSectionName(element), paragraphId: element.id || null,
            category, priority, confidence, original, replacement, rationale,
            isDirectQuote: element.isDirectQuote || false,
            requiresSourceVerification, autoFixable, globalPattern, status: 'OPEN',
            tableId, rowId, cellId, rowIndex, columnIndex,
            _element: element,
        };
    }

    function getSectionName(el) {
        if (el.type === 'heading') return el.text || '(bez naslova)';
        if (el.section) return el.section;
        return el.id || '(nepoznata lokacija)';
    }

    function getContext(text, pos, radius) {
        const s = Math.max(0, pos - radius), e = Math.min(text.length, pos + radius);
        let ctx = text.substring(s, e);
        if (s > 0) ctx = '...' + ctx;
        if (e < text.length) ctx += '...';
        return ctx;
    }

    function levenshteinDistance(a, b) {
        if (!a.length) return b.length; if (!b.length) return a.length;
        const mx = []; for (let i = 0; i <= b.length; i++) mx[i] = [i];
        for (let j = 0; j <= a.length; j++) mx[0][j] = j;
        for (let i = 1; i <= b.length; i++) for (let j = 1; j <= a.length; j++)
            mx[i][j] = b[i-1]===a[j-1] ? mx[i-1][j-1] : Math.min(mx[i-1][j-1]+1, mx[i][j-1]+1, mx[i-1][j]+1);
        return mx[b.length][a.length];
    }

    return { runAudit };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = RuleEngine; }
