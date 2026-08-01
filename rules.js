/**
 * Rule Engine Module
 * Deterministic, rule-based checks for document auditing
 * No AI - pure regex, pattern matching, and structural analysis
 */

const RuleEngine = (() => {
    'use strict';

    let findingCounter = 0;
    function resetCounter() { findingCounter = 0; }
    function nextId() {
        findingCounter++;
        return `F-${String(findingCounter).padStart(4, '0')}`;
    }

    /**
     * Run all enabled checks against the document map.
     * Each check returns { findings, scannedCount, skippedCount }.
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
            { key: 'dashes', name: 'Crtice i rasponi', fn: checkDashes },
            { key: 'bibliography', name: 'Bibliografija', fn: checkBibliography },
            { key: 'urls', name: 'URL-ovi', fn: checkUrls },
            { key: 'footnotes', name: 'Fusnote', fn: checkFootnotes },
            { key: 'repetition', name: 'Ponovljeni pasusi', fn: checkRepetition },
            { key: 'capsWords', name: 'ALL-CAPS reči', fn: checkAllCaps },
            { key: 'emptyHeadings', name: 'Prazni naslovi', fn: checkEmptyHeadings },
        ];

        for (const check of checks) {
            if (!options[check.key]) continue;

            const result = check.fn(docMap);

            // Apply direct-quote handling: findings inside quotes get demoted
            for (const f of result.findings) {
                if (f._element && f._element.isDirectQuote) {
                    f.priority = 'PROVERITI';
                    f.autoFixable = false;
                    f.requiresSourceVerification = true;
                    f.isDirectQuote = true;
                }
                // Remove internal _element reference before output
                delete f._element;
            }

            findings.push(...result.findings);

            if (result.findings.length === 0) {
                passedChecks.push({
                    area: check.name,
                    result: 'Bez grešaka',
                    count: result.scannedCount,
                });
            }
        }

        // Additional pass: table cell-level checks (respects enabled options)
        checkTableCells(docMap, findings, options);

        return { findings, passedChecks };
    }

    /**
     * Additional pass: check table cells individually.
     * Only runs checks that are enabled in options.
     * Findings include tableId/rowId/cellId/rowIndex/columnIndex.
     */
    function checkTableCells(docMap, findings, options) {
        const doSpacing = options.spacing === true;
        const doBrackets = options.brackets === true;
        const doScriptMix = options.scriptMix === true;

        if (!doSpacing && !doBrackets && !doScriptMix) return;

        for (const el of docMap.elements) {
            if (el.type !== 'table' || !el.rows) continue;

            for (const row of el.rows) {
                for (const cell of row) {
                    if (!cell.text || cell.text.trim().length === 0) continue;
                    const text = cell.text;
                    const cellMeta = {
                        tableId: cell.tableId,
                        rowId: cell.rowId,
                        cellId: cell.cellId,
                        rowIndex: cell.rowIndex,
                        columnIndex: cell.columnIndex,
                    };

                    if (doSpacing) {
                        const doubleSpace = /  +/g;
                        let m;
                        while ((m = doubleSpace.exec(text)) !== null) {
                            const ctx = getContext(text, m.index, 20);
                            findings.push(makeFinding({
                                element: el, category: 'Razmaci', priority: 'OBAVEZNO',
                                confidence: 0.99, original: ctx,
                                replacement: ctx.replace(/  +/g, ' '),
                                rationale: 'Višestruki razmak u ćeliji tabele.',
                                autoFixable: true, ...cellMeta,
                            }));
                        }
                    }

                    if (doScriptMix) {
                        const words = text.match(/[\p{L}\p{M}]+/gu) || [];
                        for (const word of words) {
                            if (word.length < 2) continue;
                            const hasCyr = /[\u0400-\u04FF]/.test(word);
                            const hasLat = /[a-zA-Z\u00C0-\u024F]/.test(word);
                            if (hasCyr && hasLat) {
                                findings.push(makeFinding({
                                    element: el, category: 'Mešanje pisama',
                                    priority: 'OBAVEZNO', confidence: 0.96,
                                    original: word,
                                    replacement: '[prebaciti celu reč na jedno pismo]',
                                    rationale: `Pomešana slova u ćeliji tabele: \u201e${word}\u201c.`,
                                    ...cellMeta,
                                }));
                            }
                        }
                    }

                    if (doBrackets) {
                        for (const [open, close] of [['(',')'],['[',']']]) {
                            let depth = 0;
                            for (let i = 0; i < text.length; i++) {
                                if (text[i] === open) depth++;
                                else if (text[i] === close) depth--;
                            }
                            if (depth !== 0) {
                                findings.push(makeFinding({
                                    element: el, category: 'Zagrade',
                                    priority: 'OBAVEZNO', confidence: 0.95,
                                    original: text.substring(0, 50),
                                    replacement: `[neuparene zagrade ${open}${close} u ćeliji]`,
                                    rationale: `Neuparene zagrade u ćeliji tabele (${cell.cellId}).`,
                                    ...cellMeta,
                                }));
                            }
                        }
                    }
                }
            }
        }
    }


    // ==========================================
    // CHECK: Unbalanced brackets
    // ==========================================
    function checkBrackets(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const pairs = [['(', ')'], ['[', ']'], ['{', '}']];
        const pairNames = { '(': 'obla zagrada', '[': 'uglasta zagrada', '{': 'vitičasta zagrada' };

        for (const el of docMap.elements) {
            if (!el.text || el.text.trim().length === 0) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; } // checked per-cell
            scannedCount++;

            for (const [open, close] of pairs) {
                let depth = 0;
                let positions = [];

                for (let i = 0; i < el.text.length; i++) {
                    if (el.text[i] === open) { depth++; positions.push(i); }
                    else if (el.text[i] === close) {
                        depth--;
                        if (depth < 0) {
                            const ctx = getContext(el.text, i, 40);
                            findings.push(makeFinding({
                                element: el,
                                category: 'Zagrade',
                                priority: 'OBAVEZNO',
                                confidence: 0.98,
                                original: ctx,
                                replacement: `[ukloniti višak \u201e${close}\u201c ili dodati \u201e${open}\u201c]`,
                                rationale: `Zatvorena ${pairNames[open]} bez odgovarajuće otvorene.`,
                            }));
                            depth = 0;
                        }
                    }
                }

                if (depth > 0) {
                    const lastOpen = positions[positions.length - 1];
                    const ctx = getContext(el.text, lastOpen, 40);
                    findings.push(makeFinding({
                        element: el,
                        category: 'Zagrade',
                        priority: 'OBAVEZNO',
                        confidence: 0.98,
                        original: ctx,
                        replacement: `[dodati \u201e${close}\u201c ili ukloniti \u201e${open}\u201c]`,
                        rationale: `Otvorena ${pairNames[open]} bez odgovarajuće zatvorene.`,
                    }));
                }
            }

            // Special check: ; where ) should be
            const semiPattern = /\([^)]*;(?=[^(]*$)/g;
            let match;
            while ((match = semiPattern.exec(el.text)) !== null) {
                const fragment = el.text.substring(match.index);
                if (fragment.indexOf(')') === -1 || fragment.indexOf(';') < fragment.indexOf(')')) {
                    const ctx = getContext(el.text, match.index, 50);
                    findings.push(makeFinding({
                        element: el,
                        category: 'Zagrade',
                        priority: 'PROVERITI',
                        confidence: 0.75,
                        original: ctx,
                        replacement: '[proveriti da li \u201e;\u201c treba da bude \u201e)\u201c]',
                        rationale: 'Tačka-zarez unutar nezatvorene zagrade \u2014 mogući artefakt find-replace operacije.',
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Quotes (straight vs typographic)
    // Consolidates once for ENTIRE DOCUMENT, not per paragraph
    // ==========================================
    function checkQuotes(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        let totalStraightQuotes = 0;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;

            const straightMatches = el.text.match(/"/g) || [];
            totalStraightQuotes += straightMatches.length;

            // Unmatched typographic quotes (per-element, this is structural)
            const openDouble = (el.text.match(/\u201E/g) || []).length;
            const closeDouble = (el.text.match(/\u201C/g) || []).length;
            if (openDouble !== closeDouble) {
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'OBAVEZNO',
                    confidence: 0.90,
                    original: `[${openDouble}\u00D7 \u201e vs ${closeDouble}\u00D7 \u201c]`,
                    replacement: '[upariti otvarajuće i zatvarajuće navodnike]',
                    rationale: `Broj otvarajućih navodnika \u201e ne odgovara broju zatvarajućih \u201c.`,
                }));
            }
        }

        // Report straight quotes once for entire document
        // Use a synthetic element (not tied to any specific paragraph)
        // to avoid accidental direct-quote priority demotion
        if (totalStraightQuotes > 0) {
            const syntheticElement = {
                id: 'doc-global',
                type: 'document',
                text: '',
                section: '(ceo dokument)',
                isDirectQuote: false,
                quoteConfidence: 0,
            };
            findings.push(makeFinding({
                element: syntheticElement,
                category: 'Tipografija',
                priority: 'OBAVEZNO',
                confidence: 0.95,
                original: `[${totalStraightQuotes} ravnih navodnika u dokumentu]`,
                replacement: '[zameniti sve ravne navodnike tipografskim: \u201e...\u201c]',
                rationale: `Pronađeno ukupno ${totalStraightQuotes} ravnih navodnika u celom dokumentu. Srpski standard: \u201e...\u201c (spoljni) ili \u2018...\u2019 (unutrašnji).`,
                autoFixable: true,
                globalPattern: true,
            }));
        }

        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Markdown artifacts
    // ==========================================
    function checkMarkdownArtifacts(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        if (docMap.type === 'markdown') return { findings, scannedCount, skippedCount };

        const patterns = [
            { re: /(?:^|\s)\*\*[^*]+\*\*(?:\s|$)/gm, desc: 'boldovani markdown (**tekst**)' },
            { re: /(?:^|\s)\*[^*]+\*(?:\s|$)/gm, desc: 'kurzivni markdown (*tekst*)' },
            { re: /(?:^|\s)__[^_]+__(?:\s|$)/gm, desc: 'boldovani markdown (__tekst__)' },
            { re: /(?:^|\s)_[^_]+_(?:\s|$)/gm, desc: 'kurzivni markdown (_tekst_)' },
            { re: /^#{1,6}\s/gm, desc: 'markdown naslov (#)' },
            { re: /^\s*[-*+]\s/gm, desc: 'markdown lista (- ili *)' },
            { re: /^\s*>\s/gm, desc: 'markdown blockquote (>)' },
            { re: /\[([^\]]+)\]\([^)]+\)/g, desc: 'markdown link [tekst](url)' },
            { re: /```/g, desc: 'markdown code block (```)' },
            { re: /(?:^|\s)`[^`]+`(?:\s|$)/g, desc: 'markdown inline code (`kod`)' },
        ];

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;

            for (const pat of patterns) {
                pat.re.lastIndex = 0;
                let m;
                while ((m = pat.re.exec(el.text)) !== null) {
                    findings.push(makeFinding({
                        element: el,
                        category: 'Markdown artefakt',
                        priority: 'OBAVEZNO',
                        confidence: 0.92,
                        original: m[0].trim(),
                        replacement: '[ukloniti markdown sintaksu, primeniti pravo formatiranje]',
                        rationale: `Ostatak ${pat.desc} \u2014 verovatno nije pretvoreno u formatiranje.`,
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Spacing and punctuation
    // ==========================================
    function checkSpacing(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; } // checked per-cell
            scannedCount++;
            const text = el.text;
            let m;

            // Double spaces
            const doubleSpace = /  +/g;
            while ((m = doubleSpace.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 25);
                findings.push(makeFinding({
                    element: el, category: 'Razmaci', priority: 'OBAVEZNO',
                    confidence: 0.99, original: ctx,
                    replacement: ctx.replace(/  +/g, ' '),
                    rationale: 'Višestruki razmak.', autoFixable: true,
                }));
            }

            // Space before punctuation
            const spaceBeforePunct = / +([,.:;!?])/g;
            while ((m = spaceBeforePunct.exec(text)) !== null) {
                const before = text.substring(Math.max(0, m.index - 5), m.index);
                if (before.match(/https?:$/) || before.match(/\d$/)) continue;
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el, category: 'Razmaci', priority: 'OBAVEZNO',
                    confidence: 0.97, original: ctx,
                    replacement: ctx.replace(/ +([,.:;!?])/, '$1'),
                    rationale: `Razmak pre interpunkcijskog znaka \u201e${m[1]}\u201c.`,
                    autoFixable: true,
                }));
            }

            // Space before closing bracket
            const spaceBeforeClose = / +([)\]}>»\u201C])/g;
            while ((m = spaceBeforeClose.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el, category: 'Razmaci', priority: 'PREPORUKA',
                    confidence: 0.90, original: ctx,
                    replacement: ctx.replace(/ +([)\]}>»\u201C])/, '$1'),
                    rationale: 'Razmak pre zatvorene zagrade/navodnika.',
                    autoFixable: true,
                }));
            }

            // Space after opening bracket
            const spaceAfterOpen = /([(\[{<«\u201E]) +/g;
            while ((m = spaceAfterOpen.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el, category: 'Razmaci', priority: 'PREPORUKA',
                    confidence: 0.90, original: ctx,
                    replacement: ctx.replace(/([(\[{<«\u201E]) +/, '$1'),
                    rationale: 'Razmak posle otvorene zagrade/navodnika.',
                    autoFixable: true,
                }));
            }

            // No space after comma/semicolon/colon
            const noSpaceAfterPunct = /([,;:])([^\s\d"'\u201C\u201D\u201E\u2019)\]])/g;
            while ((m = noSpaceAfterPunct.exec(text)) !== null) {
                const ctx5 = text.substring(Math.max(0, m.index - 10), m.index + 10);
                if (ctx5.match(/https?:/) || ctx5.match(/\w:\\/)) continue;
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el, category: 'Razmaci', priority: 'OBAVEZNO',
                    confidence: 0.85, original: ctx,
                    replacement: ctx.replace(/([,;:])(\S)/, '$1 $2'),
                    rationale: `Nedostaje razmak posle \u201e${m[1]}\u201c.`,
                    autoFixable: true,
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Script mixing (Cyrillic/Latin in same word)
    // ==========================================
    function checkScriptMixing(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const cyrillicRe = /[\u0400-\u04FF]/;
        const latinRe = /[a-zA-Z\u00C0-\u024F]/;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            if (el.type === 'table') { skippedCount++; continue; } // checked per-cell
            scannedCount++;
            const words = el.text.match(/[\p{L}\p{M}]+/gu) || [];

            for (const word of words) {
                if (word.length < 2) continue;
                const hasCyrillic = cyrillicRe.test(word);
                const hasLatin = latinRe.test(word);

                if (hasCyrillic && hasLatin) {
                    const cyrCount = (word.match(/[\u0400-\u04FF]/g) || []).length;
                    const latCount = (word.match(/[a-zA-Z\u00C0-\u024F]/g) || []).length;
                    const dominant = cyrCount > latCount ? 'ćirilica' : 'latinica';

                    findings.push(makeFinding({
                        element: el,
                        category: 'Mešanje pisama',
                        priority: 'OBAVEZNO',
                        confidence: 0.96,
                        original: word,
                        replacement: `[prebaciti celu reč na ${dominant}]`,
                        rationale: `Reč \u201e${word}\u201c sadrži pomešana slova ćirilice i latinice (${cyrCount} ćir. + ${latCount} lat.). Moguć OCR artefakt ili slučajna promena tastature.`,
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Greek text without translation
    // ==========================================
    function checkGreekWithoutTranslation(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const greekBlockRe = /[\u0370-\u03FF\u1F00-\u1FFF][\u0370-\u03FF\u1F00-\u1FFF\s,;\u00B7.'"\u0300-\u036F\u1DC0-\u1DFF]{2,}/g;
        const translationNearby = /[(\u201E\u201A""][^)"\u201C"]*[\u0400-\u04FFa-zA-Z\u00C0-\u024F]{5,}[^)"\u201C"]*[)"\u201C""]/;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            greekBlockRe.lastIndex = 0;
            let m;
            while ((m = greekBlockRe.exec(el.text)) !== null) {
                const afterGreek = el.text.substring(m.index + m[0].length, m.index + m[0].length + 200);
                const beforeGreek = el.text.substring(Math.max(0, m.index - 100), m.index);
                const hasTranslationAfter = translationNearby.test(afterGreek.substring(0, 150));
                const hasTranslationBefore = beforeGreek.match(/[(\u201E\u201A].*[\u0400-\u04FFa-zA-Z]{5,}.*[)\u201C"]/);
                const isGreekInParens = beforeGreek.match(/\(\s*$/) && afterGreek.match(/^\s*\)/);

                if (!hasTranslationAfter && !hasTranslationBefore && !isGreekInParens) {
                    const greekSnippet = m[0].length > 40 ? m[0].substring(0, 40) + '...' : m[0];
                    findings.push(makeFinding({
                        element: el,
                        category: 'Grčki bez prevoda',
                        priority: 'PROVERITI',
                        confidence: 0.75,
                        original: greekSnippet,
                        replacement: '[dodati srpski prevod u zagradi ili narednoj rečenici]',
                        rationale: 'Grčki citat/termin bez vidljivog prevoda u neposrednoj blizini (200 znakova).',
                        requiresSourceVerification: true,
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Duplicate words
    // ==========================================
    function checkDuplicateWords(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const allowedDuplicates = new Set(['ha', 'da', 'ne', 'vrlo', 'još', 'baš', 'sve']);

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            const dupeRe = /\b(\p{L}+)\s+\1\b/giu;
            let m;
            while ((m = dupeRe.exec(el.text)) !== null) {
                const word = m[1].toLowerCase();
                if (allowedDuplicates.has(word)) continue;
                if (word.length < 2) continue;
                findings.push(makeFinding({
                    element: el, category: 'Duple reči', priority: 'OBAVEZNO',
                    confidence: 0.95, original: m[0], replacement: m[1],
                    rationale: `Ponovljena reč \u201e${m[1]}\u201c \u2014 verovatno greška pri kucanju.`,
                    autoFixable: true,
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // CHECK: TOC vs actual headings
    // ==========================================
    function checkTocVsHeadings(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        const headings = docMap.elements
            .filter(el => el.type === 'heading' && el.text.trim().length > 0)
            .map(el => ({ text: el.text.trim(), level: el.headingLevel, id: el.id }));
        if (headings.length < 2) return { findings, scannedCount, skippedCount };

        const tocPattern = /^(.+?)\.{3,}\s*\d+\s*$/;
        const tocEntries = [];
        for (const el of docMap.elements) {
            if (el.type !== 'paragraph') continue;
            const m = el.text.match(tocPattern);
            if (m) { tocEntries.push({ text: m[1].trim(), element: el }); }
        }
        scannedCount = tocEntries.length;
        if (tocEntries.length === 0) return { findings, scannedCount, skippedCount };

        for (const tocEntry of tocEntries) {
            const normalizedToc = tocEntry.text.replace(/\s+/g, ' ').toLowerCase();
            const matchingHeading = headings.find(h =>
                h.text.replace(/\s+/g, ' ').toLowerCase() === normalizedToc);

            if (!matchingHeading) {
                const closeMatch = headings.find(h =>
                    levenshteinDistance(h.text.toLowerCase(), normalizedToc) <= 3);
                if (closeMatch) {
                    findings.push(makeFinding({
                        element: tocEntry.element, category: 'TOC/naslovi',
                        priority: 'OBAVEZNO', confidence: 0.88,
                        original: `TOC: \u201e${tocEntry.text}\u201c`,
                        replacement: `Uskladiti sa naslovom: \u201e${closeMatch.text}\u201c`,
                        rationale: 'Naslov u sadržaju se razlikuje od stvarnog naslova u dokumentu.',
                    }));
                } else {
                    findings.push(makeFinding({
                        element: tocEntry.element, category: 'TOC/naslovi',
                        priority: 'PROVERITI', confidence: 0.70,
                        original: `TOC: \u201e${tocEntry.text}\u201c`,
                        replacement: '[proveriti da li ovaj naslov postoji u dokumentu]',
                        rationale: 'Stavka iz sadržaja nije pronađena među naslovima u dokumentu.',
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Numbering (using OOXML data first, text fallback)
    // Removed off-by-one skip per spec
    // ==========================================
    function checkNumbering(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        // Strategy 1: Use OOXML numbering data if available
        // Only include numeric formats (exclude bullet/none)
        const ooxmlNumbered = docMap.elements.filter(el =>
            el.numId && el.displayedNumber !== null && el.displayedNumber !== undefined &&
            el.numFmt && el.numFmt !== 'bullet' && el.numFmt !== 'none');

        if (ooxmlNumbered.length > 1) {
            scannedCount = ooxmlNumbered.length;
            // Group by listInstanceId
            const byList = {};
            for (const el of ooxmlNumbered) {
                const key = el.listInstanceId || `${el.numId}-${el.numLevel || 0}`;
                if (!byList[key]) byList[key] = [];
                byList[key].push(el);
            }

            for (const [, items] of Object.entries(byList)) {
                for (let i = 1; i < items.length; i++) {
                    const prev = items[i - 1];
                    const curr = items[i];
                    const expected = prev.displayedNumber + 1;

                    if (curr.displayedNumber !== expected) {
                        const prevLabel = prev.displayedLabel || String(prev.displayedNumber);
                        const currLabel = curr.displayedLabel || String(curr.displayedNumber);
                        const expectedLabel = curr.displayedLabel
                            ? curr.displayedLabel.replace(String(curr.displayedNumber), String(expected))
                            : String(expected);
                        findings.push(makeFinding({
                            element: curr,
                            category: 'Numeracija',
                            priority: 'OBAVEZNO',
                            confidence: 0.92,
                            original: `Stavka ${currLabel} (prethodno: ${prevLabel})`,
                            replacement: `Očekivano: ${expectedLabel}`,
                            rationale: `Numeracija preskače sa ${prev.displayedNumber} na ${curr.displayedNumber}.`,
                        }));
                    }
                }
            }
            return { findings, scannedCount, skippedCount };
        }

        // Strategy 2: Fallback to text-based detection
        const numberedParas = [];
        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            const m = el.text.match(/^\s*(\d+)[.)]\s/);
            if (m) {
                numberedParas.push({ num: parseInt(m[1], 10), element: el });
            }
        }

        for (let i = 1; i < numberedParas.length; i++) {
            const prev = numberedParas[i - 1];
            const curr = numberedParas[i];
            if (curr.num === 1) continue; // New list restart

            const expected = prev.num + 1;
            if (curr.num !== expected) {
                findings.push(makeFinding({
                    element: curr.element,
                    category: 'Numeracija',
                    priority: 'OBAVEZNO',
                    confidence: 0.85,
                    original: `Stavka ${curr.num} (prethodno: ${prev.num})`,
                    replacement: `Očekivano: ${expected}`,
                    rationale: `Numeracija liste preskače sa ${prev.num} na ${curr.num}. Očekivano je ${expected}.`,
                }));
            }
        }

        if (numberedParas.length > 0 && numberedParas[0].num > 1) {
            findings.push(makeFinding({
                element: numberedParas[0].element,
                category: 'Numeracija',
                priority: 'PROVERITI',
                confidence: 0.70,
                original: `Lista počinje brojem ${numberedParas[0].num}`,
                replacement: '[proveriti da li lista treba da počne od 1]',
                rationale: 'Numerisana lista ne počinje od 1.',
            }));
        }

        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Dashes and ranges
    // ==========================================
    function checkDashes(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            let m;

            // Numeric ranges with hyphen instead of en-dash
            const rangeRe = /(\d{1,4})\s*-\s*(\d{1,4})/g;
            while ((m = rangeRe.exec(el.text)) !== null) {
                if (m.index > 0 && el.text[m.index - 1] === '(') continue;
                findings.push(makeFinding({
                    element: el, category: 'Tipografija', priority: 'OBAVEZNO',
                    confidence: 0.92, original: m[0],
                    replacement: `${m[1]}\u2013${m[2]}`,
                    rationale: 'Brojčani raspon treba pisati sa en-dash (\u2013), ne sa običnom crticom (-).',
                    autoFixable: true,
                }));
            }

            // Double hyphen as em-dash
            const doubleHyphen = /--/g;
            while ((m = doubleHyphen.exec(el.text)) !== null) {
                const ctx = getContext(el.text, m.index, 20);
                findings.push(makeFinding({
                    element: el, category: 'Tipografija', priority: 'OBAVEZNO',
                    confidence: 0.90, original: ctx,
                    replacement: ctx.replace('--', '\u2014'),
                    rationale: 'Dvostruka crtica (--) treba da bude em-dash (\u2014).',
                    autoFixable: true,
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Bibliography
    // ==========================================
    function checkBibliography(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        const bibHeadingIdx = docMap.elements.findIndex(el =>
            el.type === 'heading' &&
            el.text.match(/bibliograf|literatura|izvori|references|works cited/i));
        if (bibHeadingIdx === -1) return { findings, scannedCount, skippedCount };

        const bibEntries = [];
        for (let i = bibHeadingIdx + 1; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type === 'heading') break;
            if (el.type === 'paragraph' && el.text.trim().length > 10) {
                bibEntries.push(el);
            }
        }
        scannedCount = bibEntries.length;
        if (bibEntries.length === 0) return { findings, scannedCount, skippedCount };

        for (const entry of bibEntries) {
            const text = entry.text.trim();
            if (!text.match(/\b(1[5-9]\d{2}|20[0-2]\d)\b/)) {
                findings.push(makeFinding({
                    element: entry, category: 'Bibliografija', priority: 'PROVERITI',
                    confidence: 0.75,
                    original: text.substring(0, 60) + (text.length > 60 ? '...' : ''),
                    replacement: '[dodati godinu izdanja]',
                    rationale: 'Bibliografski zapis bez prepoznatljive godine izdanja.',
                    requiresSourceVerification: true,
                }));
            }
            if (!text.match(/:\s*[A-Z\u0400-\u04FF]/) &&
                !text.match(/\bUniversity\b|\bPress\b|\bVerlag\b|\bizdava/i)) {
                findings.push(makeFinding({
                    element: entry, category: 'Bibliografija', priority: 'PROVERITI',
                    confidence: 0.60,
                    original: text.substring(0, 60) + (text.length > 60 ? '...' : ''),
                    replacement: '[proveriti da li nedostaje izdavač/mesto izdanja]',
                    rationale: 'Moguć nedostatak izdavača ili mesta izdanja.',
                    requiresSourceVerification: true,
                }));
            }
        }

        // Cross-reference citations
        const citationPatterns = [
            /\(([A-Z\u0400-\u04FF][a-z\u0400-\u04FF]+),?\s*(\d{4})\)/g,
            /([A-Z\u0400-\u04FF][a-z\u0400-\u04FF]+)\s*\((\d{4})\)/g,
        ];
        const textElements = docMap.elements.slice(0, bibHeadingIdx);
        const citedAuthors = new Set();
        for (const el of textElements) {
            if (!el.text) continue;
            for (const pattern of citationPatterns) {
                pattern.lastIndex = 0;
                let m;
                while ((m = pattern.exec(el.text)) !== null) { citedAuthors.add(m[1]); }
            }
        }
        for (const author of citedAuthors) {
            const inBib = bibEntries.some(e => e.text.includes(author));
            if (!inBib) {
                findings.push(makeFinding({
                    element: docMap.elements[bibHeadingIdx],
                    category: 'Bibliografija', priority: 'PROVERITI', confidence: 0.70,
                    original: `Izvor citiran u tekstu: \u201e${author}\u201c`,
                    replacement: '[dodati u bibliografiju ili proveriti pisanje prezimena]',
                    rationale: `Autor \u201e${author}\u201c se citira u tekstu ali nije pronađen u bibliografiji.`,
                    requiresSourceVerification: true,
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: URLs - fixed: don't strip ) if it's part of URL
    // ==========================================
    function checkUrls(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g;

        for (const el of docMap.elements) {
            if (!el.text) { skippedCount++; continue; }
            scannedCount++;
            urlRe.lastIndex = 0;
            let m;
            while ((m = urlRe.exec(el.text)) !== null) {
                const url = m[0];

                // Count parens in URL to decide if trailing ) is part of it
                const openParens = (url.match(/\(/g) || []).length;
                const closeParens = (url.match(/\)/g) || []).length;

                // Only flag trailing punctuation if it's NOT balanced parens
                // (Wikipedia URLs often have parentheses)
                let trailingPunct = url.match(/[,.;]+$/);
                if (trailingPunct) {
                    const cleanUrl = url.replace(/[,.;]+$/, '');
                    findings.push(makeFinding({
                        element: el, category: 'URL', priority: 'PROVERITI',
                        confidence: 0.80, original: url, replacement: cleanUrl,
                        rationale: 'URL se završava interpunkcijom koja je verovatno deo rečenice.',
                    }));
                } else if (url.match(/\)$/) && closeParens > openParens) {
                    // Trailing ) is likely sentence punctuation, not part of URL
                    const cleanUrl = url.replace(/\)$/, '');
                    findings.push(makeFinding({
                        element: el, category: 'URL', priority: 'PROVERITI',
                        confidence: 0.70, original: url, replacement: cleanUrl,
                        rationale: 'URL se završava zatvorenom zagradom koja verovatno nije deo URL-a.',
                    }));
                }

                if (url.length > 100) {
                    findings.push(makeFinding({
                        element: el, category: 'URL', priority: 'PREPORUKA',
                        confidence: 0.65,
                        original: url.substring(0, 50) + '...',
                        replacement: '[proveriti da URL nije prelomljen ili odsečen]',
                        rationale: 'Veoma dugačak URL \u2014 proveriti da je kompletan.',
                    }));
                }
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // CHECK: Footnotes
    // ==========================================
    function checkFootnotes(docMap) {
        const findings = [];
        const scannedCount = docMap.footnotes.length + docMap.endnotes.length;
        let skippedCount = 0;

        for (const fn of docMap.footnotes) {
            if (fn.isEmpty) {
                findings.push(makeFinding({
                    element: { id: `fn-${fn.id}`, type: 'footnote', text: '' },
                    category: 'Fusnote', priority: 'OBAVEZNO', confidence: 0.99,
                    original: `[Fusnota ${fn.id} \u2014 prazna]`,
                    replacement: '[dodati sadržaj ili ukloniti fusnotu]',
                    rationale: 'Prazna fusnota bez sadržaja.',
                }));
            }
        }
        for (const en of docMap.endnotes) {
            if (en.isEmpty) {
                findings.push(makeFinding({
                    element: { id: `en-${en.id}`, type: 'endnote', text: '' },
                    category: 'Fusnote', priority: 'OBAVEZNO', confidence: 0.99,
                    original: `[Endnota ${en.id} \u2014 prazna]`,
                    replacement: '[dodati sadržaj ili ukloniti endnotu]',
                    rationale: 'Prazna endnota bez sadržaja.',
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // CHECK: Repeated paragraphs
    // ==========================================
    function checkRepetition(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const seen = new Map();

        for (const el of docMap.elements) {
            if (!el.text || el.text.trim().length < 50) { skippedCount++; continue; }
            if (el.type === 'heading') { skippedCount++; continue; }
            scannedCount++;
            const normalized = el.text.trim().replace(/\s+/g, ' ').toLowerCase();
            if (seen.has(normalized)) {
                const firstEl = seen.get(normalized);
                findings.push(makeFinding({
                    element: el, category: 'Ponavljanje', priority: 'OBAVEZNO',
                    confidence: 0.98,
                    original: el.text.substring(0, 80) + '...',
                    replacement: '[ukloniti ponovljeni pasus]',
                    rationale: `Identičan pasus se pojavljuje ranije u dokumentu (${firstEl.id}).`,
                }));
            } else { seen.set(normalized, el); }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // CHECK: ALL-CAPS words in body text
    // ==========================================
    function checkAllCaps(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;
        const knownAbbr = new Set([
            'UNESCO', 'NATO', 'EU', 'SAD', 'SSSR', 'DNA', 'RNA', 'URL', 'HTML',
            'CSS', 'JS', 'PDF', 'DOCX', 'ISBN', 'ISSN', 'DOI', 'NB', 'PS', 'AD',
            'BC', 'PhD', 'USA', 'UK', 'ID', 'OK', 'IT', 'PR', 'HR', 'TV', 'CD',
            'DVD', 'USB', 'PC', 'OS', 'AI',
        ]);

        for (const el of docMap.elements) {
            if (!el.text || el.type === 'heading') { skippedCount++; continue; }
            scannedCount++;
            const words = el.text.match(/\b[A-Z\u0410-\u042F]{3,}\b/g) || [];
            for (const word of words) {
                if (knownAbbr.has(word)) continue;
                if (word.match(/^[IVXLCDM]+$/)) continue;
                findings.push(makeFinding({
                    element: el, category: 'Tipografija', priority: 'PREPORUKA',
                    confidence: 0.70, original: word,
                    replacement: '[proveriti da li ALL-CAPS treba da stoji ili treba normalizovati]',
                    rationale: `Reč \u201e${word}\u201c je napisana svim velikim slovima u telu teksta.`,
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }

    // ==========================================
    // CHECK: Empty headings
    // ==========================================
    function checkEmptyHeadings(docMap) {
        const findings = [];
        let scannedCount = 0;
        let skippedCount = 0;

        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type !== 'heading') { skippedCount++; continue; }
            scannedCount++;

            if (!el.text || el.text.trim().length === 0) {
                findings.push(makeFinding({
                    element: el, category: 'Struktura', priority: 'OBAVEZNO',
                    confidence: 0.99,
                    original: `[Prazan naslov nivoa ${el.headingLevel}]`,
                    replacement: '[dodati tekst naslova ili ukloniti prazan naslov]',
                    rationale: 'Naslov bez sadržaja \u2014 moguć artefakt formatiranja.',
                }));
            }

            if (i < docMap.elements.length - 1 && docMap.elements[i + 1].type === 'heading') {
                findings.push(makeFinding({
                    element: el, category: 'Struktura', priority: 'PREPORUKA',
                    confidence: 0.75, original: el.text || '[prazan]',
                    replacement: '[dodati sadržaj ispod naslova ili ukloniti naslov]',
                    rationale: 'Naslov odmah praćen sledećim naslovom bez sadržaja između.',
                }));
            }
        }
        return { findings, scannedCount, skippedCount };
    }


    // ==========================================
    // HELPER: Create a finding object
    // Supports optional table cell metadata (tableId, rowId, cellId, rowIndex, columnIndex)
    // ==========================================
    function makeFinding({ element, category, priority, confidence, original,
        replacement, rationale, autoFixable = false, globalPattern = false,
        requiresSourceVerification = false,
        tableId = null, rowId = null, cellId = null, rowIndex = null, columnIndex = null }) {
        return {
            id: nextId(),
            section: getSectionName(element),
            paragraphId: element.id || null,
            category,
            priority,
            confidence,
            original,
            replacement,
            rationale,
            isDirectQuote: element.isDirectQuote || false,
            requiresSourceVerification,
            autoFixable,
            globalPattern,
            status: 'OPEN',
            // Table cell metadata (null for non-table findings)
            tableId,
            rowId,
            cellId,
            rowIndex,
            columnIndex,
            _element: element, // Internal ref, removed in runAudit before output
        };
    }

    function getSectionName(element) {
        if (element.type === 'heading') return element.text || '(bez naslova)';
        if (element.section) return element.section;
        return element.id || '(nepoznata lokacija)';
    }

    function getContext(text, pos, radius) {
        const start = Math.max(0, pos - radius);
        const end = Math.min(text.length, pos + radius);
        let ctx = text.substring(start, end);
        if (start > 0) ctx = '...' + ctx;
        if (end < text.length) ctx = ctx + '...';
        return ctx;
    }

    function levenshteinDistance(a, b) {
        if (a.length === 0) return b.length;
        if (b.length === 0) return a.length;
        const matrix = [];
        for (let i = 0; i <= b.length; i++) matrix[i] = [i];
        for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
        for (let i = 1; i <= b.length; i++) {
            for (let j = 1; j <= a.length; j++) {
                if (b.charAt(i - 1) === a.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        matrix[i][j - 1] + 1,
                        matrix[i - 1][j] + 1);
                }
            }
        }
        return matrix[b.length][a.length];
    }

    // Public API
    return { runAudit };
})();

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = RuleEngine;
}
