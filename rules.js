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
     * Run all enabled checks against the document map
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

            const before = findings.length;
            check.fn(docMap, findings);
            const found = findings.length - before;

            if (found === 0) {
                passedChecks.push({ area: check.name, result: 'Bez grešaka', count: docMap.elements.length });
            }
        }

        return { findings, passedChecks };
    }



    // ==========================================
    // CHECK: Unbalanced brackets
    // ==========================================
    function checkBrackets(docMap, findings) {
        const pairs = [['(', ')'], ['[', ']'], ['{', '}']];
        const pairNames = { '(': 'obla zagrada', '[': 'uglasta zagrada', '{': 'vitičasta zagrada' };

        for (const el of docMap.elements) {
            if (!el.text || el.text.trim().length === 0) continue;

            for (const [open, close] of pairs) {
                let depth = 0;
                let positions = [];

                for (let i = 0; i < el.text.length; i++) {
                    if (el.text[i] === open) {
                        depth++;
                        positions.push(i);
                    } else if (el.text[i] === close) {
                        depth--;
                        if (depth < 0) {
                            const ctx = getContext(el.text, i, 40);
                            findings.push(makeFinding({
                                element: el,
                                category: 'Zagrade',
                                priority: 'OBAVEZNO',
                                confidence: 0.98,
                                original: ctx,
                                replacement: `[ukloniti višak "${close}" ili dodati "${open}"]`,
                                rationale: `Zatvorena ${pairNames[open]} bez odgovarajuće otvorene.`,
                                autoFixable: false,
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
                        replacement: `[dodati "${close}" ili ukloniti "${open}"]`,
                        rationale: `Otvorena ${pairNames[open]} bez odgovarajuće zatvorene.`,
                        autoFixable: false,
                    }));
                }
            }

            // Special check: ; where ) should be (common find-replace artifact)
            const semiPattern = /\([^)]*;(?=[^(]*$)/g;
            let match;
            while ((match = semiPattern.exec(el.text)) !== null) {
                // Check if there's a ( without matching ) and ; near end
                const fragment = el.text.substring(match.index);
                if (fragment.indexOf(')') === -1 || fragment.indexOf(';') < fragment.indexOf(')')) {
                    const ctx = getContext(el.text, match.index, 50);
                    findings.push(makeFinding({
                        element: el,
                        category: 'Zagrade',
                        priority: 'PROVERITI',
                        confidence: 0.75,
                        original: ctx,
                        replacement: '[proveriti da li ";" treba da bude ")"]',
                        rationale: 'Tačka-zarez unutar nezatvorene zagrade — mogući artefakt find-replace operacije.',
                        autoFixable: false,
                    }));
                }
            }
        }
    }



    // ==========================================
    // CHECK: Quotes (straight vs typographic)
    // ==========================================
    function checkQuotes(docMap, findings) {
        for (const el of docMap.elements) {
            if (!el.text) continue;

            // Straight double quotes
            const straightDoubleRe = /(?<!\w)"|"(?!\w)/g;
            let m;
            let straightCount = 0;
            while ((m = straightDoubleRe.exec(el.text)) !== null) {
                straightCount++;
                if (straightCount <= 3) { // Report first few, then consolidate
                    const ctx = getContext(el.text, m.index, 30);
                    findings.push(makeFinding({
                        element: el,
                        category: 'Tipografija',
                        priority: 'OBAVEZNO',
                        confidence: 0.95,
                        original: ctx,
                        replacement: '[zameniti ravne navodnike tipografskim: „..." ili "..."]',
                        rationale: 'Ravan navodnik umesto tipografskog. Srpski standard: „..." (spoljni) ili \'...\' (unutrašnji).',
                        autoFixable: true,
                    }));
                }
            }
            // If many found, add consolidated note
            if (straightCount > 3) {
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'OBAVEZNO',
                    confidence: 0.95,
                    original: `[${straightCount} ravnih navodnika u ovom pasusu]`,
                    replacement: '[zameniti sve ravne navodnike tipografskim]',
                    rationale: `Pronađeno ukupno ${straightCount} ravnih navodnika.`,
                    autoFixable: true,
                    globalPattern: true,
                }));
            }

            // Straight single quotes used as apostrophe is fine, but as quote delimiter is not
            // Check for unmatched typographic quotes
            const openDouble = (el.text.match(/\u201E/g) || []).length; // „
            const closeDouble = (el.text.match(/\u201C/g) || []).length; // "
            if (openDouble !== closeDouble) {
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'OBAVEZNO',
                    confidence: 0.90,
                    original: `[${openDouble}× „ vs ${closeDouble}× "]`,
                    replacement: '[upariti otvarajuće i zatvarajuće navodnike]',
                    rationale: 'Broj otvarajućih navodnika „ ne odgovara broju zatvarajućih ".',
                    autoFixable: false,
                }));
            }
        }
    }



    // ==========================================
    // CHECK: Markdown artifacts
    // ==========================================
    function checkMarkdownArtifacts(docMap, findings) {
        // Only relevant for docx (markdown artifacts that leaked into final doc)
        // For .md files these are expected
        if (docMap.type === 'markdown') return;

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
            if (!el.text) continue;

            for (const pat of patterns) {
                let m;
                pat.re.lastIndex = 0;
                while ((m = pat.re.exec(el.text)) !== null) {
                    findings.push(makeFinding({
                        element: el,
                        category: 'Markdown artefakt',
                        priority: 'OBAVEZNO',
                        confidence: 0.92,
                        original: m[0].trim(),
                        replacement: '[ukloniti markdown sintaksu, primeniti pravo formatiranje]',
                        rationale: `Ostatak ${pat.desc} — verovatno nije pretvoreno u formatiranje.`,
                        autoFixable: false,
                    }));
                }
            }
        }
    }



    // ==========================================
    // CHECK: Spacing and punctuation
    // ==========================================
    function checkSpacing(docMap, findings) {
        for (const el of docMap.elements) {
            if (!el.text) continue;
            const text = el.text;

            // Double spaces
            const doubleSpace = /  +/g;
            let m;
            while ((m = doubleSpace.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 25);
                findings.push(makeFinding({
                    element: el,
                    category: 'Razmaci',
                    priority: 'OBAVEZNO',
                    confidence: 0.99,
                    original: ctx,
                    replacement: ctx.replace(/  +/g, ' '),
                    rationale: 'Višestruki razmak.',
                    autoFixable: true,
                }));
            }

            // Space before comma, period, semicolon, colon, exclamation, question
            const spaceBeforePunct = / +([,.:;!?])/g;
            while ((m = spaceBeforePunct.exec(text)) !== null) {
                // Skip if it's inside a URL or number like "3 .14"
                const before = text.substring(Math.max(0, m.index - 5), m.index);
                if (before.match(/https?:$/) || before.match(/\d$/)) continue;

                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Razmaci',
                    priority: 'OBAVEZNO',
                    confidence: 0.97,
                    original: ctx,
                    replacement: ctx.replace(/ +([,.:;!?])/, '$1'),
                    rationale: `Razmak pre interpunkcijskog znaka "${m[1]}".`,
                    autoFixable: true,
                }));
            }

            // Space before closing bracket
            const spaceBeforeClose = / +([)\]}>»"])/g;
            while ((m = spaceBeforeClose.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Razmaci',
                    priority: 'PREPORUKA',
                    confidence: 0.90,
                    original: ctx,
                    replacement: ctx.replace(/ +([)\]}>»"])/, '$1'),
                    rationale: 'Razmak pre zatvorene zagrade/navodnika.',
                    autoFixable: true,
                }));
            }

            // Space after opening bracket
            const spaceAfterOpen = /([(\[{<«„]) +/g;
            while ((m = spaceAfterOpen.exec(text)) !== null) {
                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Razmaci',
                    priority: 'PREPORUKA',
                    confidence: 0.90,
                    original: ctx,
                    replacement: ctx.replace(/([(\[{<«„]) +/, '$1'),
                    rationale: 'Razmak posle otvorene zagrade/navodnika.',
                    autoFixable: true,
                }));
            }

            // No space after comma/semicolon/colon (except at end of line or before digit in numbers)
            const noSpaceAfterPunct = /([,;:])([^\s\d"'\u201C\u201D\u201E\u2019)\]])/g;
            while ((m = noSpaceAfterPunct.exec(text)) !== null) {
                // Skip URLs, file paths
                const ctx5 = text.substring(Math.max(0, m.index - 10), m.index + 10);
                if (ctx5.match(/https?:/) || ctx5.match(/\w:\\/)) continue;

                const ctx = getContext(text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Razmaci',
                    priority: 'OBAVEZNO',
                    confidence: 0.85,
                    original: ctx,
                    replacement: ctx.replace(/([,;:])(\S)/, '$1 $2'),
                    rationale: `Nedostaje razmak posle "${m[1]}".`,
                    autoFixable: true,
                }));
            }
        }
    }



    // ==========================================
    // CHECK: Script mixing (Cyrillic/Latin in same word)
    // ==========================================
    function checkScriptMixing(docMap, findings) {
        // Unicode ranges
        const cyrillicRe = /[\u0400-\u04FF]/;
        const latinRe = /[a-zA-Z\u00C0-\u024F]/;
        // Common Latin letters that look like Cyrillic: a, e, o, p, c, x, y, A, B, C, E, H, K, M, O, P, T, X
        const confusables = /[aeopycxABCEHKMOPTX]/;

        for (const el of docMap.elements) {
            if (!el.text) continue;

            // Split into words
            const words = el.text.match(/[\p{L}\p{M}]+/gu) || [];

            for (const word of words) {
                if (word.length < 2) continue;

                const hasCyrillic = cyrillicRe.test(word);
                const hasLatin = latinRe.test(word);

                if (hasCyrillic && hasLatin) {
                    // This is a mixed-script word — likely a typo or OCR artifact
                    // Find the word in context
                    const wordIdx = el.text.indexOf(word);
                    const ctx = getContext(el.text, wordIdx, 30);

                    // Determine which script dominates
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
                        rationale: `Reč "${word}" sadrži pomešana slova ćirilice i latinice (${cyrCount} ćir. + ${latCount} lat.). Moguć OCR artefakt ili slučajna promena tastature.`,
                        autoFixable: false,
                    }));
                }
            }
        }
    }



    // ==========================================
    // CHECK: Greek text without translation
    // ==========================================
    function checkGreekWithoutTranslation(docMap, findings) {
        // Greek Unicode range (basic + extended) - match continuous Greek phrases (words + spaces/punctuation between)
        const greekBlockRe = /[\u0370-\u03FF\u1F00-\u1FFF][\u0370-\u03FF\u1F00-\u1FFF\s,;·.'"\u0300-\u036F\u1DC0-\u1DFF]{2,}/g;
        // Serbian text indicators nearby (parentheses with Cyrillic/Latin)
        const translationNearby = /[(\u201E„""][^)"\u201C"]*[\u0400-\u04FFa-zA-Z\u00C0-\u024F]{5,}[^)"\u201C"]*[)"\u201C""]/;

        for (const el of docMap.elements) {
            if (!el.text) continue;

            let m;
            greekBlockRe.lastIndex = 0;
            while ((m = greekBlockRe.exec(el.text)) !== null) {
                // Look for translation within 200 chars after the Greek text
                const afterGreek = el.text.substring(m.index + m[0].length, m.index + m[0].length + 200);
                const beforeGreek = el.text.substring(Math.max(0, m.index - 100), m.index);

                // Check if there's parenthetical text nearby
                const hasTranslationAfter = translationNearby.test(afterGreek.substring(0, 150));
                const hasTranslationBefore = beforeGreek.match(/[(\u201E„].*[\u0400-\u04FFa-zA-Z]{5,}.*[)\u201C"]/);

                // Check if the Greek itself is inside parentheses (common pattern: Serbian text (Greek original))
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
                        rationale: 'Grčki citat/termin bez vidljivog prevoda u neposrednoj blizini (200 znakova). Čitalac koji ne čita grčki neće razumeti tekst.',
                        autoFixable: false,
                        requiresSourceVerification: true,
                    }));
                }
            }
        }
    }



    // ==========================================
    // CHECK: Duplicate words
    // ==========================================
    function checkDuplicateWords(docMap, findings) {
        // Common intentional duplicates to skip
        const allowedDuplicates = new Set([
            'ha', 'da', 'ne', 'vrlo', 'još', 'baš', 'sve',
            'ha ha', 'he he', 'ho ho',
        ]);

        for (const el of docMap.elements) {
            if (!el.text) continue;

            // Match consecutive identical words (case-insensitive)
            const dupeRe = /\b(\p{L}+)\s+\1\b/giu;
            let m;
            while ((m = dupeRe.exec(el.text)) !== null) {
                const word = m[1].toLowerCase();
                if (allowedDuplicates.has(word)) continue;
                if (word.length < 2) continue;

                const ctx = getContext(el.text, m.index, 30);
                findings.push(makeFinding({
                    element: el,
                    category: 'Duple reči',
                    priority: 'OBAVEZNO',
                    confidence: 0.95,
                    original: m[0],
                    replacement: m[1],
                    rationale: `Ponovljena reč "${m[1]}" — verovatno greška pri kucanju.`,
                    autoFixable: true,
                }));
            }
        }
    }

    // ==========================================
    // CHECK: TOC vs actual headings
    // ==========================================
    function checkTocVsHeadings(docMap, findings) {
        // Extract headings from the document
        const headings = docMap.elements
            .filter(el => el.type === 'heading' && el.text.trim().length > 0)
            .map(el => ({ text: el.text.trim(), level: el.headingLevel, id: el.id }));

        if (headings.length < 2) return; // Not enough structure to check

        // Try to identify TOC section (usually first few paragraphs with page numbers or dots)
        const tocPattern = /^(.+?)\.{3,}\s*\d+\s*$/;
        const tocEntries = [];

        for (const el of docMap.elements) {
            if (el.type !== 'paragraph') continue;
            const m = el.text.match(tocPattern);
            if (m) {
                tocEntries.push({ text: m[1].trim(), element: el });
            }
        }

        if (tocEntries.length === 0) return; // No TOC detected

        // Compare TOC entries with actual headings
        for (const tocEntry of tocEntries) {
            const normalizedToc = tocEntry.text.replace(/\s+/g, ' ').toLowerCase();
            const matchingHeading = headings.find(h =>
                h.text.replace(/\s+/g, ' ').toLowerCase() === normalizedToc
            );

            if (!matchingHeading) {
                // Check for close matches (typo or minor difference)
                const closeMatch = headings.find(h =>
                    levenshteinDistance(h.text.toLowerCase(), normalizedToc) <= 3
                );

                if (closeMatch) {
                    findings.push(makeFinding({
                        element: tocEntry.element,
                        category: 'TOC/naslovi',
                        priority: 'OBAVEZNO',
                        confidence: 0.88,
                        original: `TOC: "${tocEntry.text}"`,
                        replacement: `Uskladiti sa naslovom: "${closeMatch.text}"`,
                        rationale: 'Naslov u sadržaju se razlikuje od stvarnog naslova u dokumentu.',
                        autoFixable: false,
                    }));
                } else {
                    findings.push(makeFinding({
                        element: tocEntry.element,
                        category: 'TOC/naslovi',
                        priority: 'PROVERITI',
                        confidence: 0.70,
                        original: `TOC: "${tocEntry.text}"`,
                        replacement: '[proveriti da li ovaj naslov postoji u dokumentu]',
                        rationale: 'Stavka iz sadržaja nije pronađena među naslovima u dokumentu.',
                        autoFixable: false,
                    }));
                }
            }
        }
    }



    // ==========================================
    // CHECK: Numbering (list continuity)
    // ==========================================
    function checkNumbering(docMap, findings) {
        // Detect numbered lists by pattern: "1.", "2.", etc. at start of paragraphs
        const numberedParas = [];

        for (const el of docMap.elements) {
            if (!el.text) continue;
            const m = el.text.match(/^\s*(\d+)[.)]\s/);
            if (m) {
                numberedParas.push({ num: parseInt(m[1], 10), element: el });
            }
        }

        // Check continuity
        for (let i = 1; i < numberedParas.length; i++) {
            const prev = numberedParas[i - 1];
            const curr = numberedParas[i];

            // Allow reset to 1 (new list)
            if (curr.num === 1) continue;

            const expected = prev.num + 1;
            if (curr.num !== expected && curr.num !== 1) {
                // Check if it's a sub-list or just a gap
                if (Math.abs(curr.num - expected) === 1) continue; // off by one might be sub-numbering

                findings.push(makeFinding({
                    element: curr.element,
                    category: 'Numeracija',
                    priority: 'OBAVEZNO',
                    confidence: 0.85,
                    original: `Stavka ${curr.num} (prethodno: ${prev.num})`,
                    replacement: `Očekivano: ${expected}`,
                    rationale: `Numeracija liste preskače sa ${prev.num} na ${curr.num}. Očekivano je ${expected}.`,
                    autoFixable: false,
                }));
            }
        }

        // Check for list starting with number > 1
        if (numberedParas.length > 0 && numberedParas[0].num > 1) {
            findings.push(makeFinding({
                element: numberedParas[0].element,
                category: 'Numeracija',
                priority: 'PROVERITI',
                confidence: 0.70,
                original: `Lista počinje brojem ${numberedParas[0].num}`,
                replacement: '[proveriti da li lista treba da počne od 1]',
                rationale: 'Numerisana lista ne počinje od 1.',
                autoFixable: false,
            }));
        }
    }

    // ==========================================
    // CHECK: Dashes and ranges
    // ==========================================
    function checkDashes(docMap, findings) {
        for (const el of docMap.elements) {
            if (!el.text) continue;

            // Numeric ranges with hyphen instead of en-dash: "484-425" should be "484–425"
            const rangeRe = /(\d{1,4})\s*-\s*(\d{1,4})/g;
            let m;
            while ((m = rangeRe.exec(el.text)) !== null) {
                // Skip if it's negative number
                if (m.index > 0 && el.text[m.index - 1] === '(') continue;

                const ctx = getContext(el.text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'OBAVEZNO',
                    confidence: 0.92,
                    original: m[0],
                    replacement: `${m[1]}–${m[2]}`,
                    rationale: 'Brojčani raspon treba pisati sa en-dash (–), ne sa običnom crticom (-).',
                    autoFixable: true,
                }));
            }

            // Check for double hyphen used as em-dash
            const doubleHyphen = /--/g;
            while ((m = doubleHyphen.exec(el.text)) !== null) {
                const ctx = getContext(el.text, m.index, 20);
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'OBAVEZNO',
                    confidence: 0.90,
                    original: ctx,
                    replacement: ctx.replace('--', '—'),
                    rationale: 'Dvostruka crtica (--) treba da bude em-dash (—).',
                    autoFixable: true,
                }));
            }
        }
    }



    // ==========================================
    // CHECK: Bibliography
    // ==========================================
    function checkBibliography(docMap, findings) {
        // Try to find bibliography section
        const bibHeadingIdx = docMap.elements.findIndex(el =>
            el.type === 'heading' &&
            el.text.match(/bibliograf|literatura|izvori|references|works cited/i)
        );

        if (bibHeadingIdx === -1) return; // No bibliography section found

        // Collect bibliography entries (paragraphs after bib heading until next heading)
        const bibEntries = [];
        for (let i = bibHeadingIdx + 1; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.type === 'heading') break;
            if (el.type === 'paragraph' && el.text.trim().length > 10) {
                bibEntries.push(el);
            }
        }

        if (bibEntries.length === 0) return;

        // Check each entry for common format issues
        for (const entry of bibEntries) {
            const text = entry.text.trim();

            // Check for year
            if (!text.match(/\b(1[5-9]\d{2}|20[0-2]\d)\b/)) {
                findings.push(makeFinding({
                    element: entry,
                    category: 'Bibliografija',
                    priority: 'PROVERITI',
                    confidence: 0.75,
                    original: text.substring(0, 60) + (text.length > 60 ? '...' : ''),
                    replacement: '[dodati godinu izdanja]',
                    rationale: 'Bibliografski zapis bez prepoznatljive godine izdanja.',
                    autoFixable: false,
                    requiresSourceVerification: true,
                }));
            }

            // Check for publisher (common pattern: city: publisher)
            // This is a soft check - many formats exist
            if (!text.match(/:\s*[A-Z\u0400-\u04FF]/) && !text.match(/\bUniversity\b|\bPress\b|\bVerlag\b|\bizdavač/i)) {
                findings.push(makeFinding({
                    element: entry,
                    category: 'Bibliografija',
                    priority: 'PROVERITI',
                    confidence: 0.60,
                    original: text.substring(0, 60) + (text.length > 60 ? '...' : ''),
                    replacement: '[proveriti da li nedostaje izdavač/mesto izdanja]',
                    rationale: 'Moguć nedostatak izdavača ili mesta izdanja.',
                    autoFixable: false,
                    requiresSourceVerification: true,
                }));
            }
        }

        // Cross-reference: find in-text citations and check against bibliography
        const citationPatterns = [
            /\(([A-Z\u0400-\u04FF][a-z\u0400-\u04FF]+),?\s*(\d{4})\)/g,  // (Author, 2020)
            /([A-Z\u0400-\u04FF][a-z\u0400-\u04FF]+)\s*\((\d{4})\)/g,     // Author (2020)
        ];

        const textElements = docMap.elements.slice(0, bibHeadingIdx);
        const citedAuthors = new Set();

        for (const el of textElements) {
            if (!el.text) continue;
            for (const pattern of citationPatterns) {
                pattern.lastIndex = 0;
                let m;
                while ((m = pattern.exec(el.text)) !== null) {
                    citedAuthors.add(m[1]);
                }
            }
        }

        // Check if cited authors appear in bibliography
        for (const author of citedAuthors) {
            const inBib = bibEntries.some(e => e.text.includes(author));
            if (!inBib) {
                findings.push(makeFinding({
                    element: docMap.elements[bibHeadingIdx],
                    category: 'Bibliografija',
                    priority: 'PROVERITI',
                    confidence: 0.70,
                    original: `Izvor citiran u tekstu: "${author}"`,
                    replacement: '[dodati u bibliografiju ili proveriti pisanje prezimena]',
                    rationale: `Autor "${author}" se citira u tekstu ali nije pronađen u bibliografiji.`,
                    autoFixable: false,
                    requiresSourceVerification: true,
                }));
            }
        }
    }



    // ==========================================
    // CHECK: URLs
    // ==========================================
    function checkUrls(docMap, findings) {
        const urlRe = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g;

        for (const el of docMap.elements) {
            if (!el.text) continue;

            let m;
            urlRe.lastIndex = 0;
            while ((m = urlRe.exec(el.text)) !== null) {
                const url = m[0];

                // Check for broken URL (ends with punctuation that's likely sentence-ending)
                if (url.match(/[,.)]+$/)) {
                    const cleanUrl = url.replace(/[,.)]+$/, '');
                    findings.push(makeFinding({
                        element: el,
                        category: 'URL',
                        priority: 'PROVERITI',
                        confidence: 0.80,
                        original: url,
                        replacement: cleanUrl,
                        rationale: 'URL se završava interpunkcijom koja je verovatno deo rečenice, ne URL-a.',
                        autoFixable: false,
                    }));
                }

                // Check for very long URL that might be broken across lines
                if (url.length > 100) {
                    findings.push(makeFinding({
                        element: el,
                        category: 'URL',
                        priority: 'PREPORUKA',
                        confidence: 0.65,
                        original: url.substring(0, 50) + '...',
                        replacement: '[proveriti da URL nije prelomljen ili odsečen]',
                        rationale: 'Veoma dugačak URL — proveriti da je kompletan i funkcionalan.',
                        autoFixable: false,
                    }));
                }
            }
        }
    }

    // ==========================================
    // CHECK: Footnotes
    // ==========================================
    function checkFootnotes(docMap, findings) {
        // Check for empty footnotes
        for (const fn of docMap.footnotes) {
            if (fn.isEmpty) {
                findings.push(makeFinding({
                    element: { id: `fn-${fn.id}`, type: 'footnote', text: '' },
                    category: 'Fusnote',
                    priority: 'OBAVEZNO',
                    confidence: 0.99,
                    original: `[Fusnota ${fn.id} — prazna]`,
                    replacement: '[dodati sadržaj ili ukloniti fusnotu]',
                    rationale: 'Prazna fusnota bez sadržaja.',
                    autoFixable: false,
                }));
            }
        }

        // Check for empty endnotes
        for (const en of docMap.endnotes) {
            if (en.isEmpty) {
                findings.push(makeFinding({
                    element: { id: `en-${en.id}`, type: 'endnote', text: '' },
                    category: 'Fusnote',
                    priority: 'OBAVEZNO',
                    confidence: 0.99,
                    original: `[Endnota ${en.id} — prazna]`,
                    replacement: '[dodati sadržaj ili ukloniti endnotu]',
                    rationale: 'Prazna endnota bez sadržaja.',
                    autoFixable: false,
                }));
            }
        }
    }

    // ==========================================
    // CHECK: Repeated paragraphs
    // ==========================================
    function checkRepetition(docMap, findings) {
        const seen = new Map(); // text hash -> first occurrence element

        for (const el of docMap.elements) {
            if (!el.text || el.text.trim().length < 50) continue; // Skip short paragraphs
            if (el.type === 'heading') continue;

            const normalized = el.text.trim().replace(/\s+/g, ' ').toLowerCase();

            if (seen.has(normalized)) {
                const firstEl = seen.get(normalized);
                findings.push(makeFinding({
                    element: el,
                    category: 'Ponavljanje',
                    priority: 'OBAVEZNO',
                    confidence: 0.98,
                    original: el.text.substring(0, 80) + '...',
                    replacement: '[ukloniti ponovljeni pasus]',
                    rationale: `Identičan pasus se pojavljuje ranije u dokumentu (${firstEl.id}).`,
                    autoFixable: false,
                }));
            } else {
                seen.set(normalized, el);
            }
        }
    }



    // ==========================================
    // CHECK: ALL-CAPS words in body text
    // ==========================================
    function checkAllCaps(docMap, findings) {
        // Skip headings, known abbreviations
        const knownAbbreviations = new Set([
            'UNESCO', 'NATO', 'EU', 'SAD', 'SSSR', 'DNA', 'RNA', 'URL', 'HTML', 'CSS', 'JS',
            'PDF', 'DOCX', 'ISBN', 'ISSN', 'DOI', 'NB', 'PS', 'AD', 'BC', 'PhD', 'USA', 'UK',
            'ID', 'OK', 'IT', 'PR', 'HR', 'TV', 'CD', 'DVD', 'USB', 'PC', 'OS', 'AI',
        ]);

        for (const el of docMap.elements) {
            if (!el.text || el.type === 'heading') continue;

            const words = el.text.match(/\b[A-Z\u0410-\u042F]{3,}\b/g) || [];

            for (const word of words) {
                if (knownAbbreviations.has(word)) continue;
                if (word.match(/^[IVXLCDM]+$/)) continue; // Roman numerals

                const wordIdx = el.text.indexOf(word);
                const ctx = getContext(el.text, wordIdx, 30);
                findings.push(makeFinding({
                    element: el,
                    category: 'Tipografija',
                    priority: 'PREPORUKA',
                    confidence: 0.70,
                    original: word,
                    replacement: '[proveriti da li ALL-CAPS treba da stoji ili treba normalizovati]',
                    rationale: `Reč "${word}" je napisana svim velikim slovima u telu teksta. Ako nije skraćenica, trebalo bi je napisati normalno.`,
                    autoFixable: false,
                }));
            }
        }
    }

    // ==========================================
    // CHECK: Empty headings
    // ==========================================
    function checkEmptyHeadings(docMap, findings) {
        for (const el of docMap.elements) {
            if (el.type !== 'heading') continue;

            if (!el.text || el.text.trim().length === 0) {
                findings.push(makeFinding({
                    element: el,
                    category: 'Struktura',
                    priority: 'OBAVEZNO',
                    confidence: 0.99,
                    original: `[Prazan naslov nivoa ${el.headingLevel}]`,
                    replacement: '[dodati tekst naslova ili ukloniti prazan naslov]',
                    rationale: 'Naslov bez sadržaja — moguć artefakt formatiranja.',
                    autoFixable: false,
                }));
            }

            // Check for heading without following content
            const elIdx = docMap.elements.indexOf(el);
            if (elIdx < docMap.elements.length - 1) {
                const next = docMap.elements[elIdx + 1];
                if (next.type === 'heading') {
                    findings.push(makeFinding({
                        element: el,
                        category: 'Struktura',
                        priority: 'PREPORUKA',
                        confidence: 0.75,
                        original: el.text,
                        replacement: '[dodati sadržaj ispod naslova ili ukloniti naslov]',
                        rationale: 'Naslov koji je odmah praćen sledećim naslovom bez sadržaja između.',
                        autoFixable: false,
                    }));
                }
            }
        }
    }



    // ==========================================
    // HELPER: Create a finding object
    // ==========================================
    function makeFinding({ element, category, priority, confidence, original, replacement, rationale, autoFixable = false, globalPattern = false, requiresSourceVerification = false }) {
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
            isDirectQuote: false,
            requiresSourceVerification,
            autoFixable,
            globalPattern,
            status: 'OPEN',
        };
    }

    /**
     * Get section name for an element (find nearest preceding heading)
     */
    function getSectionName(element) {
        if (element.type === 'heading') return element.text || '(bez naslova)';
        if (element.section) return element.section;
        return element.id || '(nepoznata lokacija)';
    }

    /**
     * Get context around a position in text
     */
    function getContext(text, pos, radius) {
        const start = Math.max(0, pos - radius);
        const end = Math.min(text.length, pos + radius);
        let ctx = text.substring(start, end);
        if (start > 0) ctx = '...' + ctx;
        if (end < text.length) ctx = ctx + '...';
        return ctx;
    }

    /**
     * Simple Levenshtein distance
     */
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
                        matrix[i - 1][j] + 1
                    );
                }
            }
        }
        return matrix[b.length][a.length];
    }

    // Public API
    return { runAudit };
})();
