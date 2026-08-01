/**
 * Document Parser Module
 * Handles .docx (OOXML via JSZip), .md, and .txt files
 * Extracts structured document map with hashed IDs
 */

const DocumentParser = (() => {
    'use strict';

    // Safety limits
    const MAX_ZIP_FILES = 500;
    const MAX_UNCOMPRESSED_SIZE = 100 * 1024 * 1024; // 100MB
    const MAX_SINGLE_XML_SIZE = 50 * 1024 * 1024; // 50MB

    async function parse(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        const arrayBuffer = await file.arrayBuffer();

        let docMap;
        switch (ext) {
            case 'docx': docMap = await parseDocx(arrayBuffer, file.name); break;
            case 'md': docMap = await parseMarkdown(arrayBuffer, file.name); break;
            case 'txt': case 'text': docMap = await parsePlainText(arrayBuffer, file.name); break;
            default: throw new Error(`Nepodržan format: .${ext}`);
        }

        assignHashedIds(docMap);
        detectDirectQuotes(docMap);
        return docMap;
    }


    /**
     * Helper: safely load and size-check an XML file from zip.
     * Measures byte length (not string length) for size enforcement.
     */

    async function parseDocx(arrayBuffer, fileName) {
        let zip;
        try { zip = await JSZip.loadAsync(arrayBuffer); }
        catch (e) { throw new Error('Fajl nije validan DOCX (ZIP/OOXML paket).'); }

        const fileCount = Object.keys(zip.files).length;
        if (fileCount > MAX_ZIP_FILES) {
            throw new Error(`ZIP sadrži ${fileCount} fajlova (limit: ${MAX_ZIP_FILES}).`);
        }

        // Track actual decompressed size by loading each entry
        let totalDecompressedSize = 0;

        // Override safeLoadXml locally to track cumulative size
        async function loadXml(path) {
            const file = zip.file(path);
            if (!file) return null;
            const bytes = await file.async('uint8array');
            if (bytes.length > MAX_SINGLE_XML_SIZE) {
                throw new Error(`${path} prelazi dozvoljenu veličinu (${bytes.length} > ${MAX_SINGLE_XML_SIZE} bajtova).`);
            }
            totalDecompressedSize += bytes.length;
            if (totalDecompressedSize > MAX_UNCOMPRESSED_SIZE) {
                throw new Error('Ukupna nekompresovana veličina prelazi limit.');
            }
            return new TextDecoder('utf-8').decode(bytes);
        }

        const contentTypesXml = await loadXml('[Content_Types].xml');
        if (!contentTypesXml) throw new Error('Fajl nema validnu OOXML strukturu.');
        if (contentTypesXml.includes('vbaProject') || contentTypesXml.includes('.docm')) {
            throw new Error('Makro-omogućeni dokumenti nisu podržani.');
        }

        const docContent = await loadXml('word/document.xml');
        if (!docContent) throw new Error('word/document.xml nije pronađen.');

        const stylesContent = await loadXml('word/styles.xml');
        const styles = stylesContent ? parseStyles(stylesContent) : {};

        const numberingContent = await loadXml('word/numbering.xml');
        const numbering = numberingContent ? parseNumbering(numberingContent) : {};

        const fnContent = await loadXml('word/footnotes.xml');
        const footnotes = fnContent ? parseFootnotes(fnContent) : [];

        const enContent = await loadXml('word/endnotes.xml');
        const endnotes = enContent ? parseEndnotes(enContent) : [];

        let headers = [], footers = [];
        for (const [path, file] of Object.entries(zip.files)) {
            if (path.match(/^word\/header\d+\.xml$/)) {
                const c = await loadXml(path);
                if (c) headers.push({ path, text: extractTextFromXml(c) });
            }
            if (path.match(/^word\/footer\d+\.xml$/)) {
                const c = await loadXml(path);
                if (c) footers.push({ path, text: extractTextFromXml(c) });
            }
        }

        const elements = parseDocumentXml(docContent, styles, numbering);

        let htmlPreview = '';
        try { const r = await mammoth.convertToHtml({ arrayBuffer }); htmlPreview = r.value; }
        catch (e) { /* non-critical */ }

        return {
            type: 'docx', name: fileName, elements, footnotes, endnotes,
            headers, footers, styles, numbering, htmlPreview,
            rawText: elements.map(el => el.text).join('\n'),
            wordCount: countWords(elements.map(el => el.text).join(' ')),
            paragraphCount: elements.filter(el => el.type === 'paragraph').length,
            tableCount: elements.filter(el => el.type === 'table').length,
            headingCount: elements.filter(el => el.type === 'heading').length,
        };
    }


    function parseDocumentXml(xmlStr, styles, numbering) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) {
            throw new Error('Neispravan OOXML sadržaj.');
        }

        const elements = [];
        const ns = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
        };

        const body = doc.getElementsByTagNameNS(ns.w, 'body')[0];
        if (!body) return elements;

        let paragraphIndex = 0, tableIndex = 0;
        const listInstances = {};
        let lastWasNumbered = false;

        // Only iterate direct children of body
        for (const child of body.children) {
            const ln = child.localName;
            if (ln === 'p') {
                const el = parseParagraph(child, ns, styles, numbering,
                    paragraphIndex, listInstances, lastWasNumbered, elements);
                elements.push(el);
                lastWasNumbered = !!el.numId;
                paragraphIndex++;
            } else if (ln === 'tbl') {
                const el = parseTable(child, ns, tableIndex);
                elements.push(el);
                lastWasNumbered = false;
                tableIndex++;
            } else if (ln === 'sectPr') {
                // Section break resets list instances
                for (const k of Object.keys(listInstances)) delete listInstances[k];
                lastWasNumbered = false;
            }
        }
        return elements;
    }


    function parseParagraph(pNode, ns, styles, numbering, index, listInstances, lastWasNumbered, prevElements) {
        const runs = [];
        let fullText = '';

        // Try to get w14:paraId
        const paraId = pNode.getAttributeNS(ns.w14, 'paraId') ||
            pNode.getAttribute('w14:paraId') || null;

        const pPr = pNode.getElementsByTagNameNS(ns.w, 'pPr')[0];
        let styleName = 'Normal';
        let outlineLevel = -1;
        let numId = null;
        let numLevel = null;

        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(ns.w, 'pStyle')[0];
            if (pStyle) styleName = pStyle.getAttribute('w:val') || 'Normal';

            const outlineLvl = pPr.getElementsByTagNameNS(ns.w, 'outlineLvl')[0];
            if (outlineLvl) outlineLevel = parseInt(outlineLvl.getAttribute('w:val'), 10);

            const numPr = pPr.getElementsByTagNameNS(ns.w, 'numPr')[0];
            if (numPr) {
                const ilvl = numPr.getElementsByTagNameNS(ns.w, 'ilvl')[0];
                const nId = numPr.getElementsByTagNameNS(ns.w, 'numId')[0];
                if (ilvl) numLevel = parseInt(ilvl.getAttribute('w:val'), 10);
                if (nId) numId = nId.getAttribute('w:val');
            }

            // Inherit numbering from style if not explicitly set
            if (!numId && styles && styles._numFromStyle && styles._numFromStyle[styleName]) {
                const inherited = styles._numFromStyle[styleName];
                numId = inherited.numId;
                numLevel = inherited.ilvl ?? 0;
            }
        }

        const isHeading = styleName.match(/^Heading(\d+)$/) ||
            styleName.match(/^Naslov(\d+)$/) || outlineLevel >= 0;
        let headingLevel = 0;
        if (isHeading) {
            const match = styleName.match(/(\d+)$/);
            headingLevel = match ? parseInt(match[1], 10) : (outlineLevel + 1);
        }

        for (const child of pNode.children) {
            if (child.localName === 'r') {
                const runText = extractRunText(child, ns);
                const runProps = extractRunProps(child, ns);
                if (runText) { runs.push({ text: runText, ...runProps }); fullText += runText; }
            } else if (child.localName === 'hyperlink') {
                for (const hc of child.children) {
                    if (hc.localName === 'r') {
                        const rt = extractRunText(hc, ns);
                        if (rt) { runs.push({ text: rt, isHyperlink: true }); fullText += rt; }
                    }
                }
            }
        }


        // Calculate displayed number from OOXML numbering data
        let displayedNumber = null;
        let displayedLabel = null;
        let listInstanceId = null;
        let listStart = null;
        let numFmt = null;

        if (numId && numbering && numbering.nums && numbering.nums[numId]) {
            const numDef = numbering.nums[numId];
            const abstractId = numDef.abstractNumId;
            const abstractDef = numbering.abstractNums
                ? numbering.abstractNums[abstractId] : null;
            const level = numLevel ?? 0;

            if (abstractDef && abstractDef.levels && abstractDef.levels[level]) {
                const lvlDef = abstractDef.levels[level];
                numFmt = lvlDef.numFmt || 'decimal';

                // Exclude bullet and none formats from numeric tracking
                if (numFmt === 'bullet' || numFmt === 'none') {
                    // Still record numId but no displayed number
                    const type = isHeading ? 'heading' : 'paragraph';
                    return {
                        type, index, paraId, style: styleName, text: fullText, runs,
                        headingLevel: isHeading ? headingLevel : null,
                        numId, numLevel, displayedNumber: null, displayedLabel: null,
                        listInstanceId: null, listStart: null, numFmt,
                        isEmpty: fullText.trim().length === 0,
                        isDirectQuote: false, quoteConfidence: 0,
                    };
                }

                // Apply lvlOverride/startOverride
                let effectiveStart = lvlDef.start ?? 1;
                if (numDef.lvlOverrides && numDef.lvlOverrides[level]) {
                    const ov = numDef.lvlOverrides[level];
                    if (ov.startOverride != null) effectiveStart = ov.startOverride;
                }
                listStart = effectiveStart;

                // Determine if this is a new list instance
                // Multilevel key tracks all levels for this numId
                const multiKey = `${numId}`;

                // Initialize per-numId level counters if not present
                if (!listInstances[multiKey]) {
                    listInstances[multiKey] = {
                        counters: {},  // level → current count
                        instanceSeq: 0, // increments on each restart
                        lastLevel: -1,
                    };
                }
                const inst = listInstances[multiKey];

                // Determine restart conditions
                const prevEl = prevElements.length > 0 ? prevElements[prevElements.length - 1] : null;
                const prevWasSameNum = prevEl && prevEl.numId === numId;

                // New list instance if: no previous item in this numId, or interrupted by non-list content
                if (!prevWasSameNum && !lastWasNumbered) {
                    // Reset ALL level counters — this is a brand new list
                    inst.counters = {};
                    inst.instanceSeq++;
                }

                // If higher level incremented, reset lower levels (lvlRestart logic)
                if (level < inst.lastLevel) {
                    // We went UP a level — don't reset, this is valid
                } else if (level > inst.lastLevel && inst.lastLevel >= 0) {
                    // We went DOWN — this is a sub-level, initialize it
                    // Reset only if first time at this sub-level in this sequence
                    if (inst.counters[level] === undefined) {
                        inst.counters[level] = effectiveStart - 1;
                    }
                }

                // Initialize counter for this level if needed
                if (inst.counters[level] === undefined) {
                    inst.counters[level] = effectiveStart - 1;
                }

                // Increment the current level
                inst.counters[level]++;
                inst.lastLevel = level;

                // If this level has lvlRestart pointing to a higher level,
                // and that higher level just incremented, reset this level
                // (handled by the higher-resets-lower logic above)
                // Reset lower levels when a higher level advances
                if (prevWasSameNum && prevEl.numLevel != null && prevEl.numLevel > level) {
                    // Higher level (lower number) advanced — reset all levels below
                    for (const k of Object.keys(inst.counters)) {
                        if (parseInt(k) > level) {
                            delete inst.counters[k];
                        }
                    }
                }

                displayedNumber = inst.counters[level];
                listInstanceId = `${numId}-${inst.instanceSeq}-${level}`;

                // Format displayedLabel using numFmt and lvlText
                displayedLabel = formatNumber(displayedNumber, numFmt, lvlDef.lvlText, level);
            }
        }

        const type = isHeading ? 'heading' : 'paragraph';

        return {
            type, index, paraId, style: styleName, text: fullText, runs,
            headingLevel: isHeading ? headingLevel : null,
            numId, numLevel, displayedNumber, displayedLabel,
            listInstanceId, listStart, numFmt,
            isEmpty: fullText.trim().length === 0,
            isDirectQuote: false, quoteConfidence: 0,
        };
    }

    /**
     * Format a number according to numFmt and lvlText template.
     */
    function formatNumber(num, numFmt, lvlText, level) {
        let formatted;
        switch (numFmt) {
            case 'decimal': formatted = String(num); break;
            case 'lowerLetter': formatted = String.fromCharCode(96 + ((num - 1) % 26) + 1); break;
            case 'upperLetter': formatted = String.fromCharCode(64 + ((num - 1) % 26) + 1); break;
            case 'lowerRoman': formatted = toRoman(num).toLowerCase(); break;
            case 'upperRoman': formatted = toRoman(num); break;
            default: formatted = String(num);
        }
        // Replace %N placeholder in lvlText (e.g., "%1." → "1.")
        let label = lvlText || `%${level + 1}.`;
        label = label.replace(`%${level + 1}`, formatted);
        return label;
    }

    function toRoman(num) {
        const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
        const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
        let result = '';
        for (let i = 0; i < vals.length; i++) {
            while (num >= vals[i]) { result += syms[i]; num -= vals[i]; }
        }
        return result;
    }


    function extractRunText(rNode, ns) {
        let text = '';
        for (const child of rNode.children) {
            if (child.localName === 't') text += child.textContent;
            else if (child.localName === 'tab') text += '\t';
            else if (child.localName === 'br') text += '\n';
            else if (child.localName === 'sym') text += '\u25A1';
        }
        return text;
    }

    function extractRunProps(rNode, ns) {
        const rPr = rNode.getElementsByTagNameNS(ns.w, 'rPr')[0];
        if (!rPr) return {};
        const props = {};
        if (rPr.getElementsByTagNameNS(ns.w, 'b')[0]) props.bold = true;
        if (rPr.getElementsByTagNameNS(ns.w, 'i')[0]) props.italic = true;
        if (rPr.getElementsByTagNameNS(ns.w, 'u')[0]) props.underline = true;
        const lang = rPr.getElementsByTagNameNS(ns.w, 'lang')[0];
        if (lang) props.lang = lang.getAttribute('w:val') || lang.getAttribute('w:bidi');
        return props;
    }


    /**
     * Parse table: only direct child rows/cells (not nested tables).
     * Checks w:tblHeader for actual header detection.
     */
    function parseTable(tblNode, ns, tableIdx) {
        const rows = [];
        let hasActualHeader = false;
        let rowIndex = 0;

        // Collect all cell text for table hash
        let tableTextForHash = '';

        for (const child of tblNode.children) {
            if (child.localName !== 'tr') continue;
            const tr = child;

            const trPr = tr.getElementsByTagNameNS(ns.w, 'trPr')[0];
            if (trPr) {
                const tblHeader = trPr.getElementsByTagNameNS(ns.w, 'tblHeader')[0];
                if (tblHeader) hasActualHeader = true;
            }

            const cells = [];
            let colIndex = 0;

            for (const tcChild of tr.children) {
                if (tcChild.localName !== 'tc') continue;
                const tc = tcChild;

                const cellParagraphs = [];
                let cellText = '';
                for (const tcContent of tc.children) {
                    if (tcContent.localName === 'p') {
                        const pText = extractTextFromElement(tcContent, ns);
                        cellParagraphs.push(pText);
                        cellText += pText + ' ';
                    }
                }
                cellText = cellText.trim();
                tableTextForHash += cellText;

                cells.push({
                    text: cellText,
                    rowIndex, columnIndex: colIndex,
                    paragraphs: cellParagraphs,
                    // tableId/rowId/cellId assigned after hash
                });
                colIndex++;
            }
            rows.push(cells);
            rowIndex++;
        }

        // Generate hashed table ID from content
        const tableId = hashId('table', `table|${tableIdx}|${tableTextForHash.substring(0, 200)}`);

        // Assign hashed row/cell IDs
        for (let ri = 0; ri < rows.length; ri++) {
            const rowId = hashId('table', `${tableId}|row|${ri}|${rows[ri].map(c=>c.text).join('|').substring(0,100)}`);
            for (let ci = 0; ci < rows[ri].length; ci++) {
                const cell = rows[ri][ci];
                cell.tableId = tableId;
                cell.rowId = rowId;
                cell.cellId = hashId('table', `${rowId}|cell|${ci}|${cell.text.substring(0,50)}`);
            }
        }

        return {
            type: 'table', index: tableIdx, tableId, rows,
            text: rows.map(r => r.map(c => c.text).join(' | ')).join('\n'),
            hasHeader: hasActualHeader,
            columnCount: rows.length > 0 ? rows[0].length : 0,
            rowCount: rows.length,
        };
    }

    function extractTextFromElement(node, ns) {
        let text = '';
        const tNodes = node.getElementsByTagNameNS(ns.w, 't');
        for (const t of tNodes) text += t.textContent;
        return text;
    }


    /**
     * Full parseNumbering: numId→abstractNumId, levels, lvlOverride, startOverride, lvlRestart
     */
    function parseNumbering(xmlStr) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) return {};

        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const result = { abstractNums: {}, nums: {} };

        // Parse abstractNum elements
        const abstractNumNodes = doc.getElementsByTagNameNS(ns, 'abstractNum');
        for (const an of abstractNumNodes) {
            const abstractNumId = an.getAttribute('w:abstractNumId');
            const levels = {};

            const lvlNodes = an.getElementsByTagNameNS(ns, 'lvl');
            for (const lvl of lvlNodes) {
                const ilvl = parseInt(lvl.getAttribute('w:ilvl'), 10);
                const startNode = lvl.getElementsByTagNameNS(ns, 'start')[0];
                const numFmtNode = lvl.getElementsByTagNameNS(ns, 'numFmt')[0];
                const lvlTextNode = lvl.getElementsByTagNameNS(ns, 'lvlText')[0];
                const lvlRestartNode = lvl.getElementsByTagNameNS(ns, 'lvlRestart')[0];

                levels[ilvl] = {
                    start: startNode ? parseInt(startNode.getAttribute('w:val'), 10) : null,
                    numFmt: numFmtNode ? numFmtNode.getAttribute('w:val') : 'decimal',
                    lvlText: lvlTextNode ? lvlTextNode.getAttribute('w:val') : '%1.',
                    lvlRestart: lvlRestartNode
                        ? parseInt(lvlRestartNode.getAttribute('w:val'), 10) : null,
                };
            }
            result.abstractNums[abstractNumId] = { levels };
        }

        // Parse num elements with lvlOverride support
        const numNodes = doc.getElementsByTagNameNS(ns, 'num');
        for (const num of numNodes) {
            const numId = num.getAttribute('w:numId');
            const abstractNumIdRef = num.getElementsByTagNameNS(ns, 'abstractNumId')[0];
            const entry = {
                abstractNumId: abstractNumIdRef ? abstractNumIdRef.getAttribute('w:val') : null,
                lvlOverrides: {},
            };

            // Parse lvlOverride elements
            const lvlOverrideNodes = num.getElementsByTagNameNS(ns, 'lvlOverride');
            for (const ov of lvlOverrideNodes) {
                const ilvl = parseInt(ov.getAttribute('w:ilvl'), 10);
                const startOverrideNode = ov.getElementsByTagNameNS(ns, 'startOverride')[0];
                entry.lvlOverrides[ilvl] = {
                    startOverride: startOverrideNode
                        ? parseInt(startOverrideNode.getAttribute('w:val'), 10) : null,
                };
            }

            result.nums[numId] = entry;
        }

        return result;
    }


    function parseFootnotes(xmlStr) {
        const footnotes = [];
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) return [];
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const fnNodes = doc.getElementsByTagNameNS(ns, 'footnote');
        for (const fn of fnNodes) {
            const id = fn.getAttribute('w:id');
            const type = fn.getAttribute('w:type');
            if (type === 'separator' || type === 'continuationSeparator') continue;
            const text = extractTextFromXmlNode(fn, ns);
            footnotes.push({ id, text, isEmpty: text.trim().length === 0 });
        }
        return footnotes;
    }

    function parseEndnotes(xmlStr) {
        const endnotes = [];
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) return [];
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const enNodes = doc.getElementsByTagNameNS(ns, 'endnote');
        for (const en of enNodes) {
            const id = en.getAttribute('w:id');
            const type = en.getAttribute('w:type');
            if (type === 'separator' || type === 'continuationSeparator') continue;
            const text = extractTextFromXmlNode(en, ns);
            endnotes.push({ id, text, isEmpty: text.trim().length === 0 });
        }
        return endnotes;
    }

    function extractTextFromXmlNode(node, ns) {
        let text = '';
        const tNodes = node.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) text += t.textContent;
        return text;
    }

    function extractTextFromXml(xmlStr) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        let text = '';
        const tNodes = doc.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) text += t.textContent + ' ';
        return text.trim();
    }

    function parseStyles(xmlStr) {
        const styles = {};
        styles._numFromStyle = {}; // numId/ilvl inherited from paragraph styles
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) return styles;
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const styleNodes = doc.getElementsByTagNameNS(ns, 'style');

        for (const s of styleNodes) {
            const id = s.getAttribute('w:styleId');
            const nameNode = s.getElementsByTagNameNS(ns, 'name')[0];
            if (id && nameNode) styles[id] = nameNode.getAttribute('w:val');

            // Check if style defines numbering (inherited by paragraphs using this style)
            const pPr = s.getElementsByTagNameNS(ns, 'pPr')[0];
            if (pPr && id) {
                const numPr = pPr.getElementsByTagNameNS(ns, 'numPr')[0];
                if (numPr) {
                    const numIdNode = numPr.getElementsByTagNameNS(ns, 'numId')[0];
                    const ilvlNode = numPr.getElementsByTagNameNS(ns, 'ilvl')[0];
                    if (numIdNode) {
                        styles._numFromStyle[id] = {
                            numId: numIdNode.getAttribute('w:val'),
                            ilvl: ilvlNode ? parseInt(ilvlNode.getAttribute('w:val'), 10) : 0,
                        };
                    }
                }
            }
        }
        return styles;
    }


    async function parseMarkdown(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const hm = line.match(/^(#{1,6})\s+(.*)$/);
            if (hm) {
                elements.push({
                    type: 'heading', index: index++, headingLevel: hm[1].length,
                    text: hm[2], style: `Heading${hm[1].length}`,
                    runs: [{ text: hm[2] }], lineNumber: i + 1,
                    isDirectQuote: false, quoteConfidence: 0, paraId: null,
                });
            } else if (line.trim().length > 0) {
                elements.push({
                    type: 'paragraph', index: index++, text: line, style: 'Normal',
                    runs: [{ text: line }], lineNumber: i + 1, isEmpty: false,
                    isDirectQuote: false, quoteConfidence: 0, paraId: null,
                });
            }
        }
        return {
            type: 'markdown', name: fileName, elements,
            footnotes: [], endnotes: [], headers: [], footers: [],
            styles: {}, numbering: {}, htmlPreview: '',
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.filter(el => el.type === 'paragraph').length,
            tableCount: 0, headingCount: elements.filter(el => el.type === 'heading').length,
        };
    }

    async function parsePlainText(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;
        for (let i = 0; i < lines.length; i++) {
            if (lines[i].trim().length > 0) {
                elements.push({
                    type: 'paragraph', index: index++, text: lines[i], style: 'Normal',
                    runs: [{ text: lines[i] }], lineNumber: i + 1, isEmpty: false,
                    isDirectQuote: false, quoteConfidence: 0, paraId: null,
                });
            }
        }
        return {
            type: 'txt', name: fileName, elements,
            footnotes: [], endnotes: [], headers: [], footers: [],
            styles: {}, numbering: {}, htmlPreview: '',
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.length, tableCount: 0, headingCount: 0,
        };
    }


    /**
     * Assign hashed IDs: uses w14:paraId first, then content hash.
     * Hash is based on type + style + text + previousText + nextText (no absolute index).
     */
    function assignHashedIds(docMap) {
        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];

            // Prefer w14:paraId if available
            if (el.paraId) {
                const typePrefix = el.type === 'heading' ? 'h' : el.type === 'table' ? 't' : 'p';
                el.id = `${typePrefix}-${el.paraId}`;
                continue;
            }

            const prevText = i > 0 ? (docMap.elements[i - 1].text || '').substring(0, 30) : '';
            const nextText = i < docMap.elements.length - 1
                ? (docMap.elements[i + 1].text || '').substring(0, 30) : '';
            const contextStr = `${el.type}|${el.style || ''}|${el.text || ''}|${prevText}|${nextText}`;
            el.id = hashId(el.type, contextStr);
        }
    }

    function hashId(prefix, input) {
        let hash = 0x811c9dc5;
        for (let i = 0; i < input.length; i++) {
            hash ^= input.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        const hashStr = (hash >>> 0).toString(36).padStart(7, '0');
        const typePrefix = prefix === 'heading' ? 'h' : prefix === 'table' ? 't' : 'p';
        return `${typePrefix}-${hashStr}`;
    }

    /**
     * Detect direct quotes. Uses multiline-aware regex.
     * Does NOT mark pure-italic paragraphs without an additional signal.
     */
    function detectDirectQuotes(docMap) {
        for (const el of docMap.elements) {
            if (!el.text || el.type === 'heading') continue;
            const text = el.text.trim();
            let confidence = 0;

            // Pattern 1: Typographic quotes (multiline-aware)
            if (/^\u201E[\s\S]*\u201C$/.test(text)) {
                confidence = 0.95;
            }
            // Pattern 2: Guillemets
            else if (/^\u00AB[\s\S]*\u00BB$/.test(text)) {
                confidence = 0.90;
            }
            // Pattern 3: Italic + additional signal (quote style or length > 100)
            else if (el.runs && el.runs.length > 0 && el.runs.every(r => r.italic) &&
                text.length > 20 &&
                (el.style && el.style.match(/quote|citat|blockquote/i) || text.length > 100)) {
                confidence = 0.80;
            }
            // Pattern 4: Em-dash dialogue
            else if (/^[\u2014\u2015]\s/.test(text)) {
                confidence = 0.80;
            }
            // Pattern 5: Quote/Citat style (without italic requirement)
            else if (el.style && el.style.match(/quote|citat|blockquote/i)) {
                confidence = 0.90;
            }
            // Pattern 6: Entirely Greek text
            else if (/^[\u0370-\u03FF\u1F00-\u1FFF\s,.;\u00B7'"]+$/.test(text)) {
                confidence = 0.85;
            }

            if (confidence > 0) {
                el.isDirectQuote = true;
                el.quoteConfidence = confidence;
            }
        }
    }

    function countWords(text) {
        return text.split(/\s+/).filter(w => w.length > 0).length;
    }

    // Public API
    return { parse, parseNumbering, hashId, formatNumber };
})();

if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentParser;
}
