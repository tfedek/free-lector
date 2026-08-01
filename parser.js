/**
 * Document Parser Module — Round 4
 * Full OOXML parsing with tracked changes, basedOn chain, lvlRestart, multilevel labels
 */

const DocumentParser = (() => {
    'use strict';

    const MAX_ZIP_FILES = 500;
    const MAX_UNCOMPRESSED_SIZE = 100 * 1024 * 1024;
    const MAX_SINGLE_XML_SIZE = 50 * 1024 * 1024;

    // Tracked changes mode: 'accept' | 'show_deleted' | 'ignore_deleted'
    let trackedChangesMode = 'accept';

    function setTrackedChangesMode(mode) {
        trackedChangesMode = mode;
    }

    async function parse(file, options = {}) {
        if (options.trackedChanges) trackedChangesMode = options.trackedChanges;
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


    // ==========================================
    // DOCX PARSING
    // ==========================================
    async function parseDocx(arrayBuffer, fileName) {
        let zip;
        try { zip = await JSZip.loadAsync(arrayBuffer); }
        catch (e) { throw new Error('Fajl nije validan DOCX (ZIP/OOXML paket).'); }

        const fileCount = Object.keys(zip.files).length;
        if (fileCount > MAX_ZIP_FILES) {
            throw new Error(`ZIP sadrži ${fileCount} fajlova (limit: ${MAX_ZIP_FILES}).`);
        }

        // Reject dangerous content
        if (zip.file('word/vbaProject.bin')) {
            throw new Error('Dokument sadrži VBA makroe (word/vbaProject.bin). Odbijeno.');
        }
        for (const path of Object.keys(zip.files)) {
            if (path.startsWith('word/embeddings/')) {
                throw new Error(`Dokument sadrži ugrađene objekte (${path}). Odbijeno.`);
            }
        }

        // Track total decompressed size across ALL entries (images, media, etc.)
        let totalDecompressed = 0;

        async function loadEntry(path) {
            const file = zip.file(path);
            if (!file) return null;
            const bytes = await file.async('uint8array');
            totalDecompressed += bytes.length;
            if (totalDecompressed > MAX_UNCOMPRESSED_SIZE) {
                throw new Error(`Ukupna raspakovanja veličina prelazi ${MAX_UNCOMPRESSED_SIZE} bajtova.`);
            }
            if (bytes.length > MAX_SINGLE_XML_SIZE) {
                throw new Error(`${path} prelazi ${MAX_SINGLE_XML_SIZE} bajtova.`);
            }
            return new TextDecoder('utf-8').decode(bytes);
        }

        // Also count non-XML entries toward total size
        for (const [path, entry] of Object.entries(zip.files)) {
            if (entry.dir) continue;
            if (path.endsWith('.xml') || path.endsWith('.rels')) continue; // counted when loaded
            // For binary entries (images, etc), get size
            try {
                const bytes = await entry.async('uint8array');
                totalDecompressed += bytes.length;
                if (totalDecompressed > MAX_UNCOMPRESSED_SIZE) {
                    throw new Error(`Ukupna raspakovanja veličina prelazi limit (binarne datoteke).`);
                }
            } catch (e) {
                if (e.message.includes('prelazi')) throw e;
            }
        }

        const contentTypesXml = await loadEntry('[Content_Types].xml');
        if (!contentTypesXml) throw new Error('Fajl nema validnu OOXML strukturu.');
        if (contentTypesXml.includes('vbaProject')) {
            throw new Error('Makro-omogućeni dokumenti nisu podržani.');
        }


        function parseXmlStrict(content, filename) {
            const parser = new DOMParser();
            const doc = parser.parseFromString(content, 'application/xml');
            if (doc.getElementsByTagName('parsererror').length) {
                throw new Error(`Neispravan XML: ${filename}`);
            }
            return doc;
        }

        const docContent = await loadEntry('word/document.xml');
        if (!docContent) throw new Error('word/document.xml nije pronađen.');

        const stylesContent = await loadEntry('word/styles.xml');
        const styles = stylesContent ? parseStyles(stylesContent, 'word/styles.xml', parseXmlStrict) : { _numFromStyle: {}, _basedOn: {} };

        const numberingContent = await loadEntry('word/numbering.xml');
        const numbering = numberingContent ? parseNumbering(numberingContent, 'word/numbering.xml', parseXmlStrict) : {};

        const fnContent = await loadEntry('word/footnotes.xml');
        const footnotes = fnContent ? parseFootnotes(fnContent, 'word/footnotes.xml', parseXmlStrict) : [];

        const enContent = await loadEntry('word/endnotes.xml');
        const endnotes = enContent ? parseEndnotes(enContent, 'word/endnotes.xml', parseXmlStrict) : [];

        let headers = [], footers = [];
        for (const path of Object.keys(zip.files)) {
            if (path.match(/^word\/header\d+\.xml$/)) {
                const c = await loadEntry(path);
                if (c) headers.push({ path, text: extractTextFromXml(c, path, parseXmlStrict) });
            }
            if (path.match(/^word\/footer\d+\.xml$/)) {
                const c = await loadEntry(path);
                if (c) footers.push({ path, text: extractTextFromXml(c, path, parseXmlStrict) });
            }
        }

        // Resolve basedOn chain for numbering inheritance
        resolveBasedOnNumbering(styles);

        const elements = parseDocumentXml(docContent, styles, numbering, parseXmlStrict);

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


    // ==========================================
    // DOCUMENT XML PARSING
    // ==========================================
    function parseDocumentXml(xmlStr, styles, numbering, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, 'word/document.xml');
        const elements = [];
        const ns = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
        };
        const body = doc.getElementsByTagNameNS(ns.w, 'body')[0];
        if (!body) return elements;

        let paragraphIndex = 0, tableIndex = 0;
        const listState = { instances: {}, seq: 0 };

        for (const child of body.children) {
            const ln = child.localName;
            if (ln === 'p') {
                const el = parseParagraph(child, ns, styles, numbering, paragraphIndex, listState, elements);
                elements.push(el);
                paragraphIndex++;
            } else if (ln === 'tbl') {
                elements.push(parseTable(child, ns, tableIndex));
                resetListState(listState);
                tableIndex++;
            } else if (ln === 'sectPr') {
                resetListState(listState);
            }
        }
        return elements;
    }

    function resetListState(listState) {
        listState.instances = {};
        listState.seq++;
    }


    // ==========================================
    // PARAGRAPH PARSING
    // ==========================================
    function parseParagraph(pNode, ns, styles, numbering, index, listState, prevElements) {
        const runs = [];
        let fullText = '';
        const paraId = pNode.getAttributeNS(ns.w14, 'paraId') ||
            pNode.getAttribute('w14:paraId') || null;

        const pPr = getDirectChild(pNode, ns.w, 'pPr');
        let styleName = 'Normal';
        let outlineLevel = -1;
        let numId = null;
        let numLevel = null;

        if (pPr) {
            const pStyle = getDirectChild(pPr, ns.w, 'pStyle');
            if (pStyle) styleName = pStyle.getAttribute('w:val') || 'Normal';
            const outlineLvl = getDirectChild(pPr, ns.w, 'outlineLvl');
            if (outlineLvl) outlineLevel = parseInt(outlineLvl.getAttribute('w:val'), 10);
            const numPr = getDirectChild(pPr, ns.w, 'numPr');
            if (numPr) {
                const ilvl = getDirectChild(numPr, ns.w, 'ilvl');
                const nId = getDirectChild(numPr, ns.w, 'numId');
                if (ilvl) numLevel = parseInt(ilvl.getAttribute('w:val'), 10);
                if (nId) numId = nId.getAttribute('w:val');
            }
            // Inherit numbering from style (including basedOn chain)
            if (!numId && styles._numFromStyle && styles._numFromStyle[styleName]) {
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

        // Extract text recursively (handles ins/del/sdt/smartTag/fldSimple/customXml)
        fullText = extractVisibleText(pNode, ns);
        // Build runs for formatting info
        buildRuns(pNode, ns, runs);


        // ---- NUMBERING LOGIC ----
        let displayedNumber = null;
        let displayedLabel = null;
        let listInstanceId = null;
        let listStart = null;
        let numFmt = null;

        if (numId && numbering && numbering.nums && numbering.nums[numId]) {
            const numDef = numbering.nums[numId];
            const abstractId = numDef.abstractNumId;
            const abstractDef = numbering.abstractNums ? numbering.abstractNums[abstractId] : null;
            const level = numLevel ?? 0;

            if (abstractDef && abstractDef.levels && abstractDef.levels[level]) {
                let lvlDef = abstractDef.levels[level];

                // Apply full lvlOverride (may replace entire level definition)
                if (numDef.lvlOverrides && numDef.lvlOverrides[level]) {
                    const ov = numDef.lvlOverrides[level];
                    if (ov.lvlDef) lvlDef = { ...lvlDef, ...ov.lvlDef };
                    if (ov.startOverride != null) lvlDef = { ...lvlDef, start: ov.startOverride };
                }

                numFmt = lvlDef.numFmt || 'decimal';

                // Exclude bullet/none from numeric tracking
                if (numFmt === 'bullet' || numFmt === 'none') {
                    return buildParaResult(isHeading, index, paraId, styleName, fullText, runs,
                        headingLevel, numId, numLevel, null, null, null, null, numFmt);
                }

                const effectiveStart = lvlDef.start ?? 1;
                listStart = effectiveStart;

                // Determine list instance lifecycle
                const prevEl = prevElements.length > 0 ? prevElements[prevElements.length - 1] : null;
                const prevWasSameNum = prevEl && prevEl.numId === numId;
                const interrupted = !prevWasSameNum && prevEl &&
                    (prevEl.type === 'heading' || prevEl.type === 'table' ||
                     (prevEl.type === 'paragraph' && !prevEl.numId));

                // New instance when: first occurrence, interrupted, or numId reappears after different list
                const instanceKey = numId;
                if (!listState.instances[instanceKey] || interrupted) {
                    listState.seq++;
                    listState.instances[instanceKey] = {
                        counters: {},
                        seq: listState.seq,
                        lastLevel: -1,
                    };
                }
                const inst = listState.instances[instanceKey];

                // lvlRestart: when a HIGHER level (lower number) advances, reset this level
                if (prevWasSameNum && prevEl.numLevel != null && prevEl.numLevel < level) {
                    // Higher level advanced → reset current and all lower
                    for (let l = level; l <= 9; l++) {
                        if (inst.counters[l] !== undefined) {
                            const lvlDefL = abstractDef.levels[l];
                            inst.counters[l] = (lvlDefL ? (lvlDefL.start ?? 1) : 1) - 1;
                        }
                    }
                }

                // Initialize counter if needed
                if (inst.counters[level] === undefined) {
                    inst.counters[level] = effectiveStart - 1;
                }
                inst.counters[level]++;
                inst.lastLevel = level;

                displayedNumber = inst.counters[level];
                listInstanceId = `${numId}-${inst.seq}-${level}`;

                // Format label using all level counters
                displayedLabel = formatLabel(lvlDef.lvlText || `%${level+1}.`, inst.counters, abstractDef.levels);
            }
        }

        // If heading or non-numbered paragraph, reset list state
        if (!numId && (isHeading || !fullText.trim())) {
            resetListState(listState);
        } else if (!numId) {
            // Plain paragraph interrupts lists
            listState.instances = {};
        }

        return buildParaResult(isHeading, index, paraId, styleName, fullText, runs,
            headingLevel, numId, numLevel, displayedNumber, displayedLabel,
            listInstanceId, listStart, numFmt);
    }

    function buildParaResult(isHeading, index, paraId, styleName, fullText, runs,
        headingLevel, numId, numLevel, displayedNumber, displayedLabel,
        listInstanceId, listStart, numFmt) {
        return {
            type: isHeading ? 'heading' : 'paragraph',
            index, paraId, style: styleName, text: fullText, runs,
            headingLevel: isHeading ? headingLevel : null,
            numId, numLevel, displayedNumber, displayedLabel,
            listInstanceId, listStart, numFmt,
            isEmpty: fullText.trim().length === 0,
            isDirectQuote: false, quoteConfidence: 0,
        };
    }


    // ==========================================
    // RECURSIVE VISIBLE TEXT EXTRACTION
    // Handles: w:r, w:hyperlink, w:ins, w:del, w:sdt, w:smartTag, w:fldSimple, w:customXml
    // ==========================================
    function extractVisibleText(node, ns) {
        let text = '';
        for (const child of node.children) {
            const ln = child.localName;
            if (ln === 'r') {
                text += extractRunTextFromNode(child, ns);
            } else if (ln === 'hyperlink' || ln === 'smartTag' || ln === 'customXml' || ln === 'fldSimple') {
                text += extractVisibleText(child, ns);
            } else if (ln === 'sdt') {
                // Extract from sdtContent
                const sdtContent = getDirectChild(child, ns.w, 'sdtContent');
                if (sdtContent) text += extractVisibleText(sdtContent, ns);
            } else if (ln === 'ins') {
                // Tracked change: insertion — always include
                text += extractVisibleText(child, ns);
            } else if (ln === 'del') {
                // Tracked change: deletion
                if (trackedChangesMode === 'show_deleted') {
                    text += extractVisibleText(child, ns);
                }
                // 'accept' and 'ignore_deleted' both skip deleted text
            } else if (ln === 'pPr' || ln === 'rPr' || ln === 'sectPr' || ln === 'bookmarkStart' || ln === 'bookmarkEnd') {
                // Skip non-text elements
            } else if (child.children && child.children.length > 0) {
                // Recurse into unknown containers
                text += extractVisibleText(child, ns);
            }
        }
        return text;
    }

    function extractRunTextFromNode(rNode, ns) {
        let text = '';
        for (const child of rNode.children) {
            const ln = child.localName;
            if (ln === 't') text += child.textContent;
            else if (ln === 'tab') text += '\t';
            else if (ln === 'br') text += '\n';
            else if (ln === 'sym') text += '\u25A1';
            else if (ln === 'delText') {
                if (trackedChangesMode === 'show_deleted') text += child.textContent;
            }
        }
        return text;
    }

    function buildRuns(node, ns, runs) {
        for (const child of node.children) {
            const ln = child.localName;
            if (ln === 'r') {
                const text = extractRunTextFromNode(child, ns);
                if (text) {
                    const rPr = getDirectChild(child, ns.w, 'rPr');
                    const props = {};
                    if (rPr) {
                        if (getDirectChild(rPr, ns.w, 'b')) props.bold = true;
                        if (getDirectChild(rPr, ns.w, 'i')) props.italic = true;
                        if (getDirectChild(rPr, ns.w, 'u')) props.underline = true;
                    }
                    runs.push({ text, ...props });
                }
            } else if (ln === 'hyperlink' || ln === 'ins' || ln === 'smartTag' ||
                       ln === 'customXml' || ln === 'fldSimple' || ln === 'sdt') {
                const target = ln === 'sdt' ? (getDirectChild(child, ns.w, 'sdtContent') || child) : child;
                buildRuns(target, ns, runs);
            } else if (ln === 'del') {
                if (trackedChangesMode === 'show_deleted') buildRuns(child, ns, runs);
            }
        }
    }

    function getDirectChild(parent, nsUri, localName) {
        for (const child of parent.children) {
            if (child.localName === localName &&
                (!nsUri || child.namespaceURI === nsUri)) return child;
        }
        return null;
    }


    // ==========================================
    // TABLE PARSING — hashed IDs without indices/substrings
    // ==========================================
    function parseTable(tblNode, ns, tableIdx) {
        const rows = [];
        let hasActualHeader = false;
        let allCellTexts = [];

        for (const child of tblNode.children) {
            if (child.localName !== 'tr') continue;
            const trPr = getDirectChild(child, ns.w, 'trPr');
            if (trPr && getDirectChild(trPr, ns.w, 'tblHeader')) hasActualHeader = true;

            const cells = [];
            let rowTexts = [];
            for (const tcChild of child.children) {
                if (tcChild.localName !== 'tc') continue;
                const cellParagraphs = [];
                let cellText = '';
                for (const tcContent of tcChild.children) {
                    if (tcContent.localName === 'p') {
                        const pText = extractVisibleText(tcContent, ns);
                        cellParagraphs.push(pText);
                        cellText += pText + ' ';
                    }
                }
                cellText = cellText.trim();
                rowTexts.push(cellText);
                allCellTexts.push(cellText);
                cells.push({ text: cellText, paragraphs: cellParagraphs,
                    rowIndex: rows.length, columnIndex: cells.length });
            }
            rows.push(cells);
        }

        // Generate hashed IDs from full normalized content
        const tableContent = allCellTexts.join('|');
        const tableId = hashId('table', `table|${tableContent}`);

        for (let ri = 0; ri < rows.length; ri++) {
            const rowContent = rows[ri].map(c => c.text).join('|');
            const rowId = hashId('table', `${tableId}|row|${rowContent}`);
            for (let ci = 0; ci < rows[ri].length; ci++) {
                const cell = rows[ri][ci];
                cell.tableId = tableId;
                cell.rowId = rowId;
                cell.cellId = hashId('table', `${rowId}|cell|${cell.text}`);
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


    // ==========================================
    // NUMBERING PARSING — full lvlOverride with w:lvl
    // ==========================================
    function parseNumbering(xmlStr, filename, parseXmlStrict) {
        if (!filename) filename = 'numbering.xml';
        if (!parseXmlStrict) {
            parseXmlStrict = (content, fname) => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(content, 'application/xml');
                if (doc.getElementsByTagName('parsererror').length) {
                    throw new Error(`Neispravan XML: ${fname}`);
                }
                return doc;
            };
        }
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const result = { abstractNums: {}, nums: {} };

        const abstractNumNodes = doc.getElementsByTagNameNS(ns, 'abstractNum');
        for (const an of abstractNumNodes) {
            const abstractNumId = an.getAttribute('w:abstractNumId');
            const levels = {};
            const lvlNodes = an.getElementsByTagNameNS(ns, 'lvl');
            for (const lvl of lvlNodes) {
                levels[parseInt(lvl.getAttribute('w:ilvl'), 10)] = parseLvlNode(lvl, ns);
            }
            result.abstractNums[abstractNumId] = { levels };
        }

        const numNodes = doc.getElementsByTagNameNS(ns, 'num');
        for (const num of numNodes) {
            const numId = num.getAttribute('w:numId');
            const abstractNumIdRef = num.getElementsByTagNameNS(ns, 'abstractNumId')[0];
            const entry = {
                abstractNumId: abstractNumIdRef ? abstractNumIdRef.getAttribute('w:val') : null,
                lvlOverrides: {},
            };
            // Parse lvlOverride with full w:lvl support
            const ovNodes = num.getElementsByTagNameNS(ns, 'lvlOverride');
            for (const ov of ovNodes) {
                const ilvl = parseInt(ov.getAttribute('w:ilvl'), 10);
                const startOverrideNode = ov.getElementsByTagNameNS(ns, 'startOverride')[0];
                const lvlNode = ov.getElementsByTagNameNS(ns, 'lvl')[0];
                entry.lvlOverrides[ilvl] = {
                    startOverride: startOverrideNode
                        ? parseInt(startOverrideNode.getAttribute('w:val'), 10) : null,
                    lvlDef: lvlNode ? parseLvlNode(lvlNode, ns) : null,
                };
            }
            result.nums[numId] = entry;
        }
        return result;
    }

    function parseLvlNode(lvl, ns) {
        const startNode = lvl.getElementsByTagNameNS(ns, 'start')[0];
        const numFmtNode = lvl.getElementsByTagNameNS(ns, 'numFmt')[0];
        const lvlTextNode = lvl.getElementsByTagNameNS(ns, 'lvlText')[0];
        const lvlRestartNode = lvl.getElementsByTagNameNS(ns, 'lvlRestart')[0];
        return {
            start: startNode ? parseInt(startNode.getAttribute('w:val'), 10) : null,
            numFmt: numFmtNode ? numFmtNode.getAttribute('w:val') : 'decimal',
            lvlText: lvlTextNode ? lvlTextNode.getAttribute('w:val') : '%1.',
            lvlRestart: lvlRestartNode ? parseInt(lvlRestartNode.getAttribute('w:val'), 10) : null,
        };
    }


    // ==========================================
    // MULTILEVEL LABEL FORMATTING
    // formatLabel replaces ALL %1–%9 placeholders using counters and level defs
    // ==========================================
    function formatLabel(lvlText, counters, levelDefinitions) {
        let label = lvlText || '';
        for (let lvl = 0; lvl <= 8; lvl++) {
            const placeholder = `%${lvl + 1}`;
            if (!label.includes(placeholder)) continue;
            const counter = counters[lvl] || 0;
            const lvlDef = levelDefinitions ? levelDefinitions[lvl] : null;
            const fmt = lvlDef ? (lvlDef.numFmt || 'decimal') : 'decimal';
            const formatted = formatSingleNumber(counter, fmt);
            label = label.split(placeholder).join(formatted);
        }
        return label;
    }

    /**
     * Format a single number. Handles letter overflow (26→z, 27→aa, 28→ab).
     */
    function formatSingleNumber(num, numFmt) {
        switch (numFmt) {
            case 'decimal': return String(num);
            case 'decimalZero': return num < 10 ? '0' + num : String(num);
            case 'lowerLetter': return toLetter(num, false);
            case 'upperLetter': return toLetter(num, true);
            case 'lowerRoman': return toRoman(num).toLowerCase();
            case 'upperRoman': return toRoman(num);
            case 'ordinal': return num + '.';
            default: return String(num);
        }
    }

    /**
     * Convert number to letter(s): 1→a, 26→z, 27→aa, 28→ab, 53→ba, etc.
     */
    function toLetter(num, upper) {
        let result = '';
        let n = num;
        while (n > 0) {
            n--;
            result = String.fromCharCode((upper ? 65 : 97) + (n % 26)) + result;
            n = Math.floor(n / 26);
        }
        return result;
    }

    function toRoman(num) {
        if (num <= 0) return String(num);
        const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
        const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
        let result = '';
        for (let i = 0; i < vals.length; i++) {
            while (num >= vals[i]) { result += syms[i]; num -= vals[i]; }
        }
        return result;
    }

    // Keep old API name for backward compat
    function formatNumber(num, numFmt, lvlText, level) {
        const counters = { [level]: num };
        const levelDefs = { [level]: { numFmt: numFmt || 'decimal' } };
        return formatLabel(lvlText || `%${level+1}.`, counters, levelDefs);
    }


    // ==========================================
    // STYLES PARSING with basedOn chain support
    // ==========================================
    function parseStyles(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const styles = { _numFromStyle: {}, _basedOn: {} };
        const styleNodes = doc.getElementsByTagNameNS(ns, 'style');

        for (const s of styleNodes) {
            const id = s.getAttribute('w:styleId');
            const nameNode = s.getElementsByTagNameNS(ns, 'name')[0];
            if (id && nameNode) styles[id] = nameNode.getAttribute('w:val');

            // Track basedOn relationships
            const basedOnNode = s.getElementsByTagNameNS(ns, 'basedOn')[0];
            if (basedOnNode && id) {
                styles._basedOn[id] = basedOnNode.getAttribute('w:val');
            }

            // Direct numbering from style
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

    /**
     * Resolve basedOn chain: if a style doesn't have numFromStyle but its parent does, inherit
     */
    function resolveBasedOnNumbering(styles) {
        const resolved = new Set();
        function resolve(styleId) {
            if (resolved.has(styleId)) return;
            resolved.add(styleId);
            if (styles._numFromStyle[styleId]) return; // Already has direct numbering
            const parent = styles._basedOn[styleId];
            if (!parent) return;
            resolve(parent);
            if (styles._numFromStyle[parent]) {
                styles._numFromStyle[styleId] = { ...styles._numFromStyle[parent] };
            }
        }
        for (const styleId of Object.keys(styles._basedOn)) {
            resolve(styleId);
        }
    }


    // ==========================================
    // FOOTNOTES / ENDNOTES — throw on XML error
    // ==========================================
    function parseFootnotes(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const footnotes = [];
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

    function parseEndnotes(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const endnotes = [];
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

    function extractTextFromXml(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        let text = '';
        const tNodes = doc.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) text += t.textContent + ' ';
        return text.trim();
    }


    // ==========================================
    // MARKDOWN / PLAIN TEXT
    // ==========================================
    async function parseMarkdown(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const hm = line.match(/^(#{1,6})\s+(.*)$/);
            if (hm) {
                elements.push({ type: 'heading', index: index++, headingLevel: hm[1].length,
                    text: hm[2], style: `Heading${hm[1].length}`, runs: [{ text: hm[2] }],
                    lineNumber: i+1, isDirectQuote: false, quoteConfidence: 0, paraId: null });
            } else if (line.trim().length > 0) {
                elements.push({ type: 'paragraph', index: index++, text: line, style: 'Normal',
                    runs: [{ text: line }], lineNumber: i+1, isEmpty: false,
                    isDirectQuote: false, quoteConfidence: 0, paraId: null });
            }
        }
        return { type: 'markdown', name: fileName, elements, footnotes: [], endnotes: [],
            headers: [], footers: [], styles: {}, numbering: {}, htmlPreview: '',
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.filter(e => e.type === 'paragraph').length,
            tableCount: 0, headingCount: elements.filter(e => e.type === 'heading').length };
    }

    async function parsePlainText(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;
        for (let i = 0; i < lines.length; i++) {
            if (lines[i].trim().length > 0) {
                elements.push({ type: 'paragraph', index: index++, text: lines[i], style: 'Normal',
                    runs: [{ text: lines[i] }], lineNumber: i+1, isEmpty: false,
                    isDirectQuote: false, quoteConfidence: 0, paraId: null });
            }
        }
        return { type: 'txt', name: fileName, elements, footnotes: [], endnotes: [],
            headers: [], footers: [], styles: {}, numbering: {}, htmlPreview: '',
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.length, tableCount: 0, headingCount: 0 };
    }


    // ==========================================
    // HASHED IDS — uses w14:paraId or content hash (no absolute index)
    // ==========================================
    function assignHashedIds(docMap) {
        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            if (el.paraId) {
                const prefix = el.type === 'heading' ? 'h' : el.type === 'table' ? 't' : 'p';
                el.id = `${prefix}-${el.paraId}`;
                continue;
            }
            if (el.type === 'table' && el.tableId) { el.id = el.tableId; continue; }
            const prevText = i > 0 ? (docMap.elements[i-1].text || '').substring(0,30) : '';
            const nextText = i < docMap.elements.length-1 ? (docMap.elements[i+1].text || '').substring(0,30) : '';
            const ctx = `${el.type}|${el.style||''}|${el.text||''}|${prevText}|${nextText}`;
            el.id = hashId(el.type, ctx);
        }
    }

    function hashId(prefix, input) {
        let hash = 0x811c9dc5;
        for (let i = 0; i < input.length; i++) {
            hash ^= input.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        const hashStr = (hash >>> 0).toString(36).padStart(7, '0');
        const tp = prefix === 'heading' ? 'h' : prefix === 'table' ? 't' : 'p';
        return `${tp}-${hashStr}`;
    }

    // ==========================================
    // DIRECT QUOTE DETECTION
    // ==========================================
    function detectDirectQuotes(docMap) {
        for (const el of docMap.elements) {
            if (!el.text || el.type === 'heading') continue;
            const text = el.text.trim();
            let confidence = 0;
            if (/^\u201E[\s\S]*\u201C$/.test(text)) confidence = 0.95;
            else if (/^\u00AB[\s\S]*\u00BB$/.test(text)) confidence = 0.90;
            else if (el.runs && el.runs.length > 0 && el.runs.every(r => r.italic) &&
                text.length > 20 && (el.style && el.style.match(/quote|citat|blockquote/i) || text.length > 100))
                confidence = 0.80;
            else if (/^[\u2014\u2015]\s/.test(text)) confidence = 0.80;
            else if (el.style && el.style.match(/quote|citat|blockquote/i)) confidence = 0.90;
            else if (/^[\u0370-\u03FF\u1F00-\u1FFF\s,.;\u00B7'"]+$/.test(text)) confidence = 0.85;
            if (confidence > 0) { el.isDirectQuote = true; el.quoteConfidence = confidence; }
        }
    }

    function countWords(text) { return text.split(/\s+/).filter(w => w.length > 0).length; }

    // Public API
    return { parse, parseNumbering, hashId, formatNumber, formatLabel, toLetter, setTrackedChangesMode };
})();

if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentParser;
}
