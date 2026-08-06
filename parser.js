/**
 * Document Parser Module
 * DOCX/MD/TXT parser
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
        trackedChangesMode = options.trackedChanges ?? 'accept';
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

        // Track total decompressed size across ALL entries
        let totalDecompressed = 0;

        // Count ALL entries incrementally - reject as soon as cumulative limit exceeded.
        // Also check compression ratio to prevent ZIP bombs (max 200:1 ratio)
        const MAX_COMPRESSION_RATIO = 200;
        for (const [path, entry] of Object.entries(zip.files)) {
            if (entry.dir) continue;
            const compressedSize = entry._data && entry._data.compressedSize
                ? entry._data.compressedSize : 0;
            const declaredUncompressed = entry._data && entry._data.uncompressedSize
                ? entry._data.uncompressedSize : 0;

            // PRE-decompress check: if declared sizes available, check ratio before loading
            if (compressedSize > 0 && declaredUncompressed > 0) {
                if (declaredUncompressed / compressedSize > MAX_COMPRESSION_RATIO) {
                    throw new Error(`Sumnjiv kompresioni odnos za ${path} (${Math.round(declaredUncompressed/compressedSize)}:1 deklarisan). Moguć ZIP bomb.`);
                }
                if (totalDecompressed + declaredUncompressed > MAX_UNCOMPRESSED_SIZE) {
                    throw new Error(`Ukupna raspakovana veličina bi prešla ${MAX_UNCOMPRESSED_SIZE} bajtova.`);
                }
            }

            // Decompress and verify
            const bytes = await entry.async('uint8array');
            if (compressedSize > 0 && bytes.length / compressedSize > MAX_COMPRESSION_RATIO) {
                throw new Error(`Sumnjiv kompresioni odnos za ${path} (${Math.round(bytes.length/compressedSize)}:1 stvarni). Moguć ZIP bomb.`);
            }
            totalDecompressed += bytes.length;
            if (totalDecompressed > MAX_UNCOMPRESSED_SIZE) {
                throw new Error(`Ukupna raspakovana veličina prelazi ${MAX_UNCOMPRESSED_SIZE} bajtova.`);
            }
        }

        // loadEntry for XML: re-reads (already counted above), enforces per-file limit
        async function loadEntry(path) {
            const file = zip.file(path);
            if (!file) return null;
            const bytes = await file.async('uint8array');
            if (bytes.length > MAX_SINGLE_XML_SIZE) {
                throw new Error(`${path} prelazi ${MAX_SINGLE_XML_SIZE} bajtova.`);
            }
            // BOM detection for encoding
            let encoding = 'utf-8';
            if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) encoding = 'utf-16le';
            else if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) encoding = 'utf-16be';
            else if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) encoding = 'utf-8';
            return new TextDecoder(encoding).decode(bytes);
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
        // Only load headers/footers actually referenced via w:headerReference/w:footerReference in document.xml
        const relsContent = await loadEntry('word/_rels/document.xml.rels');
        const linkedParts = new Map(); // path -> 'header'|'footer'
        if (relsContent && docContent) {
            const docDom = parseXmlStrict(docContent, 'word/document.xml');
            const nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
            const refMap = new Map(); // rid -> 'header'|'footer'
            const allElements = docDom.getElementsByTagName('*');
            for (let i = 0; i < allElements.length; i++) {
                const el = allElements[i];
                if (el.localName === 'headerReference' || el.localName === 'footerReference') {
                    const rid = el.getAttributeNS(nsR, 'id') || el.getAttribute('r:id') || '';
                    if (rid) refMap.set(rid, el.localName === 'headerReference' ? 'header' : 'footer');
                }
            }
            if (refMap.size > 0) {
                const relsDom = parseXmlStrict(relsContent, 'word/_rels/document.xml.rels');
                const relNodes = relsDom.getElementsByTagName('Relationship');
                for (let i = 0; i < relNodes.length; i++) {
                    const rel = relNodes[i];
                    const id = rel.getAttribute('Id') || '';
                    const target = rel.getAttribute('Target') || '';
                    const targetMode = rel.getAttribute('TargetMode') || '';
                    if (targetMode === 'External') continue;
                    if (refMap.has(id)) {
                        const resolved = resolvePartTarget('word/document.xml', target);
                        if (resolved) linkedParts.set(resolved, refMap.get(id));
                    }
                }
            }
        }
        for (const [partPath, kind] of linkedParts) {
            const c = await loadEntry(partPath);
            if (c) {
                const text = extractTextFromXml(c, partPath, parseXmlStrict);
                if (kind === 'header') headers.push({ path: partPath, text });
                else footers.push({ path: partPath, text });
            }
        }

        // Resolve basedOn chain for numbering inheritance
        resolveBasedOnNumbering(styles);

        // Reset table ID occurrence counters for this document
        parseTable._seen = {};
        parseTable._seenRows = {};

        const elements = parseDocumentXml(docContent, styles, numbering, parseXmlStrict);

        // Convert headers/footers to checkable pseudo-elements
        const headerElements = headers.map((h, i) => ({
            type: 'header', index: i, text: h.text, style: 'Header',
            runs: [{ text: h.text }], id: `hdr-${i}`, section: '(zaglavlje)',
            isEmpty: !h.text.trim(), isDirectQuote: false, quoteConfidence: 0, paraId: null,
        }));
        const footerElements = footers.map((f, i) => ({
            type: 'footer', index: i, text: f.text, style: 'Footer',
            runs: [{ text: f.text }], id: `ftr-${i}`, section: '(podnožje)',
            isEmpty: !f.text.trim(), isDirectQuote: false, quoteConfidence: 0, paraId: null,
        }));

        // Track unsupported/partially-supported elements
        const processingCoverage = {
            supported: ['paragraphs', 'headings', 'tables', 'numbering', 'styles', 'tracked_changes'],
            partial: [],
            unsupported: [],
        };
        if (footnotes.length > 0) processingCoverage.partial.push('footnotes');
        if (endnotes.length > 0) processingCoverage.partial.push('endnotes');
        if (headers.length > 0) processingCoverage.partial.push('headers');
        if (footers.length > 0) processingCoverage.partial.push('footers');

        // Check for textboxes, shapes, equations, charts in document XML
        if (docContent.includes('w:txbxContent') || docContent.includes('wps:txbx')) {
            processingCoverage.partial.push('textboxes');
        }
        if (docContent.includes('mc:AlternateContent') || docContent.includes('w:drawing')) {
            processingCoverage.partial.push('drawings_shapes');
        }
        if (docContent.includes('m:oMath') || docContent.includes('m:oMathPara')) {
            processingCoverage.unsupported.push('equations');
        }
        if (docContent.includes('c:chart')) {
            processingCoverage.unsupported.push('charts');
        }
        if (docContent.includes('w:altChunk')) {
            processingCoverage.unsupported.push('altChunk');
        }
        if (docContent.includes('w:object') || docContent.includes('o:OLEObject')) {
            processingCoverage.unsupported.push('ole_objects');
        }

        // Detect merged cells and nested tables for coverage report
        const hasMergedCells = elements.some(el => el.type === 'table' && el.hasMergedCells);
        const hasNestedTables = elements.some(el => el.type === 'table' && el.hasNestedTables);
        if (hasMergedCells) processingCoverage.supported.push('merged_cells');
        if (hasNestedTables) processingCoverage.supported.push('nested_tables');

        const htmlPreview = ''; // Mammoth removed - preview not used in UI

        return {
            type: 'docx', name: fileName, elements, footnotes, endnotes,
            headers, footers, headerElements, footerElements,
            styles, numbering, htmlPreview, processingCoverage,
            rawText: [
                ...elements.map(el => el.text),
                ...footnotes.map(fn => fn.text),
                ...endnotes.map(en => en.text),
                ...headers.map(h => h.text),
                ...footers.map(f => f.text),
            ].join('\n'),
            wordCount: countWords([
                ...elements.map(el => el.text),
                ...footnotes.map(fn => fn.text),
                ...endnotes.map(en => en.text),
                ...headers.map(h => h.text),
                ...footers.map(f => f.text),
            ].join(' ')),
            paragraphCount: elements.filter(el => el.type === 'paragraph').length,
            tableCount: countAllTables(elements),
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

        function processBodyChildren(parent) {
            for (const child of parent.children) {
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
                } else if (ln === 'sdt') {
                    // Block-level structured document tag - recurse into sdtContent
                    const sdtContent = getDirectChild(child, ns.w, 'sdtContent');
                    if (sdtContent) processBodyChildren(sdtContent);
                } else if (ln === 'customXml' || ln === 'ins' || ln === 'del') {
                    // Block-level tracked changes and custom XML - recurse
                    if (ln === 'del' && trackedChangesMode !== 'show_deleted') continue;
                    processBodyChildren(child);
                } else if (ln === 'AlternateContent') {
                    const branch = selectAlternateContentBranch(child);
                    if (branch) processBodyChildren(branch);
                }
            }
        }
        processBodyChildren(body);
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
            if (pStyle) styleName = getWAttr(pStyle, 'val') || 'Normal';
            const outlineLvl = getDirectChild(pPr, ns.w, 'outlineLvl');
            if (outlineLvl) outlineLevel = parseInt(getWAttr(outlineLvl, 'val'), 10);
            const numPr = getDirectChild(pPr, ns.w, 'numPr');
            if (numPr) {
                const ilvl = getDirectChild(numPr, ns.w, 'ilvl');
                const nId = getDirectChild(numPr, ns.w, 'numId');
                if (ilvl) numLevel = parseInt(getWAttr(ilvl, 'val'), 10);
                if (nId) numId = getWAttr(nId, 'val');
            }
            // Inherit numbering from style (including basedOn chain)
            if (!numId && styles._numFromStyle && styles._numFromStyle[styleName]) {
                const inherited = styles._numFromStyle[styleName];
                numId = inherited.numId;
                numLevel = inherited.ilvl ?? 0;
            }
        }

        // Detect headings: inline outlineLvl, standard style names, OR outlineLvl from style definition
        let isHeading = styleName.match(/^Heading(\d+)$/) ||
            styleName.match(/^Naslov(\d+)$/) || outlineLevel >= 0;
        // Check style definition for outlineLvl if not already detected
        if (!isHeading && styles && styles._outlineLevels && styles._outlineLevels[styleName] != null) {
            outlineLevel = styles._outlineLevels[styleName];
            isHeading = true;
        }
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
        let lvlTextPattern = null;

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
                lvlTextPattern = lvlDef.lvlText || `%${level+1}.`;

                // Exclude bullet/none from numeric tracking
                if (numFmt === 'bullet' || numFmt === 'none') {
                    return buildParaResult(isHeading, index, paraId, styleName, fullText, runs,
                        headingLevel, numId, numLevel, null, null, null, null, numFmt, null);
                }

                const effectiveStart = lvlDef.start ?? 1;
                listStart = effectiveStart;

                // Determine list instance lifecycle
                const prevEl = prevElements.length > 0 ? prevElements[prevElements.length - 1] : null;
                const prevWasSameNum = prevEl && prevEl.numId === numId;
                const interrupted = !prevWasSameNum && prevEl != null;

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

                // lvlRestart: when a higher level (lower ilvl number) advances,
                // reset only levels whose lvlRestart value matches the advancing level
                if (prevWasSameNum && prevEl.numLevel != null && prevEl.numLevel < level) {
                    // A higher level advanced - check each lower level's lvlRestart
                    for (let l = level; l <= 9; l++) {
                        const lDef = abstractDef.levels[l];
                        if (!lDef) continue;
                        // lvlRestart specifies which level's advancement triggers reset
                        // lvlRestart=0 means "never restart" (OOXML spec)
                        // lvlRestart: 1-based (0=never, 1=when level 0 advances, 2=when level 1, ...)
                        const restartAt = lDef.lvlRestart;
                        if (restartAt === 0) continue; // Never restart
                        // Convert to 0-based for comparison with numLevel
                        const effectiveRestart = restartAt != null ? (restartAt - 1) : (l > 0 ? l - 1 : null);
                        if (effectiveRestart != null && prevEl.numLevel <= effectiveRestart) {
                            // Use effective start (from lvlOverride if present)
                            let effStart = lDef.start ?? 1;
                            if (numDef.lvlOverrides && numDef.lvlOverrides[l]) {
                                const ov = numDef.lvlOverrides[l];
                                if (ov.startOverride != null) effStart = ov.startOverride;
                            }
                            inst.counters[l] = effStart - 1;
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
                displayedLabel = formatLabel(lvlDef.lvlText || `%${level+1}.`, inst.counters, abstractDef.levels, numDef.lvlOverrides);
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
            listInstanceId, listStart, numFmt, lvlTextPattern);
    }

    function buildParaResult(isHeading, index, paraId, styleName, fullText, runs,
        headingLevel, numId, numLevel, displayedNumber, displayedLabel,
        listInstanceId, listStart, numFmt, lvlTextPattern) {
        return {
            type: isHeading ? 'heading' : 'paragraph',
            index, paraId, style: styleName, text: fullText, runs,
            headingLevel: isHeading ? headingLevel : null,
            numId, numLevel, displayedNumber, displayedLabel,
            listInstanceId, listStart, numFmt, lvlTextPattern,
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
                // Tracked change: insertion - always include
                text += extractVisibleText(child, ns);
            } else if (ln === 'del') {
                // Tracked change: deletion
                if (trackedChangesMode === 'show_deleted') {
                    text += extractVisibleText(child, ns);
                }
                // 'accept' and 'ignore_deleted' both skip deleted text
            } else if (ln === 'pPr' || ln === 'rPr' || ln === 'sectPr' || ln === 'bookmarkStart' || ln === 'bookmarkEnd') {
                // Skip non-text elements
            } else if (ln === 'AlternateContent') {
                const branch = selectAlternateContentBranch(child);
                if (branch) text += extractVisibleText(branch, ns);
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
            } else if (ln === 'AlternateContent') {
                const branch = selectAlternateContentBranch(child);
                if (branch) buildRuns(branch, ns, runs);
            }
        }
    }

    function resolvePartTarget(sourcePart, target) {
        if (!target) return null;
        const base = sourcePart.includes('/') ? sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1) : '';
        const combined = target.startsWith('/') ? target.slice(1) : base + target;
        const segments = [];
        for (const segment of combined.split('/')) {
            if (!segment || segment === '.') continue;
            if (segment === '..') { if (segments.length) segments.pop(); continue; }
            segments.push(segment);
        }
        return segments.join('/');
    }

    function getDirectChild(parent, nsUri, localName) {
        for (const child of parent.children) {
            if (child.localName === localName &&
                (!nsUri || child.namespaceURI === nsUri)) return child;
        }
        return null;
    }

    function getWAttr(node, name) {
        // Namespace-aware attribute reading: try NS first, then literal prefix
        const nsW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        return node.getAttributeNS(nsW, name) || node.getAttribute(`w:${name}`) || null;
    }

    // Supported namespaces for mc:Choice Requires resolution
    const SUPPORTED_NS = new Set([
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'http://schemas.openxmlformats.org/drawingml/2006/main',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    ]);

    function selectAlternateContentBranch(acNode) {
        // Select first mc:Choice whose Requires namespaces are all supported, else mc:Fallback
        for (const child of acNode.children) {
            if (child.localName === 'Choice') {
                const requires = child.getAttribute('Requires') || '';
                if (!requires) return child; // No requirements = always valid
                const prefixes = requires.trim().split(/\s+/);
                const allSupported = prefixes.every(prefix => {
                    const uri = child.lookupNamespaceURI(prefix);
                    return uri && SUPPORTED_NS.has(uri);
                });
                if (allSupported) return child;
            }
        }
        // No supported Choice found - use Fallback
        for (const child of acNode.children) {
            if (child.localName === 'Fallback') return child;
        }
        return null;
    }


    // ==========================================
    // TABLE PARSING - gridSpan, vMerge, recursive nested tables
    // ==========================================
    function parseTable(tblNode, ns, tableIdx) {
        const rows = [];
        const rowLogicalEnds = [];
        let hasActualHeader = false;
        let allCellTexts = [];
        let hasAnyNestedTable = false;

        for (const child of tblNode.children) {
            if (child.localName !== 'tr') continue;
            const trPr = getDirectChild(child, ns.w, 'trPr');
            if (trPr && getDirectChild(trPr, ns.w, 'tblHeader')) hasActualHeader = true;

            const cells = [];
            let logicalCol = 0;
            // gridBefore: skip columns at start of row
            let rowGridBefore = 0;
            let rowGridAfter = 0;
            if (trPr) {
                const gridBeforeNode = getDirectChild(trPr, ns.w, 'gridBefore');
                if (gridBeforeNode) {
                    rowGridBefore = parseInt(getWAttr(gridBeforeNode, 'val'), 10) || 0;
                    logicalCol = rowGridBefore;
                }
                const gridAfterNode = getDirectChild(trPr, ns.w, 'gridAfter');
                if (gridAfterNode) {
                    rowGridAfter = parseInt(getWAttr(gridAfterNode, 'val'), 10) || 0;
                }
            }
            for (const tcChild of child.children) {
                if (tcChild.localName !== 'tc') continue;

                // Read gridSpan (horizontal merge)
                const tcPr = getDirectChild(tcChild, ns.w, 'tcPr');
                let gridSpan = 1;
                let vMergeType = null;
                if (tcPr) {
                    const gsNode = getDirectChild(tcPr, ns.w, 'gridSpan');
                    if (gsNode) gridSpan = parseInt(getWAttr(gsNode, 'val'), 10) || 1;
                    const vmNode = getDirectChild(tcPr, ns.w, 'vMerge');
                    if (vmNode) vMergeType = getWAttr(vmNode, 'val') || 'continue';
                }

                // Extract cell content
                const cellParagraphs = [];
                let cellText = '';
                let cellDirectText = '';
                const cellNestedTables = [];
                let cellHasNested = false;
                function processCellContent(parent) {
                    for (const tcContent of parent.children) {
                        if (tcContent.localName === 'p') {
                            const pText = extractVisibleText(tcContent, ns);
                            cellParagraphs.push(pText);
                            cellText += pText + ' ';
                            cellDirectText += pText + ' ';
                        } else if (tcContent.localName === 'tbl') {
                            const nestedTable = parseTable(tcContent, ns, tableIdx * 100 + (hasAnyNestedTable ? 1 : 0));
                            cellHasNested = true;
                            hasAnyNestedTable = true;
                            cellNestedTables.push(nestedTable);
                            cellParagraphs.push(`[Tabela: ${nestedTable.text}]`);
                            cellText += nestedTable.text + ' ';
                        } else if (tcContent.localName === 'sdt') {
                            const sdtContent = getDirectChild(tcContent, ns.w, 'sdtContent');
                            if (sdtContent) processCellContent(sdtContent);
                        } else if (tcContent.localName === 'customXml' || tcContent.localName === 'ins') {
                            processCellContent(tcContent);
                        } else if (tcContent.localName === 'del') {
                            if (trackedChangesMode === 'show_deleted') processCellContent(tcContent);
                        } else if (tcContent.localName === 'AlternateContent') {
                            const branch = selectAlternateContentBranch(tcContent);
                            if (branch) processCellContent(branch);
                        }
                    }
                }
                processCellContent(tcChild);
                cellText = cellText.trim();
                cellDirectText = cellDirectText.trim();
                allCellTexts.push(cellText);

                cells.push({
                    text: cellText, directText: cellDirectText,
                    paragraphs: cellParagraphs,
                    rowIndex: rows.length, columnIndex: logicalCol,
                    gridSpan, vMerge: vMergeType,
                    hasNestedTable: cellHasNested,
                    nestedTables: cellNestedTables.length > 0 ? cellNestedTables : undefined,
                });
                logicalCol += gridSpan;
            }
            rowLogicalEnds.push(logicalCol + rowGridAfter);
            rows.push(cells);
        }

        // Generate hashed IDs
        const tableContent = allCellTexts.join('|');
        const rawTableHash = hashId('table', `table|${tableContent}`);
        if (!parseTable._seen) parseTable._seen = {};
        parseTable._seen[rawTableHash] = (parseTable._seen[rawTableHash] || 0) + 1;
        const tableOccurrence = parseTable._seen[rawTableHash];
        const tableId = tableOccurrence > 1
            ? hashId('table', `table|${tableContent}|#${tableOccurrence}`) : rawTableHash;

        if (!parseTable._seenRows) parseTable._seenRows = {};
        for (let ri = 0; ri < rows.length; ri++) {
            const rowContent = rows[ri].map(c => c.text).join('|');
            const rawRowHash = hashId('table', `${tableId}|row|${rowContent}`);
            const rowKey = `${tableId}|${rawRowHash}`;
            parseTable._seenRows[rowKey] = (parseTable._seenRows[rowKey] || 0) + 1;
            const rowOccurrence = parseTable._seenRows[rowKey];
            const rowId = rowOccurrence > 1
                ? hashId('table', `${tableId}|row|${rowContent}|#${rowOccurrence}`) : rawRowHash;

            for (let ci = 0; ci < rows[ri].length; ci++) {
                const cell = rows[ri][ci];
                cell.tableId = tableId;
                cell.rowId = rowId;
                cell.cellId = hashId('table', `${rowId}|cell|${cell.text}|${cell.columnIndex}`);
            }
        }

        // Link vMerge continuation cells to their restart cell
        // Compute maxCols from actual cell coordinates (accounts for gridBefore/gridAfter)
        let maxCols = 0;
        for (let ri = 0; ri < rows.length; ri++) {
            if (rowLogicalEnds[ri] > maxCols) maxCols = rowLogicalEnds[ri];
            for (const cell of rows[ri]) {
                const cellEnd = cell.columnIndex + (cell.gridSpan || 1);
                if (cellEnd > maxCols) maxCols = cellEnd;
            }
        }
        for (let ci = 0; ci < maxCols; ci++) {
            let restartRow = null;
            let restartCell = null;
            for (let ri = 0; ri < rows.length; ri++) {
                const cell = rows[ri].find(c => c.columnIndex === ci || (c.columnIndex <= ci && c.columnIndex + (c.gridSpan||1) > ci));
                if (!cell) continue;
                if (cell.vMerge === 'restart') {
                    restartRow = ri;
                    restartCell = cell;
                } else if (cell.vMerge === 'continue' && restartCell && !cell.vMergeOrigin) {
                    cell.vMergeOrigin = { rowIndex: restartRow, columnIndex: restartCell.columnIndex, tableId: restartCell.tableId, rowId: restartCell.rowId, cellId: restartCell.cellId };
                } else if (!cell.vMerge) {
                    // Ordinary cell breaks the vMerge chain
                    restartRow = null;
                    restartCell = null;
                }
            }
        }

        return {
            type: 'table', index: tableIdx, tableId, rows,
            text: rows.map(r => r.map(c => c.text).join(' | ')).join('\n'),
            hasHeader: hasActualHeader,
            columnCount: maxCols,
            rowCount: rows.length,
            hasNestedTables: hasAnyNestedTable,
            hasMergedCells: rows.some(r => r.some(c => c.gridSpan > 1 || c.vMerge)),
        };
    }


    // ==========================================
    // NUMBERING PARSING - full lvlOverride with w:lvl
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
            const abstractNumId = getWAttr(an, 'abstractNumId');
            const levels = {};
            const lvlNodes = an.getElementsByTagNameNS(ns, 'lvl');
            for (const lvl of lvlNodes) {
                levels[parseInt(getWAttr(lvl, 'ilvl'), 10)] = parseLvlNode(lvl, ns);
            }
            result.abstractNums[abstractNumId] = { levels };
        }

        const numNodes = doc.getElementsByTagNameNS(ns, 'num');
        for (const num of numNodes) {
            const numId = getWAttr(num, 'numId');
            const abstractNumIdRef = num.getElementsByTagNameNS(ns, 'abstractNumId')[0];
            const entry = {
                abstractNumId: abstractNumIdRef ? getWAttr(abstractNumIdRef, 'val') : null,
                lvlOverrides: {},
            };
            // Parse lvlOverride with full w:lvl support
            const ovNodes = num.getElementsByTagNameNS(ns, 'lvlOverride');
            for (const ov of ovNodes) {
                const ilvl = parseInt(getWAttr(ov, 'ilvl'), 10);
                const startOverrideNode = ov.getElementsByTagNameNS(ns, 'startOverride')[0];
                const lvlNode = ov.getElementsByTagNameNS(ns, 'lvl')[0];
                entry.lvlOverrides[ilvl] = {
                    startOverride: startOverrideNode
                        ? parseInt(getWAttr(startOverrideNode, 'val'), 10) : null,
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
            start: startNode ? parseInt(getWAttr(startNode, 'val'), 10) : null,
            numFmt: numFmtNode ? getWAttr(numFmtNode, 'val') : 'decimal',
            lvlText: lvlTextNode ? getWAttr(lvlTextNode, 'val') : '%1.',
            lvlRestart: lvlRestartNode ? parseInt(getWAttr(lvlRestartNode, 'val'), 10) : null,
        };
    }


    // ==========================================
    // MULTILEVEL LABEL FORMATTING
    // formatLabel replaces ALL %1-%9 placeholders using counters and level defs
    // ==========================================
    function formatLabel(lvlText, counters, levelDefinitions, lvlOverrides) {
        let label = lvlText || '';
        for (let lvl = 0; lvl <= 8; lvl++) {
            const placeholder = `%${lvl + 1}`;
            if (!label.includes(placeholder)) continue;
            // Use counter if set; otherwise use start value from level definition
            let counter = counters[lvl] || 0;
            if (counter === 0 && levelDefinitions && levelDefinitions[lvl]) {
                const lvlStart = lvlOverrides && lvlOverrides[lvl] && lvlOverrides[lvl].startOverride != null
                    ? lvlOverrides[lvl].startOverride
                    : (levelDefinitions[lvl].start || 1);
                counter = lvlStart;
            }
            // Use lvlOverride definition if available, then base level definition
            let lvlDef = levelDefinitions ? levelDefinitions[lvl] : null;
            if (lvlOverrides && lvlOverrides[lvl] && lvlOverrides[lvl].lvlDef) {
                lvlDef = { ...lvlDef, ...lvlOverrides[lvl].lvlDef };
            }
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
        const styles = { _numFromStyle: {}, _basedOn: {}, _outlineLevels: {} };
        const styleNodes = doc.getElementsByTagNameNS(ns, 'style');

        for (const s of styleNodes) {
            const id = getWAttr(s, 'styleId');
            const nameNode = s.getElementsByTagNameNS(ns, 'name')[0];
            if (id && nameNode) styles[id] = getWAttr(nameNode, 'val');

            // Track outlineLvl from style definition
            const pPrStyle = s.getElementsByTagNameNS(ns, 'pPr')[0];
            if (pPrStyle && id) {
                const olvl = pPrStyle.getElementsByTagNameNS(ns, 'outlineLvl')[0];
                if (olvl) styles._outlineLevels[id] = parseInt(getWAttr(olvl, 'val'), 10);
            }

            // Track basedOn relationships
            const basedOnNode = s.getElementsByTagNameNS(ns, 'basedOn')[0];
            if (basedOnNode && id) {
                styles._basedOn[id] = getWAttr(basedOnNode, 'val');
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
                            numId: getWAttr(numIdNode, 'val'),
                            ilvl: ilvlNode ? parseInt(getWAttr(ilvlNode, 'val'), 10) : 0,
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
    // FOOTNOTES / ENDNOTES - throw on XML error
    // ==========================================
    function parseFootnotes(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const footnotes = [];
        const fnNodes = doc.getElementsByTagNameNS(ns, 'footnote');
        for (const fn of fnNodes) {
            const id = getWAttr(fn, 'id');
            const type = getWAttr(fn, 'type');
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
            const id = getWAttr(en, 'id');
            const type = getWAttr(en, 'type');
            if (type === 'separator' || type === 'continuationSeparator') continue;
            const text = extractTextFromXmlNode(en, ns);
            endnotes.push({ id, text, isEmpty: text.trim().length === 0 });
        }
        return endnotes;
    }

    function extractTextFromXmlNode(node, nsStr) {
        // Extract text per paragraph using extractVisibleText (handles tracked changes, tabs, etc.)
        const ns = { w: nsStr, w14: '' };
        const paragraphs = node.getElementsByTagNameNS(nsStr, 'p');
        if (paragraphs.length === 0) {
            // Fallback: just get all text content
            let text = '';
            const tNodes = node.getElementsByTagNameNS(nsStr, 't');
            for (const t of tNodes) text += t.textContent;
            return text;
        }
        const parts = [];
        for (const p of paragraphs) {
            parts.push(extractVisibleText(p, ns));
        }
        return parts.join('\n');
    }

    function extractTextFromXml(xmlStr, filename, parseXmlStrict) {
        const doc = parseXmlStrict(xmlStr, filename);
        const ns = { w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', w14: '' };
        // Extract text per paragraph, join with newline (no spurious spaces between runs)
        const paragraphs = doc.getElementsByTagNameNS(ns.w, 'p');
        const parts = [];
        for (const p of paragraphs) {
            parts.push(extractVisibleText(p, ns));
        }
        return parts.join('\n').trim();
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
            headers: [], footers: [], headerElements: [], footerElements: [],
            styles: {}, numbering: {}, htmlPreview: '',
            processingCoverage: { supported: ['paragraphs','headings'], partial: [], unsupported: [] },
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
            headers: [], footers: [], headerElements: [], footerElements: [],
            styles: {}, numbering: {}, htmlPreview: '',
            processingCoverage: { supported: ['paragraphs'], partial: [], unsupported: [] },
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.length, tableCount: 0, headingCount: 0 };
    }


    // ==========================================
    // HASHED IDS - uses w14:paraId or content hash (no absolute index)
    // ==========================================
    function assignHashedIds(docMap) {
        const seenIds = {};
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
            let id = hashId(el.type, ctx);
            // Dedup: if same ID already used, add occurrence counter
            seenIds[id] = (seenIds[id] || 0) + 1;
            if (seenIds[id] > 1) id = hashId(el.type, `${ctx}|#${seenIds[id]}`);
            el.id = id;
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

    function countAllTables(elements) {
        let count = 0;
        for (const el of elements) {
            if (el.type !== 'table') continue;
            count++;
            if (el.rows) {
                for (const row of el.rows) {
                    for (const cell of row) {
                        if (cell.nestedTables) count += cell.nestedTables.reduce((s, nt) => s + countAllTables([nt]), 0);
                    }
                }
            }
        }
        return count;
    }

    function countWords(text) { return text.split(/\s+/).filter(w => w.length > 0).length; }

    // Public API
    return { parse, parseNumbering, hashId, formatNumber, formatLabel, toLetter, setTrackedChangesMode };
})();

if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentParser;
}
