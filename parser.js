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

    /**
     * Main parse entry point
     * @param {File} file
     * @returns {Promise<DocumentMap>}
     */
    async function parse(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        const arrayBuffer = await file.arrayBuffer();

        let docMap;
        switch (ext) {
            case 'docx':
                docMap = await parseDocx(arrayBuffer, file.name);
                break;
            case 'md':
                docMap = await parseMarkdown(arrayBuffer, file.name);
                break;
            case 'txt':
            case 'text':
                docMap = await parsePlainText(arrayBuffer, file.name);
                break;
            default:
                throw new Error(`Nepodržan format: .${ext}`);
        }

        // Assign stable hashed IDs
        assignHashedIds(docMap);
        // Detect direct quotes
        detectDirectQuotes(docMap);
        return docMap;
    }


    /**
     * Parse DOCX using JSZip + raw OOXML
     */
    async function parseDocx(arrayBuffer, fileName) {
        let zip;
        try {
            zip = await JSZip.loadAsync(arrayBuffer);
        } catch (e) {
            throw new Error('Fajl nije validan DOCX (ZIP/OOXML paket).');
        }

        // Enforce ZIP file count limit
        const fileCount = Object.keys(zip.files).length;
        if (fileCount > MAX_ZIP_FILES) {
            throw new Error(`ZIP sadrži ${fileCount} fajlova (limit: ${MAX_ZIP_FILES}).`);
        }

        // Check total uncompressed size
        let totalSize = 0;
        for (const [, f] of Object.entries(zip.files)) {
            if (f._data && f._data.uncompressedSize) {
                totalSize += f._data.uncompressedSize;
            }
        }
        if (totalSize > MAX_UNCOMPRESSED_SIZE) {
            throw new Error('Ukupna nekompresovana veličina prelazi limit.');
        }

        // Check for required OOXML structure
        const contentTypesFile = zip.file('[Content_Types].xml');
        if (!contentTypesFile) {
            throw new Error('Fajl nema validnu OOXML strukturu.');
        }

        // Block macro-enabled documents
        const contentTypes = await contentTypesFile.async('text');
        if (contentTypes.includes('vbaProject') || contentTypes.includes('.docm')) {
            throw new Error('Makro-omogućeni dokumenti nisu podržani.');
        }

        // Parse document.xml
        const docXml = zip.file('word/document.xml');
        if (!docXml) {
            throw new Error('word/document.xml nije pronađen u DOCX paketu.');
        }
        const docContent = await docXml.async('text');
        if (docContent.length > MAX_SINGLE_XML_SIZE) {
            throw new Error('document.xml prelazi dozvoljenu veličinu.');
        }


        // Parse styles
        let styles = {};
        const stylesXml = zip.file('word/styles.xml');
        if (stylesXml) {
            const stylesContent = await stylesXml.async('text');
            styles = parseStyles(stylesContent);
        }

        // Parse numbering (full implementation)
        let numbering = {};
        const numberingXml = zip.file('word/numbering.xml');
        if (numberingXml) {
            const numberingContent = await numberingXml.async('text');
            numbering = parseNumbering(numberingContent);
        }

        // Parse footnotes
        let footnotes = [];
        const footnotesXml = zip.file('word/footnotes.xml');
        if (footnotesXml) {
            const fnContent = await footnotesXml.async('text');
            footnotes = parseFootnotes(fnContent);
        }

        // Parse endnotes
        let endnotes = [];
        const endnotesXml = zip.file('word/endnotes.xml');
        if (endnotesXml) {
            const enContent = await endnotesXml.async('text');
            endnotes = parseEndnotes(enContent);
        }

        // Parse headers and footers
        let headers = [];
        let footers = [];
        for (const [path, file] of Object.entries(zip.files)) {
            if (path.match(/^word\/header\d+\.xml$/)) {
                const content = await file.async('text');
                headers.push({ path, text: extractTextFromXml(content) });
            }
            if (path.match(/^word\/footer\d+\.xml$/)) {
                const content = await file.async('text');
                footers.push({ path, text: extractTextFromXml(content) });
            }
        }

        // Build document structure from document.xml
        const elements = parseDocumentXml(docContent, styles, numbering);


        // Also get Mammoth HTML for visual reference
        let htmlPreview = '';
        try {
            const mammothResult = await mammoth.convertToHtml({ arrayBuffer });
            htmlPreview = mammothResult.value;
        } catch (e) {
            // Mammoth failure is non-critical
        }

        return {
            type: 'docx',
            name: fileName,
            elements,
            footnotes,
            endnotes,
            headers,
            footers,
            styles,
            numbering,
            htmlPreview,
            rawText: elements.map(el => el.text).join('\n'),
            wordCount: countWords(elements.map(el => el.text).join(' ')),
            paragraphCount: elements.filter(el => el.type === 'paragraph').length,
            tableCount: elements.filter(el => el.type === 'table').length,
            headingCount: elements.filter(el => el.type === 'heading').length,
        };
    }


    /**
     * Parse document.xml into structured elements
     */
    function parseDocumentXml(xmlStr, styles, numbering) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');

        // Check for XML parse errors
        if (doc.getElementsByTagName('parsererror').length) {
            throw new Error('Neispravan OOXML sadržaj.');
        }

        const elements = [];
        const ns = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        };

        const body = doc.getElementsByTagNameNS(ns.w, 'body')[0];
        if (!body) return elements;

        let paragraphIndex = 0;
        let tableIndex = 0;
        // Track list instances for numbering
        const listInstances = {};

        for (const child of body.children) {
            const localName = child.localName;

            if (localName === 'p') {
                const el = parseParagraph(child, ns, styles, numbering,
                    paragraphIndex, listInstances);
                elements.push(el);
                paragraphIndex++;
            } else if (localName === 'tbl') {
                const el = parseTable(child, ns, tableIndex);
                elements.push(el);
                tableIndex++;
            }
        }

        return elements;
    }


    /**
     * Parse a single paragraph element
     */
    function parseParagraph(pNode, ns, styles, numbering, index, listInstances) {
        const runs = [];
        let fullText = '';

        const pPr = pNode.getElementsByTagNameNS(ns.w, 'pPr')[0];
        let styleName = 'Normal';
        let outlineLevel = -1;
        let numId = null;
        let numLevel = null;

        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(ns.w, 'pStyle')[0];
            if (pStyle) {
                styleName = pStyle.getAttribute('w:val') || 'Normal';
            }

            const outlineLvl = pPr.getElementsByTagNameNS(ns.w, 'outlineLvl')[0];
            if (outlineLvl) {
                outlineLevel = parseInt(outlineLvl.getAttribute('w:val'), 10);
            }

            const numPr = pPr.getElementsByTagNameNS(ns.w, 'numPr')[0];
            if (numPr) {
                const ilvl = numPr.getElementsByTagNameNS(ns.w, 'ilvl')[0];
                const nId = numPr.getElementsByTagNameNS(ns.w, 'numId')[0];
                if (ilvl) numLevel = parseInt(ilvl.getAttribute('w:val'), 10);
                if (nId) numId = nId.getAttribute('w:val');
            }
        }

        // Determine heading
        const isHeading = styleName.match(/^Heading(\d+)$/) ||
            styleName.match(/^Naslov(\d+)$/) || outlineLevel >= 0;
        let headingLevel = 0;
        if (isHeading) {
            const match = styleName.match(/(\d+)$/);
            headingLevel = match ? parseInt(match[1], 10) : (outlineLevel + 1);
        }

        // Extract runs
        for (const child of pNode.children) {
            if (child.localName === 'r') {
                const runText = extractRunText(child, ns);
                const runProps = extractRunProps(child, ns);
                if (runText) {
                    runs.push({ text: runText, ...runProps });
                    fullText += runText;
                }
            } else if (child.localName === 'hyperlink') {
                for (const hChild of child.children) {
                    if (hChild.localName === 'r') {
                        const runText = extractRunText(hChild, ns);
                        if (runText) {
                            runs.push({ text: runText, isHyperlink: true });
                            fullText += runText;
                        }
                    }
                }
            }
        }


        // Calculate displayed number from OOXML numbering data
        let displayedNumber = null;
        let listInstanceId = null;
        let listStart = null;

        if (numId && numbering && numbering.nums && numbering.nums[numId]) {
            const numDef = numbering.nums[numId];
            const abstractId = numDef.abstractNumId;
            const abstractDef = numbering.abstractNums
                ? numbering.abstractNums[abstractId] : null;
            const level = numLevel || 0;

            if (abstractDef && abstractDef.levels && abstractDef.levels[level]) {
                const lvlDef = abstractDef.levels[level];
                listStart = lvlDef.start || 1;

                // Track list instance counters
                const instanceKey = `${numId}-${level}`;
                if (!listInstances[instanceKey]) {
                    listInstances[instanceKey] = listStart;
                } else {
                    listInstances[instanceKey]++;
                }
                displayedNumber = listInstances[instanceKey];
                listInstanceId = instanceKey;
            }
        }

        const type = isHeading ? 'heading' : 'paragraph';

        return {
            type,
            index,
            style: styleName,
            text: fullText,
            runs,
            headingLevel: isHeading ? headingLevel : null,
            numId,
            numLevel,
            displayedNumber,
            listInstanceId,
            listStart,
            isEmpty: fullText.trim().length === 0,
            isDirectQuote: false,
            quoteConfidence: 0,
        };
    }


    function extractRunText(rNode, ns) {
        let text = '';
        for (const child of rNode.children) {
            if (child.localName === 't') {
                text += child.textContent;
            } else if (child.localName === 'tab') {
                text += '\t';
            } else if (child.localName === 'br') {
                text += '\n';
            } else if (child.localName === 'sym') {
                text += '\u25A1';
            }
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
        if (lang) {
            props.lang = lang.getAttribute('w:val') || lang.getAttribute('w:bidi');
        }
        return props;
    }


    /**
     * Parse table element with full cell metadata
     */
    function parseTable(tblNode, ns, tableIdx) {
        const rows = [];
        const trNodes = tblNode.getElementsByTagNameNS(ns.w, 'tr');
        const tableId = `tbl-${tableIdx}`;

        let rowIndex = 0;
        for (const tr of trNodes) {
            const cells = [];
            const tcNodes = tr.getElementsByTagNameNS(ns.w, 'tc');
            const rowId = `${tableId}-r${rowIndex}`;
            let colIndex = 0;

            for (const tc of tcNodes) {
                const cellText = extractTextFromElement(tc, ns);
                const cellParagraphs = [];
                const pNodes = tc.getElementsByTagNameNS(ns.w, 'p');
                for (const p of pNodes) {
                    cellParagraphs.push(extractTextFromElement(p, ns));
                }

                cells.push({
                    text: cellText,
                    tableId,
                    rowId,
                    cellId: `${rowId}-c${colIndex}`,
                    rowIndex,
                    columnIndex: colIndex,
                    paragraphs: cellParagraphs,
                });
                colIndex++;
            }
            rows.push(cells);
            rowIndex++;
        }

        return {
            type: 'table',
            index: tableIdx,
            tableId,
            rows,
            text: rows.map(r => r.map(c =>
                typeof c === 'object' ? c.text : c).join(' | ')).join('\n'),
            hasHeader: rows.length > 0,
            columnCount: rows.length > 0 ? rows[0].length : 0,
            rowCount: rows.length,
        };
    }

    function extractTextFromElement(node, ns) {
        let text = '';
        const tNodes = node.getElementsByTagNameNS(ns.w, 't');
        for (const t of tNodes) {
            text += t.textContent;
        }
        return text;
    }


    /**
     * Full parseNumbering implementation
     * Maps numId → abstractNumId → levels (ilvl → start, numFmt, lvlText)
     */
    function parseNumbering(xmlStr) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');

        if (doc.getElementsByTagName('parsererror').length) {
            return {};
        }

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

                levels[ilvl] = {
                    start: startNode
                        ? parseInt(startNode.getAttribute('w:val'), 10) : 1,
                    numFmt: numFmtNode
                        ? numFmtNode.getAttribute('w:val') : 'decimal',
                    lvlText: lvlTextNode
                        ? lvlTextNode.getAttribute('w:val') : '%1.',
                };
            }

            result.abstractNums[abstractNumId] = { levels };
        }

        // Parse num elements (numId → abstractNumId mapping)
        const numNodes = doc.getElementsByTagNameNS(ns, 'num');
        for (const num of numNodes) {
            const numId = num.getAttribute('w:numId');
            const abstractNumIdRef =
                num.getElementsByTagNameNS(ns, 'abstractNumId')[0];
            if (abstractNumIdRef) {
                result.nums[numId] = {
                    abstractNumId: abstractNumIdRef.getAttribute('w:val'),
                };
            }
        }

        return result;
    }


    /**
     * Parse footnotes.xml
     */
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

    /**
     * Parse endnotes.xml
     */
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
        for (const t of tNodes) { text += t.textContent; }
        return text;
    }

    function extractTextFromXml(xmlStr) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        let text = '';
        const tNodes = doc.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) { text += t.textContent + ' '; }
        return text.trim();
    }


    function parseStyles(xmlStr) {
        const styles = {};
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) return {};

        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const styleNodes = doc.getElementsByTagNameNS(ns, 'style');

        for (const s of styleNodes) {
            const id = s.getAttribute('w:styleId');
            const nameNode = s.getElementsByTagNameNS(ns, 'name')[0];
            if (id && nameNode) {
                styles[id] = nameNode.getAttribute('w:val');
            }
        }
        return styles;
    }

    /**
     * Parse Markdown file
     */
    async function parseMarkdown(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const headingMatch = line.match(/^(#{1,6})\s+(.*)$/);

            if (headingMatch) {
                elements.push({
                    type: 'heading',
                    index: index++,
                    headingLevel: headingMatch[1].length,
                    text: headingMatch[2],
                    style: `Heading${headingMatch[1].length}`,
                    runs: [{ text: headingMatch[2] }],
                    lineNumber: i + 1,
                    isDirectQuote: false,
                    quoteConfidence: 0,
                });
            } else if (line.trim().length > 0) {
                elements.push({
                    type: 'paragraph',
                    index: index++,
                    text: line,
                    style: 'Normal',
                    runs: [{ text: line }],
                    lineNumber: i + 1,
                    isEmpty: false,
                    isDirectQuote: false,
                    quoteConfidence: 0,
                });
            }
        }

        return {
            type: 'markdown', name: fileName, elements,
            footnotes: [], endnotes: [], headers: [], footers: [],
            styles: {}, numbering: {}, htmlPreview: '',
            rawText: text, wordCount: countWords(text),
            paragraphCount: elements.filter(el => el.type === 'paragraph').length,
            tableCount: 0,
            headingCount: elements.filter(el => el.type === 'heading').length,
        };
    }


    /**
     * Parse plain text file
     */
    async function parsePlainText(arrayBuffer, fileName) {
        const text = new TextDecoder('utf-8').decode(arrayBuffer);
        const lines = text.split('\n');
        const elements = [];
        let index = 0;

        for (let i = 0; i < lines.length; i++) {
            if (lines[i].trim().length > 0) {
                elements.push({
                    type: 'paragraph',
                    index: index++,
                    text: lines[i],
                    style: 'Normal',
                    runs: [{ text: lines[i] }],
                    lineNumber: i + 1,
                    isEmpty: false,
                    isDirectQuote: false,
                    quoteConfidence: 0,
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
     * Assign hashed IDs based on text content, style, and context
     * Instead of sequential p-0001, uses content-based hashes for stability
     */
    function assignHashedIds(docMap) {
        for (let i = 0; i < docMap.elements.length; i++) {
            const el = docMap.elements[i];
            const prevText = i > 0 ? (docMap.elements[i - 1].text || '') : '';
            const contextStr = `${el.type}|${el.style || ''}|${el.text || ''}|${prevText.substring(0, 30)}|${i}`;
            el.id = hashId(el.type, contextStr);
        }
    }

    /**
     * Generate a short hash-based ID
     */
    function hashId(prefix, input) {
        // Simple FNV-1a 32-bit hash
        let hash = 0x811c9dc5;
        for (let i = 0; i < input.length; i++) {
            hash ^= input.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
        // Convert to base36 for short representation
        const hashStr = (hash >>> 0).toString(36).padStart(7, '0');
        const typePrefix = prefix === 'heading' ? 'h' :
            prefix === 'table' ? 't' : 'p';
        return `${typePrefix}-${hashStr}`;
    }

    /**
     * Detect direct quotes in elements
     * Marks paragraphs that are entirely quoted text
     */
    function detectDirectQuotes(docMap) {
        for (const el of docMap.elements) {
            if (!el.text || el.type === 'heading') continue;

            const text = el.text.trim();
            let confidence = 0;

            // Pattern 1: Entire paragraph in typographic quotes „..."
            if (text.match(/^\u201E.*\u201C$/)) {
                confidence = 0.95;
            }
            // Pattern 2: Entire paragraph in guillemets «...»
            else if (text.match(/^\u00AB.*\u00BB$/)) {
                confidence = 0.90;
            }
            // Pattern 3: Italic run covering entire paragraph
            else if (el.runs && el.runs.length > 0 &&
                el.runs.every(r => r.italic) && text.length > 20) {
                confidence = 0.75;
            }
            // Pattern 4: Paragraph starts with em-dash (dialogue)
            else if (text.match(/^\u2014\s/) || text.match(/^—\s/)) {
                confidence = 0.80;
            }
            // Pattern 5: Block indented style (often "Quote" or "Citat")
            else if (el.style && el.style.match(/quote|citat|blockquote/i)) {
                confidence = 0.90;
            }
            // Pattern 6: Greek text (entire paragraph)
            else if (text.match(/^[\u0370-\u03FF\u1F00-\u1FFF\s,.;·'"]+$/)) {
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
    return { parse };
})();

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocumentParser;
}
