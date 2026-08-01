/**
 * Document Parser Module
 * Handles .docx (OOXML via JSZip), .md, and .txt files
 * Extracts structured document map with stable IDs
 */

const DocumentParser = (() => {
    'use strict';

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

        // Assign stable IDs
        assignStableIds(docMap);
        return docMap;
    }

    /**
     * Parse DOCX using JSZip + raw OOXML
     */
    async function parseDocx(arrayBuffer, fileName) {
        // Validate it's a valid ZIP
        let zip;
        try {
            zip = await JSZip.loadAsync(arrayBuffer);
        } catch (e) {
            throw new Error('Fajl nije validan DOCX (ZIP/OOXML paket).');
        }

        // Check for required OOXML structure
        const contentTypesFile = zip.file('[Content_Types].xml');
        if (!contentTypesFile) {
            throw new Error('Fajl nema validnu OOXML strukturu ([Content_Types].xml nedostaje).');
        }

        // Block macro-enabled documents
        const contentTypes = await contentTypesFile.async('text');
        if (contentTypes.includes('vbaProject') || contentTypes.includes('.docm')) {
            throw new Error('Makro-omogućeni dokumenti (.docm) nisu podržani iz bezbednosnih razloga.');
        }

        // Parse document.xml
        const docXml = zip.file('word/document.xml');
        if (!docXml) {
            throw new Error('word/document.xml nije pronađen u DOCX paketu.');
        }
        const docContent = await docXml.async('text');

        // Parse styles
        let styles = {};
        const stylesXml = zip.file('word/styles.xml');
        if (stylesXml) {
            const stylesContent = await stylesXml.async('text');
            styles = parseStyles(stylesContent);
        }

        // Parse numbering
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
        const elements = [];

        // Namespace handling
        const ns = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        };

        const body = doc.getElementsByTagNameNS(ns.w, 'body')[0];
        if (!body) return elements;

        let paragraphIndex = 0;
        let tableIndex = 0;

        for (const child of body.children) {
            const localName = child.localName;

            if (localName === 'p') {
                const el = parseParagraph(child, ns, styles, numbering, paragraphIndex);
                elements.push(el);
                paragraphIndex++;
            } else if (localName === 'tbl') {
                const el = parseTable(child, ns, tableIndex);
                elements.push(el);
                tableIndex++;
            } else if (localName === 'sectPr') {
                // Section break - record but don't add as element
            }
        }

        return elements;
    }

    /**
     * Parse a single paragraph element
     */
    function parseParagraph(pNode, ns, styles, numbering, index) {
        const runs = [];
        let fullText = '';

        // Get paragraph style
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

            // Check for numbering
            const numPr = pPr.getElementsByTagNameNS(ns.w, 'numPr')[0];
            if (numPr) {
                const ilvl = numPr.getElementsByTagNameNS(ns.w, 'ilvl')[0];
                const nId = numPr.getElementsByTagNameNS(ns.w, 'numId')[0];
                if (ilvl) numLevel = parseInt(ilvl.getAttribute('w:val'), 10);
                if (nId) numId = nId.getAttribute('w:val');
            }
        }

        // Determine if this is a heading
        const isHeading = styleName.match(/^Heading(\d+)$/) || styleName.match(/^Naslov(\d+)$/) || outlineLevel >= 0;
        let headingLevel = 0;
        if (isHeading) {
            const match = styleName.match(/(\d+)$/);
            headingLevel = match ? parseInt(match[1], 10) : (outlineLevel + 1);
        }

        // Extract runs (text segments with formatting)
        for (const child of pNode.children) {
            if (child.localName === 'r') {
                const runText = extractRunText(child, ns);
                const runProps = extractRunProps(child, ns);
                if (runText) {
                    runs.push({ text: runText, ...runProps });
                    fullText += runText;
                }
            } else if (child.localName === 'hyperlink') {
                // Extract text from hyperlink runs
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
            isEmpty: fullText.trim().length === 0,
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
                text += '\u25A1'; // placeholder for symbol
            }
        }
        return text;
    }

    function extractRunProps(rNode, ns) {
        const rPr = rNode.getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rPr'
        )[0];
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
     * Parse table element
     */
    function parseTable(tblNode, ns, tableIdx) {
        const rows = [];
        const trNodes = tblNode.getElementsByTagNameNS(ns.w, 'tr');

        for (const tr of trNodes) {
            const cells = [];
            const tcNodes = tr.getElementsByTagNameNS(ns.w, 'tc');
            for (const tc of tcNodes) {
                cells.push(extractTextFromElement(tc, ns));
            }
            rows.push(cells);
        }

        return {
            type: 'table',
            index: tableIdx,
            rows,
            text: rows.map(r => r.join(' | ')).join('\n'),
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
     * Parse footnotes.xml
     */
    function parseFootnotes(xmlStr) {
        const footnotes = [];
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
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
        for (const t of tNodes) {
            text += t.textContent;
        }
        return text;
    }

    function extractTextFromXml(xmlStr) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        let text = '';
        const tNodes = doc.getElementsByTagNameNS(ns, 't');
        for (const t of tNodes) {
            text += t.textContent + ' ';
        }
        return text.trim();
    }

    function parseStyles(xmlStr) {
        // Basic style extraction - map style IDs to names
        const styles = {};
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlStr, 'application/xml');
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

    function parseNumbering(xmlStr) {
        // Basic numbering extraction
        return {}; // Simplified - full implementation would track numId->abstractNumId mappings
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
                });
            }
        }

        return {
            type: 'markdown',
            name: fileName,
            elements,
            footnotes: [],
            endnotes: [],
            headers: [],
            footers: [],
            styles: {},
            numbering: {},
            htmlPreview: '',
            rawText: text,
            wordCount: countWords(text),
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
                });
            }
        }

        return {
            type: 'txt',
            name: fileName,
            elements,
            footnotes: [],
            endnotes: [],
            headers: [],
            footers: [],
            styles: {},
            numbering: {},
            htmlPreview: '',
            rawText: text,
            wordCount: countWords(text),
            paragraphCount: elements.length,
            tableCount: 0,
            headingCount: 0,
        };
    }

    /**
     * Assign stable IDs to all elements
     */
    function assignStableIds(docMap) {
        let pIdx = 0;
        let tIdx = 0;
        let hIdx = 0;

        for (const el of docMap.elements) {
            if (el.type === 'heading') {
                el.id = `h-${String(hIdx).padStart(4, '0')}`;
                hIdx++;
            } else if (el.type === 'paragraph') {
                el.id = `p-${String(pIdx).padStart(4, '0')}`;
                pIdx++;
            } else if (el.type === 'table') {
                el.id = `t-${String(tIdx).padStart(4, '0')}`;
                tIdx++;
            }
        }
    }

    function countWords(text) {
        return text.split(/\s+/).filter(w => w.length > 0).length;
    }

    // Public API
    return { parse };
})();
