/**
 * Quick CLI test - validates parser + rules work on test-document.md
 * Run: node test-cli.js
 */

const fs = require('fs');

// Mock browser globals that our modules expect
global.DOMParser = class {};
global.JSZip = {};
global.mammoth = {};
global.XLSX = { utils: { book_new() { return {}; }, aoa_to_sheet() { return {}; }, book_append_sheet() {} } };
global.document = { createElement: () => ({ textContent: '', innerHTML: '' }) };

// Load modules via eval (they use IIFE pattern that assigns to const)
// We need to make them available globally for Node
const parserCode = fs.readFileSync('./parser.js', 'utf8')
    .replace('const DocumentParser =', 'global.DocumentParser =');
const rulesCode = fs.readFileSync('./rules.js', 'utf8')
    .replace('const RuleEngine =', 'global.RuleEngine =');

eval(parserCode);
eval(rulesCode);

// Simulate parsed markdown document
const testFile = fs.readFileSync('./test-document.md', 'utf8');
const lines = testFile.split('\n');
const elements = [];
let index = 0;

for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const headingMatch = line.match(/^(#{1,6})\s+(.*)$/);
    if (headingMatch) {
        elements.push({
            type: 'heading', index: index++, headingLevel: headingMatch[1].length,
            text: headingMatch[2], style: `Heading${headingMatch[1].length}`,
            runs: [{ text: headingMatch[2] }], lineNumber: i + 1,
        });
    } else if (line.trim().length > 0) {
        elements.push({
            type: 'paragraph', index: index++, text: line, style: 'Normal',
            runs: [{ text: line }], lineNumber: i + 1, isEmpty: false,
        });
    }
}

const docMap = {
    type: 'markdown', name: 'test-document.md', elements,
    footnotes: [], endnotes: [], headers: [], footers: [],
    styles: {}, numbering: {}, htmlPreview: '', rawText: testFile,
    wordCount: testFile.split(/\s+/).length,
    paragraphCount: elements.filter(e => e.type === 'paragraph').length,
    tableCount: 0, headingCount: elements.filter(e => e.type === 'heading').length,
};

// Assign IDs
let pIdx = 0, hIdx = 0;
for (const el of docMap.elements) {
    if (el.type === 'heading') { el.id = `h-${String(hIdx++).padStart(4,'0')}`; }
    else { el.id = `p-${String(pIdx++).padStart(4,'0')}`; }
}

// Assign sections
let currentSection = '(početak)';
for (const el of docMap.elements) {
    if (el.type === 'heading') currentSection = el.text.trim();
    el.section = currentSection;
}

// Run all checks
const options = {
    brackets: true, quotes: true, markdown: false, // markdown is skipped for .md files
    spacing: true, scriptMix: true, greek: true, duplicates: true,
    toc: true, numbering: true, dashes: true, bibliography: true,
    urls: true, footnotes: true, repetition: true, capsWords: true, emptyHeadings: true,
};

const result = RuleEngine.runAudit(docMap, options);

console.log(`\n========== TEST RESULTS ==========`);
console.log(`Findings: ${result.findings.length}`);
console.log(`Passed:   ${result.passedChecks.length}`);
console.log(`\n--- FINDINGS BY CATEGORY ---`);

const byCat = {};
result.findings.forEach(f => { byCat[f.category] = (byCat[f.category] || 0) + 1; });
Object.entries(byCat).sort((a,b) => b[1]-a[1]).forEach(([cat, count]) => {
    console.log(`  ${cat}: ${count}`);
});

console.log(`\n--- SAMPLE FINDINGS ---`);
result.findings.slice(0, 10).forEach(f => {
    console.log(`  [${f.priority}] ${f.category}: "${f.original.substring(0,50)}"`);
    console.log(`    → ${f.replacement.substring(0,50)}`);
    console.log(`    (${f.rationale.substring(0,60)})`);
    console.log('');
});

console.log(`\n--- PASSED CHECKS ---`);
result.passedChecks.forEach(p => {
    console.log(`  ✓ ${p.area}`);
});

console.log(`\n========== DONE ==========\n`);
