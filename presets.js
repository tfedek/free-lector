/**
 * Free Lector - Audit Presets
 * Defines basic and full check configurations.
 */

'use strict';

const BASIC_PRESET = {
    brackets: true, spacing: true, scriptMix: true, duplicates: true, quotes: true,
    emptyNotes: true, noteContentChecks: false, footnotes: true,
    toc: false, bibliography: false, capsWords: false, numbering: false,
    greek: false, markdown: false, urls: false, emptyHeadings: false, repetition: false,
    headersFooters: false
};

const FULL_PRESET = {
    brackets: true, spacing: true, scriptMix: true, duplicates: true, quotes: true,
    emptyNotes: true, noteContentChecks: true, footnotes: true,
    toc: true, bibliography: true, capsWords: true, numbering: true,
    greek: true, markdown: true, urls: true, emptyHeadings: true, repetition: true,
    headersFooters: true
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { BASIC_PRESET, FULL_PRESET };
}
if (typeof window !== 'undefined') {
    window.BASIC_PRESET = BASIC_PRESET;
    window.FULL_PRESET = FULL_PRESET;
}
