"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/taskpane.js"),
    "utf8"
);

function sectionBetween(startToken, endToken, label) {
    const start = source.indexOf(startToken);
    assert.notEqual(start, -1, label + " start must exist");
    const end = source.indexOf(endToken, start + startToken.length);
    assert.ok(end > start, label + " end must exist after its start");
    return source.slice(start, end);
}

const referenceHandler = sectionBetween(
    "analyzeReferencesBtn.onclick =\n                            async function (e)",
    "const compareReferencesBtn =",
    "reference extraction handler"
);

assert.doesNotMatch(
    referenceHandler,
    /buildUnifiedFootnoteMap|buildFootnoteSuggestions|applyFootnoteSuggestions/,
    "reference extraction must not create or apply citation suggestions"
);

const citationHandler = sectionBetween(
    "const analyzeCitationsBtn =",
    "function sortUnifiedReferences(",
    "independent citation handler"
);

assert.match(citationHandler, /readCitationSources\(\)/);
assert.match(citationHandler, /CitationWorkflow\.analyzeCitations\(/);
assert.match(citationHandler, /buildUnifiedFootnoteMap\(/);
assert.match(citationHandler, /buildFootnoteSuggestions\(/);
assert.match(citationHandler, /applyFootnoteSuggestions\(/);
assert.doesNotMatch(
    citationHandler,
    /readReferenceSources\(|processReferenceSources\(|analyzeReferencesWithAI\(|latestUnifiedReferences/,
    "citation workflow must not depend on the bibliography extraction pipeline"
);

const citationReader = sectionBetween(
    "async function readCitationSources()",
    "// =====================================================\n// قراءة قائمة المراجع من المستند",
    "citation-only Word reader"
);

assert.match(citationReader, /body\.footnotes/);
assert.match(citationReader, /body\.endnotes/);
assert.doesNotMatch(
    citationReader,
    /body\.load\(\s*["']text["']\s*\)|readReferenceSources|readBibliographyFromCurrentDocument/,
    "citation-only reader must not load the main document text or bibliography"
);

assert.match(source, /localStorage\.setItem\("BIBLIOGRAPHY_SETTINGS"/);
assert.match(source, /localStorage\.setItem\("CITATION_SETTINGS"/);
assert.match(source, /localStorage\.getItem\("REFERENCE_STYLE"/);
assert.match(
    source,
    /CitationWorkflow\.analyzeCitations\([\s\S]*?getSavedSettings\(\)/,
    "citation service must use AI provider configuration without requiring the reference catalogue"
);

console.log("Independent bibliography/citation workflow separation tests passed.");
