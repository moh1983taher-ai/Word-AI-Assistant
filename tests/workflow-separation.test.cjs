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
    "const referenceOrderSelect =",
    "independent citation handler"
);

assert.match(citationHandler, /readCitationSources\(\)/);
assert.match(source, /function withCitationOperationTimeout\(/, "Word read stages must not leave citation analysis pending forever");
assert.match(citationHandler, /withCitationOperationTimeout\(\s*readCitationSources\(\)/);
assert.match(citationHandler, /وصل ردّ الذكاء الاصطناعي؛ جارٍ تدقيق الإحالات/, "the UI must expose completion of the AI request");
assert.match(citationHandler, /اجتاز التدقيق؛ جارٍ بناء مقترحات التوحيد/, "the UI must expose the post-AI suggestion stage");
assert.match(citationHandler, /withCitationOperationTimeout\(\s*buildFootnoteSuggestions\(/);

assert.match(citationHandler, /CitationWorkflow\.analyzeCitations\(/);
assert.match(citationHandler, /onProgress: function \(progress\)/, "citation analysis must report batch progress to the task pane");
assert.match(citationHandler, /تحليل الدفعة .* من /, "the UI must identify the active AI batch");
assert.match(citationHandler, /الحواشي المنجزة:/, "the progress panel must show completed and total footnotes");
assert.match(citationHandler, /دفعات التحليل:/, "the progress panel must show completed and total batches");
assert.match(citationHandler, /citation-progress-card[\s\S]*?role="status"/, "progress must be accessible to assistive technology");
assert.match(citationHandler, /role="progressbar"/, "the progress panel must expose a progress bar");
assert.match(citationHandler, /اكتمل تحليل التوثيقات وتجهيز المقترحات/, "the UI must display a completed state");
const citationService = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/citation-workflow.js"),
    "utf8"
);
assert.match(citationService, /sourceText/, "exact citation anchors must remain in the citation-only service");
assert.match(citationHandler, /buildUnifiedFootnoteMap\(/);
assert.match(citationHandler, /buildFootnoteSuggestions\(/);
assert.match(citationHandler, /applyFootnoteSuggestions\(/);
assert.match(source, /async function buildFootnoteSuggestions\([\s\S]*?onProgress/);
assert.match(source, /reportSuggestionProgress\(source, noteNumber\)/);
assert.match(source, /await addNotes\(\s*footnotes\.items/);
assert.match(source, /await addNotes\(\s*endnotes\.items/);
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
    /context\.document\.body\.load\(\s*["']text["']\s*\)|const\s+mainText\s*=|readReferenceSources|readBibliographyFromCurrentDocument/,
    "citation-only reader must not load the main document text or bibliography"
);

assert.match(source, /localStorage\.setItem\("BIBLIOGRAPHY_SETTINGS"/);
assert.match(source, /localStorage\.setItem\("CITATION_SETTINGS"/);
assert.match(source, /localStorage\.getItem\("REFERENCE_STYLE"/);
assert.match(source, /id="bibliography-format-select"/, "current-document bibliography format must have its own control");
assert.match(source, /id="project-reference-format-select"/, "project bibliography format must have its own control");
assert.match(source, /id="reference-format-select"/, "citation formatting must retain a separate control");
assert.match(source, /bibliographyFormatSelect\.onchange[\s\S]*?saveBibliographySettings\(\)/);
assert.match(source, /referenceFormatSelect\.onchange[\s\S]*?saveCitationSettings\(\)/);
assert.match(
    source,
    /CitationWorkflow\.analyzeCitations\([\s\S]*?getSavedSettings\(\)/,
    "citation service must use AI provider configuration without requiring the reference catalogue"
);

const citationButtonMarkup = source.indexOf('id="analyze-citations-btn"');
const citationButtonBinding = source.indexOf("const analyzeCitationsBtn =", citationButtonMarkup);
const referenceButtonBinding = source.indexOf("const analyzeReferencesBtn =", citationButtonMarkup);

assert.ok(citationButtonMarkup >= 0, "the independent citation button must be rendered");
assert.ok(
    citationButtonBinding > citationButtonMarkup &&
    citationButtonBinding < referenceButtonBinding,
    "citation click binding must occur immediately after its markup and before bibliography-button initialization"
);
assert.match(
    source.slice(citationButtonBinding, referenceButtonBinding),
    /referencesSourceWorkspace\.querySelector\("#analyze-citations-btn"\)/,
    "citation button lookup must be scoped to the workspace that just rendered it"
);

console.log("Independent bibliography/citation workflow separation tests passed.");
