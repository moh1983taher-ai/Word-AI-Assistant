"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/taskpane.js"),
    "utf8"
).replace(/\r\n/g, "\n");

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
const suggestionLoop = sectionBetween(
    "async function addNotes(",
    "await addNotes(\n                footnotes.items",
    "asynchronous suggestion progress loop"
);
assert.match(suggestionLoop, /matches\.length === 0[\s\S]*?continue;/, "unmatched footnotes must be skipped without stopping the remaining scan");
assert.doesNotMatch(suggestionLoop, /matches\.length === 0[\s\S]*?return;/, "an unmatched footnote must not terminate the whole scan");
assert.match(suggestionLoop, /await new Promise\(function \(resolve\) \{ setTimeout\(resolve, 0\); \}\)/, "the UI must be allowed to paint progress during long suggestion scans");
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
assert.match(source, /id="project-reference-format-select"/, "project bibliography formatting must remain available");
assert.equal((source.match(/id="reference-format-select"/g) || []).length, 1, "the current-document workflow must expose one shared reference-format selector");
assert.equal((source.match(/id="reference-add-missing-authors"/g) || []).length, 1, "author discovery must be controlled by one shared checkbox");
assert.doesNotMatch(source, /id="bibliography-format-select"/, "the duplicated bibliography format selector must be removed");
assert.doesNotMatch(source, /id="citation-add-missing-authors"/, "the duplicated citation author-discovery checkbox must be removed");
assert.match(source, /data-reference-operation="bibliography"/, "the bibliography operation must have its own selector");
assert.match(source, /data-reference-operation="citations"/, "the citation operation must have its own selector");
assert.match(source, /id="reference-operation-dock"/, "current-document actions must live in the persistent action dock");
assert.match(source, /actionDock\.appendChild\(applyButton\)/, "the citation apply action must move into the persistent dock");
assert.match(source, /citation-first-mention/, "the citation-only first-mention behavior must remain available");
assert.match(source, /localStorage\.setItem\("REFERENCE_SHARED_SETTINGS"/, "shared settings must be persisted as one state");
assert.match(
    source,
    /CitationWorkflow\.analyzeCitations\([\s\S]*?getSavedSettings\(\)/,
    "citation service must use AI provider configuration without requiring the reference catalogue"
);

const bibliographyHandler = sectionBetween(
    "analyzeReferencesBtn.onclick =\n                            async function (e)",
    "const compareReferencesBtn =",
    "reference extraction handler"
);
assert.equal(
    (bibliographyHandler.match(/const operationShell/g) || []).length,
    0,
    "the bibliography analysis handler must reuse the outer action-shell binding without redeclaring it"
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


const sourceNavigation = sectionBetween(
    "function setReferencesHeader(",
    "const wordDocumentPicker =",
    "reference source navigation helpers"
);

assert.match(sourceNavigation, /referencesSourcePanel\.classList\.remove\("is-hidden"\)/);
assert.match(sourceNavigation, /referencesSourceWorkspace\.innerHTML = ""/);
assert.match(sourceNavigation, /referencesSourcePanel\.classList\.add\("is-hidden"\)/);
assert.match(sourceNavigation, /backReferencesSource\.hidden = true/);
assert.match(sourceNavigation, /backReferencesSource\.hidden = false/);
assert.match(sourceNavigation, /showReferenceSourceChooser\(\)/, "the header back arrow must return to the source chooser");
assert.match(sourceNavigation, /المستند المفتوح حاليًا في Word/);
assert.match(sourceNavigation, /اختر المعالجة التي تريد تنفيذها/);

const sourceSelectionHandler = sectionBetween(
    'referencesContent\n        .querySelectorAll(\n            ".references-source-option"',
    "const referenceOrderSelect =",
    "reference source selection handler"
);

assert.match(sourceSelectionHandler, /showReferenceSourceWorkspace\(source\)/, "choosing a source must update the header and replace the source chooser");
assert.doesNotMatch(sourceSelectionHandler, /data-reference-navigation="sources"/, "the old large in-content back link must be removed");

const currentDocumentScreen = sectionBetween(
    "// المستند الحالي",
    "const referenceOrderSelect =",
    "current document workspace"
);
assert.doesNotMatch(currentDocumentScreen, /class="references-selected-source"/, "the document identification card must not duplicate the header");
assert.doesNotMatch(currentDocumentScreen, /reference-source-navigation/, "the large text back-navigation row must not consume workspace space");

const referenceMarkup = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/taskpane.html"),
    "utf8"
);
assert.match(referenceMarkup, /id="back-references-source"/, "the compact back arrow must be in the header");
assert.match(referenceMarkup, /id="references-header-title"/);
assert.match(referenceMarkup, /id="references-header-subtitle"/);

const referenceStyles = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/taskpane.css"),
    "utf8"
);
assert.match(referenceStyles, /\.references-source-panel\.is-hidden\s*\{\s*display:\s*none\s*!important;/);
assert.match(referenceStyles, /\.references-back-icon\s*\{[\s\S]*?width:\s*30px/);
assert.doesNotMatch(referenceStyles, /\.reference-source-back\s*\{/);
assert.match(source, /referencesSourceWorkspace\.innerHTML = ""/, "return navigation must clear the previous workspace rather than stack screens");

const referenceWorkspaceSection = source;
assert.doesNotMatch(referenceWorkspaceSection, /reference-operation-option-description/, "operation descriptions must not consume permanent workspace space");
assert.match(referenceWorkspaceSection, /title="استخراج المراجع وتوحيدها ومقارنتها بالقائمة الموجودة"/, "bibliography guidance must remain available on hover");
assert.match(referenceWorkspaceSection, /title="تحليل الإحالات في الحواشي واقتراح توحيدها ومراجعتها"/, "citation guidance must remain available on hover");
assert.match(referenceWorkspaceSection, /<span class="reference-operation-option-title">قائمة المراجع<\/span>/);
assert.match(referenceWorkspaceSection, /<span class="reference-operation-option-title">توثيقات الحواشي<\/span>/);
assert.doesNotMatch(referenceWorkspaceSection, /استقصاء أسماء المؤلفين عند غيابها/, "the author checkbox label must be shortened");
assert.match(referenceWorkspaceSection, /استقصاء أسماء المؤلفين/);
assert.doesNotMatch(referenceWorkspaceSection, /class="reference-author-responsibility"/, "author responsibility must not be a permanent paragraph");
assert.match(referenceWorkspaceSection, /id="reference-author-notice"[\s\S]*?hidden/, "the author responsibility notice must be hidden until enabled");
assert.match(referenceWorkspaceSection, /id="reference-operation-hint"[\s\S]*?hidden/, "operation descriptions must use a transient hint");

const operationHintLogic = sectionBetween(
    "const operationHint =",
    "const bibliographyOption =",
    "temporary operation hint logic"
);
assert.match(operationHintLogic, /setTimeout/);
assert.match(operationHintLogic, /3600/);

const authorToggleLogic = sectionBetween(
    "referenceAddMissingAuthors.onchange = function ()",
    "if (citationFirstMention)",
    "temporary author responsibility notice"
);
assert.match(authorToggleLogic, /this\.checked/);
assert.match(authorToggleLogic, /تنبيه:/);
assert.match(authorToggleLogic, /4800/);

console.log("Compact reference header and workspace tests passed.");


const referenceProgressHelper = sectionBetween(
    "function renderReferenceAnalysisProgress(",
    "// قراءة مستندات المشروع المحددة",
    "shared reference analysis progress helper"
);
assert.match(referenceProgressHelper, /reference-analysis-progress-card/, "reference progress must use a dedicated card class");
assert.match(referenceProgressHelper, /role="progressbar"/, "reference progress must expose an accessible progress bar");
assert.match(referenceProgressHelper, /نسبة إنجاز تحليل المراجع/, "reference progress must identify bibliography analysis");
assert.match(referenceProgressHelper, /aria-valuenow/, "reference progress must expose the current value accessibly");

const projectBibliographyHandler = sectionBetween(
    "analyzeProjectReferencesBtn.onclick =",
    "// قراءة مستندات المشروع المحددة",
    "project bibliography analysis handler"
);
assert.match(projectBibliographyHandler, /updateProjectReferenceProgress/, "project bibliography analysis must report progress");
assert.match(projectBibliographyHandler, /totalProjectBatches/, "project progress must use actual batch totals");
assert.match(projectBibliographyHandler, /completedProjectBatches \+= 1/, "project progress must advance after completed batches");
assert.match(projectBibliographyHandler, /projectInput.documentName/, "project progress must identify the active document");
assert.match(projectBibliographyHandler, /percent: 100/, "project bibliography analysis must expose completion");

const currentBibliographyHandler = sectionBetween(
    "analyzeReferencesBtn.onclick =\n                            async function (e)",
    "const compareReferencesBtn =",
    "current-document bibliography analysis handler"
);
assert.match(currentBibliographyHandler, /updateReferenceProgress/, "current-document bibliography analysis must report progress");
assert.match(currentBibliographyHandler, /تحليل الدفعة/, "current-document progress must identify the active batch");
assert.match(currentBibliographyHandler, /completedReferenceBatches \+= 1/, "current-document progress must advance after batch attempts");
assert.match(currentBibliographyHandler, /percent: 100/, "current-document bibliography analysis must expose completion");
assert.match(currentBibliographyHandler, /state: "error"/, "current-document bibliography analysis must expose an error state");

const unifiedBibliographyTypography = referenceStyles.slice(
    referenceStyles.indexOf("/* Consistent typography for bibliography analysis"),
);
assert.ok(unifiedBibliographyTypography.length > 0, "the shared bibliography typography rules must exist");
assert.match(unifiedBibliographyTypography, /\.references-comparison-title[\s\S]*?font-size: 12px;/, "comparison status must use the common result font size");
assert.match(unifiedBibliographyTypography, /\.bibliography-build-title/, "bibliography update status must use the shared style");
assert.match(unifiedBibliographyTypography, /\.ai-reference-type[\s\S]*?font-size: 12px;/, "reference type text must use the common row size");
assert.match(unifiedBibliographyTypography, /\.project-reference-author[\s\S]*?font-size: 12px;/, "project reference rows must use the common text size");
