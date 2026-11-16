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

const providerLayer = sectionBetween(
    "function parseAIRetryDelay(value)",
    "// API Error",
    "shared AI provider request layer"
);

assert.match(providerLayer, /function parseAIRetryDelay\(/);
assert.match(providerLayer, /function getAIProviderRetryDelayMs\(/);
assert.match(providerLayer, /function createAIProviderRequestError\(/);
assert.match(providerLayer, /async function requestAIProviderJSON\(/);
assert.match(providerLayer, /retry-after/i, "Retry-After must be respected");
assert.match(providerLayer, /x-ratelimit-reset-tokens/i, "token reset header must be respected");
assert.match(providerLayer, /x-ratelimit-reset-requests/i, "request reset header must be respected");
assert.match(providerLayer, /response\.status === 429/, "429 responses must be retried");
assert.match(providerLayer, /maxAttempts = 3/, "transient requests must have a bounded retry count");
assert.match(providerLayer, /output tokens per minute/i, "Groq OTPM errors must be recognized");
assert.match(providerLayer, /requestedOutputMatch/, "the provider's requested output estimate must guide adaptive budgeting");
assert.match(providerLayer, /scaledBudget/, "the output budget must be reduced proportionally to the reported OTPM limit");
assert.match(providerLayer, /outputBudgetWithinLimit/, "OTPM failures after a budget reduction must still honor provider reset headers");
assert.match(providerLayer, /max_completion_tokens/, "Groq must use max_completion_tokens");
assert.match(providerLayer, /delete requestBody\.max_tokens/, "the deprecated max_tokens field must be removed for Groq");
assert.match(providerLayer, /input tokens per minute|TPM/i, "input token-rate limits must be distinguished from request-rate limits");
assert.match(providerLayer, /shouldSplit: outputLimitExceeded \|\| requestTooLarge/, "token/context limits must mark a batch for splitting");

const analyzer = sectionBetween(
    "async function analyzeReferencesWithAI(",
    "function analyzeReferenceBatchWithRecovery(",
    "reference extraction analyzer"
);
assert.match(analyzer, /requestAIProviderJSON\(/, "reference extraction must use the shared request layer");
assert.match(analyzer, /requestBody\.max_completion_tokens = 850/, "Groq reference extraction must use a bounded output budget");
assert.match(analyzer, /truncatedError\.shouldSplit = true/, "truncated provider responses must trigger batch recovery");
assert.match(analyzer, /emptyAnswerError\.shouldSplit = true/, "empty provider replies must be recoverable by splitting the batch");

const recovery = sectionBetween(
    "async function analyzeReferenceBatchWithRecovery(",
    "function normalizeOccurrence(",
    "shared reference batch recovery"
);
assert.match(recovery, /error\?\.shouldSplit === true/, "the recovery layer must distinguish size/truncation failures");
assert.match(recovery, /recover\(leftBatch/, "failed batches must be split recursively");
assert.match(recovery, /recover\(rightBatch/, "both split halves must be recovered");
assert.match(recovery, /materialIds:/, "failed source material IDs must be retained");
assert.match(recovery, /failed: failed/, "unrecovered sub-batches must remain visible to callers");

const normalizer = sectionBetween(
    "async function canonicalizeReferenceNamesWithAI(",
    "// AI REFERENCE MERGER",
    "document-wide name normalization"
);
assert.match(normalizer, /batchSize = provider === "groq" \? 6 : 24/, "Groq normalization batches must be bounded");
assert.match(normalizer, /relevantFrequencyForBatch/, "normalization must retain document-wide frequency context without repeating all entries in every request");
assert.match(normalizer, /requestBody\.max_completion_tokens = 700/, "Groq normalization must use its own conservative output budget");
assert.match(normalizer, /itemsByIndex\.size !== batch\.length/, "partial normalization replies must not be applied");
assert.match(normalizer, /Apply only after the complete batch has been validated/, "normalization changes must be atomic per batch");
assert.match(normalizer, /normalizationFailures/, "failed normalization groups must be retained and reported");
assert.match(normalizer, /error\?\.shouldSplit === true/, "normalization must split on output and request-size limits");

const projectHandler = sectionBetween(
    "analyzeProjectReferencesBtn.onclick =",
    "// =====================================================\n                // 5. التحقق من النتائج",
    "project reference analysis handler"
);
assert.match(projectHandler, /await analyzeReferenceBatchWithRecovery\(/, "project extraction must use shared batch recovery");
assert.match(projectHandler, /failedProjectBatches\.push/, "project batch failures must be retained");
const projectCompletion = sectionBetween(
    "const completedProgressHTML = updateProjectReferenceProgress({",
    "container.innerHTML =",
    "project final completion state"
);
assert.match(projectCompletion, /اكتمل تحليل مراجع المشروع جزئيًا/, "project partial results must not be labeled as complete");
assert.ok(projectCompletion.includes('state: failedProjectBatches.length ? "error" : "success"'), "project progress must expose partial failure state");
assert.match(source, /راجع الدفعات المتعثرة/, "the project action must expose partial completion");

const currentHandler = sectionBetween(
    "analyzeReferencesBtn.onclick =\n                            async function (e)",
    "const compareReferencesBtn =",
    "current-document bibliography handler"
);
assert.match(currentHandler, /await analyzeReferenceBatchWithRecovery\(/, "current-document extraction must use shared batch recovery");
assert.match(currentHandler, /normalizationFailures/, "current-document normalization failures must be tracked");
assert.match(currentHandler, /تعذر توحيد أسماء بعض المجموعات/, "normalization failure must be visible in the result");
assert.match(currentHandler, /اكتمل تحليل المراجع جزئيًا/, "partial extraction/normalization must not be labeled successful");
assert.match(currentHandler, /normalizationStatus/, "the diagnostic result must distinguish completed and partial normalization");

const citationWorkflow = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/citation-workflow.js"),
    "utf8"
);
assert.match(citationWorkflow, /analyzeCitations/, "the citation analysis service must remain separate");
assert.doesNotMatch(citationWorkflow, /analyzeReferenceBatchWithRecovery|canonicalizeReferenceNamesWithAI/);
assert.match(source, /CitationWorkflow\.analyzeCitations\(/, "the existing independent footnote workflow must remain connected");

console.log("AI provider resilience and reference recovery regression tests passed.");
