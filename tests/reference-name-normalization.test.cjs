const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(
    path.join(__dirname, "../src/taskpane/taskpane.js"),
    "utf8"
);

const normalizeFunction = source.indexOf("async function canonicalizeReferenceNamesWithAI(");
const normalizeCall = source.indexOf("await canonicalizeReferenceNamesWithAI(");
const finalMerge = source.indexOf("mergeEquivalentReferences(", normalizeCall);

assert.notEqual(normalizeFunction, -1, "Document-wide AI name normalizer must exist");
assert.notEqual(normalizeCall, -1, "Final pipeline must invoke the name normalizer");
assert.ok(finalMerge > normalizeCall, "Name normalization must run before deterministic reference merging");

assert.match(source, /تكرار الصيغ في المستند كله/, "Normalization must use document-wide frequency");
assert.match(source, /اللقب أو الشهرة أو النسبة المعروفة/, "Author names should prefer the recognized scholarly surname");
assert.match(source, /خيار الاستقصاء مفعّل/, "Author investigation must be conditional on the user option");
assert.match(source, /occurrences وvariants ومواضع الصفحات والأجزاء/, "The prompt must preserve bibliography occurrences, variants and page metadata");

const referenceAnalyzerStart = source.indexOf("async function analyzeReferencesWithAI(");
const referenceAnalyzerEnd = source.indexOf("function parseUnifiedReferenceAIResult(", referenceAnalyzerStart);
assert.ok(referenceAnalyzerStart >= 0 && referenceAnalyzerEnd > referenceAnalyzerStart);
const referenceAnalyzerSource = source.slice(referenceAnalyzerStart, referenceAnalyzerEnd);
assert.doesNotMatch(
    referenceAnalyzerSource,
    /sourceText|sourceTextConfidence/,
    "Bibliography extraction must not ask AI to resolve exact citation replacement spans"
);
assert.match(source, /لا تنشئ سجلًا ولا تحذف سجلًا ولا تدمج السجلات/, "The normalization pass must not change reference count or identity");
assert.match(
    source,
    /const normalized = references\.map\(function \(reference\)/,
    "The original author spelling must be captured before canonicalization"
);
assert.match(
    source,
    /reference\.authorFull \|\| reference\.author/,
    "First-mention formatting must use the preserved full author when available"
);
assert.match(
    source,
    /FirstMentionCitation/,
    "The first-mention path must use the isolated source-preserving formatter"
);

console.log("Reference name normalization policy tests passed.");
