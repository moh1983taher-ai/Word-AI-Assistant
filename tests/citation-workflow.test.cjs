"use strict";

const assert = require("node:assert/strict");
const CitationWorkflow = require("../src/taskpane/citation-workflow.js");

const notes = [
    {
        id: "footnote-1",
        source: "footnote",
        noteNumber: 1,
        originalText: "اقتباس محفوظ؛ السيوطي، الأشباه والنظائر، ط1، ص98؛ تعليق لاحق محفوظ"
    },
    {
        id: "footnote-2",
        source: "footnote",
        noteNumber: 2,
        originalText: "السيوطي، الأشباه والنظائر، ط1، ص100"
    },
    {
        id: "endnote-1",
        source: "endnote",
        noteNumber: 1,
        originalText: "الكتاب، الغزالي، ص22"
    }
];

const citations = [
    {
        materialId: "footnote-1",
        sourceText: "السيوطي، الأشباه والنظائر، ط1، ص98",
        author: "السيوطي",
        authorFull: "جلال الدين عبد الرحمن بن أبي بكر السيوطي",
        title: "الأشباه والنظائر",
        type: "book",
        volume: "",
        page: "98",
        pageRange: "",
        confidence: 0.98
    },
    {
        materialId: "footnote-2",
        sourceText: "السيوطي، الأشباه والنظائر، ط1، ص100",
        author: "السيوطي",
        title: "الأشباه والنظائر",
        type: "book",
        page: "100",
        confidence: 0.96
    },
    {
        materialId: "endnote-1",
        sourceText: "الكتاب، الغزالي، ص22",
        author: "الغزالي",
        title: "الكتاب",
        type: "book",
        page: "22",
        confidence: 0.9
    },
    {
        materialId: "unknown-note",
        sourceText: "السيوطي، الأشباه والنظائر، ص1",
        author: "السيوطي",
        title: "الأشباه والنظائر"
    }
];

const validateSpan = (original, candidate) =>
    original.includes(candidate) ? candidate : "";

const grouped = CitationWorkflow.buildUnifiedCitationReferences(
    citations,
    notes,
    validateSpan
);

assert.equal(grouped.references.length, 2, "the same author/title must be grouped, while different works remain separate");
assert.equal(grouped.skippedCount, 1, "unknown material IDs must be discarded");

const syuti = grouped.references.find((item) => item.author === "السيوطي");
assert.ok(syuti);
assert.equal(syuti.occurrences.length, 2, "separate note occurrences must remain separate");
assert.equal(syuti.variants.length, 2);
assert.equal(syuti.authorFull, "جلال الدين عبد الرحمن بن أبي بكر السيوطي");
assert.equal(syuti.occurrences[0].sourceText, "السيوطي، الأشباه والنظائر، ط1، ص98");
assert.equal(syuti.occurrences[0].materialId, "footnote-1");
assert.equal(syuti.occurrences[1].page, "100");

const sameTitleDifferentAuthors = CitationWorkflow.buildUnifiedCitationReferences(
    [
        {
            materialId: "footnote-1",
            sourceText: "الأشباه، السيوطي، ص1",
            author: "السيوطي",
            title: "الأشباه",
            type: "book"
        },
        {
            materialId: "footnote-2",
            sourceText: "الأشباه، ابن نجيم، ص2",
            author: "ابن نجيم",
            title: "الأشباه",
            type: "book"
        }
    ],
    [
        { id: "footnote-1", source: "footnote", originalText: "الأشباه، السيوطي، ص1" },
        { id: "footnote-2", source: "footnote", originalText: "الأشباه، ابن نجيم، ص2" }
    ],
    validateSpan
);
assert.equal(sameTitleDifferentAuthors.references.length, 2, "different authors must never be merged only because titles match");


const similarButDistinctTitles = CitationWorkflow.buildUnifiedCitationReferences(
    [
        {
            materialId: "eye-title-1",
            sourceText: "الخليل بن أحمد، كتاب العين",
            author: "الخليل بن أحمد",
            title: "كتاب العين",
            type: "book"
        },
        {
            materialId: "eye-title-2",
            sourceText: "الخليل بن أحمد، العين",
            author: "الخليل بن أحمد",
            title: "العين",
            type: "book"
        }
    ],
    [
        { id: "eye-title-1", source: "footnote", noteNumber: 1, originalText: "الخليل بن أحمد، كتاب العين" },
        { id: "eye-title-2", source: "footnote", noteNumber: 2, originalText: "الخليل بن أحمد، العين" }
    ],
    validateSpan
);
assert.equal(
    similarButDistinctTitles.references.length,
    2,
    "كتاب العين and العين must remain distinct title identities"
);


const invalidSpan = CitationWorkflow.buildUnifiedCitationReferences(
    [{
        materialId: "footnote-1",
        sourceText: "نص غير موجود في الحاشية",
        author: "السيوطي",
        title: "الأشباه والنظائر"
    }],
    notes,
    validateSpan
);
assert.equal(invalidSpan.references.length, 0, "unverified AI-selected spans must fail closed");
assert.equal(invalidSpan.skippedCount, 1);

const batches = CitationWorkflow.createCitationBatches(notes, 120, 1);
assert.equal(batches.length, 3, "batching must be bounded by both note count and serialized size");

const prompt = CitationWorkflow.createCitationSystemPrompt();
assert.match(prompt, /sourceText/);
assert.match(prompt, /المنخول 454/);
assert.match(prompt, /كما نقل عنه الإسنوي/);
assert.ok(prompt.includes("المنثور، 2/362"));
assert.match(prompt, /كتاب العين/);
assert.match(prompt, /لا تنشئ قائمة مراجع/);
assert.doesNotMatch(prompt, /قائمة المراجع النهائية الموحدة/);
assert.match(prompt, /خيار استقصاء المؤلف غير مفعّل/);

const authorDiscoveryPrompt = CitationWorkflow.createCitationSystemPrompt({
    addMissingAuthors: true
});
assert.match(authorDiscoveryPrompt, /خيار استقصاء المؤلف مفعّل/);
assert.match(authorDiscoveryPrompt, /لا تخمّن النسبة/);

const fence = String.fromCharCode(96).repeat(3);
const parsed = CitationWorkflow.parseCitationJSON(
    fence + "json\n" + JSON.stringify({ citations: [{ materialId: "footnote-1", title: "الكتاب" }] }) + "\n" + fence
);
assert.equal(parsed.length, 1);

(async function testPollinationsProvider() {
    let request = null;
    const responseData = {
        choices: [{
            message: { content: JSON.stringify({ citations: [] }) },
            finish_reason: "stop"
        }]
    };

    const result = await CitationWorkflow.analyzeCitations(
        [{
            id: "footnote-pollinations",
            source: "footnote",
            noteNumber: 1,
            originalText: "السيوطي، الأشباه والنظائر، ص98"
        }],
        {
            provider: "pollinations",
            key: "test-pollinations-key",
            model: "openai/gpt-5.4-nano"
        },
        {
            fetch: async function (url, options) {
                request = { url: url, options: options };
                return { ok: true };
            },
            readJSON: async function () { return responseData; },
            getAPIError: function () { return "unexpected API error"; },
            extractOpenAIStyleAnswer: function (data) {
                return data.choices[0].message.content;
            }
        }
    );

    assert.ok(request, "Pollinations must be called");
    assert.equal(request.url, "https://gen.pollinations.ai/v1/chat/completions");
    assert.equal(request.options.headers.Authorization, "Bearer test-pollinations-key");
    assert.equal(JSON.parse(request.options.body).model, "openai/gpt-5.4-nano");
    assert.equal(result.citations.length, 0);

    let timeoutSignal = null;
    await assert.rejects(
        CitationWorkflow.analyzeCitations(
            [{
                id: "footnote-timeout",
                source: "footnote",
                noteNumber: 1,
                originalText: "السيوطي، الأشباه والنظائر، ص98"
            }],
            {
                provider: "pollinations",
                key: "test-pollinations-key",
                model: "openai/gpt-5.4-nano"
            },
            {
                fetch: function (_url, options) {
                    timeoutSignal = options.signal;
                    return new Promise(function () {
                        // Simulate a provider connection that never settles.
                    });
                },
                readJSON: async function () { return {}; },
                getAPIError: function () { return "unexpected API error"; },
                extractOpenAIStyleAnswer: function () { return "{}"; },
                requestTimeoutMs: 20
            }
        ),
        /تجاوز الاتصال بمزوّد Pollinations مهلة الانتظار/
    );
    assert.ok(timeoutSignal && timeoutSignal.aborted, "a stalled provider request must be aborted at the timeout");

    await assert.rejects(
        CitationWorkflow.analyzeCitations(
            [{
                id: "footnote-body-timeout",
                source: "footnote",
                noteNumber: 1,
                originalText: "الغزالي، المستصفى، ص22"
            }],
            {
                provider: "pollinations",
                key: "test-pollinations-key",
                model: "openai/gpt-5.4-nano"
            },
            {
                fetch: async function () { return { ok: true }; },
                readJSON: function () {
                    return new Promise(function () {
                        // Simulate an HTTP response whose body never finishes.
                    });
                },
                getAPIError: function () { return "unexpected API error"; },
                extractOpenAIStyleAnswer: function () { return "{}"; },
                requestTimeoutMs: 20
            }
        ),
        /تجاوز الاتصال بمزوّد Pollinations مهلة الانتظار/
    );

    const progressEvents = [];
    const thirteenNotes = Array.from({ length: 13 }, function (_, index) {
        return {
            id: "progress-note-" + (index + 1),
            source: "footnote",
            noteNumber: index + 1,
            originalText: "إحالة اختبارية رقم " + (index + 1)
        };
    });
    const progressResult = await CitationWorkflow.analyzeCitations(
        thirteenNotes,
        {
            provider: "pollinations",
            key: "test-pollinations-key",
            model: "openai/gpt-5.4-nano"
        },
        {
            fetch: async function () { return { ok: true }; },
            readJSON: async function () {
                return {
                    choices: [{
                        message: { content: JSON.stringify({ citations: [] }) },
                        finish_reason: "stop"
                    }]
                };
            },
            getAPIError: function () { return "unexpected API error"; },
            extractOpenAIStyleAnswer: function (data) {
                return data.choices[0].message.content;
            },
            onProgress: function (progress) {
                progressEvents.push(progress);
            }
        }
    );
    assert.equal(progressResult.batchCount, 2, "13 notes must be analyzed in two bounded batches");
    assert.equal(progressResult.noteCount, 13);
    assert.ok(progressEvents.some(function (item) {
        return item.stage === "batch-start" && item.currentBatch === 1 && item.totalBatches === 2;
    }), "the UI must be notified when each batch starts");
    assert.ok(progressEvents.some(function (item) {
        return item.stage === "batch-complete" && item.currentBatch === 1 && item.completedNotes === 12;
    }), "the UI must be notified when the first batch completes");
    assert.ok(progressEvents.some(function (item) {
        return item.stage === "batch-complete" && item.currentBatch === 2 && item.completedNotes === 13;
    }), "the UI must be notified when the final batch completes");
    assert.equal(progressEvents[progressEvents.length - 1].stage, "complete");

    console.log("Independent citation workflow tests passed, including timeouts and live batch progress.");
})().catch(function (error) {
    console.error(error);
    process.exitCode = 1;
});
