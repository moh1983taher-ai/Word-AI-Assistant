"use strict";

const assert = require("node:assert/strict");
const revision = require("../src/taskpane/document-revision");

function applyDiff(source, operations) {
    let result = source;
    operations.slice().sort((a, b) => b.start - a.start).forEach(operation => {
        result = result.slice(0, operation.start) +
            operation.replacement +
            result.slice(operation.start + operation.length);
    });
    return result;
}

assert.equal(revision.isRequest("صحح الأخطاء اللغوية في النص المحدد"), true);
assert.equal(revision.isRequest("جع النص المحدد لغويا"), true);
assert.equal(revision.isRequest("راجع النص المحدد لغويًا"), true);
assert.equal(revision.isRequest("تصحيح لغوي"), true);
assert.equal(revision.isRequest("أعد صياغة هذه الفقرة بأسلوب أكاديمي"), true);
assert.equal(revision.isRequest("ما الفرق بين القياس والاستحسان؟"), false);

const samples = [
    ["هذا نص صحيح.", "هذا النص صحيح."],
    ["وتظهر هذه النتيجة في البحث", "وتتضح هذه النتيجة في البحث."],
    ["الأمة الإسلامية والعولمة", "الأمة الإسلامية وعصر العولمة"],
    ["فقرة أولى.\r\nفقرة ثانية.", "فقرة أولى، وفقرة ثانية."],
    ["لا تغيير هنا", "لا تغيير هنا"],
    ["النص القديم", "النص الجديد مع إضافة"]
];

for (const [original, revised] of samples) {
    const operations = revision.diffText(original, revised);
    assert.equal(applyDiff(original, operations), revised);
}

const longDocument = ("فقرة أكاديمية تتناول الفكرة وتحافظ على تسلسلها.\r\n").repeat(180);
const chunks = revision.splitText(longDocument, 1200);
assert.ok(chunks.length > 1);
assert.equal(chunks.join(""), longDocument);

assert.equal(revision.hasSameLineBreakCount("فقرة أولى.\r\nفقرة ثانية.", "صياغة أولى.\nصياغة ثانية."), true);
assert.equal(revision.normalizeLineBreaks("فقرة أولى.\r\nفقرة ثانية.", "صياغة أولى.\nصياغة ثانية."), "صياغة أولى.\r\nصياغة ثانية.");
assert.equal(revision.hasSameLineBreakCount("فقرة واحدة", "فقرة أولى\nفقرة ثانية"), false);

const paragraphSource = "فقرة أولى قديمة.\r\nفقرة ثانية قديمة.";
const paragraphRevision = "فقرة أولى منقحة.\nفقرة ثانية مصححة.";
const paragraphOps = revision.diffTextPreservingParagraphs(
    paragraphSource,
    revision.normalizeLineBreaks(paragraphSource, paragraphRevision)
);
assert.equal(applyDiff(paragraphSource, paragraphOps), "فقرة أولى منقحة.\r\nفقرة ثانية مصححة.");
for (const operation of paragraphOps) {
    assert.equal(/[\r\n]/.test(paragraphSource.slice(operation.start, operation.start + operation.length)), false);
    assert.equal(/[\r\n]/.test(operation.replacement), false);
}
assert.throws(
    () => revision.diffTextPreservingParagraphs("فقرة واحدة", "فقرة أولى\nفقرة ثانية"),
    /عدد فواصل الفقرات/
);

const replacement = revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "rewrite",
    text: "صياغة جديدة"
}));
assert.deepEqual(replacement, {
    status: "ready",
    mode: "rewrite",
    text: "صياغة جديدة"
});

assert.deepEqual(
    revision.parseResponse(JSON.stringify({
        status: "clarify",
        question: "هل تريد تصحيح الأخطاء أم إعادة صياغة الفقرة؟"
    })),
    {
        status: "clarify",
        question: "هل تريد تصحيح الأخطاء أم إعادة صياغة الفقرة؟"
    }
);

assert.throws(() => revision.parseResponse("not json"), /تعذر التحقق/);
console.log("Document revision tests passed.");
