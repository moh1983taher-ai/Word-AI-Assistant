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
