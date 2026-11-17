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

// Word search-result resolution must work without the desktop-only Range.start/end API.
// Repeated Arabic words are mapped by their exact occurrence order within the captured range.
const repeatedArabicSource = "هذا بأن الأمر واضح، ثم وردت كلمة بأن في موضع آخر.";
const repeatedArabicMatches = [
    { text: "بأن" },
    { text: "بأن" }
];
assert.equal(
    revision.resolveSearchMatchByOccurrence(
        repeatedArabicMatches,
        repeatedArabicSource,
        "بأن",
        repeatedArabicSource.lastIndexOf("بأن")
    ),
    repeatedArabicMatches[1]
);
assert.equal(
    revision.resolveSearchMatchByOccurrence(
        repeatedArabicMatches,
        repeatedArabicSource,
        "بأن",
        repeatedArabicSource.indexOf("بأن")
    ),
    repeatedArabicMatches[0]
);

const repeatedWordSource = "الملاك في المثال الأول، والملاك في المثال الثاني.";
const repeatedWordMatches = [{ text: "الملاك" }, { text: "الملاك" }];
assert.equal(
    revision.resolveSearchMatchByOccurrence(
        repeatedWordMatches,
        repeatedWordSource,
        "الملاك",
        repeatedWordSource.lastIndexOf("الملاك")
    ),
    repeatedWordMatches[1]
);
assert.throws(
    () => revision.resolveSearchMatchByOccurrence(
        [{ text: "بأن" }],
        repeatedArabicSource,
        "بأن",
        repeatedArabicSource.lastIndexOf("بأن")
    ),
    /نتائج Word/
);
assert.throws(
    () => revision.resolveSearchMatchByOccurrence(
        repeatedArabicMatches,
        repeatedArabicSource,
        "بأن",
        repeatedArabicSource.indexOf("غير موجود")
    ),
    /موضع التعديل المحسوب/
);


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

const paragraphBoundaries = "الفقرة الأولى.\r\nالفقرة الثانية.\n\nالفقرة الثالثة.";
const paragraphChunks = revision.splitText(paragraphBoundaries, 1200);
assert.deepEqual(paragraphChunks, [paragraphBoundaries]);
assert.equal(paragraphChunks.join(""), paragraphBoundaries);

// Many short paragraphs are packed into a few bounded requests, not one request per paragraph.
const paragraphBatchSource = ("فقرة قصيرة للمراجعة اللغوية.\r\n").repeat(180);
const paragraphBatchChunks = revision.splitText(paragraphBatchSource, 1200);
assert.ok(paragraphBatchChunks.length > 1);
assert.ok(paragraphBatchChunks.length < 20);
assert.equal(paragraphBatchChunks.join(""), paragraphBatchSource);
assert.ok(paragraphBatchChunks.every(chunk => chunk.length <= 1200));

// Select larger batches for localized proofreading and smaller batches for full rewrites.
assert.equal(revision.getRecommendedChunkSize("صحح الأخطاء النحوية والإملائية"), 16000);
assert.equal(revision.getRecommendedChunkSize("أعد صياغة النص بأسلوب أكاديمي"), 8000);
assert.equal(revision.getRecommendedChunkSize("صحح النص وأعد صياغته"), 8000);
assert.equal(revision.getRecommendedChunkSize("راجع هذا المستند"), 12000);

const oversizedParagraph = "كلمة ".repeat(600);
const oversizedChunks = revision.splitText(oversizedParagraph, 1200);
assert.ok(oversizedChunks.length > 1);
assert.equal(oversizedChunks.join(""), oversizedParagraph);
assert.ok(oversizedChunks.every(chunk => !/[\r\n]/.test(chunk)));

const longDocument = ("فقرة أكاديمية تتناول الفكرة وتحافظ على تسلسلها.\r\n").repeat(180);
const chunks = revision.splitText(longDocument, 1200);
assert.ok(chunks.length > 1);
assert.equal(chunks.join(""), longDocument);

assert.equal(revision.hasSameLineBreakCount("فقرة أولى.\r\nفقرة ثانية.", "صياغة أولى.\nصياغة ثانية."), true);
assert.equal(revision.normalizeLineBreaks("فقرة أولى.\r\nفقرة ثانية.", "صياغة أولى.\nصياغة ثانية."), "صياغة أولى.\r\nصياغة ثانية.");
assert.equal(revision.hasSameLineBreakCount("فقرة واحدة", "فقرة أولى\nفقرة ثانية"), false);

// Regression: batched chunks can contain multiple paragraphs. Preserve all
// internal paragraph separators instead of flattening them as in single-paragraph mode.
const batchedParagraphSource = "فقرة أولى.\r\nفقرة ثانية.\n\nفقرة ثالثة.";
const batchedModelOutput = "فقرة أولى مصححة.\nفقرة ثانية مصححة.\n\nفقرة ثالثة.";
const normalizedBatch = revision.normalizeBatchLineBreaks(
    batchedParagraphSource,
    batchedModelOutput
);
assert.equal(
    normalizedBatch,
    "فقرة أولى مصححة.\r\nفقرة ثانية مصححة.\n\nفقرة ثالثة."
);
assert.equal(
    revision.hasSameLineBreakCount(batchedParagraphSource, normalizedBatch),
    true
);
assert.throws(
    () => revision.normalizeBatchLineBreaks(
        batchedParagraphSource,
        "فقرة واحدة مدمجة."
    ),
    /فواصل الفقرات الأصلية/
);

assert.equal(
    revision.normalizeChunkLineBreaks("فقرة أولى.\r\n", "فقرة أولى منقحة."),
    "فقرة أولى منقحة.\r\n"
);
assert.equal(
    revision.normalizeChunkLineBreaks("\r\nفقرة ثانية.", "فقرة ثانية منقحة."),
    "\r\nفقرة ثانية منقحة."
);
assert.equal(
    revision.normalizeChunkLineBreaks("أولى.\r\n", "أولى منقحة.\nسطر تابع"),
    "أولى منقحة. سطر تابع\r\n"
);
assert.equal(
    revision.normalizeChunkLineBreaks("فقرة أصلية.", "فقرة منقحة.\nسطر إضافي"),
    "فقرة منقحة. سطر إضافي"
);


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


const proofreadingSource = "تتطلب هذه المسألة لايمكن إغفالها في البحث، وهي مهمة.";
const proofreadingResponse = JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: "لايمكن",
        replacement: "لا يمكن",
        contextBefore: "تتطلب هذه المسألة ",
        contextAfter: " إغفالها",
        reason: "فصل أداة النفي عن الفعل"
    }]
});
const proofreadingResult = revision.parseResponse(proofreadingResponse, proofreadingSource);
assert.equal(proofreadingResult.text, "تتطلب هذه المسألة لا يمكن إغفالها في البحث، وهي مهمة.");
assert.equal(proofreadingResult.edits.length, 1);
assert.equal(proofreadingResult.edits[0].start, proofreadingSource.indexOf("لايمكن"));
assert.equal(proofreadingResult.edits[0].reason, "فصل أداة النفي عن الفعل");

const repeatedSource = "هذا خطأ لايمكن قبوله، وموضع آخر لايمكن إغفاله.";
const repeatedResult = revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: "لايمكن",
        replacement: "لا يمكن",
        contextBefore: "موضع آخر ",
        contextAfter: " إغفاله",
        reason: "تصحيح الفصل"
    }]
}), repeatedSource);
assert.equal(repeatedResult.text, "هذا خطأ لايمكن قبوله، وموضع آخر لا يمكن إغفاله.");
const uniqueWithImpreciseContext = revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: "أن المصلحة",
        replacement: "أن المصلحةَ",
        contextBefore: "سياق غير مطابق أرسله النموذج",
        contextAfter: "سياق لاحق غير مطابق",
        reason: "ضبط العبارة"
    }]
}), "وتبين الدراسة أن المصلحة العامة تقتضي ذلك.");
assert.equal(uniqueWithImpreciseContext.text, "وتبين الدراسة أن المصلحةَ العامة تقتضي ذلك.");
assert.equal(uniqueWithImpreciseContext.edits[0].start, "وتبين الدراسة أن المصلحة العامة تقتضي ذلك.".indexOf("أن المصلحة"));


// Regression: a phrase can occur more than once in a full-document batch.
// A slightly inaccurate long context must not abort the whole batch when the
// adjacent boundary words uniquely identify the intended occurrence.
const repeatedPhrase = "أن تعارف الناس أمراً ما";
const repeatedPhraseSource =
    "كان من عادة أهل القرية أن تعارف الناس أمراً ما في المجالس القديمة، ثم يتناقلونه. " +
    "وفي سياق آخر، سمعنا عبارة أن تعارف الناس أمراً ما بين الأصدقاء، فتنتشر سريعاً.";
const repeatedPhraseResult = revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: repeatedPhrase,
        replacement: "أن يتعارف الناس على أمر ما",
        contextBefore: "وذكر النص أنه كان من عادة أهل القرية",
        contextAfter: "في المجالس القديمة ثم يتداولونه",
        reason: "تصحيح التركيب النحوي"
    }]
}), repeatedPhraseSource);
const firstRepeatedPhraseStart = repeatedPhraseSource.indexOf(repeatedPhrase);
assert.equal(repeatedPhraseResult.edits.length, 1);
assert.equal(repeatedPhraseResult.edits[0].start, firstRepeatedPhraseStart);
assert.ok(repeatedPhraseResult.text.startsWith(
    "كان من عادة أهل القرية أن يتعارف الناس على أمر ما في المجالس القديمة"
));

// Arabic diacritics in the model's boundary context do not prevent locating
// the exact undiacritized source phrase.
const diacriticContextResult = revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: repeatedPhrase,
        replacement: "أن يتعارف الناس أمراً ما",
        contextBefore: "من عادة أهلُ القريةِ",
        contextAfter: "في المجالس القديمة ثم",
        reason: "تصحيح اختباري"
    }]
}), repeatedPhraseSource);
assert.equal(diacriticContextResult.edits[0].start, firstRepeatedPhraseStart);

// When two identical phrases also have indistinguishable adjacent context,
// the resolver must refuse rather than choose the first occurrence.
const indistinguishableSource =
    "وكان أهل القرية يرون أن تعارف الناس أمراً ما بينهم. " +
    "وكان أهل القرية يرون أن تعارف الناس أمراً ما بينهم.";
assert.throws(() => revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: repeatedPhrase,
        replacement: "أن يتعارف الناس أمراً ما",
        contextBefore: "وكان أهل القرية يرون",
        contextAfter: "بينهم",
        reason: "تصحيح اختباري"
    }]
}), indistinguishableSource), /وردت 2 مرات/);

assert.throws(() => revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: "لايمكن",
        replacement: "لا يمكن",
        contextBefore: "",
        contextAfter: "",
        reason: "تصحيح"
    }]
}), repeatedSource), /بصورة فريدة/);

assert.throws(() => revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: [{
        original: "عبارة\nأخرى",
        replacement: "عبارة أخرى",
        contextBefore: "",
        contextAfter: "",
        reason: "اختبار"
    }]
}), "عبارة\nأخرى"), /فاصل فقرة/);

assert.deepEqual(revision.parseResponse(JSON.stringify({
    status: "ready",
    mode: "correction",
    edits: []
}), "نص سليم."), {
    status: "ready",
    mode: "correction",
    text: "نص سليم.",
    edits: []
});

// Response parser regressions: tolerate harmless wrappers and trailing commas,
// but never fabricate a closing brace for a truncated model response.
const fencedResponse = "```json\n" + JSON.stringify({
    status: "ready",
    mode: "rewrite",
    text: "صياغة تتضمن قوسًا } داخل النص."
}) + "\n```";
assert.equal(
    revision.parseResponse(fencedResponse).text,
    "صياغة تتضمن قوسًا } داخل النص."
);

const wrappedResponse = "إليك النتيجة المطلوبة:\n" + JSON.stringify({
    status: "ready",
    mode: "rewrite",
    text: "نص منقح"
}) + "\nانتهى.";
assert.equal(revision.parseResponse(wrappedResponse).text, "نص منقح");

const trailingCommaResponse =
    '{"status":"ready","mode":"correction","edits":[{"original":"خطا","replacement":"خطأ","reason":"تصويب إملائي",}]}';
assert.equal(
    revision.parseResponse(trailingCommaResponse, "هذا خطا واضح.").text,
    "هذا خطأ واضح."
);

assert.throws(
    () => revision.parseResponse('{"status":"ready","mode":"rewrite","text":"استجابة مبتورة'),
    /تفاصيل: الاستجابة مبتورة/
);
assert.throws(() => revision.parseResponse("not json"), /تعذر قراءة استجابة/);

// Regression: model output may contain literal newlines, tabs, CRLF, or other
// U+0000–U+001F controls inside a quoted JSON value. Preserve them as text.
const rawControlCharResponse =
    '{"status":"ready","mode":"rewrite","text":"السطر الأول\nالسطر الثاني\tمع اقتباس \\"داخلي\\" وعودة\r\nسطر ثالث"}';
assert.equal(
    revision.parseResponse(rawControlCharResponse).text,
    'السطر الأول\nالسطر الثاني\tمع اقتباس "داخلي" وعودة\r\nسطر ثالث'
);

const rawReasonLineBreakResponse =
    '{"status":"ready","mode":"correction","edits":[{"original":"خطا","replacement":"خطأ","reason":"سبب أول\nسبب ثانٍ"}]}';
const rawReasonResult = revision.parseResponse(
    rawReasonLineBreakResponse,
    "هذا خطا واضح."
);
assert.equal(rawReasonResult.edits[0].reason, "سبب أول\nسبب ثانٍ");

const rawLowControlResponse =
    '{"status":"ready","mode":"rewrite","text":"قبل\u0001بعد"}';
assert.equal(
    revision.parseResponse(rawLowControlResponse).text,
    "قبل\u0001بعد"
);
console.log("Document revision tests passed.");
