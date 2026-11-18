"use strict";

const assert = require("node:assert/strict");
const preflight = require("../src/taskpane/document-revision-preflight");

assert.equal(preflight.hasTrackedChanges("<w:p><w:ins><w:r/></w:ins></w:p>"), true);
assert.equal(preflight.hasTrackedChanges("<w:p><w:del><w:r/></w:del></w:p>"), true);
assert.equal(preflight.hasTrackedChanges("<x:p><x:moveFrom/></x:p>"), true);
assert.equal(preflight.hasTrackedChanges("<w:p><w:delText>نص محذوف</w:delText></w:p>"), false);
assert.equal(preflight.hasTrackedChanges("<w:p><w:r><w:t>نص مستقر</w:t></w:r></w:p>"), false);

const xml = [
    '<w:p w14:paraId="1"><w:r><w:t>فقرة أولى</w:t></w:r></w:p>',
    '<w:p w14:paraId="2"><w:ins><w:r><w:t>فقرة ثانية</w:t></w:r></w:ins></w:p>'
].join("");
assert.equal(preflight.extractParagraphXml(xml).length, 2);

const mappedParagraphText = "هذا نص مكرر.\r\nهذا نص مكرر.\r\nفقرة أخيرة.";
const paragraphMap = preflight.mapParagraphRanges(mappedParagraphText, [
    "هذا نص مكرر.\r",
    "هذا نص مكرر.",
    "فقرة أخيرة."
]);
assert.equal(paragraphMap.length, 3);
assert.deepEqual(
    preflight.findParagraphRange(paragraphMap, mappedParagraphText.indexOf("هذا نص مكرر.") + 2, 3),
    paragraphMap[0]
);
assert.deepEqual(
    preflight.findParagraphRange(paragraphMap, mappedParagraphText.lastIndexOf("هذا نص مكرر.") + 2, 3),
    paragraphMap[1]
);
assert.equal(
    preflight.findParagraphRange(paragraphMap, mappedParagraphText.indexOf("\r\n") - 1, 4),
    null
);
assert.equal(
    preflight.mapParagraphRanges("فقرة أولى.\r\n\r\nفقرة ثالثة.", [
        "فقرة أولى.",
        "",
        "فقرة ثالثة."
    ]).length,
    2
);

const source = "مقدمة سليمة.\r\nفقرة قديمة وفقرة مضافة.\nخاتمة سليمة.";
const paragraphRecords = [
    { index: 0, text: "مقدمة سليمة.\r", ooxml: "<w:p><w:r/></w:p>" },
    { index: 1, text: "فقرة قديمة وفقرة مضافة.\r", ooxml: "<w:p><w:ins><w:r/></w:ins></w:p>" },
    { index: 2, text: "خاتمة سليمة.", ooxml: "<w:p><w:r/></w:p>" }
];
const protectedRanges = preflight.buildProtectedParagraphRanges(source, paragraphRecords);
assert.deepEqual(protectedRanges, [{
    paragraphIndex: 1,
    start: source.indexOf("فقرة قديمة وفقرة مضافة."),
    end: source.indexOf("فقرة قديمة وفقرة مضافة.") + "فقرة قديمة وفقرة مضافة.".length,
    text: "فقرة قديمة وفقرة مضافة."
}]);
assert.deepEqual(
    preflight.buildProtectedParagraphRanges(source, [
        { index: 0, text: "مقدمة سليمة.", ooxml: "<w:p/>" }
    ]),
    []
);

// Native WordApi 1.6 tracked-change metadata is preferred when available;
// OOXML remains the compatibility fallback for older Word hosts.
assert.deepEqual(
    preflight.buildProtectedParagraphRanges(source, [
        { index: 0, text: "مقدمة سليمة.", trackedChangesCount: 0 },
        { index: 1, text: "فقرة قديمة وفقرة مضافة.", trackedChangesCount: 2 },
        { index: 2, text: "خاتمة سليمة.", trackedChangesCount: 0 }
    ]),
    protectedRanges
);
assert.deepEqual(
    preflight.buildProtectedParagraphRanges(source, [
        { index: 0, text: "مقدمة سليمة.", hasTrackedChanges: false },
        { index: 1, text: "فقرة قديمة وفقرة مضافة.", hasTrackedChanges: true },
        { index: 2, text: "خاتمة سليمة.", hasTrackedChanges: false }
    ]),
    protectedRanges
);
assert.throws(
    () => preflight.buildProtectedParagraphRanges(source, [
        { index: 0, text: "فقرة غير موجودة", ooxml: "<w:p><w:ins/></w:p>" }
    ]),
    /مطابقة بنية فقرات Word/
);

const chunk = source;
const maskState = preflight.maskChunk(chunk, protectedRanges, 0, "revision-123");
assert.notEqual(maskState.text, source);
assert.equal((maskState.text.match(/\r\n|\r|\n/g) || []).length, (source.match(/\r\n|\r|\n/g) || []).length);
assert.equal(maskState.segments.length, 1);
assert.equal(maskState.text.includes(maskState.segments[0].token), true);
assert.equal(
    preflight.restoreChunk(maskState.text, maskState),
    source
);

const offsetSource = "قبل خطأ.\r\nفقرة عليها تعديل.\r\nهذا خطأ هنا، ثم خطأ هناك.";
const offsetProtectedRange = [{
    paragraphIndex: 1,
    start: offsetSource.indexOf("فقرة عليها تعديل."),
    end: offsetSource.indexOf("فقرة عليها تعديل.") + "فقرة عليها تعديل.".length,
    text: "فقرة عليها تعديل."
}];
const offsetMask = preflight.maskChunk(offsetSource, offsetProtectedRange, 0, "revision-offset");
const modelVisibleTarget = offsetMask.text.lastIndexOf("خطأ");
const mappedTarget = preflight.mapMaskedRangeToSource(offsetMask, modelVisibleTarget, "خطأ".length);
assert.equal(offsetSource.slice(mappedTarget.start, mappedTarget.end), "خطأ");
assert.equal(
    mappedTarget.start,
    offsetSource.lastIndexOf("خطأ")
);
const markerStart = offsetMask.text.indexOf(offsetMask.segments[0].token);
assert.throws(
    () => preflight.mapMaskedRangeToSource(offsetMask, markerStart + 2, 5),
    /علامة حماية/
);
assert.equal(
    preflight.mapMaskedOffsetToSource(offsetMask, offsetMask.text.length),
    offsetSource.length
);

const revisedWithProtectedToken = maskState.text.replace("مقدمة سليمة.", "مقدمة مصححة.");
assert.equal(
    preflight.restoreChunk(revisedWithProtectedToken, maskState),
    "مقدمة مصححة.\r\nفقرة قديمة وفقرة مضافة.\nخاتمة سليمة."
);
assert.throws(
    () => preflight.restoreChunk(maskState.text.replace(maskState.segments[0].token, ""), maskState),
    /لم يحافظ النموذج/
);
assert.throws(
    () => preflight.restoreChunk(
        maskState.text.replace(maskState.segments[0].token, maskState.segments[0].token + maskState.segments[0].token),
        maskState
    ),
    /لم يحافظ النموذج/
);

const twoProtectedSource = "أولى.\r\nثانية.\r\nثالثة.";
const twoProtectedRanges = [
    { paragraphIndex: 0, start: 0, end: "أولى.".length, text: "أولى." },
    {
        paragraphIndex: 2,
        start: twoProtectedSource.indexOf("ثالثة."),
        end: twoProtectedSource.indexOf("ثالثة.") + "ثالثة.".length,
        text: "ثالثة."
    }
];
const orderedMask = preflight.maskChunk(twoProtectedSource, twoProtectedRanges, 0, "revision-order");
const reversedOutput = orderedMask.text
    .replace(orderedMask.segments[0].token, "__FIRST__")
    .replace(orderedMask.segments[1].token, orderedMask.segments[0].token)
    .replace("__FIRST__", orderedMask.segments[1].token);
assert.throws(
    () => preflight.restoreChunk(reversedOutput, orderedMask),
    /نقل النموذج|ترتيب الفقرات المحمية/
);

const movedWithinLine = orderedMask.text.replace(
    orderedMask.segments[0].token,
    "نص أُضيف قبل الفقرة " + orderedMask.segments[0].token
);
assert.throws(
    () => preflight.restoreChunk(movedWithinLine, orderedMask),
    /موضع الفقرة المحمية داخل سطرها/
);

const movedProtectedToken = orderedMask.text
    .replace(orderedMask.segments[0].token, "__TOKEN__")
    .replace(orderedMask.segments[1].token, orderedMask.segments[0].token)
    .replace("__TOKEN__", orderedMask.segments[1].token);
assert.throws(
    () => preflight.restoreChunk(movedProtectedToken, orderedMask),
    /نقل النموذج|ترتيب الفقرات المحمية/
);

assert.throws(
    () => preflight.maskChunk(
        "فقرة طويلة مقسمة",
        [{ paragraphIndex: 4, start: 3, end: 20, text: "نص محمي أطول من الدفعة" }],
        0,
        "revision-partial"
    ),
    /تقاطع حدّ الدفعة/
);

assert.doesNotThrow(() => preflight.assertNoProtectedEdits([
    { start: 0, original: "مقدمة", replacement: "مدخل" }
], maskState));
assert.throws(
    () => preflight.assertNoProtectedEdits([
        { start: protectedRanges[0].start, original: "فقرة قديمة", replacement: "فقرة حديثة" }
    ], maskState),
    /داخل فقرة لها تغييرات متعقبة/
);

const noteSource = "فقرة أصلية تنتهي بإحالة.\u0002";
const noteStart = noteSource.indexOf("\u0002");
const noteMask = preflight.maskChunk(noteSource, [{
    paragraphIndex: 0,
    start: noteStart,
    end: noteStart + 1,
    text: "\u0002",
    kind: "note-reference",
    atParagraphEnd: true
}], 0, "note-test");
assert.equal(noteMask.text.includes("\u0002"), false);
assert.equal(noteMask.segments.length, 1);
assert.equal(noteMask.segments[0].kind, "note-reference");
const noteToken = noteMask.segments[0].token;
assert.equal(
    preflight.restoreChunk("صياغة جديدة تنتهي بالإحالة." + noteToken, noteMask),
    "صياغة جديدة تنتهي بالإحالة.\u0002"
);
assert.throws(
    () => preflight.restoreChunk("صياغة جديدة" + noteToken + " ثم نص بعد الإحالة", noteMask),
    /نهاية فقرتها/
);

const inlineNoteSource = "الجملة الأولى.\u0002 ثم الجملة الثانية.";
const inlineNoteStart = inlineNoteSource.indexOf("\u0002");
const inlineNoteMask = preflight.maskChunk(inlineNoteSource, [{
    paragraphIndex: 0,
    start: inlineNoteStart,
    end: inlineNoteStart + 1,
    text: "\u0002",
    kind: "note-reference",
    atParagraphEnd: false
}], 0, "inline-note-test");
const inlineToken = inlineNoteMask.segments[0].token;
assert.equal(
    preflight.restoreChunk("صياغة أولى." + inlineToken + " ثم صياغة ثانية.", inlineNoteMask),
    "صياغة أولى.\u0002 ثم صياغة ثانية."
);

console.log("Document revision preflight tests passed.");
