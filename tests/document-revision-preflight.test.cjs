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
    { paragraphIndex: 0, start: 0, end: 4, text: "أولى." },
    { paragraphIndex: 2, start: 13, end: 18, text: "ثالثة." }
];
const orderedMask = preflight.maskChunk(twoProtectedSource, twoProtectedRanges, 0, "revision-order");
const reversedOutput = orderedMask.text
    .replace(orderedMask.segments[0].token, "__FIRST__")
    .replace(orderedMask.segments[1].token, orderedMask.segments[0].token)
    .replace("__FIRST__", orderedMask.segments[1].token);
assert.throws(
    () => preflight.restoreChunk(reversedOutput, orderedMask),
    /ترتيب الفقرات المحمية/
);

assert.throws(
    () => preflight.maskChunk(
        "فقرة طويلة مقسمة",
        [{ paragraphIndex: 4, start: 3, end: 12, text: "رة طويلة" }],
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

console.log("Document revision preflight tests passed.");
