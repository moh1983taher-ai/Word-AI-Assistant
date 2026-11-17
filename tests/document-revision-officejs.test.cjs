"use strict";

const assert = require("node:assert/strict");
const revision = require("../src/taskpane/document-revision");

let documentText = "هذا خطا واضح";
let selectionText = documentText;
let collectionLoads = 0;
let itemLoads = 0;

function makeParagraphCollection(textGetter, baseOffsetGetter, context) {
    const source = String(textGetter());
    const items = [];
    let start = 0;

    while (start < source.length) {
        let end = start;
        while (end < source.length && source[end] !== "\\r" && source[end] !== "\\n") end++;

        const localStart = start;
        const localEnd = end;
        const paragraph = makeRange(
            () => String(textGetter()).slice(localStart, localEnd),
            () => Number(baseOffsetGetter ? baseOffsetGetter() : 0) + localStart
        );
        paragraph.context = context;
        paragraph.getOoxml = () => ({
            value: "<w:p><w:r><w:t>" + String(textGetter()).slice(localStart, localEnd) + "</w:t></w:r></w:p>"
        });
        items.push(paragraph);

        if (end >= source.length) break;
        const separatorLength = source[end] === "\\r" && source[end + 1] === "\\n" ? 2 : 1;
        start = end + separatorLength;
    }

    if (!items.length) {
        const empty = makeRange(() => "", () => Number(baseOffsetGetter ? baseOffsetGetter() : 0));
        empty.context = context;
        empty.getOoxml = () => ({ value: "<w:p/>" });
        items.push(empty);
    }

    let itemsLoaded = false;
    return {
        context: context,
        load(property) {
            assert.equal(property, "items");
            itemsLoaded = true;
        },
        get items() {
            assert.equal(itemsLoaded, true, "paragraph items must be loaded before access");
            return items;
        }
    };
}

function makeRange(textGetter, absoluteStartGetter) {
    return {
        context: null,
        text: "",
        load(property) {
            if (property === "text") this.text = String(textGetter());
        },
        track() {},
        untrack() {},
        getOoxml() { return { value: "<w:document/>" }; },
        getRange(location) {
            const baseOffset = Number(absoluteStartGetter ? absoluteStartGetter() : 0);
            const rangeText = String(textGetter());
            const absolutePosition = location === "End"
                ? baseOffset + rangeText.length
                : baseOffset;
            return {
                context: this.context,
                insertText(value) {
                    const insertedText = String(value || "");
                    const current = String(documentText);
                    documentText = current.slice(0, absolutePosition) + insertedText +
                        current.slice(absolutePosition);
                    return {
                        load() {},
                        get text() { return insertedText; },
                        font: {}
                    };
                }
            };
        },
        search(query) {
            const source = String(textGetter());
            const baseOffset = Number(absoluteStartGetter ? absoluteStartGetter() : 0);
            const matches = [];
            let cursor = 0;
            while ((cursor = source.indexOf(query, cursor)) !== -1) {
                matches.push(makeMatch(query, baseOffset + cursor));
                cursor += query.length;
            }
            let itemsLoaded = false;
            return {
                context: this.context,
                load(property) {
                    assert.equal(property, "items");
                    itemsLoaded = true;
                    collectionLoads++;
                },
                get items() {
                    assert.equal(itemsLoaded, true, "must load collection items before reading .items");
                    return matches;
                }
            };
        },
        get paragraphs() {
            return makeParagraphCollection(textGetter, absoluteStartGetter, this.context);
        }
    };
}

function makeMatch(initialText, absoluteStart) {
    let text = initialText;
    let textLoaded = false;
    return {
        load(property) {
            assert.equal(property, "text");
            textLoaded = true;
            itemLoads++;
        },
        get text() {
            assert.equal(textLoaded, true, "must load each match's text before reading it");
            return text;
        },
        insertText(replacement, location) {
            const current = String(documentText);
            if (current.slice(absoluteStart, absoluteStart + text.length) !== text) {
                throw new Error("mock Word range no longer matches its source text");
            }
            const insertedText = String(replacement || "");
            let insertedStart = absoluteStart;
            if (location === "Before") {
                documentText = current.slice(0, absoluteStart) + insertedText +
                    current.slice(absoluteStart);
            } else if (location === "After") {
                insertedStart = absoluteStart + text.length;
                documentText = current.slice(0, insertedStart) + insertedText +
                    current.slice(insertedStart);
            } else {
                documentText = current.slice(0, absoluteStart) + insertedText +
                    current.slice(absoluteStart + text.length);
            }
            return {
                load() {},
                get text() { return insertedText; },
                font: {}
            };
        }
    };
}

const selectedRange = makeRange(() => selectionText, () => 0);
const bodyRange = makeRange(() => documentText, () => 0);

global.Office = {
    context: {
        requirements: {
            isSetSupported(name, version) {
                return name === "WordApi" && version === "1.3";
            }
        }
    }
};

let requestContextCount = 0;
global.Word = {
    run(objectOrCallback, maybeCallback) {
        if (typeof objectOrCallback === "function") {
            const context = {
                id: ++requestContextCount,
                document: {
                    getSelection() {
                        selectedRange.context = context;
                        return selectedRange;
                    },
                    body: {
                        getRange() {
                            bodyRange.context = context;
                            return bodyRange;
                        },
                        get paragraphs() {
                            return makeParagraphCollection(() => documentText, () => 0, context);
                        }
                    }
                },
                trackedObjects: {
                    add(object) {
                        assert.equal(
                            object.context,
                            context,
                            "tracked object must belong to the active RequestContext"
                        );
                    }
                },
                async sync() {}
            };
            return objectOrCallback(context);
        }

        const objects = Array.isArray(objectOrCallback)
            ? objectOrCallback
            : [objectOrCallback];
        const context = objects[0] && objects[0].context;
        assert.ok(context, "Word.run(object, callback) must receive a proxy with its owner context");
        assert.ok(objects.every(object => object.context === context));
        assert.equal(typeof maybeCallback, "function");
        return maybeCallback(context);
    }
};

(async function () {
    assert.equal(selectedRange.text, "", "mock range text starts unloaded");
    const captured = await revision.capture();
    assert.equal(selectedRange.text, documentText, "capture must load selection text before reading it");
    assert.equal(captured.targetType, "selection");
    assert.equal(captured.originalText, "هذا خطا واضح");
    const captureContext = selectedRange.context;
    assert.ok(captureContext, "capture should attach the selected Range to its RequestContext");

    const selectedRequest = revision.buildRequest(
        "صحح النص لغويًا فقط",
        captured.originalText,
        { sourceOffset: 0 }
    );
    assert.equal(
        selectedRequest.revision.structuredCorrection,
        false,
        "selected-text corrections must retain the established legacy parsing path"
    );
    assert.doesNotMatch(
        selectedRequest.systemPrompt,
        /paragraphId/,
        "selected-text correction prompts must not require paragraph identifiers"
    );
    assert.doesNotMatch(
        selectedRequest.userPrompt,
        /paragraphId/,
        "selected-text user prompts must not contain a paragraph-identifier contract"
    );

    const parsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [{
            original: "خطا",
            replacement: "خطأ",
            contextBefore: "هذا ",
            contextAfter: " واضح",
            reason: "تصويب إملائي"
        }]
    }), captured.originalText, selectedRequest.revision);

    const result = await revision.apply(Object.assign({}, captured, parsed));
    assert.equal(
        selectedRange.context,
        captureContext,
        "apply must reuse the selected Range's owning RequestContext"
    );
    assert.equal(result.applied, true);
    assert.equal(documentText, "هذا خطأ واضح");
    assert.ok(collectionLoads > 0, "search collection items should be explicitly loaded");
    assert.ok(itemLoads > 0, "search result text should be explicitly loaded");
    // A selected single-paragraph rewrite must parse and apply through the
    // same captured Office.js range, without a paragraph-ID contract.
    documentText = "كانت الفكرة قديمة وغير واضحة.";
    selectionText = documentText;
    const rewriteCapture = await revision.capture({ scope: "selection" });
    const rewriteRequest = revision.buildRequest(
        "أعد صياغة النص المحدد بأسلوب أكاديمي",
        rewriteCapture.originalText,
        { sourceOffset: 0 }
    );
    assert.equal(rewriteRequest.revision.rewriteSegments.length, 1);
    assert.doesNotMatch(rewriteRequest.systemPrompt, /paragraphId/);
    const rewriteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: [{
            segmentId: rewriteRequest.revision.rewriteSegments[0].segmentId,
            text: "تتسم الفكرة بطابع تقليدي ولا تتضح دلالتها بما يكفي."
        }]
    }), rewriteCapture.originalText, rewriteRequest.revision);
    const rewriteResult = await revision.apply(Object.assign({}, rewriteCapture, rewriteParsed));
    assert.equal(rewriteResult.applied, true);
    assert.equal(documentText, "تتسم الفكرة بطابع تقليدي ولا تتضح دلالتها بما يكفي.");

    // Rewriting several paragraphs must preserve the original CRLF separators
    // and apply every rewritten paragraph to the full-document target.
    documentText = "الفقرة الأولى تحتاج إلى تحسين.\r\nالفقرة الثانية تحتاج إلى توضيح.";
    selectionText = "";
    const documentRewriteCapture = await revision.capture({ scope: "document" });
    const documentRewriteRequest = revision.buildRequest(
        "أعد صياغة المستند كاملًا بأسلوب أكاديمي",
        documentRewriteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    assert.equal(documentRewriteRequest.revision.rewriteSegments.length, 2);
    const documentRewriteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: documentRewriteRequest.revision.rewriteSegments.map((segment, index) => ({
            segmentId: segment.segmentId,
            text: index === 0
                ? "تحتاج الفقرة الأولى إلى تحرير أدق."
                : "تحتاج الفقرة الثانية إلى مزيد من البيان."
        }))
    }), documentRewriteCapture.originalText, documentRewriteRequest.revision);
    assert.equal(
        documentRewriteParsed.text,
        "تحتاج الفقرة الأولى إلى تحرير أدق.\r\nتحتاج الفقرة الثانية إلى مزيد من البيان."
    );
    const documentRewriteResult = await revision.apply(Object.assign(
        {},
        documentRewriteCapture,
        documentRewriteParsed
    ));
    assert.equal(documentRewriteResult.applied, true);
    assert.equal(
        documentText,
        "تحتاج الفقرة الأولى إلى تحرير أدق.\r\nتحتاج الفقرة الثانية إلى مزيد من البيان."
    );

    // Restore the base fixture for the existing document-correction tests.
    documentText = "هذا خطأ واضح";
    selectionText = "";

    // The document-wide path must use the captured target's RequestContext too.
    const documentCapture = await revision.capture();
    assert.equal(documentCapture.targetType, "document");
    assert.equal(documentCapture.originalText, documentText);
    const documentContext = bodyRange.context;
    const documentRequest = revision.buildRequest(
        "صحح المستند لغويًا ونحويًا فقط",
        documentCapture.originalText,
        {
            sourceOffset: 0,
            paragraphBaseIndex: 0
        }
    );
    const documentParagraphId = documentRequest.revision.paragraphRecords[0].paragraphId;

    const documentParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [{
            paragraphId: documentParagraphId,
            original: "واضح",
            replacement: "بيّن",
            contextBefore: "هذا خطأ ",
            contextAfter: "",
            reason: "تصويب اختباري"
        }]
    }), documentCapture.originalText);

    const documentResult = await revision.apply(Object.assign({}, documentCapture, documentParsed));
    assert.equal(bodyRange.context, documentContext);
    assert.equal(documentResult.applied, true);
    assert.equal(documentText, "هذا خطأ بيّن");

    // Explicit whole-document intent overrides a non-empty lingering selection.
    documentText = "هذا الأطراد الأول.\r\nوتكرر الأطراد في الفقرة الثانية.";
    selectionText = "تحديد سابق لا ينبغي أن يقيّد مراجعة المستند كاملًا.";
    assert.equal(revision.getRequestedScope("صحح المستند لغويًا ونحويًا فقط"), "document");
    const multiParagraphCapture = await revision.capture({
        scope: revision.getRequestedScope("صحح المستند لغويًا ونحويًا فقط")
    });
    assert.equal(multiParagraphCapture.targetType, "document");
    assert.equal(multiParagraphCapture.originalText, documentText);
    assert.equal(selectionText, "تحديد سابق لا ينبغي أن يقيّد مراجعة المستند كاملًا.");

    const multiParagraphRequest = revision.buildRequest(
        "صحح المستند لغويًا ونحويًا فقط",
        multiParagraphCapture.originalText,
        {
            sourceOffset: 0,
            paragraphBaseIndex: revision.countParagraphsBefore(
                multiParagraphCapture.originalText,
                0
            )
        }
    );
    const ids = multiParagraphRequest.revision.paragraphRecords.map(item => item.paragraphId);
    assert.match(
        multiParagraphRequest.systemPrompt,
        /paragraphId/,
        "full-document correction prompts must retain the paragraph-identifier contract"
    );
    const multiParagraphParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [
            {
                paragraphId: ids[0],
                original: "الأطراد",
                replacement: "الاطراد",
                contextBefore: "هذا ",
                contextAfter: " الأول",
                reason: "تصويب إملائي"
            },
            {
                paragraphId: ids[1],
                original: "الأطراد",
                replacement: "الاطراد",
                contextBefore: "وتكرر ",
                contextAfter: " في الفقرة",
                reason: "تصويب إملائي"
            }
        ]
    }), multiParagraphCapture.originalText);

    assert.equal(multiParagraphParsed.edits.length, 2);
    assert.equal(multiParagraphParsed.rejected.length, 0);
    const multiParagraphResult = await revision.apply(Object.assign(
        {},
        multiParagraphCapture,
        multiParagraphParsed
    ));
    assert.equal(multiParagraphResult.applied, true);
    assert.equal(
        documentText,
        "هذا الاطراد الأول.\r\nوتكرر الاطراد في الفقرة الثانية."
    );
    assert.equal(documentText.match(/الاطراد/g).length, 2);

    const beforeUnsafeApply = documentText;
    await assert.rejects(
        revision.apply(Object.assign({}, documentCapture, {
            status: "ready",
            mode: "correction",
            text: "هذا نص مختلف",
            edits: [{
                start: 0,
                original: "هذا",
                replacement: "ذلك"
            }]
        })),
        /لا تفسر النص الناتج بالكامل/
    );
    assert.equal(documentText, beforeUnsafeApply, "invalid direct edits must not mutate the Word document");

    process.stdout.write("document-revision-officejs: collection loading, selection, paragraph targeting, and direct-edit integrity tests passed\n");
})().catch(error => {
    process.stderr.write(String(error && error.stack || error) + "\n");
    process.exitCode = 1;
});
