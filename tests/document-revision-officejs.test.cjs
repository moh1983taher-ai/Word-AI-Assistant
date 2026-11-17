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
        getRange() { return this; },
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
        insertText(replacement) {
            const current = String(documentText);
            if (current.slice(absoluteStart, absoluteStart + text.length) !== text) {
                throw new Error("mock Word range no longer matches its source text");
            }
            documentText = current.slice(0, absoluteStart) + replacement +
                current.slice(absoluteStart + text.length);
            text = replacement;
            return {
                load() {},
                get text() { return text; },
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

    // The document-wide path must use the captured target's RequestContext too.
    selectionText = "";
    const documentCapture = await revision.capture();
    assert.equal(documentCapture.targetType, "document");
    assert.equal(documentCapture.originalText, documentText);
    const documentContext = bodyRange.context;

    const documentParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [{
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

    // Paragraph-aware document correction must target each identical phrase
    // within its own Word paragraph, not the first occurrence in the entire body.
    documentText = "هذا الأطراد الأول.\r\nوتكرر الأطراد في الفقرة الثانية.";
    selectionText = "";
    const multiParagraphCapture = await revision.capture();
    assert.equal(multiParagraphCapture.targetType, "document");

    const multiParagraphRequest = revision.buildRequest(
        "صحح المستند لغويًا ونحويًا فقط",
        multiParagraphCapture.originalText,
        { sourceOffset: 0 }
    );
    const ids = multiParagraphRequest.revision.paragraphRecords.map(item => item.paragraphId);
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
    }), multiParagraphCapture.originalText, multiParagraphRequest.revision);

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
