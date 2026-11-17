"use strict";

const assert = require("node:assert/strict");
const revision = require("../src/taskpane/document-revision");

let documentText = "هذا خطا واضح";
let selectionText = documentText;
let collectionLoads = 0;
let itemLoads = 0;

function makeRange(textGetter) {
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
            const index = source.indexOf(query);
            const matches = index < 0 ? [] : [makeMatch(query)];
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
        }
    };
}

function makeMatch(initialText) {
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
            documentText = documentText.replace(text, replacement);
            text = replacement;
            return {
                load() {},
                get text() { return text; },
                font: {}
            };
        }
    };
}

const selectedRange = makeRange(() => selectionText);
const bodyRange = makeRange(() => documentText);

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
                            const paragraph = {
                                text: "",
                                load(property) {
                                    assert.equal(property, "text");
                                    this.text = documentText;
                                },
                                getOoxml() {
                                    return {
                                        value: "<w:p><w:r><w:t>" + documentText + "</w:t></w:r></w:p>"
                                    };
                                }
                            };
                            let itemsLoaded = false;
                            return {
                                context,
                                load(property) {
                                    assert.equal(property, "items");
                                    itemsLoaded = true;
                                },
                                get items() {
                                    assert.equal(itemsLoaded, true, "paragraph items must be loaded before access");
                                    return [paragraph];
                                }
                            };
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
    }), captured.originalText);

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

    process.stdout.write("document-revision-officejs: collection loading, selection and document context tests passed\n");
})().catch(error => {
    process.stderr.write(String(error && error.stack || error) + "\n");
    process.exitCode = 1;
});
