"use strict";

const assert = require("node:assert/strict");
const revision = require("../src/taskpane/document-revision");

let documentText = "هذا خطا واضح";
let collectionLoads = 0;
let itemLoads = 0;

function makeRange(textGetter) {
    return {
        load() {},
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

const selectedRange = makeRange(() => documentText);
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

global.Word = {
    run(callback) {
        const context = {
            document: {
                getSelection() { return selectedRange; },
                body: { getRange() { return bodyRange; } }
            },
            trackedObjects: { add() {} },
            async sync() {}
        };
        return callback(context);
    }
};

(async function () {
    const captured = await revision.capture();
    assert.equal(captured.targetType, "selection");
    assert.equal(captured.originalText, "هذا خطا واضح");

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
    assert.equal(result.applied, true);
    assert.equal(documentText, "هذا خطأ واضح");
    assert.ok(collectionLoads > 0, "search collection items should be explicitly loaded");
    assert.ok(itemLoads > 0, "search result text should be explicitly loaded");
    process.stdout.write("document-revision-officejs: collection loading and selected-text apply passed\n");
})().catch(error => {
    process.stderr.write(String(error && error.stack || error) + "\n");
    process.exitCode = 1;
});
