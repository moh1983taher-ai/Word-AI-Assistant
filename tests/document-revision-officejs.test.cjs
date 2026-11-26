"use strict";

const assert = require("node:assert/strict");
const revision = require("../src/taskpane/document-revision");

let documentText = "هذا خطا واضح";
let selectionText = documentText;
let collectionLoads = 0;
let itemLoads = 0;
let searchQueries = [];
let footnoteText = "حاشية قديمة فيها خطا.";
let endnoteText = "";
let documentTrackingMode = "Off";
let noteTrackedChanges = false;


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

function makeRangeSegment(startOffset, endOffset, context) {
    let textLoaded = false;
    return {
        context: context,
        startOffset: startOffset,
        endOffset: endOffset,
        load(property) {
            assert.equal(property, "text");
            textLoaded = true;
        },
        get text() {
            assert.equal(textLoaded, true, "range text must be loaded before access");
            return String(documentText).slice(startOffset, endOffset);
        },
        getRange(location) {
            const point = location === "End" ? endOffset : startOffset;
            return makeRangeSegment(point, point, context);
        },
        expandTo(other) {
            return makeRangeSegment(
                Math.min(startOffset, other.startOffset),
                Math.max(endOffset, other.endOffset),
                context
            );
        },
        insertText(value) {
            const insertedText = String(value || "");
            const current = String(documentText);
            documentText = current.slice(0, startOffset) + insertedText +
                current.slice(endOffset);
            return {
                load() {},
                get text() { return insertedText; },
                font: {}
            };
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
            return makeRangeSegment(absolutePosition, absolutePosition, this.context);
        },
        search(query) {
            searchQueries.push(String(query));
            const source = String(textGetter());
            const baseOffset = Number(absoluteStartGetter ? absoluteStartGetter() : 0);
            const matches = [];
            let cursor = 0;
            if (String(query).trim() &&
                String(query).length <= 255 &&
                !String(query).includes("\u0002") &&
                !/^[\s\p{P}\p{S}]+$/u.test(String(query))) {
                while ((cursor = source.indexOf(query, cursor)) !== -1) {
                    matches.push(makeMatch(query, baseOffset + cursor, this.context));
                    cursor += query.length;
                }
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

function makeMatch(initialText, absoluteStart, context) {
    let text = initialText;
    let textLoaded = false;
    return {
        context: context,
        load(property) {
            assert.equal(property, "text");
            textLoaded = true;
            itemLoads++;
        },
        get text() {
            assert.equal(textLoaded, true, "must load each match's text before reading it");
            return text;
        },
        getRange(location) {
            const point = location === "End" ? absoluteStart + text.length : absoluteStart;
            return makeRangeSegment(point, point, context);
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

function makeNoteRange(textGetter, textSetter, context) {
    const range = makeRange(textGetter, () => 0);
    range.context = context;
    range.getOoxml = () => ({
        value: noteTrackedChanges
            ? "<w:p><w:ins><w:r><w:t>" + String(textGetter()) + "</w:t></w:r></w:ins></w:p>"
            : "<w:p><w:r><w:t>" + String(textGetter()) + "</w:t></w:r></w:p>"
    });
    range.search = function (query) {
        searchQueries.push(String(query));
        const source = String(textGetter());
        const matches = [];
        let cursor = 0;
        if (String(query).trim() && String(query).length <= 255 &&
            !String(query).includes("\u0002") &&
            !/^[\s\p{P}\p{S}]+$/u.test(String(query))) {
            while ((cursor = source.indexOf(query, cursor)) !== -1) {
                const matchStart = cursor;
                let textLoaded = false;
                matches.push({
                    context: context,
                    load(property) {
                        assert.equal(property, "text");
                        textLoaded = true;
                        itemLoads++;
                    },
                    get text() {
                        assert.equal(textLoaded, true);
                        return query;
                    },
                    insertText(replacement) {
                        const current = String(textGetter());
                        if (current.slice(matchStart, matchStart + query.length) !== query) {
                            throw new Error("note range no longer matches its source text");
                        }
                        const insertedText = String(replacement || "");
                        assert.equal(
                            documentTrackingMode,
                            "TrackAll",
                            "footnote replacements must be recorded as Word tracked changes"
                        );
                        textSetter(
                            current.slice(0, matchStart) + insertedText +
                            current.slice(matchStart + query.length)
                        );
                        return {
                            load() {},
                            get text() { return insertedText; },
                            font: {}
                        };
                    }
                });
                cursor += query.length;
            }
        }
        let itemsLoaded = false;
        return {
            context: context,
            load(property) {
                assert.equal(property, "items");
                itemsLoaded = true;
                collectionLoads++;
            },
            get items() {
                assert.equal(itemsLoaded, true);
                return matches;
            }
        };
    };
    return range;
}

function makeNoteCollection(context, type) {
    const isFootnote = type === "footnote";
    const noteRange = makeNoteRange(
        () => isFootnote ? footnoteText : endnoteText,
        value => { if (isFootnote) footnoteText = value; else endnoteText = value; },
        context
    );
    const noteItem = {
        context: context,
        type: isFootnote ? "Footnote" : "Endnote",
        load(property) { assert.equal(property, "type"); },
        body: {
            getRange() { return noteRange; }
        },
        reference: {
            context: context,
            load(property) { assert.equal(property, "text"); },
            text: isFootnote ? "1" : "1"
        },
        track() {},
        untrack() {}
    };
    let itemsLoaded = false;
    return {
        context: context,
        track() {},
        untrack() {},
        load(property) {
            assert.equal(property, "items");
            itemsLoaded = true;
        },
        get items() {
            assert.equal(itemsLoaded, true, "note items must be loaded before access");
            const noteText = isFootnote ? footnoteText : endnoteText;
            return noteText.trim() ? [noteItem] : [];
        }
    };
}

const selectedRange = makeRange(() => selectionText, () => 0);
const bodyRange = makeRange(() => documentText, () => 0);
Object.defineProperties(bodyRange, {
    footnotes: { get() { return makeNoteCollection(this.context, "footnote"); } },
    endnotes: { get() { return makeNoteCollection(this.context, "endnote"); } }
});
Object.defineProperties(selectedRange, {
    footnotes: { get() { return makeNoteCollection(this.context, "footnote"); } },
    endnotes: { get() { return makeNoteCollection(this.context, "endnote"); } }
});

global.Office = {
    context: {
        requirements: {
            isSetSupported(name, version) {
                return name === "WordApi" &&
                    ["1.3", "1.4", "1.5"].includes(version);
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
                    load(property) {
                        assert.equal(property, "changeTrackingMode");
                    },
                    get changeTrackingMode() {
                        return documentTrackingMode;
                    },
                    set changeTrackingMode(value) {
                        documentTrackingMode = String(value);
                    },
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

    // A whitespace-only diff operation is not reliably returned by Word.search.
    // The contextual fallback must resolve it between the adjacent word ranges.
    documentText = "كلمة كلمة";
    selectionText = "";
    const whitespaceCapture = await revision.capture({ scope: "document" });
    const whitespaceRequest = revision.buildRequest(
        "أعد صياغة المستند بأسلوب أكاديمي",
        whitespaceCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    const whitespaceParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: whitespaceRequest.revision.rewriteSegments.map(segment => ({
            segmentId: segment.segmentId,
            text: "كلمة\tكلمة"
        }))
    }), whitespaceCapture.originalText, whitespaceRequest.revision);
    const whitespaceResult = await revision.apply(Object.assign(
        {}, whitespaceCapture, whitespaceParsed
    ));
    assert.equal(whitespaceResult.applied, true);
    assert.equal(documentText, "كلمة\tكلمة");

    // Multi-paragraph rewrites can produce changed spans beyond Word's
    // direct-search limit. They must be applied through verified paragraph
    // boundaries without searching for the entire long phrase.
    const makeLongToken = seed => Array.from({ length: 320 }, (_, index) =>
        (index + seed).toString(36).padStart(3, "0")
    ).join("");
    const longOriginalOne = "بداية " + makeLongToken(1) + " نهاية";
    const longOriginalTwo = "بداية " + makeLongToken(500) + " نهاية";
    const longRevisedOne = "بداية " + makeLongToken(900) + " نهاية";
    const longRevisedTwo = "بداية " + makeLongToken(1300) + " نهاية";
    documentText = longOriginalOne + "\r\n" + longOriginalTwo;
    selectionText = "";
    const longRewriteCapture = await revision.capture({ scope: "document" });
    const longRewriteRequest = revision.buildRequest(
        "أعد صياغة المستند كاملًا بأسلوب أكاديمي",
        longRewriteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    const longRewriteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: longRewriteRequest.revision.rewriteSegments.map((segment, index) => ({
            segmentId: segment.segmentId,
            text: index === 0 ? longRevisedOne : longRevisedTwo
        }))
    }), longRewriteCapture.originalText, longRewriteRequest.revision);
    assert.ok(
        revision.diffTextPreservingParagraphs(longRewriteCapture.originalText, longRewriteParsed.text)
            .some(operation => operation.length > 255),
        "the regression fixture must include a long Word search span"
    );
    const longRewriteResult = await revision.apply(Object.assign(
        {}, longRewriteCapture, longRewriteParsed
    ));
    assert.equal(longRewriteResult.applied, true);
    assert.equal(documentText, longRevisedOne + "\r\n" + longRevisedTwo);

    // Pure punctuation insertions must not depend on Word.search finding
    // the punctuation mark itself. The neighboring exact text is the anchor.
    documentText = "كلمة.";
    selectionText = "";
    const punctuationCapture = await revision.capture({ scope: "document" });
    const punctuationRequest = revision.buildRequest(
        "أعد صياغة المستند بأسلوب أكاديمي",
        punctuationCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    const punctuationParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: punctuationRequest.revision.rewriteSegments.map(segment => ({
            segmentId: segment.segmentId,
            text: "كلمة،."
        }))
    }), punctuationCapture.originalText, punctuationRequest.revision);
    assert.ok(
        revision.diffTextPreservingParagraphs(
            punctuationCapture.originalText,
            punctuationParsed.text
        ).some(operation => operation.length === 0),
        "the regression fixture must include a pure insertion"
    );
    const punctuationResult = await revision.apply(Object.assign(
        {}, punctuationCapture, punctuationParsed
    ));
    assert.equal(punctuationResult.applied, true);
    assert.equal(documentText, "كلمة،.");

    // Footnote references are structural Word anchors (U+0002), not text.
    // The model receives a protected token, while Word replacement searches
    // only the text around the reference and leaves the reference itself intact.
    documentText = "هذه فقرة أصلية تنتهي بحاشية.\u0002";
    selectionText = "";
    const originalFootnoteBody = "مرجع المصدر، بيانات النشر والصفحة.";
    const queryLogStart = searchQueries.length;
    const footnoteCapture = await revision.capture({ scope: "document" });
    const footnoteProofreadingRequest = revision.buildRequest(
        "صحح المستند لغويًا ونحويًا فقط",
        footnoteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    assert.equal(footnoteProofreadingRequest.revision.structuredCorrection, false);
    assert.doesNotMatch(footnoteProofreadingRequest.systemPrompt, /paragraphId/);
    assert.ok(footnoteProofreadingRequest.userPrompt.includes("هذه فقرة أصلية تنتهي بحاشية."));
    const footnoteRequest = revision.buildRequest(
        "أعد صياغة المستند كاملًا بأسلوب أكاديمي",
        footnoteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    assert.equal(footnoteRequest.userPrompt.includes("\u0002"), false);
    const noteTokenMatch = footnoteRequest.userPrompt.match(/\[\[WB_KEEP_[^\]]+\]\]/);
    assert.ok(noteTokenMatch, "the footnote reference must be represented by a protected token");
    const noteToken = noteTokenMatch[0];
    const footnoteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: footnoteRequest.revision.rewriteSegments.map(segment => ({
            segmentId: segment.segmentId,
            text: "صياغة جديدة تحفظ الإحالة." + noteToken
        }))
    }), footnoteCapture.originalText, footnoteRequest.revision);
    assert.equal(footnoteParsed.text, "صياغة جديدة تحفظ الإحالة.\u0002");
    const footnoteResult = await revision.apply(Object.assign(
        {}, footnoteCapture, footnoteParsed
    ));
    assert.equal(footnoteResult.applied, true);
    assert.equal(documentText, "صياغة جديدة تحفظ الإحالة.\u0002");
    assert.equal(originalFootnoteBody, "مرجع المصدر، بيانات النشر والصفحة.");
    assert.equal(
        searchQueries.slice(queryLogStart).some(query => query.includes("\u0002")),
        false,
        "Word.search must never be asked to resolve a footnote/endnote placeholder"
    );

    // A pure insertion immediately before a note reference must anchor to
    // the preceding word, never to the U+0002 control character itself.
    documentText = "كلمة\u0002";
    selectionText = "";
    const inlineFootnoteCapture = await revision.capture({ scope: "document" });
    const inlineFootnoteRequest = revision.buildRequest(
        "أعد صياغة المستند بأسلوب أكاديمي",
        inlineFootnoteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    const inlineFootnoteToken = inlineFootnoteRequest.userPrompt.match(/\[\[WB_KEEP_[^\]]+\]\]/)[0];
    const inlineFootnoteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "rewrite",
        paragraphs: inlineFootnoteRequest.revision.rewriteSegments.map(segment => ({
            segmentId: segment.segmentId,
            text: "كلمة،" + inlineFootnoteToken
        }))
    }), inlineFootnoteCapture.originalText, inlineFootnoteRequest.revision);
    const inlineQueryStart = searchQueries.length;
    const inlineFootnoteResult = await revision.apply(Object.assign(
        {}, inlineFootnoteCapture, inlineFootnoteParsed
    ));
    assert.equal(inlineFootnoteResult.applied, true);
    assert.equal(documentText, "كلمة،\u0002");
    assert.equal(
        searchQueries.slice(inlineQueryStart).some(query => query.includes("\u0002")),
        false
    );

    // Exercise the complete proofreading path in a note-bearing paragraph.
    // Some model endpoints return "review" instead of the requested literal
    // mode "correction"; the parser may normalize that label only because the
    // request metadata explicitly identifies a proofreading operation and
    // the payload contains the structured edits array.
    documentText = "هذه فقرة فيها خطا.\u0002";
    selectionText = "";
    const correctionNoteCapture = await revision.capture({ scope: "document" });
    const correctionNoteRequest = revision.buildRequest(
        "صحح المستند لغويًا ونحويًا فقط",
        correctionNoteCapture.originalText,
        { sourceOffset: 0, paragraphBaseIndex: 0 }
    );
    assert.equal(correctionNoteRequest.revision.structuredCorrection, false);
    assert.doesNotMatch(correctionNoteRequest.systemPrompt, /paragraphId/);
    const correctionNoteParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "review",
        edits: [{
            original: "خطا",
            replacement: "خطأ",
            contextBefore: "هذه فقرة فيها ",
            contextAfter: ".",
            reason: "تصويب إملائي"
        }]
    }), correctionNoteCapture.originalText, correctionNoteRequest.revision);
    assert.equal(correctionNoteParsed.mode, "correction");
    assert.equal(correctionNoteParsed.text, "هذه فقرة فيها خطأ.\u0002");
    const correctionNoteResult = await revision.apply(Object.assign(
        {}, correctionNoteCapture, correctionNoteParsed
    ));
    assert.equal(correctionNoteResult.applied, true);
    assert.equal(documentText, "هذه فقرة فيها خطأ.\u0002");

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
    assert.equal(multiParagraphRequest.revision.structuredCorrection, false);
    assert.doesNotMatch(multiParagraphRequest.systemPrompt, /paragraphId/);
    assert.doesNotMatch(multiParagraphRequest.userPrompt, /paragraphId/);
    const multiParagraphParsed = revision.parseResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [
            {
                original: "الأطراد",
                replacement: "الاطراد",
                contextBefore: "هذا ",
                contextAfter: " الأول",
                reason: "تصويب إملائي"
            },
            {
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

    // Whole-document revision now captures footnotes/endnotes as independent
    // targets. A note correction is applied to its own Word story while the
    // main document text and the reference itself remain unchanged.
    documentText = "نص المتن يبقى كما هو.";
    selectionText = "";
    footnoteText = "حاشية سفلية قديمة فيها خطا.";
    endnoteText = "حاشية ختامية قديمة فيها خطا.";
    const noteWorkspaceCapture = await revision.capture({ scope: "document" });
    assert.equal(noteWorkspaceCapture.notesSupported, true);
    assert.equal(noteWorkspaceCapture.notes.length, 2);
    assert.equal(noteWorkspaceCapture.notes[0].kind, "footnote");
    assert.equal(noteWorkspaceCapture.notes[0].label, "حاشية سفلية 1");
    assert.equal(noteWorkspaceCapture.notes[0].originalText, footnoteText);
    assert.equal(noteWorkspaceCapture.notes[1].kind, "endnote");
    assert.equal(noteWorkspaceCapture.notes[1].label, "حاشية ختامية 1");
    assert.equal(noteWorkspaceCapture.notes[1].originalText, endnoteText);

    const capturedFootnote = noteWorkspaceCapture.notes[0];
    const capturedEndnote = noteWorkspaceCapture.notes[1];
    const correctionInstruction = "صحح المستند لغويًا ونحويًا فقط";
    const noteBatchRequest = revision.buildNotesBatchRequest(correctionInstruction, [
        {
            noteId: capturedFootnote.id,
            notePartId: capturedFootnote.id + "-part-1",
            label: capturedFootnote.label,
            sourceText: capturedFootnote.originalText,
            sourceOffset: 0
        },
        {
            noteId: capturedEndnote.id,
            notePartId: capturedEndnote.id + "-part-1",
            label: capturedEndnote.label,
            sourceText: capturedEndnote.originalText,
            sourceOffset: 0
        }
    ]);
    assert.match(noteBatchRequest.systemPrompt, /notePartId/);
    assert.doesNotMatch(noteBatchRequest.systemPrompt, /paragraphId|من كلمتين إلى ست كلمات/);
    assert.match(noteBatchRequest.userPrompt, /BEGIN_NOTES/);
    assert.equal(noteBatchRequest.revision.noteParts.length, 2);
    const noteBatchParsed = revision.parseNotesBatchResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [
            {
                noteId: capturedFootnote.id,
                notePartId: capturedFootnote.id + "-part-1",
                original: "خطا",
                replacement: "خطأ",
                contextBefore: "حاشية سفلية قديمة فيها ",
                contextAfter: ".",
                reason: "تصويب إملائي"
            },
            {
                noteId: capturedEndnote.id,
                notePartId: capturedEndnote.id + "-part-1",
                original: "خطا",
                replacement: "خطأ",
                contextBefore: "حاشية ختامية قديمة فيها ",
                contextAfter: ".",
                reason: "تصويب إملائي"
            }
        ]
    }), noteBatchRequest.revision);
    assert.equal(noteBatchParsed.parts.length, 2);
    assert.equal(noteBatchParsed.parts[0].text, "حاشية سفلية قديمة فيها خطأ.");
    assert.equal(noteBatchParsed.parts[1].text, "حاشية ختامية قديمة فيها خطأ.");
    assert.equal(noteBatchParsed.parts[0].edits[0].noteId, capturedFootnote.id);
    assert.equal(noteBatchParsed.parts[1].edits[0].noteId, capturedEndnote.id);

    const badRouteResult = revision.parseNotesBatchResponse(JSON.stringify({
        status: "ready",
        mode: "correction",
        edits: [{
            noteId: capturedEndnote.id,
            notePartId: capturedFootnote.id + "-part-1",
            original: "خطا",
            replacement: "خطأ",
            contextBefore: "حاشية سفلية قديمة فيها ",
            contextAfter: ".",
            reason: "محاولة إسناد خاطئة"
        }]
    }), noteBatchRequest.revision);
    assert.equal(badRouteResult.parts[0].edits.length, 0);
    assert.equal(badRouteResult.rejected.length, 1);
    assert.equal(badRouteResult.rejected[0].rejectionCode, "invalid_note_target");

    assert.throws(() => revision.parseNotesBatchResponse(JSON.stringify({
        status: "partial",
        mode: "correction",
        edits: []
    }), noteBatchRequest.revision), /status=partial/);

    const workspaceRevision = {
        id: noteWorkspaceCapture.id,
        targetType: "document",
        originalText: noteWorkspaceCapture.originalText,
        text: noteWorkspaceCapture.originalText,
        mode: "correction",
        edits: [],
        status: "ready",
        notes: noteBatchParsed.parts.map(function (part, index) {
            const sourceNote = noteWorkspaceCapture.notes[index];
            return {
                id: noteWorkspaceCapture.id,
                targetPartId: part.noteId,
                targetType: "note",
                kind: sourceNote.kind,
                index: sourceNote.index,
                label: sourceNote.label,
                originalText: part.originalText,
                text: part.text,
                mode: part.mode,
                edits: part.edits,
                rejectedEdits: part.rejected,
                protected: false
            };
        })
    };
    const workspaceApplied = await revision.applyWorkspace(workspaceRevision);
    assert.equal(workspaceApplied.applied, true);
    assert.equal(workspaceApplied.noteCount, 2);
    assert.equal(documentText, "نص المتن يبقى كما هو.");
    assert.equal(footnoteText, "حاشية سفلية قديمة فيها خطأ.");
    assert.equal(endnoteText, "حاشية ختامية قديمة فيها خطأ.");
    assert.equal(documentTrackingMode, "Off", "the user's original change-tracking mode must be restored after applying notes");
    assert.equal(searchQueries.some(query => query.includes("\u0002")), false);

    // A new tracked edit introduced after capture must invalidate the note
    // preview, even when the visible text still matches the reviewed source.
    noteTrackedChanges = false;
    footnoteText = "حاشية سفلية عاد فيها خطا بعد التقاط المراجعة.";
    endnoteText = "حاشية ختامية عاد فيها خطا بعد التقاط المراجعة.";
    const staleNoteCapture = await revision.capture({ scope: "document" });
    assert.equal(staleNoteCapture.notes[0].protected, false);
    noteTrackedChanges = true;
    const staleNoteReview = {
        id: staleNoteCapture.id,
        targetType: "document",
        originalText: staleNoteCapture.originalText,
        text: staleNoteCapture.originalText,
        mode: "correction",
        status: "ready",
        notes: staleNoteCapture.notes.map(function (note) {
            return {
                targetPartId: note.id,
                kind: note.kind,
                originalText: note.originalText,
                text: String(note.originalText).replace("خطا", "خطأ"),
                mode: "correction",
                status: "ready",
                edits: [{
                    start: note.originalText.indexOf("خطا"),
                    original: "خطا",
                    replacement: "خطأ"
                }],
                protected: false
            };
        })
    };
    const beforeStaleApplyFootnote = footnoteText;
    const beforeStaleApplyEndnote = endnoteText;
    await assert.rejects(
        revision.applyWorkspace(staleNoteReview),
        /تغييرات متعقبة جديدة في الحاشية/
    );
    assert.equal(footnoteText, beforeStaleApplyFootnote);
    assert.equal(endnoteText, beforeStaleApplyEndnote);
    assert.equal(documentTrackingMode, "Off");

    process.stdout.write("document-revision-officejs: collection loading, selection, paragraph targeting, batched note capture, tracked application, and stale-note protection tests passed\n");
})().catch(error => {
    process.stderr.write(String(error && error.stack || error) + "\n");
    process.exitCode = 1;
});
