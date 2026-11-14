const assert = require("node:assert/strict");
const engine = require("../src/taskpane/footnote-citation-engine.js");

const reference = { author: "سيبويه", title: "الكتاب" };
const repeatedPages = [
    { reference, occurrence: { volume: "1", page: "3" } },
    { reference, occurrence: { volume: "2", page: "4" } }
];
assert.equal(engine.uniqueReferenceEntries(repeatedPages).length, 1,
    "one author/title identity should be emitted once per note, regardless of page or volume");

const distinctWork = [
    ...repeatedPages,
    { reference: { author: "سيبويه", title: "شرح الكتاب" }, occurrence: { page: "5" } }
];
assert.equal(engine.uniqueReferenceEntries(distinctWork).length, 2,
    "a distinct title by the same author must remain a separate reference");

assert.equal(
    engine.referenceKey({ author: "سيبويه", title: "الكتاب، 1/3" }),
    engine.referenceKey({ author: "سيبويه", title: "الكتاب، 2/4" }),
    "page suffixes embedded in model title fields must not create separate book identities"
);


assert.equal(
    engine.getIdentityText("الكتاب، سيبويه، 1/3، 2/4"),
    "الكتاب، سيبويه",
    "multiple volume/page pairs must be excluded from the replaceable identity"
);

assert.equal(
    engine.getIdentityText("الكتاب، سيبويه، ص ١٢–١٤"),
    "الكتاب، سيبويه",
    "Arabic-Indic page ranges must be excluded from the replaceable identity"
);

assert.equal(
    engine.getIdentityText("المجموع شرح المهذب، النووي، ج ١، ص ٣"),
    "المجموع شرح المهذب، النووي",
    "volume and page labels must remain outside the replaceable identity"
);

const note = "ورد في الحاشية: الكتاب، سيبويه، 1/3، 2/4، ثم شرح المهذب.";
const pairs = engine.buildPairs(note, {
    references: [reference, reference],
    sourceTexts: ["الكتاب، سيبويه، 1/3", "الكتاب، سيبويه، 2/4"],
    suggestedTexts: ["سيبويه، الكتاب، 1/3", "سيبويه، الكتاب، 2/4"]
});
assert.equal(pairs.length, 1,
    "duplicate AI occurrences for the same work and author must produce one replacement");
assert.equal(pairs[0].expectedText, "الكتاب، سيبويه");
assert.equal(pairs[0].replacementText, "سيبويه، الكتاب");
assert.equal(note.slice(pairs[0].start, pairs[0].end), "الكتاب، سيبويه");

assert.equal(
    engine.buildPairs("ذكر المؤلف في موضع آخر، 1/3", {
        references: [reference],
        sourceTexts: ["عنوان مختلف، 1/3"],
        suggestedTexts: ["سيبويه، الكتاب"]
    }).length,
    0,
    "an unanchored source citation must be skipped rather than guessed"
);

assert.equal(
    engine.buildPairs("الكتاب، سيبويه، 1/3", {
        references: [{ author: "شخص آخر", title: "الكتاب" }],
        sourceTexts: ["عنوان غير مطابق"],
        suggestedTexts: ["شخص آخر، الكتاب"]
    }).length,
    0,
    "a matching title alone must not authorize a fallback replacement"
);


const repeatedNote = "الكتاب، سيبويه، 1/3؛ ثم الكتاب، سيبويه، 1/3.";
const repeatedPairs = engine.buildPairs(repeatedNote, {
    references: [reference, reference],
    sourceTexts: ["الكتاب، سيبويه، 1/3", "الكتاب، سيبويه، 1/3"],
    suggestedTexts: ["سيبويه، الكتاب، 1/3", "سيبويه، الكتاب، 1/3"]
});
assert.equal(repeatedPairs.length, 2,
    "the same work cited twice in one note must yield two independent anchored replacements");
assert.deepEqual(repeatedPairs.map(pair => pair.sourceOccurrenceIndex), [0, 1],
    "repeated source anchors must keep their stable occurrence ordinals");
assert.equal(repeatedNote.slice(repeatedPairs[0].start, repeatedPairs[0].end), "الكتاب، سيبويه");
assert.equal(repeatedNote.slice(repeatedPairs[1].start, repeatedPairs[1].end), "الكتاب، سيبويه");

const oneCitationSeveralPages = "الكتاب، سيبويه، 1/3، 2/4";
const oneCitationPairs = engine.buildPairs(oneCitationSeveralPages, {
    references: [reference, reference],
    sourceTexts: ["الكتاب، سيبويه، 1/3", "الكتاب، سيبويه، 2/4"],
    suggestedTexts: ["سيبويه، الكتاب، 1/3", "سيبويه، الكتاب، 2/4"]
});
assert.equal(oneCitationPairs.length, 1,
    "multiple page locations attached to one citation must not become duplicate identity replacements");

assert.equal(
    engine.normalizeBibliographicTitle("السيوطي، الأشباه والنظائر والنظائر، ط1، ص98", "السيوطي", "ص98"),
    "الأشباه والنظائر، ط1",
    "the formatter must remove a duplicated author prefix and adjacent repeated title word while preserving edition data"
);

assert.equal(
    engine.normalizeBibliographicTitle("الأشباه والنظائر، السيوطي، ط1، ص98", "السيوطي", "ص98"),
    "الأشباه والنظائر، ط1",
    "the author must be removed from a separate bibliographic segment even when it appears between title and edition"
);
assert.equal(
    engine.normalizeBibliographicTitle("السيوطي، الأشباه والنظائر، ص98", "السيوطي"),
    "الأشباه والنظائر، ص98",
    "a page suffix must remain when no separate location value is available to re-append"
);
assert.equal(
    engine.normalizeBibliographicTitle("السيوطي، الأشباه والنظائر، ص98", "السيوطي", "ص97"),
    "الأشباه والنظائر، ص98",
    "a page suffix must not be removed when it disagrees with the separately recovered page"
);

assert.equal(
    engine.findUniqueBibliographicAnchor(
        "المرجع: السيوطي، الأشباه والنظائر، ص98.",
        { author: "السيوطي", title: "الأشباه والنظائر" }
    ),
    "السيوطي، الأشباه والنظائر",
    "a unique author/title anchor must be recovered when the model omitted citation variants"
);
assert.equal(
    engine.findUniqueBibliographicAnchor(
        "السيوطي، الأشباه والنظائر، ثم السيوطي، الأشباه والنظائر",
        { author: "السيوطي", title: "الأشباه والنظائر" }
    ),
    "",
    "the fallback must refuse ambiguous repeated identities rather than guess an occurrence"
);

const occurrenceMatches = [{ id: 1 }, { id: 2 }];
assert.equal(engine.selectOccurrence(occurrenceMatches, 1).id, 2,
    "Word search results must be selected by the validated occurrence ordinal");
assert.equal(engine.selectOccurrence(occurrenceMatches, 2), null,
    "a missing occurrence ordinal must never fall back to the first match");


const validatedSpanNote = "السيوطي، الأشباه والنظائر، ط1، ص98";
assert.equal(engine.findValidatedSourceSpan(validatedSpanNote, "السيوطي: الأشباه"), "",
    "a short AI variant must be rejected when the note continues with the rest of the title");
assert.equal(engine.findValidatedSourceSpan(validatedSpanNote, "السيوطي، الأشباه والنظائر"),
    "السيوطي، الأشباه والنظائر",
    "a complete title anchor may end before edition/page metadata");
assert.equal(engine.findValidatedSourceSpan(validatedSpanNote, "الأشباه والنظائر"), "الأشباه والنظائر",
    "a complete title may be anchored despite author omission");
assert.equal(engine.findValidatedSourceSpan("السيوطي، الأشباه؛ ثم النووي، المجموع", "السيوطي، الأشباه"),
    "السيوطي، الأشباه",
    "a complete short citation may end at a clear delimiter");


assert.equal(
    engine.findUniqueBibliographicAnchor(
        "السيوطي، الأشباه والنظائر، ص98",
        { author: "السيوطي", title: "الأشباه" }
    ),
    "",
    "the unique-anchor fallback must also reject a truncated work title"
);
assert.equal(
    engine.findUniqueBibliographicAnchor(
        "السيوطي، الأشباه والنظائر، ص98",
        { author: "السيوطي", title: "الأشباه والنظائر" }
    ),
    "السيوطي، الأشباه والنظائر",
    "the fallback should retain a complete title while leaving location metadata untouched"
);

// Editorial lead-ins remain outside the replaceable bibliographic identity.
const lookAtNote = "انظر: المجموع، النووي، ج 1، ص 25";
const lookAtPairs = engine.buildPairs(lookAtNote, {
    references: [{ author: "النووي", title: "المجموع" }],
    sourceTexts: ["انظر: المجموع، النووي، ج 1، ص 25"],
    suggestedTexts: ["النووي، المجموع"]
});
assert.equal(lookAtPairs.length, 1);
assert.equal(lookAtNote.slice(lookAtPairs[0].start, lookAtPairs[0].end), "المجموع، النووي",
    "buildPairs must identify the original source span, not pretend it is already replaced");
assert.equal(lookAtPairs[0].replacementText, "النووي، المجموع",
    "the canonical order belongs in replacementText");
assert.ok(lookAtNote.slice(0, lookAtPairs[0].start).startsWith("انظر:"));

// Page ranges followed by "وما بعدها" remain outside the replacement span.
const extendedPageNote = "أساس القياس، الغزالي، ص37 - 39، وما بعدها.";
assert.equal(engine.getIdentityText("الغزالي، أساس القياس، ص37 - 39، وما بعدها"),
    "الغزالي، أساس القياس");
const extendedPagePairs = engine.buildPairs(extendedPageNote, {
    references: [{ author: "الغزالي", title: "أساس القياس" }],
    sourceTexts: ["أساس القياس، الغزالي، ص37 - 39، وما بعدها"],
    suggestedTexts: ["الغزالي، أساس القياس"]
});
assert.equal(extendedPagePairs.length, 1);
assert.equal(extendedPageNote.slice(extendedPagePairs[0].start, extendedPagePairs[0].end),
    "أساس القياس، الغزالي",
    "the source span retains the author's original order");
assert.equal(extendedPagePairs[0].replacementText, "الغزالي، أساس القياس");
assert.ok(extendedPageNote.slice(extendedPagePairs[0].end).includes("ص37 - 39، وما بعدها"));

// A standalone attached kaf before a definite-article author is a valid anchor boundary.
const kafNote = "كالزركشي، البحر المحيط، ج 8، ص 85";
const kafPairs = engine.buildPairs(kafNote, {
    references: [{ author: "الزركشي", title: "البحر المحيط" }],
    sourceTexts: ["كالزركشي، البحر المحيط"],
    suggestedTexts: ["البحر المحيط، الزركشي"]
});
assert.equal(kafPairs.length, 1);
assert.equal(kafNote.slice(kafPairs[0].start, kafPairs[0].end), "الزركشي، البحر المحيط",
    "the author anchor is found after the attached kaf");
assert.equal(kafNote[kafPairs[0].start - 1], "ك",
    "the attached preposition must remain outside the replacement span");
assert.equal(kafPairs[0].replacementText, "البحر المحيط، الزركشي",
    "the replacement may reorder the author and title without consuming the kaf");

// Hadith takhrij is excluded, while a normal bibliographic reference to al-Bukhari is not.
assert.equal(engine.isHadithTakhrijReference({ type: "hadith", author: "البخاري", title: "صحيح البخاري" }), true);
assert.equal(engine.isHadithTakhrijReference({
    type: "book", author: "البخاري", title: "حديث رقم 123",
    variants: ["أخرجه البخاري، حديث رقم 123"]
}), true);
assert.equal(engine.isHadithTakhrijReference({
    type: "book", author: "البخاري", title: "صحيح البخاري",
    variants: ["البخاري، صحيح البخاري، ج 1، ص 20"]
}), false);


assert.equal(
    engine.getIdentityText("(1) انظر: الزركشي، المنثور، ج 2، ص 145"),
    "الزركشي، المنثور",
    "a note counter and editorial lead-in must stay outside the bibliographic replacement"
);
assert.equal(
    engine.getIdentityText("1 انظر: الزركشي، المنثور، ج 2، ص 145"),
    "الزركشي، المنثور",
    "an unbracketed note counter before an editorial lead-in must remain outside the replacement"
);
assert.equal(
    engine.getIdentityText("(انظر: الزركشي، المنثور، ج 2، ص 145)"),
    "الزركشي، المنثور",
    "outer citation parentheses and editorial lead-ins must remain outside the replacement"
);

const protectedContextNote =
    "(1) انظر: الزركشي، المنثور، ج 2، ص 145؛ ويراجع الباحث ذلك في موضع لاحق.";
const protectedContextPairs = engine.buildPairs(protectedContextNote, {
    references: [{ author: "الزركشي", title: "المنثور" }],
    sourceTexts: ["(1) انظر: الزركشي، المنثور، ج 2، ص 145"],
    suggestedTexts: ["الزركشي، المنثور المحقق"]
});
assert.equal(protectedContextPairs.length, 1);
assert.equal(
    protectedContextNote.slice(protectedContextPairs[0].start, protectedContextPairs[0].end),
    "الزركشي، المنثور",
    "only the bibliographic identity may be included in the replacement range"
);
assert.equal(
    protectedContextNote.slice(0, protectedContextPairs[0].start),
    "(1) انظر: ",
    "the note counter and editorial phrase must remain byte-for-byte unchanged"
);
assert.equal(
    protectedContextNote.slice(protectedContextPairs[0].end),
    "، ج 2، ص 145؛ ويراجع الباحث ذلك في موضع لاحق.",
    "volume, pages, separators, and all following prose must remain unchanged"
);

// End-to-end Office.js adapter simulation: the exact source span is located,
// only the bibliographic identity is replaced, and surrounding text survives.
async function testExactWordReplacementPreservesContext() {
    const note = { text: protectedContextNote };
    const searchCalls = [];

    function makeRange(start, end) {
        return {
            get text() { return note.text.slice(start, end); },
            load() { return this; },
            search(needle, options) {
                searchCalls.push({ needle, options: { ...options } });
                const scopeText = this.text;
                const items = [];
                let from = 0;
                while (from <= scopeText.length - needle.length) {
                    const found = scopeText.indexOf(needle, from);
                    if (found < 0) break;
                    items.push(makeRange(start + found, start + found + needle.length));
                    from = found + Math.max(1, needle.length);
                }
                return { items, load() { return this; } };
            },
            insertText(replacement, location) {
                assert.equal(location, "Replace");
                note.text =
                    note.text.slice(0, start) +
                    replacement +
                    note.text.slice(end);
                return { get text() { return replacement; }, load() { return this; } };
            }
        };
    }

    const footnotes = {
        items: [{ body: { getRange: () => makeRange(0, note.text.length) } }],
        load() { return this; }
    };
    const endnotes = { items: [], load() { return this; } };
    const previousWord = global.Word;
    global.Word = {
        async run(callback) {
            return callback({
                document: { body: { footnotes, endnotes } },
                async sync() {}
            });
        }
    };

    try {
        const result = await engine.apply([{
            source: "footnote",
            materialId: "footnote-1",
            noteNumber: 1,
            originalText: protectedContextNote,
            references: [{ author: "الزركشي", title: "المنثور" }],
            sourceTexts: ["(1) انظر: الزركشي، المنثور، ج 2، ص 145"],
            suggestedTexts: ["الزركشي، المنثور المحقق"]
        }]);

        assert.deepEqual(result, { applied: 1, skipped: 0, unchanged: 0 });
        assert.equal(
            note.text,
            "(1) انظر: الزركشي، المنثور المحقق، ج 2، ص 145؛ ويراجع الباحث ذلك في موضع لاحق.",
            "the live note must change only inside the exact bibliographic identity"
        );
        assert.ok(searchCalls.length >= 2);
        for (const call of searchCalls) {
            assert.equal(call.options.matchCase, true);
            assert.equal(call.options.ignorePunct, false);
            assert.equal(call.options.ignoreSpace, false);
        }
    } finally {
        if (previousWord === undefined) delete global.Word;
        else global.Word = previousWord;
    }
}

testExactWordReplacementPreservesContext()
    .then(() => console.log("Footnote citation engine tests passed."))
    .catch(error => {
        console.error(error);
        process.exitCode = 1;
    });

