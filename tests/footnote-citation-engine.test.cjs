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
    engine.getIdentityText("1) انظر: الزركشي، المنثور، ج 2، ص 145"),
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
    sourceTexts: ["1) انظر: الزركشي، المنثور، ج 2، ص 145"],
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
            sourceTexts: ["1) انظر: الزركشي، المنثور، ج 2، ص 145"],
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


// Full first-mention citations must preserve their original bibliographic metadata.
// Only the contiguous title/author identity is eligible for replacement.
const detailedCitation =
    "نزهة الألباء في طبقات الأدباء، الأنباري، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)";
const detailedReference = {
    title: "نزهة الألباء في طبقات الأدباء",
    author: "الأنباري",
    authorFull: "عبد الرحمن بن محمد بن عبيد الله الأنصاري أبو البركات كمال الدين الأنباري"
};
const detailedIdentity = engine.resolveBibliographicIdentitySpan(detailedCitation, detailedReference);
assert.ok(detailedIdentity, "a full author/title pair should be found inside a detailed citation");
assert.equal(
    detailedCitation.slice(detailedIdentity.start, detailedIdentity.end),
    "نزهة الألباء في طبقات الأدباء، الأنباري"
);
assert.ok(detailedCitation.slice(detailedIdentity.end).startsWith("، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي"));
assert.ok(detailedCitation.slice(detailedIdentity.end).includes("مكتبة المنار، الأردن، الطبعة الثالثة:"));
assert.ok(detailedCitation.slice(detailedIdentity.end).includes("(١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)"));

const firstMentionIdentity = engine.formatBibliographicIdentity(
    detailedReference,
    "title-author-first",
    true,
    detailedReference.authorFull
);
assert.equal(
    firstMentionIdentity,
    "نزهة الألباء في طبقات الأدباء، عبد الرحمن بن محمد بن عبيد الله الأنصاري أبو البركات كمال الدين الأنباري"
);
assert.equal(
    engine.formatBibliographicIdentity(detailedReference, "title-author-first", false),
    "نزهة الألباء في طبقات الأدباء"
);


assert.equal(
    engine.formatBibliographicIdentity(detailedReference, "reference-only"),
    "نزهة الألباء في طبقات الأدباء"
);
assert.equal(
    engine.formatBibliographicIdentity(detailedReference, "author-title"),
    "الأنباري، نزهة الألباء في طبقات الأدباء"
);
assert.equal(
    engine.formatBibliographicIdentity(detailedReference, "title-author"),
    "نزهة الألباء في طبقات الأدباء، الأنباري"
);

const detailedPreview = engine.buildCitationPreview(
    detailedCitation,
    detailedReference,
    firstMentionIdentity
);
assert.ok(detailedPreview.includes("(المتوفى سنة: ٥٧٧ه‍)"));
assert.ok(detailedPreview.includes("المحقق: إبراهيم السامرائي"));
assert.ok(detailedPreview.includes("مكتبة المنار، الأردن"));
assert.ok(detailedPreview.includes("الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م)"));
assert.ok(detailedPreview.endsWith("(١/٨١)"));

const alreadyFullCitation =
    "نزهة الألباء في طبقات الأدباء، عبد الرحمن بن محمد بن عبيد الله الأنصاري أبو البركات كمال الدين الأنباري، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)";
const alreadyFullSpan = engine.resolveBibliographicIdentitySpan(alreadyFullCitation, detailedReference);
assert.equal(
    alreadyFullCitation.slice(alreadyFullSpan.start, alreadyFullSpan.end),
    "نزهة الألباء في طبقات الأدباء، عبد الرحمن بن محمد بن عبيد الله الأنصاري أبو البركات كمال الدين الأنباري"
);
const alreadyFullPairs = engine.buildPairs(alreadyFullCitation, {
    references: [detailedReference],
    sourceTexts: [alreadyFullCitation],
    replacementTexts: [firstMentionIdentity]
});
assert.equal(alreadyFullPairs.length, 0, "an already canonical full first mention should remain unchanged");
assert.equal(alreadyFullPairs.unchangedCount, 1);


const detailedNote = "([1]) ترجمة مطولة للفراء، ينظر: " + detailedCitation + "؛ ثم يراجع الباحث ذلك لاحقًا.";
const detailedPairs = engine.buildPairs(detailedNote, {
    references: [detailedReference],
    sourceTexts: [detailedCitation],
    suggestedTexts: [detailedPreview],
    replacementTexts: [firstMentionIdentity]
});
assert.equal(detailedPairs.length, 1);
assert.equal(
    detailedNote.slice(detailedPairs[0].start, detailedPairs[0].end),
    detailedCitation.slice(detailedIdentity.start, detailedIdentity.end)
);
assert.ok(detailedNote.slice(0, detailedPairs[0].start).startsWith("([1]) ترجمة مطولة للفراء، ينظر: "));
assert.ok(detailedNote.slice(detailedPairs[0].end).startsWith("، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي"));
assert.ok(detailedNote.slice(detailedPairs[0].end).includes("مكتبة المنار، الأردن، الطبعة الثالثة:"));
assert.ok(detailedNote.slice(detailedPairs[0].end).includes("؛ ثم يراجع الباحث ذلك لاحقًا."));

const dictionaryCitation =
    "تهذيب اللغة، أبي منصور محمد بن أحمد الأزهري، (المتوفى سنة:٣٧٠ه‍) (مادة: قيم)، دار إحياء التراث العربي، بيروت، الطبعة الأولى: (١٤٢١ه‍ /٢٠٠١م): (٩/٣٦٩).";
const dictionaryReference = {
    title: "تهذيب اللغة",
    author: "الأزهري",
    authorFull: "أبي منصور محمد بن أحمد الأزهري"
};
const dictionaryIdentity = engine.resolveBibliographicIdentitySpan(dictionaryCitation, dictionaryReference);
assert.ok(dictionaryIdentity?.includesAuthor);
assert.equal(
    dictionaryCitation.slice(dictionaryIdentity.start, dictionaryIdentity.end),
    "تهذيب اللغة، أبي منصور محمد بن أحمد الأزهري"
);
assert.ok(dictionaryCitation.slice(dictionaryIdentity.end).includes("(مادة: قيم)"));
assert.ok(dictionaryCitation.slice(dictionaryIdentity.end).includes("دار إحياء التراث العربي، بيروت"));
assert.ok(dictionaryCitation.slice(dictionaryIdentity.end).includes("الطبعة الأولى: (١٤٢١ه‍ /٢٠٠١م)"));
assert.ok(dictionaryCitation.slice(dictionaryIdentity.end).includes("(٩/٣٦٩)"));

const prefixedShortCitation = "تهذيب اللغة، للأزهري: (مادة: خلق)، (٧/١٧)";
const prefixedShortReference = {
    title: "تهذيب اللغة",
    author: "الأزهري",
    authorFull: "أبي منصور محمد بن أحمد الأزهري"
};
const prefixedShortIdentity = engine.resolveBibliographicIdentitySpan(
    prefixedShortCitation,
    prefixedShortReference
);
assert.ok(prefixedShortIdentity?.includesAuthor);
assert.equal(
    prefixedShortCitation.slice(prefixedShortIdentity.start, prefixedShortIdentity.end),
    "تهذيب اللغة، للأزهري",
    "attached lam should be included in the matched short author"
);

const ambiguousAuthorCitation =
    "تهذيب اللغة، أبي منصور محمد بن أحمد الأزهري، (المتوفى سنة:٣٧٠ه‍) (مادة: قيم)، دار إحياء التراث العربي";
const ambiguousAuthorReference = { title: "تهذيب اللغة", author: "الأزهري" };
const ambiguousAuthorIdentity = engine.resolveBibliographicIdentitySpan(
    ambiguousAuthorCitation,
    ambiguousAuthorReference
);
assert.ok(ambiguousAuthorIdentity?.authorFound);
assert.equal(ambiguousAuthorIdentity.includesAuthor, false);
assert.equal(
    engine.buildPairs(ambiguousAuthorCitation, {
        references: [ambiguousAuthorReference],
        sourceTexts: [ambiguousAuthorCitation],
        replacementTexts: ["تهذيب اللغة، الأزهري"]
    }).length,
    0,
    "ambiguous author boundaries must be skipped instead of consuming the full author data"
);

// Independent detailed references in one note must receive independent spans.
const secondDetailedCitation =
    "معجم الأدباء، الحموي، (المتوفى سنة:٦٢٦ه‍) تحقيق: إحسان عباس، دار الغرب الإسلامي، بيروت، الطبعة الأولى: (١٤١٤ه‍ / ١٩٩٣): (٦/٢٨١٢ -٢٨١٤)";
const secondDetailedReference = {
    title: "معجم الأدباء",
    author: "الحموي",
    authorFull: "شهاب الدين أبو عبدالله ياقوت بن عبد الله الرومي الحموي"
};
const twoCitationNote = detailedCitation + "، و" + secondDetailedCitation + "؛";
const twoCitationPairs = engine.buildPairs(twoCitationNote, {
    references: [detailedReference, secondDetailedReference],
    sourceTexts: [detailedCitation, secondDetailedCitation],
    replacementTexts: [
        firstMentionIdentity,
        "معجم الأدباء، شهاب الدين أبو عبدالله ياقوت بن عبد الله الرومي الحموي"
    ]
});
assert.equal(twoCitationPairs.length, 2);
assert.ok(twoCitationNote.slice(twoCitationPairs[0].end).includes("(المتوفى سنة: ٥٧٧ه‍)"));
assert.ok(twoCitationNote.slice(twoCitationPairs[1].end).includes("تحقيق: إحسان عباس، دار الغرب الإسلامي"));

async function testDetailedCitationApplyPreservesAllMetadata() {
    const originalNote = "([1]) ترجمة مطولة للفراء، ينظر: " + detailedCitation + "؛ ثم يراجع الباحث ذلك لاحقًا.";
    const note = { text: originalNote };
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
                note.text = note.text.slice(0, start) + replacement + note.text.slice(end);
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
            originalText: originalNote,
            references: [detailedReference],
            sourceTexts: [detailedCitation],
            suggestedTexts: [detailedPreview],
            replacementTexts: [firstMentionIdentity]
        }]);
        assert.deepEqual(result, { applied: 1, skipped: 0, unchanged: 0 });
        const expectedNote = "([1]) ترجمة مطولة للفراء، ينظر: " + firstMentionIdentity +
            "، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)؛ ثم يراجع الباحث ذلك لاحقًا.";
        assert.equal(note.text, expectedNote,
            "Word apply must replace only the identity and preserve all metadata and surrounding prose");
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


// A full author name found in metadata must not prevent the adjacent short
// author form from being used to identify the replaceable book/author span.
const authorRepeatedInMetadataCitation =
    "الكتاب، السيوطي، المحقق: جلال الدين عبد الرحمن بن أبي بكر السيوطي، دار النشر، بيروت، ط2، ص45";
const authorRepeatedInMetadataReference = {
    title: "الكتاب",
    author: "السيوطي",
    authorFull: "جلال الدين عبد الرحمن بن أبي بكر السيوطي"
};
const authorRepeatedIdentity = engine.resolveBibliographicIdentitySpan(
    authorRepeatedInMetadataCitation,
    authorRepeatedInMetadataReference
);
assert.ok(authorRepeatedIdentity?.includesAuthor);
assert.equal(
    authorRepeatedInMetadataCitation.slice(authorRepeatedIdentity.start, authorRepeatedIdentity.end),
    "الكتاب، السيوطي",
    "the short adjacent author must be selected when the full name occurs only in later metadata"
);

const firstMentionFallbackText =
    "الكتاب، جلال الدين عبد الرحمن بن أبي بكر السيوطي";
const firstMentionFallbackPairs = engine.buildPairs(authorRepeatedInMetadataCitation, {
    references: [authorRepeatedInMetadataReference],
    sourceTexts: [authorRepeatedInMetadataCitation],
    replacementTexts: [firstMentionFallbackText]
});
assert.equal(firstMentionFallbackPairs.length, 1);
assert.equal(
    authorRepeatedInMetadataCitation.slice(
        firstMentionFallbackPairs[0].start,
        firstMentionFallbackPairs[0].end
    ),
    "الكتاب، السيوطي"
);
assert.ok(
    authorRepeatedInMetadataCitation.slice(firstMentionFallbackPairs[0].end)
        .startsWith("، المحقق: جلال الدين عبد الرحمن بن أبي بكر السيوطي"),
    "the full name in the editor/metadata segment must remain untouched"
);

const authorTitleFallbackPairs = engine.buildPairs(authorRepeatedInMetadataCitation, {
    references: [authorRepeatedInMetadataReference],
    sourceTexts: [authorRepeatedInMetadataCitation],
    replacementTexts: ["السيوطي، الكتاب"]
});
assert.equal(authorTitleFallbackPairs.length, 1);
assert.equal(
    authorRepeatedInMetadataCitation.slice(
        authorTitleFallbackPairs[0].start,
        authorTitleFallbackPairs[0].end
    ),
    "الكتاب، السيوطي",
    "the author-title format must use the same validated identity span"
);

// Location data accidentally included in the AI title field must be removed
// from the match target so that volume/page information is preserved.
const titleWithLocationCitation = "النووي، المجموع، ج 2، ص 145";
const titleWithLocationReference = {
    author: "النووي",
    title: "المجموع، ج 2، ص 145"
};
const titleWithLocationIdentity = engine.resolveBibliographicIdentitySpan(
    titleWithLocationCitation,
    titleWithLocationReference
);
assert.ok(titleWithLocationIdentity?.includesAuthor);
assert.equal(
    titleWithLocationCitation.slice(
        titleWithLocationIdentity.start,
        titleWithLocationIdentity.end
    ),
    "النووي، المجموع"
);
const titleWithLocationPairs = engine.buildPairs(titleWithLocationCitation, {
    references: [titleWithLocationReference],
    sourceTexts: [titleWithLocationCitation],
    replacementTexts: ["النووي، المجموع المحقق"]
});
assert.equal(titleWithLocationPairs.length, 1);
assert.equal(
    titleWithLocationCitation.slice(titleWithLocationPairs[0].end),
    "، ج 2، ص 145",
    "the volume and page suffix must remain outside the replacement span"
);

const titleWithLocationAuthorTitle = engine.formatBibliographicIdentity(
    titleWithLocationReference,
    "author-title"
);
assert.equal(
    titleWithLocationAuthorTitle,
    "النووي، المجموع",
    "formatted author-title output must not carry page data from the model title field"
);
assert.equal(
    engine.formatBibliographicIdentity(titleWithLocationReference, "title-author"),
    "المجموع، النووي",
    "formatted title-author output must also exclude location metadata"
);

const fullAuthorWithLocationReference = {
    ...titleWithLocationReference,
    authorFull: "أبو زكريا محيي الدين يحيى بن شرف النووي"
};
const firstMentionWithLocation = engine.formatBibliographicIdentity(
    fullAuthorWithLocationReference,
    "title-author-first",
    true,
    fullAuthorWithLocationReference.authorFull
);
assert.equal(
    firstMentionWithLocation,
    "المجموع، أبو زكريا محيي الدين يحيى بن شرف النووي",
    "the first-mention formatter must use the full author without appending page data"
);
const firstMentionWithLocationPairs = engine.buildPairs(titleWithLocationCitation, {
    references: [fullAuthorWithLocationReference],
    sourceTexts: [titleWithLocationCitation],
    replacementTexts: [firstMentionWithLocation]
});
assert.equal(firstMentionWithLocationPairs.length, 1);
assert.equal(
    titleWithLocationCitation.slice(firstMentionWithLocationPairs[0].end),
    "، ج 2، ص 145",
    "first-mention replacement must preserve the original location suffix"
);


testExactWordReplacementPreservesContext()
    .then(() => testDetailedCitationApplyPreservesAllMetadata())
    .then(() => console.log("Footnote citation engine tests passed."))
    .catch(error => {
        console.error(error);
        process.exitCode = 1;
    });
