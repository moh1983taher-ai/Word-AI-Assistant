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
    engine.buildPairs("الكتاب، سيبويه، 1/3", {
        references: [reference],
        sourceTexts: ["عنوان مختلف، 1/3"],
        suggestedTexts: ["سيبويه، الكتاب"]
    }).length,
    0,
    "an unanchored source citation must be skipped rather than guessed"
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

console.log("Footnote citation engine tests passed.");
