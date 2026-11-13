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

console.log("Footnote citation engine tests passed.");
