const assert = require("node:assert/strict");
const citation = require("../src/taskpane/first-mention-citation.js");
const engine = require("../src/taskpane/footnote-citation-engine.js");

const fullAuthor = "عبد الرحمن بن محمد بن عبيد الله الأنصاري أبو البركات كمال الدين الأنباري";
const title = "نزهة الألباء في طبقات الأدباء";
const reference = { title, author: "الأنباري" };

const detailedCitation =
    "نزهة الألباء في طبقات الأدباء، الأنباري، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)";
const detailedExpected =
    "نزهة الألباء في طبقات الأدباء، " + fullAuthor +
    "، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)";
assert.equal(
    citation.build(detailedCitation, reference, fullAuthor, false),
    detailedExpected,
    "only the short author span changes; death, editor, publisher, city, edition, date, volume and page remain unchanged"
);

// The helper output must pass through the stable citation engine as a complete
// suggestion. This protects both forms of AI-selected source spans: a narrow
// author/title identity and a complete citation carrying publication metadata.
const detailedPairs = engine.buildPairs(detailedCitation, {
    references: [reference],
    sourceTexts: [detailedCitation],
    suggestedTexts: [citation.build(detailedCitation, reference, fullAuthor, false)]
});
assert.equal(detailedPairs.length, 1);
assert.equal(detailedPairs[0].expectedText, detailedCitation);
assert.equal(detailedPairs[0].replacementText, detailedExpected,
    "a complete selected citation must retain all bibliographic metadata in its replacement");

const authorFirstCitation = "السيوطي، الأشباه والنظائر، ط1، ص98";
const authorFirstReference = { title: "الأشباه والنظائر", author: "السيوطي" };
const authorFull = "جلال الدين عبد الرحمن بن أبي بكر السيوطي";
assert.equal(
    citation.build(authorFirstCitation, authorFirstReference, authorFull, false),
    "الأشباه والنظائر، " + authorFull + "، ط1، ص98",
    "the same limited replacement works when the author precedes the title"
);

const repeatedMetadataCitation =
    "الكتاب، السيوطي، المحقق: جلال الدين عبد الرحمن بن أبي بكر السيوطي، دار النشر، بيروت، ط2، ص45";
assert.equal(
    citation.build(
        repeatedMetadataCitation,
        { title: "الكتاب", author: "السيوطي" },
        authorFull,
        false
    ),
    "الكتاب، " + authorFull +
        "، المحقق: جلال الدين عبد الرحمن بن أبي بكر السيوطي، دار النشر، بيروت، ط2، ص45",
    "a full author mentioned later as metadata must not block replacement of the short author beside the title"
);

const alreadyFullCitation =
    "الكتاب، " + authorFull + "، المحقق: إبراهيم السامرائي، دار النشر، ص45";
assert.equal(
    citation.build(alreadyFullCitation, { title: "الكتاب", author: "السيوطي" }, authorFull, false),
    alreadyFullCitation,
    "an already complete first mention remains byte-for-byte unchanged"
);

const ambiguousCitation = "الكتاب، السيوطي؛ ثم الكتاب، السيوطي، ط1";
assert.equal(
    citation.build(ambiguousCitation, { title: "الكتاب", author: "السيوطي" }, authorFull, false),
    ambiguousCitation,
    "ambiguous repeated book/author pairs are left unchanged"
);

const missingAuthorCitation = "المجموع، ج 2، ص 145";
assert.equal(
    citation.build(missingAuthorCitation, { title: "المجموع", author: "" }, authorFull, true),
    "المجموع، " + authorFull + "، ج 2، ص 145",
    "when the user enables author discovery, the name is inserted after the title without losing location data"
);
assert.equal(
    citation.build(missingAuthorCitation, { title: "المجموع", author: "" }, authorFull, false),
    missingAuthorCitation,
    "the helper does not add an absent author unless the user enabled discovery"
);

const separatedAuthorCitation =
    "الكتاب، المحقق: " + authorFull + "، دار النشر، ص45";
assert.equal(
    citation.build(separatedAuthorCitation, { title: "الكتاب", author: "السيوطي" }, authorFull, true),
    separatedAuthorCitation,
    "a non-adjacent author mention is not moved or duplicated"
);

async function testIntegratedWordReplacementPreservesMetadata() {
    const citationSpan = title + "، الأنباري";
    const noteText =
        "اقتباس سابق محفوظ؛ " + citationSpan +
        "، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)؛ تعليق لاحق محفوظ";
    const expectedNote =
        "اقتباس سابق محفوظ؛ " + title + "، " + fullAuthor +
        "، (المتوفى سنة: ٥٧٧ه‍) المحقق: إبراهيم السامرائي، مكتبة المنار، الأردن، الطبعة الثالثة: (١٤٠٥ه‍ / ١٩٨٥م): (١/٨١)؛ تعليق لاحق محفوظ";
    const suggestionText = citation.build(citationSpan, reference, fullAuthor, false);
    const note = { text: noteText };

    function makeRange(start, end) {
        return {
            get text() {
                return note.text.slice(start, end);
            },
            load() {
                return this;
            },
            search(needle) {
                const scopeText = this.text;
                const items = [];
                let from = 0;
                while (from <= scopeText.length - needle.length) {
                    const found = scopeText.indexOf(needle, from);
                    if (found < 0) break;
                    items.push(makeRange(start + found, start + found + needle.length));
                    from = found + Math.max(1, needle.length);
                }
                return {
                    items,
                    load() {
                        return this;
                    }
                };
            },
            insertText(replacement, location) {
                assert.equal(location, "Replace");
                note.text =
                    note.text.slice(0, start) +
                    replacement +
                    note.text.slice(end);
                return {
                    get text() {
                        return replacement;
                    },
                    load() {
                        return this;
                    }
                };
            }
        };
    }

    const footnotes = {
        items: [{
            body: {
                getRange() {
                    return makeRange(0, note.text.length);
                }
            }
        }],
        load() {
            return this;
        }
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
            originalText: noteText,
            references: [reference],
            sourceTexts: [citationSpan],
            suggestedTexts: [suggestionText]
        }]);

        assert.deepEqual(result, { applied: 1, skipped: 0, unchanged: 0 });
        assert.equal(
            note.text,
            expectedNote,
            "Word application must change only the selected author/title identity and retain the quote, editor, publisher, city, edition, date, volume, page and following prose"
        );
    } finally {
        if (previousWord === undefined) delete global.Word;
        else global.Word = previousWord;
    }
}

testIntegratedWordReplacementPreservesMetadata()
    .then(() => console.log("Limited first-mention citation tests passed."))
    .catch(error => {
        console.error(error);
        process.exitCode = 1;
    });
