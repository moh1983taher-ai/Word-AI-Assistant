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




// AI-selected source spans: replace only the bibliographic citation and preserve
// explanatory prose, quotations, and page locators around it.
const quotedAlMankhul = 'يقول: "وللشافعي رضي الله عنه مسلكان يحصر في أحدهما التمسك في الشبه أو المخيل الذي يشهد له أصل معين ويرد كل استدلال مرسل " المنخول 454.';
const alMankhulPairs = engine.buildPairs(quotedAlMankhul, {
    references: [{ author: "الغزالي", title: "المنخول من تعليقات الأصول" }],
    sourceTexts: ["المنخول 454"],
    suggestedTexts: ["الغزالي، المنخول من تعليقات الأصول"]
});
assert.equal(alMankhulPairs.length, 1);
assert.equal(quotedAlMankhul.slice(alMankhulPairs[0].start, alMankhulPairs[0].end), "المنخول");
assert.equal(alMankhulPairs[0].replacementText, "الغزالي، المنخول من تعليقات الأصول");
assert.ok(quotedAlMankhul.slice(0, alMankhulPairs[0].start).includes("وللشافعي"));
assert.ok(quotedAlMankhul.slice(alMankhulPairs[0].end).includes("454"));

const attributedIsnawi = "كما نقل عنه الإسنوي، نهاية السول في شرح منهاج الوصول،";
const isnawiPairs = engine.buildPairs(attributedIsnawi, {
    references: [{ author: "الإسنوي", title: "نهاية السول في شرح منهاج الوصول" }],
    sourceTexts: ["الإسنوي، نهاية السول في شرح منهاج الوصول"],
    suggestedTexts: ["الإسنوي، نهاية السول في شرح منهاج الوصول"]
});
assert.equal(isnawiPairs.length, 0);
assert.equal(isnawiPairs.unchangedCount, 1);
assert.ok(attributedIsnawi.startsWith("كما نقل عنه"));

const quotedZarkashi = 'ومثلها قاعدة الزركشي في العقود: "العادة إذا اطردت ينزل اللفظ في العقود عليها، وإذا اضطربت لم تعتبر ووجب البيان، (وإذا) تعارضت الظنون في اعتبارها فخلاف " المنثور، 2/362.';
const zarkashiPairs = engine.buildPairs(quotedZarkashi, {
    references: [{ author: "الزركشي", title: "المنثور" }],
    sourceTexts: ["المنثور، 2/362"],
    suggestedTexts: ["الزركشي، المنثور"]
});
assert.equal(zarkashiPairs.length, 1);
assert.equal(quotedZarkashi.slice(zarkashiPairs[0].start, zarkashiPairs[0].end), "المنثور");
assert.ok(quotedZarkashi.slice(0, zarkashiPairs[0].start).startsWith("ومثلها قاعدة الزركشي"));
assert.ok(quotedZarkashi.slice(zarkashiPairs[0].end).includes("2/362"));

console.log("Footnote citation engine tests passed.");
