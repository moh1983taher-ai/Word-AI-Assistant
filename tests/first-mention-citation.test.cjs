const assert = require("node:assert/strict");
const citation = require("../src/taskpane/first-mention-citation.js");

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

console.log("Limited first-mention citation tests passed.");
