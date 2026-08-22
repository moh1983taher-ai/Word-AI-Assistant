// ======================================
// Word AI Assistant
// Main Application Controller
// PART 1
// التخزين + المستندات + بنية المستند
// ======================================

// ======================================
// Application State
// يجب أن يكون خارج Office.onReady
// ======================================

const AppState = {

    wordConnected: false,

    currentDocument: null,

    currentChat: null,

    currentProject: null,

    aiSettings: {

        provider: "openrouter",

        key: "",

        model: ""

    },

    currentCitationSources: [],

    searchState: {

        query: "",

        results: [],

        isSearching: false

    },

    streaming: {

        active: false,

        provider: "",

        text: ""

    }

};


let currentProject = null;
let currentDocument = null;
let currentChat = null;
let referencesSourceDocument = null;

let researchScope = {

    type: "project",

    scope: "all",

    id: null,

    name: "جميع المشاريع"

};

// ======================================
// Office Initialization
// ======================================

Office.onReady(function () {


// ======================================
// Elements
// ======================================

const projectsBtn =
    document.getElementById("projects-btn");

const projectsPopup =
    document.getElementById("projects-popup");

const projectsList =
    document.getElementById("projects-list");

const referencesBtn =
    document.getElementById("references-btn");

const referencesWorkspace =
    document.getElementById("references-workspace");

const closeReferences =
    document.getElementById("close-references");

const referencesContent =
    document.getElementById("references-content");

const referencesSourceWorkspace =
    document.getElementById(
        "references-source-workspace"
    );


const wordDocumentPicker =
    document.getElementById(
        "word-document-picker"
    );

const newProjectBtn =
    document.getElementById("new-project-btn");

const chatBtn =
    document.getElementById("chat-btn");

const newChatBtn =
    document.getElementById("new-chat-btn");




// =====================================================
// تثبيت موضع زر الشريط الجانبي الأصلي
// =====================================================


const input =
    document.getElementById("user-input");

const inputArea =
    document.querySelector(".input-area");

const sendBtn =
    document.getElementById("send-btn");

const chatArea =
    document.getElementById("chat-area");

const documentTitle =
    document.getElementById("document-title");


// ======================================
// AI Settings Elements
// ======================================

const settingsBtn =
    document.getElementById("settings-btn");

const settingsWindow =
    document.getElementById("settings-window");

const closeSettings =
    document.getElementById("close-settings");

const showKey =
    document.getElementById("show-key");

const apiKey =
    document.getElementById("api-key");

const provider =
    document.getElementById("provider-select");

const modelSelect =
    document.getElementById("model-select");

const refreshModels =
    document.getElementById("refresh-models");

const saveSettings =
    document.getElementById("save-settings");

const testConnection =
    document.getElementById("test-connection");

const settingsStatus =
    document.getElementById("settings-status");

const providerInfo =
    document.getElementById("provider-info");


// ======================================
// Chat popup
// ======================================

const chatPopup =
    document.getElementById("chat-popup");

const recentChatList =
    document.getElementById("recent-chat-list");


// ======================================
// Search
// ======================================

const searchPopup =
    document.getElementById("search-popup");

const searchInput =
    document.getElementById("search-input");

const searchResults =
    document.getElementById("search-results");

const searchBtn =
    document.getElementById("search-btn");


// ======================================
// Projects System
// ======================================

let projects = [];

try {

    projects =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_PROJECTS"
            )
        ) || [];

}
catch (e) {

    projects = [];

}


projects =
    projects
        .filter(function (project) {

            return (
                project &&
                typeof project === "object"
            );

        })
        .map(function (project) {

            const now =
                new Date().toISOString();

            return {

                id:
                    project.id ||
                    Date.now(),

                name:
                    project.name ||
                    "مشروع جديد",

                createdAt:
                    project.createdAt ||
                    now,

                updatedAt:
                    project.updatedAt ||
                    now,

                documents:
                    Array.isArray(
                        project.documents
                    )
                        ? project.documents
                        : [],

                references:
                    Array.isArray(
                        project.references
                    )
                        ? project.references
                        : [],

                chatIds:
                    Array.isArray(
                        project.chatIds
                    )
                        ? project.chatIds
                        : [],

                settings:
                    project.settings &&
                    typeof project.settings === "object"
                        ? project.settings
                        : {
                            citationStyle: "",
                            notes: ""
                        }

            };

        });





// ======================================
// Orama Retrieval Cache
// ذاكرة محرك البحث الجديد
// ======================================

let oramaRetrievalDb =
    null;

let oramaRetrievalCacheKey =
    "";


// ======================================
// Orama Document Cache
// فهرس Orama للمستند الحالي
// ======================================

let oramaDocumentDb =
    null;

let oramaDocumentId =
    null;


// ======================================
// Documents System
// ======================================

let documents = [];

try {

    documents =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_DOCUMENTS"
            )
        ) || [];


    documents =
        documents
            .filter(function (documentItem) {

                return (
                    documentItem &&
                    typeof documentItem === "object"
                );

            })
            .map(function (documentItem) {

                if (
                    typeof documentItem.indexTokenCount !==
                    "number"
                ) {

                    documentItem.indexTokenCount =
                        0;

                }

                if (
                    typeof documentItem.indexUniqueTerms !==
                    "number"
                ) {

                    documentItem.indexUniqueTerms =
                        0;

                }

                if (
                    typeof documentItem.indexUniqueFamilies !==
                    "number"
                ) {

                    documentItem.indexUniqueFamilies =
                        0;

                }

                return documentItem;

            });

}
catch (e) {

    documents = [];

}


// ======================================
// Save Documents
// ======================================

function saveDocuments() {

    localStorage.setItem(
        "WORD_AI_DOCUMENTS",
        JSON.stringify(
            documents
        )
    );

}


// ======================================
// Working Documents Storage
// ======================================

const DOCUMENT_DB_NAME =
    "WORD_AI_DOCUMENT_STORAGE";

const DOCUMENT_DB_VERSION =
    4;

const DOCUMENT_STORE_NAME =
    "files";

const DOCUMENT_TEXT_STORE_NAME =
    "texts";

const DOCUMENT_INDEX_STORE_NAME =
    "indexes";

const DOCUMENT_STRUCTURE_STORE_NAME =
    "structures";


// ======================================
// Orama Schema Version
// مهم: تغييره يجبر Orama على
// إعادة بناء الفهرس عند الحاجة
// ======================================

const ORAMA_SCHEMA_VERSION =
    1;


// ======================================
// Open Documents Database
// ======================================

function openDocumentDatabase() {

    return new Promise(
        function (
            resolve,
            reject
        ) {

            const request =
                indexedDB.open(
                    DOCUMENT_DB_NAME,
                    DOCUMENT_DB_VERSION
                );


            request.onupgradeneeded =
                function () {

                    const db =
                        request.result;


                    if (
                        !db.objectStoreNames.contains(
                            DOCUMENT_STORE_NAME
                        )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames.contains(
                            DOCUMENT_INDEX_STORE_NAME
                        )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_INDEX_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames.contains(
                            DOCUMENT_TEXT_STORE_NAME
                        )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_TEXT_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames.contains(
                            DOCUMENT_STRUCTURE_STORE_NAME
                        )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_STRUCTURE_STORE_NAME
                        );

                    }

                };


            request.onsuccess =
                function () {

                    resolve(
                        request.result
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };

        }
    );

}


// ======================================
// Save Working Word File
// ======================================

async function saveWorkingWordFile(
    fileId,
    file
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readwrite"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );

            const request =
                store.put(
                    file,
                    fileId
                );


            request.onsuccess =
                function () {

                    resolve(
                        fileId
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}

const referenceStyle = {

    order: "appearance",
    // appearance | author | title

    format: "author-title"
    // author-title | title-author

};

function formatReferenceForOutput(reference) {

    const author =
        String(reference?.author || "").trim();

    const title =
        String(reference?.title || "").trim();

    if (
        referenceStyle.format === "title-author"
    ) {

        return [
            title,
            author
        ]
            .filter(Boolean)
            .join("، ");

    }

    return [
        author,
        title
    ]
        .filter(Boolean)
        .join("، ");
}

function sortReferencesForOutput(references) {

    const list =
        Array.isArray(references)
            ? [...references]
            : [];

    if (
        referenceStyle.order === "author"
    ) {

        return list.sort(
            function (a, b) {

                return String(
                    a?.author || ""
                ).localeCompare(
                    String(
                        b?.author || ""
                    ),
                    "ar"
                );

            }
        );

    }

    if (
        referenceStyle.order === "title"
    ) {

        return list.sort(
            function (a, b) {

                return String(
                    a?.title || ""
                ).localeCompare(
                    String(
                        b?.title || ""
                    ),
                    "ar"
                );

            }
        );

    }

    return list;
}
// ======================================
// Get Working Word File
// ======================================

async function getWorkingWordFile(
    fileId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readonly"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );

            const request =
                store.get(
                    fileId
                );


            request.onsuccess =
                function () {

                    resolve(
                        request.result ||
                        null
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Delete Working Word File
// ======================================

async function deleteWorkingWordFile(
    fileId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readwrite"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );

            const request =
                store.delete(
                    fileId
                );


            request.onsuccess =
                function () {

                    resolve();

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Convert Blob/File to Base64
// ======================================

function fileToBase64(
    file
) {

    return new Promise(
        function (
            resolve,
            reject
        ) {

            const reader =
                new FileReader();


            reader.onload =
                function () {

                    const result =
                        String(
                            reader.result || ""
                        );

                    const commaIndex =
                        result.indexOf(",");


                    if (
                        commaIndex ===
                        -1
                    ) {

                        reject(
                            new Error(
                                "تعذر تحويل ملف Word إلى Base64."
                            )
                        );

                        return;

                    }


                    resolve(
                        result.substring(
                            commaIndex + 1
                        )
                    );

                };


            reader.onerror =
                function () {

                    reject(
                        reader.error ||
                        new Error(
                            "فشل قراءة ملف Word."
                        )
                    );

                };


            reader.readAsDataURL(
                file
            );

        }
    );

}


// ======================================
// Create Document
// ======================================

function createDocument(
    file,
    projectId,
    order
) {

    const now =
        new Date().toISOString();


    const documentId =
        Date.now();


    const documentItem = {

        id:
            documentId,

        projectId:
            projectId,

        name:
            file.name.replace(
                /\.docx$/i,
                ""
            ),

        fileName:
            file.name,

        fileType:
            file.type ||
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",

        storageId:
            String(
                documentId
            ),

        order:
            typeof order === "number"
                ? order
                : 0,

        type:
            "word",

        readStatus:
            "new",

        indexStatus:
            "new",

        indexTokenCount:
            0,

        indexUniqueTerms:
            0,

        indexUniqueFamilies:
            0,

        indexSchemaVersion:
            0,

        oramaSchemaVersion:
            0,

        createdAt:
            now,

        updatedAt:
            now

    };


    documents.push(
        documentItem
    );


    saveDocuments();


    return documentItem;

}


// ======================================
// Update Document Read Status
// ======================================

function updateDocumentReadStatus(
    documentItem,
    status
) {

    if (!documentItem)
        return;


    documentItem.readStatus =
        status;


    documentItem.updatedAt =
        new Date().toISOString();


    if (
        status ===
        "read"
    ) {

        documentItem.readAt =
            new Date().toISOString();

    }


    saveDocuments();

}


// ======================================
// Update Document Index Status
// ======================================

function updateDocumentIndexStatus(
    documentItem,
    status
) {

    if (!documentItem)
        return;


    documentItem.indexStatus =
        status;


    documentItem.updatedAt =
        new Date().toISOString();


    saveDocuments();

}


// ======================================
// Get Project Documents
// ======================================

function getProjectDocuments(
    projectId
) {

    if (!projectId)
        return [];


    return documents
        .filter(function (
            documentItem
        ) {

            return (
                documentItem &&
                documentItem.projectId ===
                    projectId
            );

        })
        .sort(function (
            a,
            b
        ) {

            return (
                a.order -
                b.order
            );

        });

}


// ======================================
// Save Indexed Document Text
// ======================================

async function saveDocumentText(
    documentId,
    text
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_TEXT_STORE_NAME,
                    "readwrite"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_TEXT_STORE_NAME
                );


            const record = {

                documentId:
                    String(
                        documentId
                    ),

                text:
                    String(
                        text || ""
                    ),

                updatedAt:
                    new Date().toISOString()

            };


            const request =
                store.put(
                    record,
                    String(
                        documentId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        record
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Get Indexed Document Text
// ======================================

async function getDocumentText(
    documentId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_TEXT_STORE_NAME,
                    "readonly"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_TEXT_STORE_NAME
                );


            const request =
                store.get(
                    String(
                        documentId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        request.result ||
                        null
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// =====================================================
// READ DOCUMENT SOURCES FOR REFERENCES
// قراءة المتن والحواشي والحواشي الختامية للمراجع
// =====================================================

async function readReferenceSources() {

    if (
        !Office.context.requirements.isSetSupported(
            "WordApi",
            "1.5"
        )
    ) {

        throw new Error(
            "قراءة الحواشي تحتاج إلى Word API 1.5 أو أحدث."
        );

    }


    return await Word.run(
        async function (
            context
        ) {

            const document =
                context.document;


            // =============================================
            // المتن
            // =============================================

            const body =
                document.body;

            body.load(
                "text"
            );


            // =============================================
            // الحواشي السفلية
            // =============================================

            const footnotes =
                body.footnotes;

            footnotes.load(
                "items"
            );


            // =============================================
            // الحواشي الختامية
            // =============================================

            const endnotes =
                body.endnotes;

            endnotes.load(
                "items"
            );


            await context.sync();


            // =============================================
            // تحميل تفاصيل الحواشي
            // =============================================

            footnotes.items.forEach(
                function (
                    note
                ) {

                    note.body.load(
                        "text"
                    );

                    note.reference.load(
                        "text"
                    );

                }
            );


            endnotes.items.forEach(
                function (
                    note
                ) {

                    note.body.load(
                        "text"
                    );

                    note.reference.load(
                        "text"
                    );

                }
            );


            await context.sync();


            return {

                mainText:
                    String(
                        body.text || ""
                    ),

                footnotes:
                    footnotes.items.map(
                        function (
                            note,
                            index
                        ) {

                            return {

                                number:
                                    index + 1,

                                reference:
                                    String(
                                        note.reference?.text || ""
                                    ),

                                text:
                                    String(
                                        note.body?.text || ""
                                    ),

                                rawText:
                                    String(
                                        note.body?.text || ""
                                    )

                            };

                        }
                    ),

                endnotes:
                    endnotes.items.map(
                        function (
                            note,
                            index
                        ) {

                            return {

                                number:
                                    index + 1,

                                reference:
                                    String(
                                        note.reference?.text || ""
                                    ),

                                text:
                                    String(
                                        note.body?.text || ""
                                    ),

                                rawText:
                                    String(
                                        note.body?.text || ""
                                    )

                            };

                        }
                    )

            };

        }
    );

}
// =====================================================
// قراءة قائمة المراجع من المستند
// =====================================================
async function readBibliographyFromCurrentDocument() {

    return await Word.run(
        async function (context) {

            const paragraphs =
                context.document.body.paragraphs;

            paragraphs.load(
                "items/text,items/isListItem"
            );

            await context.sync();

            const items =
                paragraphs.items;

            // ---------------------------------------------
            // البحث عن عنوان "المراجع والمصادر"
            // ---------------------------------------------

            let startIndex = -1;

            for (
                let i = 0;
                i < items.length;
                i++
            ) {

                const text =
                    String(
                        items[i].text || ""
                    )
                        .replace(/\s+/g, " ")
                        .trim();

                if (
                    text === "المراجع والمصادر" ||
                    text === "المصادر والمراجع"
                ) {

                    startIndex = i + 1;
                    break;

                }

            }

            if (
                startIndex === -1
            ) {

                console.warn(
                    "لم يتم العثور على عنوان المراجع والمصادر."
                );

                return [];

            }

            // ---------------------------------------------
            // تحميل أرقام عناصر القائمة
            // ---------------------------------------------

            const listItems = [];

            for (
                let i = startIndex;
                i < items.length;
                i++
            ) {

                if (
                    items[i].isListItem
                ) {

                    const listItem =
                        items[i].listItemOrNullObject;

                    listItem.load(
                        "listString"
                    );

                    listItems.push({
                        index: i,
                        paragraph: items[i],
                        listItem: listItem
                    });

                }

            }

            await context.sync();

            // ---------------------------------------------
            // استخراج المراجع
            // ---------------------------------------------

            const bibliography = [];

            listItems.forEach(
                function (
                    item
                ) {

                    const listString =
                        String(
                            item.listItem.listString || ""
                        ).trim();

                    const text =
                        String(
                            item.paragraph.text || ""
                        )
                            .replace(/\s+/g, " ")
                            .trim();

                    if (
                        !text
                    ) {

                        return;

                    }

                    bibliography.push({

                        id:
                            `bibliography-${bibliography.length + 1}`,

                        number:
                            bibliography.length + 1,

                        listString:
                            listString,

                        text:
                            text

                    });

                }
            );

            console.log(
                "قائمة المراجع والمصادر المستخرجة:",
                bibliography
            );

            return bibliography;

        }
    );

}

function compareUnifiedReferencesWithBibliography(
    unifiedReferences,
    bibliography
) {

    const references =
        Array.isArray(unifiedReferences)
            ? unifiedReferences
            : [];

    const list =
        Array.isArray(bibliography)
            ? bibliography
            : [];

    function normalize(value) {

        return String(value || "")
            .toLowerCase()
            .replace(/[ًٌٍَُِّْـ]/g, "")
            .replace(/[.,،:؛()[\]{}"'`]/g, " ")
            .replace(/\s+/g, " ")
            .trim();

    }

    function titleWords(value) {

        return normalize(value)
            .split(" ")
            .filter(function (word) {

                return (
                    word.length > 1 &&
                    word !== "في" &&
                    word !== "من" &&
                    word !== "على" &&
                    word !== "و"
                );

            });

    }

    function authorMatch(author, text) {

        const value =
            normalize(author);

        if (!value) {
            return true;
        }

        const words =
            value
                .split(" ")
                .filter(Boolean);

        return words.every(
            function (word) {

                return text.includes(word);

            }
        );

    }

    function titleMatch(title, text) {

        const words =
            titleWords(title);

        if (
            words[0] === "كتاب"
        ) {
            words.shift();
        }

        if (!words.length) {
            return false;
        }

        const matched =
            words.filter(
                function (word) {

                    return text.includes(word);

                }
            ).length;

        /*
         * يكفي تطابق معظم كلمات العنوان.
         * وهذا يسمح مثلًا:
         *
         * المجموع
         * المجموع شرح المهذب
         *
         * و:
         *
         * قواطع الأدلة
         * قواطع الأدلة في الأصول
         */

        return (
            matched >=
            Math.max(
                1,
                Math.ceil(
                    words.length * 0.6
                )
            )
        );

    }

    function referenceMatches(
        reference,
        bibliographyText
    ) {

        const text =
            normalize(
                bibliographyText
            );

        const author =
            normalize(
                reference?.author === "غير محدد"
                    ? ""
                    : reference?.author || ""
            );

        const title =
            normalize(
                reference?.title ||
                ""
            );

        /*
         * إذا كان المؤلف والعنوان موجودين:
         * يجب أن يتطابقا معًا.
         */

        if (
            author &&
            title
        ) {

            return (
                authorMatch(
                    author,
                    text
                ) &&
                titleMatch(
                    title,
                    text
                )
            );

        }

        /*
         * إذا كان العنوان فقط معروفًا.
         */

        if (title) {

            return titleMatch(
                title,
                text
            );

        }

        /*
         * إذا كان المؤلف فقط معروفًا.
         * لا نعتبره تطابقًا كافيًا إلا إذا كان النص قصيرًا جدًا.
         */

        return false;

    }

    const bibliographyResults =
        list.map(
            function (item) {

                const text =
                    String(
                        item?.text || ""
                    ).trim();

                const matches =
                    references.filter(
                        function (
                            reference
                        ) {

                            return referenceMatches(
                                reference,
                                text
                            );

                        }
                    );

                return {

                    id:
                        item?.id || "",

                    text:
                        text,

                    matched:
                        matches.length > 0,

                    references:
                        matches

                };

            }
        );

    const matchedReferences =
        new Set();

    bibliographyResults.forEach(
        function (item) {

            item.references.forEach(
                function (reference) {

                    matchedReferences.add(
                        reference
                    );

                }
            );

        }
    );

    const missingFromBibliography =
        references.filter(
            function (reference) {

                return !matchedReferences.has(
                    reference
                );

            }
        );

    const unusedBibliography =
        bibliographyResults.filter(
            function (item) {

                return !item.matched;

            }
        );

    return {

        totalUnifiedReferences:
            references.length,

        totalBibliographyEntries:
            list.length,

        matchedCount:
            matchedReferences.size,

        missingFromBibliography:
            missingFromBibliography,

        unusedBibliography:
            unusedBibliography,

        matches:
            bibliographyResults

    };

}

async function writeFinalBibliographyToDocument(
    finalBibliography,
    bibliographyComparison
) {

    if (
        !Array.isArray(finalBibliography) ||
        finalBibliography.length === 0
    ) {
        throw new Error("لا توجد مراجع.");
    }

    // =============================================
    // تطبيق نمط الإخراج المشترك
    // =============================================

    const outputBibliography =
        sortReferencesForOutput(
            finalBibliography
        );

    return await Word.run(
        async function (context) {

            const paragraphs =
                context.document.body.paragraphs;

            paragraphs.load(
                "items/text,items/isListItem"
            );

            await context.sync();

            let headingIndex = -1;

            for (
                let i = 0;
                i < paragraphs.items.length;
                i++
            ) {

                const text =
                    String(
                        paragraphs.items[i].text || ""
                    )
                        .replace(/\s+/g, " ")
                        .trim();

                if (
                    text === "المراجع والمصادر" ||
                    text === "المصادر والمراجع"
                ) {

                    headingIndex = i;
                    break;

                }

            }

            // =============================================
            // لا توجد قائمة: إنشاء القائمة كاملة
            // =============================================

            if (
                headingIndex === -1
            ) {

                const end =
                    context.document.body.getRange("End");

                end.insertParagraph(
                    "",
                    "Before"
                );

                end.insertParagraph(
                    "المراجع والمصادر",
                    "Before"
                );

                let addedCount = 0;

                outputBibliography.forEach(
                    function (
                        reference
                    ) {

                        const text =
                            formatReferenceForOutput(
                                reference
                            );

                        if (!text) {
                            return;
                        }

                        end.insertParagraph(
                            text,
                            "Before"
                        );

                        addedCount++;

                    }
                );

                await context.sync();

                return {

                    created: true,

                    added: addedCount

                };

            }

            // =============================================
            // توجد قائمة: إضافة المراجع الناقصة فقط
            // =============================================

            const missing =
                bibliographyComparison &&
                Array.isArray(
                    bibliographyComparison.missingFromBibliography
                )
                    ? bibliographyComparison.missingFromBibliography
                    : [];

            if (
                missing.length === 0
            ) {

                return {

                    created: false,

                    added: 0

                };

            }

            // =============================================
            // ترتيب المراجع الناقصة وفق النمط المختار
            // =============================================

            const outputMissing =
                sortReferencesForOutput(
                    missing
                );

            // =============================================
            // آخر فقرة في المستند
            // =============================================

            const lastParagraph =
                paragraphs.items[
                    paragraphs.items.length - 1
                ];

            let addedCount = 0;

            outputMissing.forEach(
                function (
                    reference
                ) {

                    const text =
                        formatReferenceForOutput(
                            reference
                        );

                    if (!text) {
                        return;
                    }

                    lastParagraph.insertParagraph(
                        text,
                        "After"
                    );

                    addedCount++;

                }
            );

            await context.sync();

            return {

                created: false,

                added: addedCount

            };

        }
    );

}

function buildFinalBibliography(
    unifiedReferences,
    comparison
) {

    const references =
        Array.isArray(unifiedReferences)
            ? unifiedReferences
            : [];

    const existing =
        Array.isArray(comparison?.matches)
            ? comparison.matches
            : [];

    const result = [];

    // المراجع الموجودة في القائمة والمطابقة
    existing.forEach(function (item) {

        if (
            item.matched &&
            item.references &&
            item.references.length
        ) {

            result.push(
                item.references[0]
            );

        }

    });

    // المراجع المستشهد بها وغير الموجودة
    (comparison?.missingFromBibliography || [])
        .forEach(function (reference) {

            if (
                !result.includes(reference)
            ) {

                result.push(reference);

            }

        });

    return result;
}
// =====================================================
// REFERENCE PROCESSOR
// المحرك الموحد لتنظيف وتحليل مواد المراجع
// =====================================================

function processReferenceSources(referenceSources) {

    const records = [];

    if (
        !referenceSources ||
        typeof referenceSources !== "object"
    ) {
        return records;
    }

    // =====================================================
    // أدوات التنظيف الأساسية
    // =====================================================

    function cleanRawText(value) {

        let text = String(value || "");

        // إزالة رموز التحكم غير المرغوبة
        text = text.replace(
            /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g,
            " "
        );

        // إزالة رمز الاستبدال
        text = text.replace(
            /\uFFFD/g,
            " "
        );

        // إزالة الرمز الغريب الذي قد يظهر من Word
        text = text.replace(
            /\u0002/g,
            " "
        );

        // توحيد الأسطر والمسافات
        text = text.replace(
            /\s+/g,
            " "
        );

        return text.trim();
    }


    function normalizeArabicDigits(value) {

        return String(value || "")
            .replace(
                /[٠-٩]/g,
                function (digit) {

                    return String(
                        "٠١٢٣٤٥٦٧٨٩".indexOf(
                            digit
                        )
                    );

                }
            );

    }


    function normalizeReferenceText(value) {

        let text =
            cleanRawText(value);

        text =
            normalizeArabicDigits(text);

        // توحيد بعض المسافات حول علامات الترقيم
        text = text.replace(
            /\s*([،,:؛])\s*/g,
            "$1 "
        );

        text = text.replace(
            /\s*\/\s*/g,
            "/"
        );

        text = text.replace(
            /\s*([–—-])\s*/g,
            "$1"
        );

        text =
            text.trim();

        return text;
    }


    // =====================================================
    // إضافة مادة خام واحدة
    //
    // لا نحكم هنا هل هي:
    // مرجع / شرح / إحالة / حديث / مصدر نفسه
    //
    // الذكاء الاصطناعي هو الذي سيقرر ذلك.
    // =====================================================

    function addRawMaterial(
        source,
        text,
        metadata = {}
    ) {

        const cleanedText =
            normalizeReferenceText(
                text
            );

        if (!cleanedText) {
            return;
        }

        records.push({

            id:
                `reference-${records.length + 1}`,

            source:
                source,

            kind:
                "raw",

            originalText:
                String(
                    text || ""
                ).trim(),

            cleanedText:
                cleanedText,

            context:
                cleanedText,

            ...metadata

        });

    }


    // =====================================================
    // 1. المتن
    //
    // نلتقط فقط المواضع التي تبدو كإحالات،
    // ونترك الحكم النهائي للذكاء الاصطناعي.
    // =====================================================

    const mainText =
        String(
            referenceSources.mainText || ""
        );


    if (mainText) {

        // -----------------------------------------------
        // الإحالات الموجودة بين الأقواس
        //
        // (الخصائص، 1/33-34)
        // （الكتاب، 3/22）
        // [الصحاح، 2/619]
        // -----------------------------------------------

        const parentheticalPatterns = [

            /[([]([^()[\]\n]{2,400})[)\]]/g,

            /（([^（）\n]{2,400})）/g,

            /【([^【】\n]{2,400})】/g

        ];


        parentheticalPatterns.forEach(
            function (pattern) {

                let match;

                while (
                    (
                        match =
                            pattern.exec(
                                mainText
                            )
                    ) !== null
                ) {

                    const value =
                        String(
                            match[1] || ""
                        ).trim();

                    if (!value) {
                        continue;
                    }

                    addRawMaterial(
                        "main-text",
                        value,
                        {
                            position:
                                match.index,

                            extraction:
                                "parenthetical"
                        }
                    );

                }

            }
        );


        // -----------------------------------------------
        // الإحالات اللفظية
        //
        // انظر:
        // ينظر:
        // راجع:
        // نقلاً عن:
        // نقلًا عن:
        // المصدر:
        // -----------------------------------------------

        const verbalPattern =
            /(?:انظر|ينظر|راجع|المصدر|نقلاً عن|نقلًا عن)\s*[:：]?\s*([^.\n؛]{3,400})/gi;


        let verbalMatch;

        while (
            (
                verbalMatch =
                    verbalPattern.exec(
                        mainText
                    )
            ) !== null
        ) {

            const fullText =
                String(
                    verbalMatch[0] || ""
                ).trim();

            if (!fullText) {
                continue;
            }

            addRawMaterial(
                "main-text",
                fullText,
                {
                    position:
                        verbalMatch.index,

                    extraction:
                        "verbal"
                }
            );

        }

    }


    // =====================================================
    // 2. الحواشي السفلية
    //
    // الحاشية نفسها هي وحدة التحليل.
    //
    // لا نهتم:
    // ( )
    // [ ]
    // { }
    // 【 】
    // أو أي علامة تحيط بالنص.
    //
    // نأخذ نص الحاشية كاملًا كما أعاده Word.
    // =====================================================

    if (
        Array.isArray(
            referenceSources.footnotes
        )
    ) {

        referenceSources.footnotes.forEach(
            function (
                note,
                index
            ) {

                const noteText =
                    note?.rawText ||
                    note?.text ||
                    "";


                addRawMaterial(
                    "footnote",
                    noteText,
                    {
                        noteNumber:
                            note?.number ||
                            index + 1,

                        marker:
                            note?.reference ||
                            "",

                        extraction:
                            "whole-footnote"
                    }
                );

            }
        );

    }


    // =====================================================
    // 3. الحواشي الختامية
    //
    // نفس الفكرة تمامًا:
    // الحاشية وحدة واحدة،
    // والذكاء الاصطناعي هو الذي يقرر ما بداخلها.
    // =====================================================

    if (
        Array.isArray(
            referenceSources.endnotes
        )
    ) {

        referenceSources.endnotes.forEach(
            function (
                note,
                index
            ) {

                const noteText =
                    note?.rawText ||
                    note?.text ||
                    "";


                addRawMaterial(
                    "endnote",
                    noteText,
                    {
                        noteNumber:
                            note?.number ||
                            index + 1,

                        marker:
                            note?.reference ||
                            "",

                        extraction:
                            "whole-endnote"
                    }
                );

            }
        );

    }


    // =====================================================
    // 4. إزالة التكرار التقني فقط
    //
    // لا ندمج مراجع متشابهة هنا.
    // لا نحذف "المصدر نفسه".
    // لا نحذف الإحالات.
    //
    // نحذف فقط نفس المادة المكررة تقنيًا داخل
    // نفس المصدر ونفس رقم الحاشية.
    // =====================================================

    const seen =
        new Set();


    return records.filter(
        function (
            record
        ) {

            const key =
                [
                    record.source,

                    record.noteNumber ||
                        "",

                    record.cleanedText

                ].join(
                    "|"
                );


            if (
                seen.has(key)
            ) {

                return false;

            }


            seen.add(key);

            return true;

        }
    );

}

// =====================================================
// AI UNIFIED REFERENCE ANALYZER
// تحليل المواد وبناء المراجع الموحدة في طلب واحد
// =====================================================

async function analyzeAndBuildFinalReferencesWithAI(
    processedReferences
) {

    if (
        !Array.isArray(processedReferences) ||
        processedReferences.length === 0
    ) {

        return [];

    }


    const settings =
        getSavedSettings();


    const provider =
        String(
            settings.provider || "openrouter"
        ).toLowerCase();


    const key =
        String(
            settings.key || ""
        ).trim();


    const model =
        String(
            settings.model || ""
        ).trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
        );

    }


    // =================================================
    // إعداد المواد التي سترسل إلى الذكاء الاصطناعي
    // =================================================

    const materials =
        processedReferences.map(
            function (
                material,
                index
            ) {

                return {

                    materialId:
                        String(
                            material.id ||
                            `material-${index + 1}`
                        ),

                    source:
                        String(
                            material.source || ""
                        ),

                    kind:
                        String(
                            material.kind || ""
                        ),

                    noteNumber:
                        material.noteNumber ??
                        null,

                    marker:
                        String(
                            material.marker || ""
                        ),

                    text:
                        String(
                            material.originalText ||
                            material.cleanedText ||
                            material.text ||
                            ""
                        ),

                    context:
                        String(
                            material.context ||
                            ""
                        )

                };

            }
        );


    // =================================================
    // التعليمات الرئيسية
    // =================================================

    const systemPrompt = `

أنت محلل مراجع أكاديمي متخصص في الدراسات العربية
والدراسات الإسلامية والفقه وأصول الفقه.

ستتلقى مواد خام مأخوذة من:
- متن مستند Word
- حواشي سفلية
- حواشي ختامية

مهمتك أن تنتج مباشرة قائمة المراجع الموحدة النهائية.

لا أريد منك قائمة بالمواد الخام.
لا أريد منك نتيجة تحليل أولي لكل مادة.
أريد فقط السجلات المرجعية الموحدة النهائية.

القواعد:

1. استخرج المراجع الببليوغرافية الحقيقية فقط.

2. وحّد الصيغ المختلفة التي تشير إلى الكتاب نفسه.
مثال:
"الزركشي، البحر المحيط 8/85"
و
"الزركشي، البحر المحيط في أصول الفقه، ط1، 8/85"
هما مرجع واحد إذا كان السياق يسمح بذلك.

3. اختلاف الصفحة أو الجزء لا يجعل الكتاب مرجعًا جديدًا.
يجب وضع جميع مواضع الاستشهاد داخل locations.

4. اختلاف طريقة كتابة المرجع لا يجعله مرجعًا جديدًا.
ضع الصيغ الأصلية في variants.

5. احتفظ بكل materialId التي استُخدمت في المرجع داخل occurrences.

6. "المصدر نفسه" و"المرجع نفسه" ليست مراجع جديدة.
اربطها بالمرجع السابق المناسب عند الإمكان.

7. لا تربط "المصدر نفسه" بمرجع سابق إذا كان الربط غير موثوق.
عند الشك اجعل needsReview = true.

8. إذا احتوت مادة واحدة على عدة مراجع:
أنشئ سجلات مستقلة لكل مرجع.

9. لا تعتبر إحالات صفحات الدراسة مثل:
"ص 413"
"ص 25-26"
"في هذه الدراسة ص 67"
مراجع ببليوغرافية.

10. لا تعتبر المصطلحات أو التواريخ أو الجمل الشارحة مراجع.

11. تخريج الحديث يجب أن يبقى من النوع hadith وليس book.

12. المجلات والدوريات تكون من النوع journal.

13. لا تخترع:
- اسم مؤلف
- عنوان كتاب
- دار نشر
- مدينة
- سنة
- رقم مجلد
- رقم صفحة

إذا لم تكن المعلومة موجودة فلا تخترعها.

14. إذا كان رقم أو معلومة تبدو خطأ مطبعيًا في المصدر الأصلي:
احتفظ بها كما هي، وضع ملاحظة في notes.

15. لا تدمج مرجعين لمجرد تشابه العنوان.
يجب أن تتوفر قرائن كافية على أنهما الكتاب نفسه.

16. عند وجود اسم مختصر للكتاب واسم كامل لنفس الكتاب، يجوز توحيدهما إذا كان السياق واضحًا.

17. أعد جميع المراجع الموحدة، حتى المرجع الذي ظهر مرة واحدة.

18. confidence بين 0 و1.

19. needsReview = true عندما يكون التعرف أو الدمج غير مؤكد.

20. أعد JSON فقط.
لا تستخدم Markdown.
لا تضف أي شرح قبل JSON أو بعده.

الصيغة:

{
  "references": [
    {
      "id": "ref-1",
      "type": "book",
      "author": "",
      "title": "",
      "edition": "",
      "editor": "",
      "publisher": "",
      "city": "",
      "year": "",
      "locations": [
        {
          "volume": "",
          "page": "",
          "pageRange": ""
        }
      ],
      "variants": [],
      "occurrences": [],
      "notes": "",
      "confidence": 0.0,
      "needsReview": false
    }
  ]
}

أنواع type المسموح بها:
book
journal
hadith

لا تضع explanatory أو internal-reference أو review داخل references.
هذه مواد لا تدخل قائمة المراجع النهائية.

في occurrences استخدم materialId فقط.
`;


    const userPrompt =
        [
            "حلل مواد المستند التالية وابنِ قائمة المراجع الموحدة مباشرة.",
            "",
            JSON.stringify(
                materials,
                null,
                2
            )
        ].join("\n");


    let result;


    // =================================================
    // Gemini
    // =================================================

    if (
        provider ===
        "gemini"
    ) {

        const response =
            await fetch(
                `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(normalizeGeminiModel(model))}:generateContent?key=${encodeURIComponent(key)}`,
                {

                    method:
                        "POST",

                    headers:
                        {
                            "Content-Type":
                                "application/json"
                        },

                    body:
                        JSON.stringify({

                            systemInstruction:
                                {
                                    parts:
                                        [
                                            {
                                                text:
                                                    systemPrompt
                                            }
                                        ]
                                },

                            contents:
                                [
                                    {
                                        role:
                                            "user",

                                        parts:
                                            [
                                                {
                                                    text:
                                                        userPrompt
                                                }
                                            ]
                                    }
                                ],

                            generationConfig:
                                {
                                    temperature:
                                        0.1,

                                    maxOutputTokens:
                                        12000
                                }

                        })

                }
            );


        result =
            await readJSON(
                response
            );


        if (
            !response.ok
        ) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل تحليل المراجع باستخدام Gemini."
                )
            );

        }


        const answer =
            extractGeminiAnswer(
                result
            );


        return parseUnifiedReferenceJSON(
            answer
        );

    }


    // =================================================
    // OpenAI / OpenRouter / Groq
    // =================================================

    const endpoint =
        provider === "openai"

            ? "https://api.openai.com/v1/chat/completions"

            : provider === "groq"

                ? "https://api.groq.com/openai/v1/chat/completions"

                : "https://openrouter.ai/api/v1/chat/completions";


    const headers = {

        "Content-Type":
            "application/json",

        "Authorization":
            "Bearer " + key

    };


    if (
        provider ===
        "openrouter"
    ) {

        headers[
            "HTTP-Referer"
        ] =
            window.location.href;

        headers[
            "X-Title"
        ] =
            "Research Tools";

    }


    const response =
        await fetch(
            endpoint,
            {

                method:
                    "POST",

                headers:
                    headers,

                body:
                    JSON.stringify({

                        model:
                            model,

                        messages:
                            [
                                {
                                    role:
                                        "system",

                                    content:
                                        systemPrompt
                                },

                                {
                                    role:
                                        "user",

                                    content:
                                        userPrompt
                                }
                            ],

                        temperature:
                            0.1,

                        max_tokens:
                            provider ===
                            "groq"
                                ? 12000
                                : 14000

                    })

            }
        );


    result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                `فشل تحليل المراجع باستخدام ${provider}.`
            )
        );

    }


    const answer =
        extractOpenAIStyleAnswer(
            result,
            provider
        );


    return parseUnifiedReferenceJSON(
        answer
    );

}
// =====================================================
// PARSE UNIFIED REFERENCE JSON
// =====================================================

function parseUnifiedReferenceJSON(
    text
) {

    let raw =
        String(
            text || ""
        ).trim();


    if (!raw) {

        throw new Error(
            "لم يُعد الذكاء الاصطناعي أي بيانات للمراجع."
        );

    }


    raw =
        raw.replace(
            /^```json\s*/i,
            ""
        );


    raw =
        raw.replace(
            /^```\s*/i,
            ""
        );


    raw =
        raw.replace(
            /\s*```$/i,
            ""
        );


    let parsed;


    try {

        parsed =
            JSON.parse(
                raw
            );

    }
    catch (
        error
    ) {

        console.error(
            "النص الذي أعاده الذكاء الاصطناعي:",
            raw
        );


        throw new Error(
            "تعذر قراءة نتيجة المراجع الموحدة التي أعادها الذكاء الاصطناعي."
        );

    }


    if (
        !parsed ||
        !Array.isArray(
            parsed.references
        )
    ) {

        throw new Error(
            "نتيجة الذكاء الاصطناعي لا تحتوي على قائمة مراجع صحيحة."
        );

    }


    return parsed.references;

}
// =====================================================
// PARSE REFERENCE RECORD
// تحليل مكونات المرجع الواحد
// =====================================================

function parseReferenceRecord(record) {

    if (
        !record ||
        typeof record !== "object"
    ) {
        return null;
    }


    const originalText =
        String(
            record.cleanedText ||
            record.originalText ||
            ""
        ).trim();


    if (!originalText) {
        return null;
    }


    // =================================================
    // السجل الأساسي
    // =================================================

    const parsed = {

        id:
            record.id || "",

        source:
            record.source || "",

        kind:
            record.kind || "",

        noteNumber:
            record.noteNumber ||
            null,

        originalText:
            originalText,

        author:
            "",

        title:
            "",

        volume:
            "",

        page:
            "",

        pageRange:
            "",

        edition:
            "",

        editor:
            "",

        publisher:
            "",

        city:
            "",

        year:
            "",

        referenceType:
            "unknown",

        confidence:
            0,

        unresolved:
            []

    };


    // =================================================
    // الإحالات التي لا تحتاج تفكيكًا الآن
    // =================================================

    if (
        record.kind === "internal"
    ) {

        parsed.referenceType =
            "internal";

        parsed.confidence =
            1;

        return parsed;

    }


    if (
        record.kind === "ibid"
    ) {

        parsed.referenceType =
            "ibid";

        parsed.confidence =
            1;

        return parsed;

    }


    if (
        record.kind === "hadith"
    ) {

        parsed.referenceType =
            "hadith";

    }


    // =================================================
    // نسخة للعمل
    // =================================================

    let text =
        originalText;


    // إزالة "انظر:" وما شابهها
    text =
        text.replace(
            /^(?:انظر|ينظر|نقلاً عن|نقلًا عن)\s*[:：]?\s*/i,
            ""
        ).trim();


    // =================================================
    // استخراج الطبعة
    // =================================================

    const editionMatch =
        text.match(
            /(?:ط|الطبعة)\s*\.?\s*([0-9٠-٩]+)/i
        );


    if (editionMatch) {

        parsed.edition =
            editionMatch[1];

    }


    // =================================================
    // استخراج الجزء/الصفحة
    //
    // 8/85
    // 2/161
    // =================================================

    const volumePageMatch =
        text.match(
            /(?:ج\s*)?([0-9٠-٩]+)\s*\/\s*([0-9٠-٩]+)/
        );


    if (volumePageMatch) {

        parsed.volume =
            volumePageMatch[1];

        parsed.page =
            volumePageMatch[2];

    }


    // =================================================
    // استخراج ص 115
    // =================================================

    const pageMatch =
        text.match(
            /(?:ص|صفحة)\s*\.?\s*([0-9٠-٩]+)(?:\s*[-–—]\s*([0-9٠-٩]+))?/i
        );


    if (pageMatch) {

        if (
            !parsed.page
        ) {

            parsed.page =
                pageMatch[1];

        }


        if (
            pageMatch[2]
        ) {

            parsed.pageRange =
                `${pageMatch[1]}-${pageMatch[2]}`;

        }

    }


    // =================================================
    // استخراج السنة
    // =================================================

    const yearMatch =
        text.match(
            /\b([0-9٠-٩]{4})\s*(?:هـ|ه|م)\b/
        );


    if (yearMatch) {

        parsed.year =
            yearMatch[1];

    }


    // =================================================
    // المحقق
    // =================================================

    const editorMatch =
        text.match(
            /تحقيق\s*[:：]?\s*([^،؛.]+)/
        );


    if (editorMatch) {

        parsed.editor =
            editorMatch[1].trim();

    }


    // =================================================
    // الناشر
    // =================================================

    const publisherMatch =
        text.match(
            /(?:دار|مؤسسة)\s+([^،؛.]+)/
        );


    if (publisherMatch) {

        parsed.publisher =
            `${publisherMatch[0]}`.trim();

    }


    // =================================================
    // مدينة النشر
    // =================================================

    const cityMatch =
        text.match(
            /(?:بيروت|القاهرة|دمشق|الرياض|بغداد|مكة|المدينة|عمان|جدة|الكويت|الدوحة|دبي)/
        );


    if (cityMatch) {

        parsed.city =
            cityMatch[0];

    }


    // =================================================
    // تحديد عنوان الكتاب والمؤلف
    //
    // القاعدة الأولية:
    // "المؤلف، العنوان، البيانات..."
    // =================================================

    const parts =
        text
            .split("،")
            .map(
                function (item) {

                    return item.trim();

                }
            )
            .filter(
                Boolean
            );


    if (
        parts.length >= 2
    ) {

        // الجزء الأول مرشح للمؤلف
        parsed.author =
            parts[0];


        // الجزء الثاني مرشح لعنوان الكتاب
        parsed.title =
            parts[1];


        parsed.confidence =
            0.65;

    }
    else {

        // إذا لم توجد فاصلة عربية،
        // نحاول البحث عن بنية واضحة

        const colonParts =
            text.split(":");


        if (
            colonParts.length >= 2
        ) {

            parsed.author =
                colonParts[0].trim();

            parsed.title =
                colonParts[1]
                    .trim();

            parsed.confidence =
                0.5;

        }

    }


    // =================================================
    // تنظيف العنوان من بيانات النشر
    // =================================================

    if (
        parsed.title
    ) {

        parsed.title =
            parsed.title
                .replace(
                    /(?:ط|الطبعة)\s*\.?\s*[0-9٠-٩]+/i,
                    ""
                )
                .trim();

    }


    // =================================================
    // تحديد نوع المرجع
    // =================================================

    if (
        parsed.referenceType !==
        "hadith"
    ) {

        parsed.referenceType =
            "book";

    }


    // =================================================
    // رفع درجة الثقة
    // =================================================

    if (
        parsed.volume &&
        parsed.page
    ) {

        parsed.confidence +=
            0.15;

    }


    if (
        parsed.author
    ) {

        parsed.confidence +=
            0.1;

    }


    if (
        parsed.title
    ) {

        parsed.confidence +=
            0.1;

    }


    parsed.confidence =
        Math.min(
            1,
            parsed.confidence
        );


    // =================================================
    // تحديد ما لم نستطع فهمه
    // =================================================

    if (
        !parsed.author
    ) {

        parsed.unresolved.push(
            "author"
        );

    }


    if (
        !parsed.title
    ) {

        parsed.unresolved.push(
            "title"
        );

    }


    if (
        !parsed.page &&
        !parsed.pageRange
    ) {

        parsed.unresolved.push(
            "page"
        );

    }


    return parsed;

}

// =====================================================
// PARSE ALL REFERENCES
// =====================================================

function parseAllReferenceRecords(
    records
) {

    if (
        !Array.isArray(records)
    ) {

        return [];

    }


    return records
        .map(
            function (
                record
            ) {

                return parseReferenceRecord(
                    record
                );

            }
        )
        .filter(
            Boolean
        );

}

// =====================================================
// NORMALIZE REFERENCE
// توحيد صيغة المرجع للمقارنة
// =====================================================

function normalizeParsedReference(
    record
) {

    if (
        !record ||
        typeof record !== "object"
    ) {
        return null;
    }


    function clean(value) {

        return String(
            value || ""
        )
        .toLowerCase()
        .replace(
            /[\u064B-\u065F\u0670]/g,
            ""
        )
        .replace(
            /[أإآ]/g,
            "ا"
        )
        .replace(
            /ة/g,
            "ه"
        )
        .replace(
            /ى/g,
            "ي"
        )
        .replace(
            /ـ/g,
            ""
        )
        .replace(
            /[^a-z0-9\u0600-\u06FF]+/gi,
            " "
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim();

    }


    const author =
        clean(
            record.author
        );

    const title =
        clean(
            record.title
        );

    const volume =
        clean(
            record.volume
        );

    const page =
        clean(
            record.page
        );

    const pageRange =
        clean(
            record.pageRange
        );


    const referenceKey =
        [
            author,
            title,
            volume,
            page,
            pageRange
        ]
        .filter(
            Boolean
        )
        .join("|");


    return {

        ...record,

        normalized: {

            author:
                author,

            title:
                title,

            volume:
                volume,

            page:
                page,

            pageRange:
                pageRange

        },

        matchKey:
            referenceKey

    };

}
// =====================================================
// NORMALIZE ALL REFERENCES
// =====================================================

function normalizeAllParsedReferences(
    references
) {

    if (
        !Array.isArray(
            references
        )
    ) {
        return [];
    }


    return references
        .map(
            function (
                reference
            ) {

                return normalizeParsedReference(
                    reference
                );

            }
        )
        .filter(
            Boolean
        );

}
// =====================================================
// RESOLVE IBID REFERENCES
// حل «المصدر نفسه» بالرجوع إلى المرجع السابق
// =====================================================

function resolveIbidReferences(
    references
) {

    if (
        !Array.isArray(
            references
        )
    ) {

        return [];

    }


    const resolved = [];

    // آخر مرجع معروف لكل مصدر
    const lastReferenceBySource =
        new Map();


    references.forEach(
        function (
            record
        ) {

            if (!record) {
                return;
            }


            // -----------------------------------------
            // مرجع حقيقي
            // -----------------------------------------

            if (
                record.kind ===
                "reference"
            ) {

                lastReferenceBySource.set(
                    record.source,
                    record
                );

                resolved.push({
                    ...record,
                    resolvedFromIbid: false
                });

                return;

            }


            // -----------------------------------------
            // «المصدر نفسه»
            // -----------------------------------------

            if (
                record.kind ===
                "ibid"
            ) {

                const previous =
                    lastReferenceBySource.get(
                        record.source
                    );


                if (previous) {

                    resolved.push({

                        ...record,

                        resolvedFromIbid:
                            true,

                        resolvedReferenceId:
                            previous.id,

                        resolvedAuthor:
                            previous.author ||
                            "",

                        resolvedTitle:
                            previous.title ||
                            "",

                        resolvedVolume:
                            record.volume ||
                            previous.volume ||
                            "",

                        resolvedPage:
                            record.page ||
                            previous.page ||
                            "",

                        resolvedPageRange:
                            record.pageRange ||
                            previous.pageRange ||
                            "",

                        matchKey:
                            previous.matchKey ||
                            ""

                    });

                }
                else {

                    resolved.push({

                        ...record,

                        resolvedFromIbid:
                            false,

                        unresolvedReason:
                            "لا يوجد مرجع سابق واضح"

                    });

                }

                return;

            }


            // -----------------------------------------
            // أنواع أخرى
            // -----------------------------------------

            resolved.push(record);

        }
    );


    return resolved;

}
// =====================================================
// APPLY IBID RESOLUTION
// =====================================================

function applyIbidResolution(
    references
) {

    return resolveIbidReferences(
        references
    );

}
// =====================================================
// GROUP DUPLICATE REFERENCES
// تجميع الإحالات التي تشير إلى المرجع نفسه
// =====================================================

function groupDuplicateReferences(
    references
) {

    if (
        !Array.isArray(
            references
        )
    ) {

        return [];

    }


    const groups =
        new Map();


    function cleanIdentityPart(
        value
    ) {

        return String(
            value || ""
        )
        .toLowerCase()
        .replace(
            /[\u064B-\u065F\u0670]/g,
            ""
        )
        .replace(
            /[أإآ]/g,
            "ا"
        )
        .replace(
            /ة/g,
            "ه"
        )
        .replace(
            /ى/g,
            "ي"
        )
        .replace(
            /ـ/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim();

    }


    function getIdentityKey(
        record
    ) {

        // ---------------------------------------------
        // إذا كان «المصدر نفسه» قد حُلّ
        // ---------------------------------------------

        if (
            record.resolvedFromIbid &&
            record.resolvedReferenceId
        ) {

            return (
                "resolved:" +
                record.resolvedReferenceId
            );

        }


        const author =
            cleanIdentityPart(
                record.author
            );


        const title =
            cleanIdentityPart(
                record.title
            );


        // ---------------------------------------------
        // الهوية الأساسية:
        // المؤلف + العنوان
        // ولا ندخل الجزء والصفحة في الهوية
        // لأن المرجع الواحد قد يُحال إليه في صفحات مختلفة.
        // ---------------------------------------------

        if (
            author ||
            title
        ) {

            return [
                "book",
                author,
                title
            ].join("|");

        }


        // ---------------------------------------------
        // لا توجد مكونات كافية:
        // نستخدم النص المنظف كحل احتياطي.
        // ---------------------------------------------

        const text =
            cleanIdentityPart(
                record.cleanedText
            );


        return [
            "text",
            text
        ].join("|");

    }


    references.forEach(
        function (
            record
        ) {

            if (!record) {
                return;
            }


            // لا نضع الإحالات الداخلية
            // في قائمة المراجع.
            if (
                record.kind ===
                "internal"
            ) {

                return;

            }


            // مواد المراجعة تبقى منفصلة
            if (
                record.kind ===
                "review"
            ) {

                const reviewKey =
                    "review:" +
                    cleanIdentityPart(
                        record.cleanedText
                    );


                if (
                    !groups.has(
                        reviewKey
                    )
                ) {

                    groups.set(
                        reviewKey,
                        {

                            id:
                                `reference-group-${groups.size + 1}`,

                            status:
                                "review",

                            representative:
                                record,

                            variants:
                                [record],

                            occurrences:
                                [record],

                            pages:
                                [],

                            sources:
                                new Set()

                        }
                    );

                }
                else {

                    const group =
                        groups.get(
                            reviewKey
                        );

                    group.variants.push(
                        record
                    );

                    group.occurrences.push(
                        record
                    );

                }

                return;

            }


            const key =
                getIdentityKey(
                    record
                );


            if (
                !groups.has(
                    key
                )
            ) {

                groups.set(
                    key,
                    {

                        id:
                            `reference-group-${groups.size + 1}`,

                        status:
                            "reference",

                        representative:
                            record,

                        variants:
                            [],

                        occurrences:
                            [],

                        pages:
                            [],

                        sources:
                            new Set()

                    }
                );

            }


            const group =
                groups.get(
                    key
                );


            group.variants.push(
                record
            );


            group.occurrences.push(
                record
            );


            if (
                record.source
            ) {

                group.sources.add(
                    record.source
                );

            }


            // -----------------------------------------
            // حفظ مواضع الإحالة
            // -----------------------------------------

            const page =
                record.page ||
                record.resolvedPage ||
                "";


            const pageRange =
                record.pageRange ||
                record.resolvedPageRange ||
                "";


            const volume =
                record.volume ||
                record.resolvedVolume ||
                "";


            if (
                page ||
                pageRange ||
                volume
            ) {

                const locationKey =
                    [
                        volume,
                        page,
                        pageRange
                    ].join("|");


                const alreadyExists =
                    group.pages.some(
                        function (
                            item
                        ) {

                            return (
                                item.key ===
                                locationKey
                            );

                        }
                    );


                if (
                    !alreadyExists
                ) {

                    group.pages.push({

                        key:
                            locationKey,

                        volume:
                            volume,

                        page:
                            page,

                        pageRange:
                            pageRange,

                        source:
                            record.source,

                        noteNumber:
                            record.noteNumber ||
                            null

                    });

                }

            }

        }
    );


    // =================================================
    // تحويل Set إلى Array
    // =================================================

    return Array.from(
        groups.values()
    )
    .map(
        function (
            group
        ) {

            return {

                ...group,

                sources:
                    Array.from(
                        group.sources
                    ),

                occurrenceCount:
                    group.occurrences.length,

                variantCount:
                    group.variants.length,

                locationCount:
                    group.pages.length

            };

        }
    );

}
// =====================================================
// BUILD FINAL REFERENCE RECORD
// بناء السجل النهائي للمرجع
// =====================================================

function buildFinalReferenceRecords(
    groupedReferences
) {

    if (
        !Array.isArray(
            groupedReferences
        )
    ) {

        return [];
    }


    function clean(value) {

        return String(
            value || ""
        ).trim();

    }


    return groupedReferences.map(
        function (
            group,
            index
        ) {

            const representative =
                group.representative ||
                {};


            const finalRecord = {

                // -----------------------------------------
                // هوية السجل
                // -----------------------------------------

                id:
                    `final-reference-${index + 1}`,

                groupId:
                    group.id || "",


                // -----------------------------------------
                // حالة المرجع
                // -----------------------------------------

                status:
                    group.status ||
                    "reference",


                // -----------------------------------------
                // البيانات الببليوغرافية
                // -----------------------------------------

                author:
                    clean(
                        representative.author
                    ),

                title:
                    clean(
                        representative.title
                    ),

                volume:
                    clean(
                        representative.volume
                        ||
                        representative.resolvedVolume
                    ),

                page:
                    clean(
                        representative.page
                        ||
                        representative.resolvedPage
                    ),

                pageRange:
                    clean(
                        representative.pageRange
                        ||
                        representative.resolvedPageRange
                    ),

                edition:
                    clean(
                        representative.edition
                    ),

                editor:
                    clean(
                        representative.editor
                    ),

                publisher:
                    clean(
                        representative.publisher
                    ),

                city:
                    clean(
                        representative.city
                    ),

                year:
                    clean(
                        representative.year
                    ),


                // -----------------------------------------
                // نوع المصدر
                // -----------------------------------------

                referenceType:
                    representative.referenceType
                    ||
                    (
                        representative.kind === "hadith"
                            ? "hadith"
                            : "book"
                    ),


                // -----------------------------------------
                // النصوص الأصلية
                // -----------------------------------------

                originalText:
                    clean(
                        representative.originalText
                    ),

                cleanedText:
                    clean(
                        representative.cleanedText
                    ),


                // -----------------------------------------
                // الثقة
                // -----------------------------------------

                confidence:
                    Number(
                        representative.confidence || 0
                    ),


                // -----------------------------------------
                // جميع الصيغ التي ظهرت في المستند
                // -----------------------------------------

                variants:
                    Array.isArray(
                        group.variants
                    )
                        ? group.variants.map(
                            function (
                                item
                            ) {

                                return {

                                    text:
                                        clean(
                                            item.originalText ||
                                            item.cleanedText
                                        ),

                                    source:
                                        item.source ||
                                        "",

                                    kind:
                                        item.kind ||
                                        "",

                                    noteNumber:
                                        item.noteNumber ||
                                        null

                                };

                            }
                        )
                        : [],


                // -----------------------------------------
                // جميع مواضع الإحالة
                // -----------------------------------------

                occurrences:
                    Array.isArray(
                        group.occurrences
                    )
                        ? group.occurrences.map(
                            function (
                                item
                            ) {

                                return {

                                    source:
                                        item.source ||
                                        "",

                                    kind:
                                        item.kind ||
                                        "",

                                    noteNumber:
                                        item.noteNumber ||
                                        null,

                                    position:
                                        Number(
                                            item.position || 0
                                        ),

                                    text:
                                        clean(
                                            item.originalText ||
                                            item.cleanedText
                                        )

                                };

                            }
                        )
                        : [],


                // -----------------------------------------
                // مواضع الجزء/الصفحة
                // -----------------------------------------

                locations:
                    Array.isArray(
                        group.pages
                    )
                        ? group.pages.map(
                            function (
                                item
                            ) {

                                return {

                                    volume:
                                        clean(
                                            item.volume
                                        ),

                                    page:
                                        clean(
                                            item.page
                                        ),

                                    pageRange:
                                        clean(
                                            item.pageRange
                                        ),

                                    source:
                                        item.source ||
                                        "",

                                    noteNumber:
                                        item.noteNumber ||
                                        null

                                };

                            }
                        )
                        : [],


                // -----------------------------------------
                // مصادر ظهور المرجع
                // -----------------------------------------

                sources:
                    Array.isArray(
                        group.sources
                    )
                        ? [...group.sources]
                        : [],


                // -----------------------------------------
                // إحصاءات
                // -----------------------------------------

                occurrenceCount:
                    Number(
                        group.occurrenceCount || 0
                    ),

                variantCount:
                    Number(
                        group.variantCount || 0
                    ),

                locationCount:
                    Number(
                        group.locationCount || 0
                    ),


                // -----------------------------------------
                // معلومات المراجعة
                // -----------------------------------------

                unresolved:
                    Array.isArray(
                        representative.unresolved
                    )
                        ? [
                            ...representative.unresolved
                        ]
                        : [],

                needsReview:
                    group.status === "review"
                    ||
                    (
                        Array.isArray(
                            representative.unresolved
                        )
                        &&
                        representative.unresolved.length > 0
                    )

            };


            return finalRecord;

        }
    );

}
// =====================================================
// RENDER FINAL REFERENCE RECORDS
// عرض السجلات النهائية للمراجع
// =====================================================

function renderFinalReferenceRecords(
    finalReferences
) {

    if (
        !Array.isArray(finalReferences)
        ||
        finalReferences.length === 0
    ) {

        return `
            <div class="references-empty">
                لم يتم العثور على مراجع نهائية.
            </div>
        `;

    }


    return finalReferences
        .map(
            function (
                record,
                index
            ) {

                const author =
                    record.author ||
                    "مؤلف غير محدد";

                const title =
                    record.title ||
                    record.cleanedText ||
                    "مرجع غير محدد";


                const location =
                    record.volume
                    &&
                    record.page

                        ? `ج ${record.volume} / ص ${record.page}`

                        : record.pageRange

                            ? `ص ${record.pageRange}`

                            : record.page

                                ? `ص ${record.page}`

                                : "";


                const reviewClass =
                    record.needsReview
                        ? " reference-needs-review"
                        : "";


                return `
                    <div
                        class="final-reference-card${reviewClass}"
                        data-reference-id="${record.id}"
                    >

                        <div class="final-reference-head">

                            <span class="final-reference-number">
                                ${index + 1}
                            </span>

                            <span class="final-reference-type">
                                ${
                                    record.referenceType === "hadith"
                                        ? "حديث"
                                        : "كتاب"
                                }
                            </span>

                        </div>


                        <div class="final-reference-author">
                            ${author}
                        </div>


                        <div class="final-reference-title">
                            ${title}
                        </div>


                        ${
                            location
                                ? `
                                <div class="final-reference-location">
                                    ${location}
                                </div>
                                `
                                : ""
                        }


                        <div class="final-reference-meta">

                            تكرار:
                            ${record.occurrenceCount}

                            &nbsp;·&nbsp;

                            صيغ:
                            ${record.variantCount}

                            ${
                                record.needsReview
                                    ? `
                                        &nbsp;·&nbsp;
                                        <span class="reference-review-label">
                                            يحتاج مراجعة
                                        </span>
                                    `
                                    : ""
                            }

                        </div>

                    </div>
                `;

            }
        )
        .join("");

}
// ======================================
// Normalize Search Text
// العربية
// ======================================

function normalizeSearchText(
    text
) {

    return String(
        text || ""
    )

        // إزالة التشكيل
        .replace(
            /[\u064B-\u065F\u0670]/g,
            ""
        )

        // إزالة التطويل
        .replace(
            /\u0640/g,
            ""
        )

        // توحيد الألف
        .replace(
            /[أإآٱ]/g,
            "ا"
        )

        // توحيد الياء
        .replace(
            /ى/g,
            "ي"
        )

        // توحيد بعض أشكال الهمزة
        .replace(
            /ؤ/g,
            "و"
        )

        .replace(
            /ئ/g,
            "ي"
        )

        // تنظيف المسافات
        .replace(
            /\s+/g,
            " "
        )

        .trim()
        .toLowerCase();

}


// ======================================
// Tokenize Document Text
// ======================================

function tokenizeDocumentText(
    text
) {

    const normalized =
        normalizeSearchText(
            text
        );


    const matches =
        normalized.match(
            /[\p{L}\p{N}]+/gu
        );


    return (
        matches ||
        []
    );

}


// ======================================
// Build Document Structure
// البنية الأساسية ستبقى كما هي
// لأنها حلقة الربط مع Word
// ======================================

async function buildDocumentStructure(
    documentItem
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    const file =
        await getWorkingWordFile(
            documentItem.storageId
        );


    if (!file) {

        throw new Error(
            "لم يتم العثور على نسخة العمل."
        );

    }


    const base64 =
        await fileToBase64(
            file
        );


    return await Word.run(
        async function (
            context
        ) {

            if (
                !Office.context.requirements
                    .isSetSupported(
                        "WordApiHiddenDocument",
                        "1.3"
                    )
            ) {

                throw new Error(
                    "إصدار Word الحالي لا يدعم تحليل بنية المستند."
                );

            }


            const workingDocument =
                context.application
                    .createDocument(
                        base64
                    );


            const paragraphs =
                workingDocument.body.paragraphs;


            paragraphs.load(
                [
                    "items/text",
                    "items/styleBuiltIn",
                    "items/tableNestingLevel"
                ]
            );


            const tables =
                workingDocument.body.tables;


            tables.load(
                [
                    "items/rowCount",
                    "items/columnCount",
                    "items/styleBuiltIn"
                ]
            );


            await context.sync();


            const paragraphItems =
                paragraphs.items.map(
                    function (
                        paragraph,
                        index
                    ) {

                        const text =
                            String(
                                paragraph.text ||
                                ""
                            ).trim();


                        return {

                            index:
                                index,

                            id:
                                String(
                                    index
                                ),

                            text:
                                text,

                            style:
                                paragraph.styleBuiltIn ||
                                "",

                            tableNestingLevel:
                                paragraph.tableNestingLevel ||
                                0

                        };

                    }
                );


            const headings =
                paragraphItems.filter(
                    function (
                        paragraph
                    ) {

                        return (
                            paragraph.style &&
                            /^Heading[1-9]$/i.test(
                                paragraph.style
                            )
                        );

                    }
                );


            const tableItems =
                tables.items.map(
                    function (
                        table,
                        index
                    ) {

                        return {

                            index:
                                index,

                            rows:
                                table.rowCount,

                            columns:
                                table.columnCount,

                            style:
                                table.styleBuiltIn ||
                                ""

                        };

                    }
                );


            return {

                documentId:
                    String(
                        documentItem.id
                    ),

                paragraphCount:
                    paragraphItems.length,

                headingCount:
                    headings.length,

                tableCount:
                    tableItems.length,

                paragraphs:
                    paragraphItems,

                headings:
                    headings,

                tables:
                    tableItems,

                updatedAt:
                    new Date().toISOString()

            };

        }
    );

}


// ======================================
// Save Document Structure
// ======================================

async function saveDocumentStructure(
    documentId,
    structureData
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_STRUCTURE_STORE_NAME,
                    "readwrite"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_STRUCTURE_STORE_NAME
                );


            const request =
                store.put(
                    structureData,
                    String(
                        documentId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        structureData
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Get Document Structure
// ======================================

async function getDocumentStructure(
    documentId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (
            resolve,
            reject
        ) {

            const transaction =
                db.transaction(
                    DOCUMENT_STRUCTURE_STORE_NAME,
                    "readonly"
                );

            const store =
                transaction.objectStore(
                    DOCUMENT_STRUCTURE_STORE_NAME
                );


            const request =
                store.get(
                    String(
                        documentId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        request.result ||
                        null
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Ensure Document Structure
// ======================================

async function ensureDocumentStructure(
    documentItem
) {

    let structure =
        await getDocumentStructure(
            documentItem.id
        );


    if (
        structure &&
        Array.isArray(
            structure.paragraphs
        )
    ) {

        return structure;

    }


    structure =
        await buildDocumentStructure(
            documentItem
        );


    await saveDocumentStructure(
        documentItem.id,
        structure
    );


    return structure;

}


// ======================================
// Read Current Working Word Document
// القراءة فقط
// ======================================

async function readCurrentWordDocument(
    documentItem
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    updateDocumentReadStatus(
        documentItem,
        "reading"
    );


    try {

        if (
            !Office.context.requirements
                .isSetSupported(
                    "WordApiHiddenDocument",
                    "1.3"
                )
        ) {

            throw new Error(
                "إصدار Word الحالي لا يدعم قراءة نسخة العمل."
            );

        }


        const file =
            await getWorkingWordFile(
                documentItem.storageId
            );


        if (!file) {

            throw new Error(
                "لم يتم العثور على نسخة العمل."
            );

        }


        const base64 =
            await fileToBase64(
                file
            );


        const text =
            await Word.run(
                async function (
                    context
                ) {

                    const workingDocument =
                        context.application
                            .createDocument(
                                base64
                            );


                    const body =
                        workingDocument.body;


                    body.load(
                        "text"
                    );


                    await context.sync();


                    return (
                        body.text ||
                        ""
                    );

                }
            );


        await saveDocumentText(
            documentItem.id,
            text
        );


        const structureData =
            await buildDocumentStructure(
                documentItem
            );


        await saveDocumentStructure(
            documentItem.id,
            structureData
        );


        updateDocumentReadStatus(
            documentItem,
            "read"
        );


        return text;

    }
    catch (error) {

        updateDocumentReadStatus(
            documentItem,
            "error"
        );


        console.error(
            "فشل قراءة مستند Word:",
            error
        );


        throw error;

    }

}




// ======================================
// Set Active Document
//
// currentDocument هو المستند المفعّل
// للدردشة والبحث.
// ======================================

function setCurrentDocument(
    documentItem
) {

    // ==================================
    // إلغاء التفعيل
    // ==================================

    if (
        !documentItem
    ) {

        currentDocument =
            null;


        currentCitationSources =
            [];


        // إعادة ضبط فهرس Orama الحالي
        if (
            typeof oramaRetrievalDb !==
            "undefined"
        ) {

            oramaRetrievalDb =
                null;

        }


        if (
            typeof oramaRetrievalCacheKey !==
            "undefined"
        ) {

            oramaRetrievalCacheKey =
                "";

        }


        if (
            typeof oramaRetrievalDocumentId !==
            "undefined"
        ) {

            oramaRetrievalDocumentId =
                null;

        }


        if (
            documentTitle
        ) {

            documentTitle.textContent =
                "لا يوجد مستند مفتوح";

        }


        renderProjects();


        return;

    }


    // ==================================
    // تفعيل المستند
    // ==================================

    currentDocument =
        documentItem;


    currentCitationSources =
        [];


    if (
        documentTitle
    ) {

        documentTitle.textContent =
            documentItem.name;

    }


    // ==================================
    // تبديل فهرس Orama
    // ==================================

    if (
        typeof oramaRetrievalDb !==
        "undefined"
    ) {

        oramaRetrievalDb =
            null;

    }


    if (
        typeof oramaRetrievalCacheKey !==
        "undefined"
    ) {

        oramaRetrievalCacheKey =
            "";

    }


    if (
        typeof oramaRetrievalDocumentId !==
        "undefined"
    ) {

        oramaRetrievalDocumentId =
            null;

    }


    // ==================================
    // المستند مقروء مسبقًا
    //
    // لا نستدعي ensureDocumentIndex
    // لأنها غير موجودة في النسخة الحالية.
    // ==================================

    if (
        documentItem.readStatus ===
        "read"
    ) {

        renderProjects();

        return;

    }


    // ==================================
    // المستند جديد
    // ==================================

    readCurrentWordDocument(
        documentItem
    )
    .then(
        function (
            text
        ) {

            console.log(
                "محتوى نسخة العمل:",
                text
            );


            renderProjects();

        }
    )
    .catch(
        function (
            error
        ) {

            console.error(
                "تعذر قراءة نسخة العمل:",
                error
            );


            renderProjects();

        }
    );


    renderProjects();

}


// ======================================
// Update Document Timestamp
// ======================================

function touchDocument(
    documentItem
) {

    if (!documentItem)
        return;


    documentItem.updatedAt =
        new Date().toISOString();

}


// ======================================
// Attach Document To Project
// ======================================

function attachDocumentToProject(
    project,
    documentItem
) {

    if (
        !project ||
        !documentItem
    ) {

        return;

    }


    if (
        !Array.isArray(
            project.documents
        )
    ) {

        project.documents =
            [];

    }


    if (
        !project.documents.includes(
            documentItem.id
        )
    ) {

        project.documents.push(
            documentItem.id
        );


        project.updatedAt =
            new Date().toISOString();


        saveProjects();

    }

}


// ======================================
// Set Active Project
// ======================================

function setCurrentProject(
    project
) {

    currentProject =
        project ||
        null;

}


// ======================================
// Add Document
// ======================================

if (
    wordDocumentPicker
) {

    


    wordDocumentPicker.onchange =
        async function () {

            try {

                const file =
                    wordDocumentPicker.files &&
                    wordDocumentPicker.files[0];


                if (!file)
                    return;


                if (
                    !/\.docx$/i.test(
                        file.name
                    )
                ) {

                    console.warn(
                        "الملف المختار ليس DOCX."
                    );

                    return;

                }


                if (!currentProject)
                    return;


                const projectDocuments =
                    getProjectDocuments(
                        currentProject.id
                    );


                const nextOrder =
                    projectDocuments.length +
                    1;


                const documentItem =
                    createDocument(
                        file,
                        currentProject.id,
                        nextOrder
                    );


                await saveWorkingWordFile(
                    documentItem.storageId,
                    file
                );


                attachDocumentToProject(
                    currentProject,
                    documentItem
                );


                setCurrentDocument(
                    documentItem
                );


            

                renderProjects();


                console.log(
                    "تم استيراد مستند Word:",
                    {
                        name:
                            documentItem.name,

                        fileName:
                            documentItem.fileName,

                        storageId:
                            documentItem.storageId
                    }
                );

            }
            catch (error) {

                console.error(
                    "فشل استيراد مستند Word:",
                    error
                );


                renderProjects();

            }

        };

}
// =====================================================
// Orama Search Engine
// PART 2
// محرك البحث الجديد + الفهرسة + العناوين + العائلات
// =====================================================


// =====================================================
// Conservative Family Key
// الحفاظ على ميزة العائلات العربية
// =====================================================

function getConservativeFamilyKey(
    word,
    surfaceSet
) {

    let w =
        normalizeSearchText(
            word
        );


    if (!w) {

        return "";

    }


    // ==================================
    // تنظيف الرموز
    // ==================================

    w =
        w.replace(
            /[^\p{L}\p{N}]+/gu,
            ""
        );


    if (!w) {

        return "";

    }


    // ==================================
    // الكلمات القصيرة
    // ==================================

    if (
        w.length <=
        3
    ) {

        return w;

    }


    // ==================================
    // كلمات محمية
    // ==================================

    const protectedWords =
        new Set([

            "الله",
            "القران",
            "اسلام",
            "اسلامي",
            "اسلامية"

        ]);


    if (
        protectedWords.has(
            w
        )
    ) {

        return w;

    }


    // ==================================
    // السوابق
    // ==================================

    const prefixes = [

        "وال",
        "بال",
        "كال",
        "فال",
        "لل",

        "ول",
        "بل",
        "فل",

        "ال",

        "وا",
        "با",
        "كا",
        "فا"

    ];


    // ==================================
    // اللواحق
    // ==================================

    const suffixes = [

        "يات",
        "ات",

        "هما",
        "هم",
        "هن",
        "ها",

        "ية",
        "يا",

        "ون",
        "ين",
        "ان",

        "كم",
        "كن",

        "ه",
        "ك",

        "ي",
        "ة",

        "ا"

    ];


    const MIN_ROOT_LENGTH =
        3;


    // ==================================
    // إزالة السوابق
    // ==================================

    let prefixChanged =
        true;


    while (
        prefixChanged
    ) {

        prefixChanged =
            false;


        for (
            let i = 0;
            i < prefixes.length;
            i++
        ) {

            const prefix =
                prefixes[i];


            if (
                w.startsWith(
                    prefix
                ) &&
                (
                    w.length -
                    prefix.length
                ) >=
                    MIN_ROOT_LENGTH
            ) {

                w =
                    w.substring(
                        prefix.length
                    );


                prefixChanged =
                    true;


                break;

            }

        }

    }


    // ==================================
    // إزالة اللواحق
    // ==================================

    let suffixChanged =
        true;


    while (
        suffixChanged
    ) {

        suffixChanged =
            false;


        for (
            let i = 0;
            i < suffixes.length;
            i++
        ) {

            const suffix =
                suffixes[i];


            if (
                w.endsWith(
                    suffix
                ) &&
                (
                    w.length -
                    suffix.length
                ) >=
                    MIN_ROOT_LENGTH
            ) {

                w =
                    w.substring(
                        0,
                        w.length -
                        suffix.length
                    );


                suffixChanged =
                    true;


                break;

            }

        }

    }


    // ==================================
    // الحماية النهائية
    // ==================================

    if (
        w.length <
        MIN_ROOT_LENGTH
    ) {

        return normalizeSearchText(
            word
        );

    }


    return w;

}


// =====================================================
// Get Heading Level Number
// =====================================================

function getHeadingLevelNumber(
    style
) {

    const match =
        String(
            style ||
            ""
        ).match(
            /Heading\s*([1-9])/i
        );


    if (match) {

        return Number(
            match[1]
        );

    }


    return 9;

}


// =====================================================
// Get Orama Library
// =====================================================

function getOramaEngine() {

    if (
        !window.Orama
    ) {

        throw new Error(
            "لم يتم تحميل مكتبة Orama."
        );

    }


    return window.Orama;

}


// =====================================================
// Build Orama Retrieval Index
// محرك Orama الرئيسي للاسترجاع
// =====================================================

async function ensureOramaRetrievalIndex(
    documentItem,
    structureData
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    if (!structureData) {

        throw new Error(
            "لا توجد بنية للمستند."
        );

    }


    if (
        !window.Orama
    ) {

        throw new Error(
            "مكتبة Orama غير متاحة."
        );

    }


    const {
        create,
        insertMultiple,
        pluginQPS
    } =
        window.Orama;


    // ==================================
    // مفتاح ذاكرة الفهرس
    // ==================================

    const cacheKey =
        [

            String(
                documentItem.id
            ),

            String(
                documentItem.indexUpdatedAt ||
                ""
            ),

            String(
                (
                    structureData.paragraphs ||
                    []
                ).length
            ),

            String(
                (
                    structureData.headings ||
                    []
                ).length
            ),

            "qps-v1"

        ].join(
            "|"
        );


    // ==================================
    // إعادة استخدام الفهرس
    // ==================================

    if (
        oramaRetrievalDb &&
        oramaRetrievalCacheKey ===
            cacheKey
    ) {

        return oramaRetrievalDb;

    }


    // ==================================
    // إنشاء قاعدة Orama
    // ==================================

    const db =
        create({

            schema: {

                id:
                    "string",

                paragraphIndex:
                    "number",

                paragraphId:
                    "string",

                text:
                    "string",

                heading:
                    "string",

                headingIndex:
                    "number",

                headingLevel:
                    "string",

                isHeading:
                    "boolean"

            },

            language:
                "arabic",

            plugins: [

                pluginQPS()

            ]

        });


    // ==================================
    // الفقرات
    // ==================================

    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    // ==================================
    // العناوين
    // ==================================

    const headings =
        Array.isArray(
            structureData.headings
        )
            ? structureData.headings
                .filter(
                    function (
                        heading
                    ) {

                        return (
                            heading &&
                            typeof heading.index !==
                                "undefined" &&
                            String(
                                heading.text ||
                                ""
                            ).trim()
                        );

                    }
                )
                .sort(
                    function (
                        a,
                        b
                    ) {

                        return (
                            Number(
                                a.index
                            ) -
                            Number(
                                b.index
                            )
                        );

                    }
                )
            : [];


    // ==================================
    // تحديد العنوان الأقرب
    // ==================================

    function getParagraphHeading(
        paragraphIndex
    ) {

        let nearest =
            null;


        for (
            let i =
                headings.length - 1;

            i >= 0;

            i--
        ) {

            const heading =
                headings[i];


            if (
                Number(
                    heading.index
                ) <=
                Number(
                    paragraphIndex
                )
            ) {

                nearest =
                    heading;

                break;

            }

        }


        return nearest;

    }


    // ==================================
    // سجلات Orama
    // ==================================

    const records =
        [];


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                !paragraph
            ) {

                return;

            }


            const text =
                String(
                    paragraph.text ||
                    ""
                ).trim();


            if (!text) {

                return;

            }


            const paragraphIndex =
                Number(
                    paragraph.index
                );


            const heading =
                getParagraphHeading(
                    paragraphIndex
                );


            records.push({

                id:
                    "p-" +
                    String(
                        paragraphIndex
                    ),

                paragraphIndex:
                    paragraphIndex,

                paragraphId:
                    String(
                        paragraph.id ||
                        paragraphIndex
                    ),

                text:
                    text,

                heading:
                    heading
                        ? String(
                            heading.text ||
                            ""
                        ).trim()
                        : "",

                headingIndex:
                    heading
                        ? Number(
                            heading.index
                        )
                        : -1,

                headingLevel:
                    heading
                        ? String(
                            heading.style ||
                            ""
                        )
                        : "",

                isHeading:
                    Boolean(
                        heading &&
                        Number(
                            heading.index
                        ) ===
                        paragraphIndex
                    )

            });

        }
    );


    // ==================================
    // إدخال البيانات
    // ==================================

    if (
        records.length >
        0
    ) {

        insertMultiple(
            db,
            records
        );

    }


    // ==================================
    // حفظ الفهرس
    // ==================================

    oramaRetrievalDb =
        db;


    oramaRetrievalCacheKey =
        cacheKey;


    console.log(
        "تم بناء فهرس Orama + QPS:",
        records.length,
        "فقرة"
    );


    return db;

}


// =====================================================
// Search Orama Headings
// البحث في عناوين المباحث والمطالب
// =====================================================

async function searchOramaHeadings(
    documentItem,
    query
) {

    if (
        !documentItem ||
        !query
    ) {

        return [];

    }


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData ||
        !Array.isArray(
            structureData.headings
        )
    ) {

        return [];

    }


    const headings =
        structureData.headings
            .filter(
                function (
                    heading
                ) {

                    return (
                        heading &&
                        typeof heading.index !==
                            "undefined" &&
                        String(
                            heading.text ||
                            ""
                        ).trim()
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        Number(
                            a.index
                        ) -
                        Number(
                            b.index
                        )
                    );

                }
            );


    if (
        headings.length ===
        0
    ) {

        return [];

    }


    const {
        create,
        insertMultiple,
        search
    } =
        getOramaEngine();


    const db =
        create({

            schema: {

                id:
                    "string",

                headingIndex:
                    "number",

                heading:
                    "string",

                headingNormalized:
                    "string",

                style:
                    "string",

                level:
                    "number"

            },

            language:
                "arabic"

        });


    const records =
        headings.map(
            function (
                heading
            ) {

                return {

                    id:
                        "h-" +
                        String(
                            documentItem.id
                        ) +
                        "-" +
                        String(
                            heading.index
                        ),

                    headingIndex:
                        Number(
                            heading.index
                        ),

                    heading:
                        String(
                            heading.text
                        ).trim(),

                    headingNormalized:
                        normalizeSearchText(
                            heading.text
                        ),

                    style:
                        String(
                            heading.style ||
                            ""
                        ),

                    level:
                        getHeadingLevelNumber(
                            heading.style
                        )

                };

            }
        );


    insertMultiple(
        db,
        records
    );


    const normalizedQuery =
        normalizeSearchText(
            query
        );


    if (!normalizedQuery) {

        return [];

    }


    const result =
        search(
            db,
            {

                term:
                    normalizedQuery,

                properties: [
                    "heading",
                    "headingNormalized"
                ],

                tolerance:
                    1,

                limit:
                    20

            }
        );


    const hits =
        Array.isArray(
            result.hits
        )
            ? result.hits
            : [];


    return hits.map(
        function (
            hit
        ) {

            const heading =
                hit.document;


            const headingText =
                normalizeSearchText(
                    heading.heading
                );


            let score =
                Number(
                    hit.score ||
                    0
                );


            // ==================================
            // تطابق عنوان كامل
            // ==================================

            if (
                headingText ===
                normalizedQuery
            ) {

                score +=
                    40;

            }


            // ==================================
            // احتواء العبارة
            // ==================================

            if (
                headingText.includes(
                    normalizedQuery
                )
            ) {

                score +=
                    25;

            }


            // ==================================
            // تغطية كلمات السؤال
            // ==================================

            const queryTokens =
                getSearchQueryTokens(
                    normalizedQuery
                );


            let matchedTokens =
                0;


            queryTokens.forEach(
                function (
                    token
                ) {

                    if (
                        headingText.includes(
                            token
                        )
                    ) {

                        matchedTokens +=
                            1;

                    }

                }
            );


            const coverage =
                queryTokens.length >
                0
                    ? matchedTokens /
                      queryTokens.length
                    : 0;


            score +=
                matchedTokens *
                6;


            score +=
                coverage *
                20;


            return {

                score:
                    score,

                headingIndex:
                    heading.headingIndex,

                heading:
                    heading.heading,

                style:
                    heading.style,

                level:
                    heading.level,

                matchedTokens:
                    matchedTokens,

                coverage:
                    coverage

            };

        }
    )
    .sort(
        function (
            a,
            b
        ) {

            if (
                b.score !==
                a.score
            ) {

                return (
                    b.score -
                    a.score
                );

            }


            return (
                a.headingIndex -
                b.headingIndex
            );

        }
    );

}


// =====================================================
// Search Query Tokens
// استخراج الكلمات المهمة دون حذف الكلمات الدلالية
// =====================================================

function getSearchQueryTokens(
    query
) {

    const tokens =
        tokenizeDocumentText(
            query
        );


    const stopWords =
        new Set([

            "ما",
            "ماذا",
            "من",
            "في",
            "على",
            "عن",
            "الى",
            "إلى",

            "منه",
            "منها",
            "به",
            "بها",
            "له",
            "لها",

            "هذا",
            "هذه",
            "ذلك",
            "تلك",

            "الذي",
            "التي",
            "الذين",
            "اللاتي",
            "اللائي",

            "بين",
            "مع",
            "ثم",

            "و",
            "ف",
            "ب",
            "ك",
            "ل",

            "أن",
            "إن",
            "هل",

            "كان",
            "كانت",
            "كانوا",
            "تكون",
            "يكون",

            "او",
            "أو",

            "أي",
            "اي"

        ]);


    const filtered =
        tokens.filter(
            function (
                token
            ) {

                return (
                    token.length > 2 &&
                    !stopWords.has(
                        token
                    )
                );

            }
        );


    // ==================================
    // لا نحذف الكلمات الدلالية
    // مثل:
    // المقصود
    // تعريف
    // المشكلة
    // الأثر
    // الأسباب
    // الفرق
    // ==================================

    return (
        filtered.length >
        0
            ? filtered
            : tokens
    );

}


// =====================================================
// Search Orama Document
// البحث الحقيقي بالمحرك الجديد
// =====================================================

async function searchOramaDocument(
    documentItem,
    query,
    options
) {

    if (
        !documentItem
    ) {

        return [];

    }


    const normalizedQuery =
        normalizeSearchText(
            query
        );


    if (!normalizedQuery) {

        return [];

    }


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData
    ) {

        return [];

    }


    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    const {
        search
    } =
        getOramaEngine();


    const settings =
        options || {};


    const limit =
        typeof settings.limit ===
            "number"
                ? settings.limit
                : 50;


    const tolerance =
        typeof settings.tolerance ===
            "number"
                ? settings.tolerance
                : 1;


    // ==================================
    // البحث الأساسي
    // ==================================

    let oramaResult =
        search(
            db,
            {

                term:
                    normalizedQuery,

                properties: [

                    "text",

                    "searchText",

                    "heading"

                ],

                tolerance:
                    tolerance,

                limit:
                    limit

            }
        );


    // ==================================
    // إذا لم توجد نتائج:
    // محاولة مطابقة أدق
    // ==================================

    if (
        (
            !oramaResult ||
            !oramaResult.hits ||
            oramaResult.hits.length ===
                0
        ) &&
        normalizedQuery.length >=
            3
    ) {

        oramaResult =
            search(
                db,
                {

                    term:
                        normalizedQuery,

                    properties: [

                        "text",

                        "searchText",

                        "heading"

                    ],

                    exact:
                        true,

                    limit:
                        limit

                }
            );

    }


    const hits =
        oramaResult &&
        Array.isArray(
            oramaResult.hits
        )
            ? oramaResult.hits
            : [];


    // ==================================
    // نتائج العناوين
    // ==================================

    const headingResults =
        await searchOramaHeadings(
            documentItem,
            normalizedQuery
        );


    const headingMap =
        new Map();


    headingResults.forEach(
        function (
            heading
        ) {

            headingMap.set(
                Number(
                    heading.headingIndex
                ),
                heading
            );

        }
    );


    // ==================================
    // الفقرات
    // ==================================

    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    const paragraphMap =
        new Map();


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                paragraph &&
                typeof paragraph.index !==
                    "undefined"
            ) {

                paragraphMap.set(
                    Number(
                        paragraph.index
                    ),
                    paragraph
                );

            }

        }
    );


    // ==================================
    // تحويل نتائج Orama
    // ==================================

    const results =
        [];


    hits.forEach(
        function (
            hit
        ) {

            if (
                !hit ||
                !hit.document
            ) {

                return;

            }


            const doc =
                hit.document;


            const paragraphIndex =
                Number(
                    doc.paragraphIndex
                );


            const paragraph =
                paragraphMap.get(
                    paragraphIndex
                );


            if (!paragraph) {

                return;

            }


            const originalText =
                String(
                    paragraph.text ||
                    doc.text ||
                    ""
                );


            if (!originalText.trim()) {

                return;

            }


            const normalizedText =
                normalizeSearchText(
                    originalText
                );


            const heading =
                String(
                    doc.heading ||
                    ""
                ).trim();


            const normalizedHeading =
                normalizeSearchText(
                    heading
                );


            let score =
                Number(
                    hit.score ||
                    0
                );


            // ==================================
            // المطابقة الكاملة
            // ==================================

            if (
                normalizedText.includes(
                    normalizedQuery
                )
            ) {

                score +=
                    12;

            }


            // ==================================
            // مطابقة بداية الفقرة
            // ==================================

            if (
                normalizedText.startsWith(
                    normalizedQuery
                )
            ) {

                score +=
                    2;

            }


            // ==================================
            // مطابقة العنوان
            // ==================================

            const headingData =
                headingMap.get(
                    Number(
                        doc.headingIndex
                    )
                );


            let headingMatch =
                false;


            let headingScore =
                0;


            let headingCoverage =
                0;


            let matchedHeadingTokens =
                0;


            if (
                headingData
            ) {

                headingScore =
                    Number(
                        headingData.score ||
                        0
                    );


                headingCoverage =
                    Number(
                        headingData.coverage ||
                        0
                    );


                matchedHeadingTokens =
                    Number(
                        headingData.matchedTokens ||
                        0
                    );


            }


            if (
                normalizedHeading ===
                normalizedQuery
            ) {

                headingMatch =
                    true;

                score +=
                    30;

            }
            else if (
                normalizedHeading.includes(
                    normalizedQuery
                )
            ) {

                headingMatch =
                    true;

                score +=
                    18;

            }


            if (
                headingData
            ) {

                score +=
                    headingScore;

                score +=
                    headingCoverage *
                    10;

            }


            // ==================================
            // كلمات السؤال
            // ==================================

            const queryTokens =
                getSearchQueryTokens(
                    normalizedQuery
                );


            const meaningfulTokens =
                queryTokens.filter(
                    function (
                        token
                    ) {

                        return (
                            token &&
                            token.length >=
                                3
                        );

                    }
                );


            let exactWordMatches =
                0;


            meaningfulTokens.forEach(
                function (
                    token
                ) {

                    if (
                        normalizedText.includes(
                            token
                        )
                    ) {

                        exactWordMatches +=
                            1;

                    }

                }
            );


            score +=
                exactWordMatches *
                4;


            // ==================================
            // تغطية السؤال
            // ==================================

            const queryCoverage =
                meaningfulTokens.length >
                0
                    ? exactWordMatches /
                      meaningfulTokens.length
                    : 0;


            score +=
                queryCoverage *
                8;


            // ==================================
            // العائلة الصرفية داخل الفقرة
            // ==================================

            const matchedParagraphFamilies =
                [];


            const paragraphTokens =
                tokenizeDocumentText(
                    originalText
                );


            const paragraphFamilies =
                new Set();


            paragraphTokens.forEach(
                function (
                    token
                ) {

                    const family =
                        getConservativeFamilyKey(
                            token,
                            null
                        );


                    if (family) {

                        paragraphFamilies.add(
                            family
                        );

                    }

                }
            );


            const queryFamilies =
                [];


            meaningfulTokens.forEach(
                function (
                    token
                ) {

                    const family =
                        getConservativeFamilyKey(
                            token,
                            null
                        );


                    if (
                        family &&
                        !queryFamilies.includes(
                            family
                        )
                    ) {

                        queryFamilies.push(
                            family
                        );

                    }

                }
            );


            queryFamilies.forEach(
                function (
                    family
                ) {

                    if (
                        paragraphFamilies.has(
                            family
                        )
                    ) {

                        matchedFamilies.push(
                            family
                        );

                    }

                }
            );


            score +=
                matchedFamilies.length *
                4;


            // ==================================
            // كثافة العائلات
            // ==================================

            let familyOccurrencesInParagraph =
                0;


            paragraphTokens.forEach(
                function (
                    token
                ) {

                    const family =
                        getConservativeFamilyKey(
                            token,
                            null
                        );


                    if (
                        matchedFamilies.includes(
                            family
                        )
                    ) {

                        familyOccurrencesInParagraph +=
                            1;

                    }

                }
            );


            score +=
                Math.min(
                    familyOccurrencesInParagraph,
                    8
                ) *
                0.75;


            // ==================================
            // نوع المطابقة
            // ==================================

            let matchType =
                "family";


            if (
                headingMatch
            ) {

                matchType =
                    "heading";

            }
            else if (
                normalizedText.includes(
                    normalizedQuery
                )
            ) {

                matchType =
                    "exact";

            }
            else if (
                exactWordMatches >
                    0
            ) {

                matchType =
                    "word";

            }


            // ==================================
            // موضع المطابقة
            // ==================================

            let searchPosition =
                normalizedText.indexOf(
                    normalizedQuery
                );


            if (
                searchPosition ===
                -1
            ) {

                for (
                    let i = 0;
                    i < meaningfulTokens.length;
                    i++
                ) {

                    const position =
                        normalizedText.indexOf(
                            meaningfulTokens[i]
                        );


                    if (
                        position !==
                        -1
                    ) {

                        searchPosition =
                            position;

                        break;

                    }

                }

            }


            // ==================================
            // بناء السياق
            // ==================================

            let context =
                originalText;


            if (
                searchPosition !==
                    -1
            ) {

                const contextStart =
                    Math.max(
                        0,
                        searchPosition -
                        120
                    );


                const contextEnd =
                    Math.min(
                        normalizedText.length,
                        searchPosition +
                        normalizedQuery.length +
                        300
                    );


                context =
                    normalizedText.substring(
                        contextStart,
                        contextEnd
                    );

            }
            else {

                context =
                    originalText.substring(
                        0,
                        420
                    );

            }


            // ==================================
            // النتيجة
            // ==================================

            results.push({

                paragraphIndex:
                    paragraphIndex,

                paragraphId:
                    String(
                        doc.paragraphId ||
                        paragraph.id ||
                        paragraphIndex
                    ),

                text:
                    originalText,

                context:
                    context,

                heading:
                    heading,

                headingLevel:
                    doc.headingLevel ||
                    "",

                matchedTerms:
                    meaningfulTokens.filter(
                        function (
                            token
                        ) {

                            return normalizedText.includes(
                                token
                            );

                        }
                    ),

                matchedFamilies:
                    matchedFamilies,

                matchedFamilyCount:
                    matchedFamilies.length,

                familyOccurrencesInParagraph:
                    familyOccurrencesInParagraph,

                exactWordMatches:
                    exactWordMatches,

                totalQueryTerms:
                    meaningfulTokens.length,

                headingMatch:
                    headingMatch,

                headingScore:
                    headingScore,

                headingCoverage:
                    headingCoverage,

                score:
                    score,

                matchType:
                    matchType

            });

        }
    );


    // =================================================
    // إذا كانت المطابقة القوية في العنوان
    // لم تظهر ضمن hits الرئيسية
    // نضيف العنوان وفقراته الأولى
    // =================================================

    headingResults
        .filter(
            function (
                heading
            ) {

                return (
                    Number(
                        heading.score
                    ) >
                    10
                );

            }
        )
        .forEach(
            function (
                heading
            ) {

                const range =
                    getHeadingParagraphRange(
                        structureData,
                        heading.headingIndex
                    );


                range
                    .slice(
                        0,
                        3
                    )
                    .forEach(
                        function (
                            paragraph
                        ) {

                            if (
                                !paragraph ||
                                !String(
                                    paragraph.text ||
                                    ""
                                ).trim()
                            ) {

                                return;

                            }


                            const paragraphIndex =
                                Number(
                                    paragraph.index
                                );


                            const exists =
                                results.some(
                                    function (
                                        item
                                    ) {

                                        return (
                                            Number(
                                                item.paragraphIndex
                                            ) ===
                                            paragraphIndex
                                        );

                                    }
                                );


                            if (exists) {

                                return;

                            }


                            results.push({

                                paragraphIndex:
                                    paragraphIndex,

                                paragraphId:
                                    String(
                                        paragraph.id ||
                                        paragraphIndex
                                    ),

                                text:
                                    String(
                                        paragraph.text ||
                                        ""
                                    ),

                                context:
                                    String(
                                        paragraph.text ||
                                        ""
                                    )
                                        .substring(
                                            0,
                                            420
                                        ),

                                heading:
                                    heading.heading,

                                headingLevel:
                                    heading.style,

                                matchedTerms:
                                    [],

                                matchedFamilies:
                                    [],

                                matchedFamilyCount:
                                    0,

                                familyOccurrencesInParagraph:
                                    0,

                                exactWordMatches:
                                    0,

                                totalQueryTerms:
                                    meaningfulTokensCount(
                                        normalizedQuery
                                    ),

                                headingMatch:
                                    true,

                                headingScore:
                                    Number(
                                        heading.score ||
                                        0
                                    ),

                                headingCoverage:
                                    Number(
                                        heading.coverage ||
                                        0
                                    ),

                                score:
                                    Number(
                                        heading.score ||
                                        0
                                    ) +
                                    20,

                                matchType:
                                    "heading"

                            });

                        }
                    );

            }
        );


    // ==================================
    // ترتيب النتائج
    // ==================================

    results.sort(
        function (
            a,
            b
        ) {

            if (
                b.score !==
                a.score
            ) {

                return (
                    b.score -
                    a.score
                );

            }


            if (
                Boolean(
                    b.headingMatch
                ) !==
                Boolean(
                    a.headingMatch
                )
            ) {

                return (
                    b.headingMatch
                        ? 1
                        : -1
                );

            }


            if (
                b.matchedFamilyCount !==
                a.matchedFamilyCount
            ) {

                return (
                    b.matchedFamilyCount -
                    a.matchedFamilyCount
                );

            }


            if (
                b.exactWordMatches !==
                a.exactWordMatches
            ) {

                return (
                    b.exactWordMatches -
                    a.exactWordMatches
                );

            }


            return (
                a.paragraphIndex -
                b.paragraphIndex
            );

        }
    );

    // ==================================
    // حد المرشحين الأولي
    // ==================================

    const candidateLimit =
        retrievalProfile === "comparison"
            ? 30
            : retrievalProfile === "definition"
                ? 20
                : 25;

    results.splice(
        candidateLimit
    );

    return results;

}


// =====================================================
// Meaningful Tokens Count
// =====================================================

function meaningfulTokensCount(
    query
) {

    return getSearchQueryTokens(
        query
    ).filter(
        function (
            token
        ) {

            return (
                token &&
                token.length >=
                    3
            );

        }
    ).length;

}


// =====================================================
// Get Heading Paragraph Range
// الحصول على فقرات المطلب
// =====================================================

function getHeadingParagraphRange(
    structureData,
    headingIndex
) {

    if (
        !structureData ||
        !Array.isArray(
            structureData.paragraphs
        ) ||
        !Array.isArray(
            structureData.headings
        )
    ) {

        return [];

    }


    const paragraphs =
        structureData.paragraphs;


    const headings =
        structureData.headings
            .filter(
                function (
                    heading
                ) {

                    return (
                        heading &&
                        typeof heading.index !==
                            "undefined"
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        Number(
                            a.index
                        ) -
                        Number(
                            b.index
                        )
                    );

                }
            );


    const currentHeadingPosition =
        headings.findIndex(
            function (
                heading
            ) {

                return (
                    Number(
                        heading.index
                    ) ===
                    Number(
                        headingIndex
                    )
                );

            }
        );


    if (
        currentHeadingPosition ===
        -1
    ) {

        return [];

    }


    const startIndex =
        Number(
            headings[
                currentHeadingPosition
            ].index
        );


    const nextHeading =
        headings[
            currentHeadingPosition + 1
        ];


    const endIndex =
        nextHeading
            ? Number(
                nextHeading.index
            ) - 1
            : paragraphs.length - 1;


    return paragraphs.filter(
        function (
            paragraph
        ) {

            const index =
                Number(
                    paragraph.index
                );


            return (
                index >
                startIndex &&
                index <=
                endIndex
            );

        }
    );

}


// =====================================================
// Analyze Search Query
// المحلل العام للاستعلام
//
// الهدف:
// 1) فصل أدوات السؤال عن مفاهيمه.
// 2) استخراج المفهوم الأساسي من الكلمات الملصقة.
// 3) الاحتفاظ بالصيغ الأصلية والبديلة للبحث.
// 4) اكتشاف العائلات الصرفية.
// 5) اكتشاف العبارات المركبة الموجودة فعليًا.
// 6) حساب وزن المفهوم بحسب وجوده في المستند.
//
// لا يحتوي على قاموس فقهي أو أصولي.
// =====================================================

function analyzeSearchQuery(
    query,
    documentData
) {

    // ==================================
    // السؤال الأصلي
    // ==================================

    const originalQuery =
        String(
            query ||
            ""
        ).trim();


    // ==================================
    // نتيجة فارغة
    // ==================================

    if (
        !originalQuery
    ) {

        return {

            query:
                "",

            normalizedQuery:
                "",

            allTokens:
                [],

            contentTokens:
                [],

            functionTokens:
                [],

            phraseCandidates:
                [],

            families:
                [],

            weightedTerms:
                [],

            contentCount:
                0,

            functionCount:
                0

        };

    }


    // ==================================
    // تطبيع السؤال
    // ==================================

    const normalizedQuery =
        normalizeSearchText(
            originalQuery
        );


    // ==================================
    // تقسيم السؤال
    // ==================================

    const allTokens =
        tokenizeDocumentText(
            normalizedQuery
        );


    // ==================================
    // الكلمات الوظيفية العامة
    // ==================================

    const commonFunctionWords =
        new Set([

            "ما",
            "ماذا",
            "من",
            "هل",
            "كيف",
            "لماذا",
            "أين",
            "متى",
            "أي",
            "اي",

            "في",
            "على",
            "عن",
            "الى",
            "إلى",
            "من",
            "مع",
            "بين",

            "ثم",
            "أو",
            "او",

            "و",
            "ف",
            "ب",
            "ك",
            "ل",

            "أن",
            "ان",
            "إن",

            "حتى",
            "قد",
            "لقد",
            "لم",
            "لن",
            "لا",

            "ليس",
            "ليست",

            "هذا",
            "هذه",
            "ذلك",
            "تلك",

            "الذي",
            "التي",
            "الذين",
            "اللاتي",

            "به",
            "بها",
            "له",
            "لها",
            "منه",
            "منها",
            "عنه",
            "عنها"

        ]);


    // ==================================
    // كلمات توجه السؤال
    //
    // لا نحذفها من السؤال،
    // لكن لا نجعلها مفاهيم موضوعية.
    // ==================================

    const queryInstructionWords =
        new Set([

            "المقصود",
            "مقصود",

            "معنى",

            "تعريف",
            "تعريفه",
            "تعريفها",

            "المراد",
            "مراد",

            "يقصد",

            "تفسير",
            "شرح",

            "أثر",
            "اثر",

            "تأثير",
            "تاثير",

            "نتيجة",
            "نتائج",

            "حكم",
            "أحكام",

            "حجية",

            "سبب",
            "أسباب",
            "اسباب",

            "علة",
            "علل",

            "شرط",
            "شروط",

            "فرق",
            "الفرق",

            "فروق",
            "الفروق",

            "مقارنة",
            "مقارن",

            "علاقة",
            "العلاقة",

            "صلة",
            "صلته",
            "صلتها",

            "كيفية",

            "موضع",
            "موضعه",
            "موضعها"

        ]);


    // ==================================
    // النص الكامل للمستند
    // ==================================

    let documentText =
        "";


    if (
        documentData
    ) {

        if (
            Array.isArray(
                documentData.paragraphs
            )
        ) {

            documentText =
                documentData.paragraphs
                    .map(
                        function (
                            paragraph
                        ) {

                            return (
                                paragraph &&
                                paragraph.text
                                    ? paragraph.text
                                    : ""
                            );

                        }
                    )
                    .join(
                        "\n"
                    );

        }
        else if (
            typeof documentData.text ===
            "string"
        ) {

            documentText =
                documentData.text;

        }

    }


    const normalizedDocumentText =
        normalizeSearchText(
            documentText
        );


    // ==================================
    // توكنات المستند
    // ==================================

    const documentTokens =
        normalizedDocumentText
            ? tokenizeDocumentText(
                normalizedDocumentText
            )
            : [];


    // ==================================
    // تكرار الكلمات في المستند
    // ==================================

    const termFrequency =
        new Map();


    documentTokens.forEach(
        function (
            token
        ) {

            const cleanToken =
                String(
                    token ||
                    ""
                ).trim();


            if (
                !cleanToken
            ) {

                return;

            }


            termFrequency.set(
                cleanToken,
                (
                    termFrequency.get(
                        cleanToken
                    ) ||
                    0
                ) +
                1
            );

        }
    );


    // ==================================
    // استخراج صيغ الكلمة
    //
    // أمثلة:
    //
    // بالاستصلاح
    // → بالاستصلاح
    // → الاستصلاح
    //
    // والاستحسان
    // → والاستحسان
    // → الاستحسان
    //
    // وبالاستصلاح
    // → وبالاستصلاح
    // → بالاستصلاح
    // → الاستصلاح
    //
    // للاستمصال / للاجتهاد ...
    // نحاول الحفاظ على "ال" قدر الإمكان.
    // ==================================

    function getSearchTermVariants(
        token
    ) {

        const variants =
            [];


        const originalToken =
            String(
                token ||
                ""
            ).trim();


        if (
            !originalToken
        ) {

            return variants;

        }


        function addVariant(
            value
        ) {

            const clean =
                String(
                    value ||
                    ""
                ).trim();


            if (
                !clean
            ) {

                return;

            }


            if (
                clean.length <
                3
            ) {

                return;

            }


            if (
                !variants.includes(
                    clean
                )
            ) {

                variants.push(
                    clean
                );

            }

        }


        addVariant(
            originalToken
        );


        let current =
            originalToken;


        // ==================================
        // التركيبات الشائعة مع ال التعريف
        // ==================================

        const attachedDefinitePatterns = [

            "وال",
            "فال",
            "بال",
            "كال",
            "لل"

        ];


        attachedDefinitePatterns.forEach(
            function (
                prefix
            ) {

                if (
                    current.startsWith(
                        prefix
                    ) &&
                    current.length >
                    prefix.length +
                    1
                ) {

                    const remainder =
                        current.substring(
                            prefix.length
                        );


                    // ----------------------------------
                    // مثل:
                    // والاستحسان
                    // بالاستصلاح
                    // كالقياس
                    // فالاجتهاد
                    // للاجتهاد
                    //
                    // نعيد "ال" إلى أصل الكلمة.
                    // ----------------------------------

                    const baseWithArticle =
                        "ال" +
                        remainder;


                    addVariant(
                        baseWithArticle
                    );

                }

            }
        );


        // ==================================
        // نزع الواو والفاء والباء والكاف واللام
        // إذا بقيت كلمة صالحة.
        // ==================================

        let working =
            current;


        let changed =
            true;


        let safety =
            0;


        while (
            changed &&
            safety <
            4
        ) {

            changed =
                false;


            safety +=
                1;


            const firstChar =
                working.charAt(
                    0
                );


            if (
                (
                    firstChar ===
                    "و" ||

                    firstChar ===
                    "ف" ||

                    firstChar ===
                    "ب" ||

                    firstChar ===
                    "ك" ||

                    firstChar ===
                    "ل"
                ) &&
                working.length >
                3
            ) {

                const next =
                    working.substring(
                        1
                    );


                if (
                    next.length >=
                    3
                ) {

                    working =
                        next;


                    addVariant(
                        working
                    );


                    changed =
                        true;

                }

            }

        }


        // ==================================
        // إذا وصلنا إلى صيغة تبدأ بـ "ا"
        // وحُذفت اللام بسبب "لل..."
        // نعيد احتمال "ال..."
        // ==================================

        if (
            !variants.some(
                function (
                    value
                ) {

                    return value.startsWith(
                        "ال"
                    );

                }
            )
        ) {

            const candidateWithoutFirstPrefix =
                originalToken.substring(
                    1
                );


            if (
                candidateWithoutFirstPrefix.startsWith(
                    "ال"
                )
            ) {

                addVariant(
                    candidateWithoutFirstPrefix
                );

            }

        }


        return variants;

    }


    // ==================================
    // تحديد المفهوم الأساسي
    //
    // نريد:
    //
    // بالاستصلاح
    // → الاستصلاح
    //
    // والاستحسان
    // → الاستحسان
    //
    // والقياس
    // → القياس
    // ==================================

    function getBaseContentTerm(
        token,
        variants
    ) {

        if (
            Array.isArray(
                variants
            )
        ) {

            // ----------------------------------
            // الأفضلية لصيغة تبدأ بـ "ال"
            // لأنها غالبًا الصيغة الأساسية
            // ----------------------------------

            for (
                let i =
                    0;

                i <
                    variants.length;

                i++
            ) {

                const variant =
                    variants[i];


                if (
                    variant &&
                    variant.startsWith(
                        "ال"
                    ) &&
                    variant.length >
                    3
                ) {

                    return variant;

                }

            }


            // ----------------------------------
            // إذا لم نجد صيغة بالـ
            // نستخدم آخر صيغة مجردة
            // ----------------------------------

            if (
                variants.length >
                1
            ) {

                return variants[
                    variants.length - 1
                ];

            }

        }


        return String(
            token ||
            ""
        ).trim();

    }


    // ==================================
    // تصنيف كلمات السؤال
    // ==================================

    const functionTokens =
        [];


    const rawContentTerms =
        [];


    allTokens.forEach(
        function (
            token
        ) {

            const cleanToken =
                String(
                    token ||
                    ""
                ).trim();


            if (
                !cleanToken
            ) {

                return;

            }


            if (
                commonFunctionWords.has(
                    cleanToken
                )
            ) {

                functionTokens.push({

                    term:
                        cleanToken,

                    weight:
                        0.10,

                    category:
                        "function",

                    variants:
                        getSearchTermVariants(
                            cleanToken
                        )

                });

                return;

            }


            if (
                queryInstructionWords.has(
                    cleanToken
                )
            ) {

                functionTokens.push({

                    term:
                        cleanToken,

                    weight:
                        0.35,

                    category:
                        "instruction",

                    variants:
                        getSearchTermVariants(
                            cleanToken
                        )

                });

                return;

            }


            const variants =
                getSearchTermVariants(
                    cleanToken
                );


            const baseTerm =
                getBaseContentTerm(
                    cleanToken,
                    variants
                );


            rawContentTerms.push({

                original:
                    cleanToken,

                base:
                    baseTerm,

                variants:
                    variants

            });

        }
    );


    // ==================================
    // إزالة التكرار بعد استخراج الأصل
    // ==================================

    const contentTokens =
        [];


    rawContentTerms.forEach(
        function (
            item
        ) {

            if (
                !item ||
                !item.base
            ) {

                return;

            }


            if (
                !contentTokens.includes(
                    item.base
                )
            ) {

                contentTokens.push(
                    item.base
                );

            }

        }
    );


    // ==================================
    // بناء lookup للمصطلحات
    // ==================================

    const contentTermMap =
        new Map();


    rawContentTerms.forEach(
        function (
            item
        ) {

            if (
                !item ||
                !item.base
            ) {

                return;

            }


            const existing =
                contentTermMap.get(
                    item.base
                );


            if (
                !existing
            ) {

                contentTermMap.set(
                    item.base,
                    {

                        term:
                            item.base,

                        variants:
                            Array.isArray(
                                item.variants
                            )
                                ? item.variants.slice()
                                : []

                    }
                );

            }
            else {

                item.variants.forEach(
                    function (
                        variant
                    ) {

                        if (
                            !existing.variants.includes(
                                variant
                            )
                        ) {

                            existing.variants.push(
                                variant
                            );

                        }

                    }
                );

            }

        }
    );


    // ==================================
    // تكرار العائلات
    // ==================================

    const familyFrequency =
        new Map();


    documentTokens.forEach(
        function (
            token
        ) {

            const family =
                getConservativeFamilyKey(
                    token,
                    null
                );


            if (
                !family
            ) {

                return;

            }


            familyFrequency.set(
                family,
                (
                    familyFrequency.get(
                        family
                    ) ||
                    0
                ) +
                1
            );

        }
    );


    // ==================================
    // العائلات الموجودة في السؤال
    // ==================================

    const families =
        [];


    contentTokens.forEach(
        function (
            baseTerm
        ) {

            const item =
                contentTermMap.get(
                    baseTerm
                );


            if (
                !item
            ) {

                return;

            }


            let family =
                getConservativeFamilyKey(
                    baseTerm,
                    null
                );


            // ----------------------------------
            // تجربة الصيغ إن لم نحصل على عائلة
            // ----------------------------------

            if (
                !family
            ) {

                item.variants.forEach(
                    function (
                        variant
                    ) {

                        if (
                            family
                        ) {

                            return;

                        }


                        family =
                            getConservativeFamilyKey(
                                variant,
                                null
                            );

                    }
                );

            }


            if (
                family &&
                !families.includes(
                    family
                )
            ) {

                families.push(
                    family
                );

            }

        }
    );


    // ==================================
    // المصطلحات الموزونة
    // ==================================

    const weightedTerms =
        [];


    contentTokens.forEach(
        function (
            baseTerm
        ) {

            const item =
                contentTermMap.get(
                    baseTerm
                );


            if (
                !item
            ) {

                return;

            }


            const variants =
                Array.isArray(
                    item.variants
                ) &&
                item.variants.length >
                0

                    ? item.variants

                    : [
                        baseTerm
                    ];


            // ----------------------------------
            // تكرار المفهوم الأساسي
            // ----------------------------------

            const exactFrequency =
                termFrequency.get(
                    baseTerm
                ) ||
                0;


            // ----------------------------------
            // أعلى تكرار للصيغ
            // ----------------------------------

            let variantFrequency =
                0;


            variants.forEach(
                function (
                    variant
                ) {

                    const frequency =
                        termFrequency.get(
                            variant
                        ) ||
                        0;


                    if (
                        frequency >
                        variantFrequency
                    ) {

                        variantFrequency =
                            frequency;

                    }

                }
            );


            // ----------------------------------
            // العائلة
            // ----------------------------------

            let family =
                getConservativeFamilyKey(
                    baseTerm,
                    null
                );


            if (
                !family
            ) {

                variants.forEach(
                    function (
                        variant
                    ) {

                        if (
                            family
                        ) {

                            return;

                        }


                        family =
                            getConservativeFamilyKey(
                                variant,
                                null
                            );

                    }
                );

            }


            const familyFreq =
                familyFrequency.get(
                    family
                ) ||
                0;


            // ----------------------------------
            // وجود المفهوم في المستند
            // ----------------------------------

            let documentPresence =
                0;


            if (
                exactFrequency >
                0
            ) {

                documentPresence =
                    1;

            }
            else if (
                variantFrequency >
                0
            ) {

                documentPresence =
                    0.90;

            }
            else if (
                familyFreq >
                0
            ) {

                documentPresence =
                    0.75;

            }
            else {

                documentPresence =
                    0.35;

            }


            // ----------------------------------
            // وزن التكرار
            // ----------------------------------

            const strongestFrequency =
                Math.max(
                    exactFrequency,
                    variantFrequency,
                    familyFreq
                );


            let frequencyWeight =
                0.60;


            if (
                strongestFrequency >
                0
            ) {

                frequencyWeight =
                    Math.min(
                        1.25,
                        0.75 +
                        Math.log10(
                            strongestFrequency +
                            1
                        ) *
                        0.35
                    );

            }


            const finalWeight =
                documentPresence *
                frequencyWeight;


            weightedTerms.push({

                // المفهوم الأساسي
                term:
                    baseTerm,

                // الأصل الذي ظهر في السؤال
                originalTerms:
                    rawContentTerms
                        .filter(
                            function (
                                rawItem
                            ) {

                                return (
                                    rawItem.base ===
                                    baseTerm
                                );

                            }
                        )
                        .map(
                            function (
                                rawItem
                            ) {

                                return rawItem.original;

                            }
                        ),

                // صيغ البحث
                variants:
                    variants,

                family:
                    family ||
                    "",

                exactFrequency:
                    exactFrequency,

                variantFrequency:
                    variantFrequency,

                familyFrequency:
                    familyFreq,

                documentPresence:
                    documentPresence,

                weight:
                    Number(
                        finalWeight.toFixed(
                            4
                        )
                    ),

                category:
                    "content"

            });

        }
    );


    // ==================================
    // اكتشاف العبارات المركبة
    //
    // مهم:
    // لا نخترع عبارة مركبة.
    // لا نسجلها إلا إذا كانت موجودة
    // فعليًا في المستند.
    // ==================================

    const phraseCandidates =
        [];


    // ----------------------------------
    // ترتيب المفاهيم كما وردت في السؤال
    // ----------------------------------

    const phraseTerms =
        contentTokens.slice();


    // ==================================
    // البحث عن عبارات من 3 كلمات
    // ثم كلمتين
    // ==================================

    for (
        let phraseLength = 3;

        phraseLength >= 2;

        phraseLength--
    ) {

        if (
            phraseTerms.length <
            phraseLength
        ) {

            continue;

        }


        for (
            let i = 0;

            i +
                phraseLength <=
                phraseTerms.length;

            i++
        ) {

            const sequence =
                phraseTerms.slice(
                    i,
                    i +
                    phraseLength
                );


            if (
                sequence.length <
                2
            ) {

                continue;

            }


            // ==================================
            // تجربة الصيغ الأساسية
            // ==================================

            const candidatePhrase =
                sequence.join(
                    " "
                );


            const normalizedCandidate =
                normalizeSearchText(
                    candidatePhrase
                );


            if (
                normalizedCandidate &&
                normalizedDocumentText.includes(
                    normalizedCandidate
                )
            ) {

                if (
                    !phraseCandidates.some(
                        function (
                            item
                        ) {

                            return (
                                item.phrase ===
                                normalizedCandidate
                            );

                        }
                    )
                ) {

                    phraseCandidates.push({

                        phrase:
                            normalizedCandidate,

                        tokens:
                            sequence,

                        length:
                            sequence.length,

                        weight:
                            phraseLength ===
                            3
                                ? 1.20
                                : 1.00,

                        source:
                            "base"

                    });

                }

            }


            // ==================================
            // تجربة الصيغ البديلة
            // لكل مفهوم
            // ==================================

            const variantsSets =
                sequence.map(
                    function (
                        term
                    ) {

                        const item =
                            contentTermMap.get(
                                term
                            );


                        return (
                            item &&
                            Array.isArray(
                                item.variants
                            ) &&
                            item.variants.length >
                                0

                                ? item.variants

                                : [
                                    term
                                ]
                        );

                    }
                );


            // ==================================
            // نأخذ الصيغة الأقرب إلى الأصل
            // التي تبدأ بـ "ال" إن وجدت.
            // ==================================

            const normalizedSequence =
                variantsSets.map(
                    function (
                        variants
                    ) {

                        for (
                            let j = 0;

                            j <
                                variants.length;

                            j++
                        ) {

                            if (
                                variants[j] &&
                                variants[j].startsWith(
                                    "ال"
                                )
                            ) {

                                return variants[j];

                            }

                        }


                        return variants[
                            variants.length - 1
                        ];

                    }
                );


            const normalizedPhrase =
                normalizeSearchText(
                    normalizedSequence.join(
                        " "
                    )
                );


            if (
                normalizedPhrase &&
                normalizedPhrase !==
                    normalizedCandidate &&
                normalizedDocumentText.includes(
                    normalizedPhrase
                )
            ) {

                if (
                    !phraseCandidates.some(
                        function (
                            item
                        ) {

                            return (
                                item.phrase ===
                                normalizedPhrase
                            );

                        }
                    )
                ) {

                    phraseCandidates.push({

                        phrase:
                            normalizedPhrase,

                        tokens:
                            sequence,

                        length:
                            sequence.length,

                        weight:
                            phraseLength ===
                            3
                                ? 1.20
                                : 1.00,

                        source:
                            "normalized"

                    });

                }

            }

        }

    }


    // ==================================
    // ترتيب المصطلحات
    // ==================================

    weightedTerms.sort(
        function (
            a,
            b
        ) {

            if (
                b.weight !==
                a.weight
            ) {

                return (
                    b.weight -
                    a.weight
                );

            }


            if (
                b.familyFrequency !==
                a.familyFrequency
            ) {

                return (
                    b.familyFrequency -
                    a.familyFrequency
                );

            }


            if (
                b.variantFrequency !==
                a.variantFrequency
            ) {

                return (
                    b.variantFrequency -
                    a.variantFrequency
                );

            }


            return (
                b.term.length -
                a.term.length
            );

        }
    );


    // ==================================
    // ترتيب العبارات
    // ==================================

    phraseCandidates.sort(
        function (
            a,
            b
        ) {

            if (
                b.weight !==
                a.weight
            ) {

                return (
                    b.weight -
                    a.weight
                );

            }


            return (
                b.length -
                a.length
            );

        }
    );


    // ==================================
    // النتيجة
    // ==================================

    return {

        query:
            originalQuery,

        normalizedQuery:
            normalizedQuery,

        allTokens:
            allTokens,

        // الآن هذه هي المفاهيم الأساسية
        // وليس الصيغ الملحقة
        contentTokens:
            contentTokens,

        functionTokens:
            functionTokens,

        phraseCandidates:
            phraseCandidates,

        families:
            families,

        weightedTerms:
            weightedTerms,

        contentCount:
            contentTokens.length,

        functionCount:
            functionTokens.length

    };

}

window.testAnalyzeSearchQuery =
    async function (
        query
    ) {

        if (
            !currentDocument
        ) {

            console.warn(
                "لا يوجد مستند نشط."
            );

            return null;

        }


        try {

            const structureData =
                await getDocumentStructure(
                    currentDocument.id
                );


            const result =
                analyzeSearchQuery(
                    query,
                    structureData
                );


            console.log(
                "تحليل الاستعلام:",
                query
            );


            console.log(
                result
            );


            return result;

        }
        catch (
            error
        ) {

            console.error(
                "فشل تحليل الاستعلام:",
                error
            );


            return null;

        }

    };
// =====================================================
// Search Indexed Document
// المحرك العام للاسترجاع
//
// يعتمد على:
// 1) analyzeSearchQuery()
// 2) Orama
// 3) QPS داخل فهرس Orama
// 4) الصيغ العربية التي استخرجها المحلل
// 5) العائلات الصرفية
// 6) العبارات المركبة
// 7) تغطية مفاهيم السؤال
// 8) أولوية العنوان
//
// الهدف:
// استخراج أفضل المرشحين للذكاء الاصطناعي.
// لا يحاول المحرك الإجابة عن السؤال.
// =====================================================

async function searchIndexedDocument(
    documentId,
    query,
    options
) {

    // ==================================
    // الإعدادات
    // ==================================

    const settings =
        options || {};


    // ==================================
    // نوع السؤال
    // يستخدم كعامل ترجيح خفيف فقط
    // ==================================

    const profileInfo =
        getRetrievalProfile(
            query
        );


    const retrievalProfile =
        settings.profile ||
        (
            profileInfo &&
            profileInfo.type
                ? profileInfo.type
                : "general"
        );


    // ==================================
    // تطبيع السؤال
    // ==================================

    const searchTerm =
        normalizeSearchText(
            query
        );


    if (
        !searchTerm
    ) {

        return {

            query:
                "",

            profile:
                retrievalProfile,

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            phraseCandidates:
                [],

            queryAnalysis:
                null,

            totalQueryTerms:
                0,

            contentTermCount:
                0,

            functionTermCount:
                0,

            indexTokenCount:
                0,

            indexUniqueTerms:
                0,

            indexUniqueFamilies:
                0,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // العثور على المستند
    // ==================================

    const documentItem =
        documents.find(
            function (
                doc
            ) {

                return (
                    doc &&
                    String(
                        doc.id
                    ) ===
                    String(
                        documentId
                    )
                );

            }
        ) ||
        currentDocument;


    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم العثور على المستند."
        );

    }


    // ==================================
    // بنية المستند
    // ==================================

    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData
    ) {

        throw new Error(
            "لا توجد بنية محفوظة لهذا المستند."
        );

    }


    // ==================================
    // تحليل الاستعلام
    // ==================================

    const queryAnalysis =
        analyzeSearchQuery(
            searchTerm,
            structureData
        );


    // ==================================
    // حماية مخرجات المحلل
    // ==================================

    const queryTokens =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.allTokens
        )
            ? queryAnalysis.allTokens
            : [];


    const contentTokens =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.contentTokens
        )
            ? queryAnalysis.contentTokens
            : [];


    const functionTokens =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.functionTokens
        )
            ? queryAnalysis.functionTokens
            : [];


    const phraseCandidates =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.phraseCandidates
        )
            ? queryAnalysis.phraseCandidates
            : [];


    const weightedTerms =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.weightedTerms
        )
            ? queryAnalysis.weightedTerms
            : [];


    const queryFamilies =
        Array.isArray(
            queryAnalysis &&
            queryAnalysis.families
        )
            ? queryAnalysis.families
            : [];


    // ==================================
    // سجل التحليل
    // ==================================

    console.log(
        "======================================"
    );

    console.log(
        "تحليل الاستعلام:",
        searchTerm
    );

    console.log(
        "كل الكلمات:",
        queryTokens
    );

    console.log(
        "كلمات المحتوى:",
        contentTokens
    );

    console.log(
        "الكلمات الوظيفية:",
        functionTokens
    );

    console.log(
        "العبارات المركبة:",
        phraseCandidates
    );

    console.log(
        "العائلات:",
        queryFamilies
    );

    console.log(
        "المصطلحات الموزونة:",
        weightedTerms
    );

    console.log(
        "======================================"
    );


    if (
        queryTokens.length ===
        0
    ) {

        return {

            query:
                searchTerm,

            profile:
                retrievalProfile,

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            phraseCandidates:
                phraseCandidates,

            queryAnalysis:
                queryAnalysis,

            totalQueryTerms:
                0,

            contentTermCount:
                contentTokens.length,

            functionTermCount:
                functionTokens.length,

            indexTokenCount:
                documentItem.indexTokenCount ||
                0,

            indexUniqueTerms:
                documentItem.indexUniqueTerms ||
                0,

            indexUniqueFamilies:
                documentItem.indexUniqueFamilies ||
                0,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // فهرس Orama
    // ==================================

    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    if (
        !db
    ) {

        throw new Error(
            "تعذر إنشاء فهرس Orama."
        );

    }


    // ==================================
    // تشغيل بحث Orama
    // ==================================

    function runOramaSearch(
        term,
        exact,
        limit
    ) {

        const cleanTerm =
            String(
                term ||
                ""
            ).trim();


        if (
            !cleanTerm
        ) {

            return {

                count:
                    0,

                hits:
                    []

            };

        }


        try {

            if (
                !db ||
                !window.Orama ||
                typeof window.Orama.search !==
                    "function"
            ) {

                return {

                    count:
                        0,

                    hits:
                        []

                };

            }


            return (
                window.Orama.search(
                    db,
                    {

                        term:
                            cleanTerm,

                        properties: [

                            "text",

                            "heading"

                        ],

                        boost: {

                            heading:
                                4,

                            text:
                                1

                        },

                        tolerance:
                            exact
                                ? 0
                                : 1,

                        exact:
                            Boolean(
                                exact
                            ),

                        limit:
                            typeof limit ===
                            "number"
                                ? limit
                                : 100

                    }
                ) ||
                {

                    count:
                        0,

                    hits:
                        []

                }
            );

        }
        catch (
            error
        ) {

            console.warn(
                "فشل بحث Orama:",
                cleanTerm,
                error
            );


            return {

                count:
                    0,

                hits:
                    []

            };

        }

    }


    // ==================================
    // النتائج الخام
    // ==================================

    const rawHits =
        [];


    // ==================================
    // 1) البحث في السؤال كاملًا
    // ==================================

    const fullQueryResult =
        runOramaSearch(
            searchTerm,
            false,
            100
        );


    if (
        fullQueryResult &&
        Array.isArray(
            fullQueryResult.hits
        )
    ) {

        fullQueryResult.hits.forEach(
            function (
                hit
            ) {

                rawHits.push({

                    hit:
                        hit,

                    source:
                        "full-query",

                    searchTerm:
                        searchTerm,

                    queryWeight:
                        1

                });

            }
        );

    }


    // ==================================
    // 2) العبارات المركبة
    // ==================================

    phraseCandidates.forEach(
        function (
            phraseItem
        ) {

            if (
                !phraseItem ||
                !phraseItem.phrase
            ) {

                return;

            }


            const phraseResult =
                runOramaSearch(
                    phraseItem.phrase,
                    false,
                    100
                );


            if (
                !phraseResult ||
                !Array.isArray(
                    phraseResult.hits
                )
            ) {

                return;

            }


            const phraseWeight =
                Number(
                    phraseItem.weight ||
                    1
                );


            phraseResult.hits.forEach(
                function (
                    hit
                ) {

                    rawHits.push({

                        hit:
                            hit,

                        source:
                            "phrase",

                        searchTerm:
                            phraseItem.phrase,

                        queryWeight:
                            phraseWeight

                    });

                }
            );

        }
    );


    // ==================================
    // 3) المصطلحات الموضوعية وصيغها
    //
    // لا نعيد تحليل الكلمة هنا.
    // نستخدم variants التي خرجت من
    // analyzeSearchQuery().
    // ==================================

    weightedTerms.forEach(
        function (
            termItem
        ) {

            if (
                !termItem ||
                !termItem.term
            ) {

                return;

            }


            const baseWeight =
                Number(
                    termItem.weight ||
                    0.5
                );


            const variants =
                Array.isArray(
                    termItem.variants
                ) &&
                termItem.variants.length >
                    0

                    ? termItem.variants

                    : [
                        termItem.term
                    ];


            const searchedVariants =
                new Set();


            variants.forEach(
                function (
                    variant,
                    variantIndex
                ) {

                    const cleanVariant =
                        String(
                            variant ||
                            ""
                        ).trim();


                    if (
                        !cleanVariant ||
                        searchedVariants.has(
                            cleanVariant
                        )
                    ) {

                        return;

                    }


                    searchedVariants.add(
                        cleanVariant
                    );


                    const variantResult =
                        runOramaSearch(
                            cleanVariant,
                            false,
                            100
                        );


                    if (
                        !variantResult ||
                        !Array.isArray(
                            variantResult.hits
                        )
                    ) {

                        return;

                    }


                    let variantWeight =
                        baseWeight;


                    // الصيغة الأساسية أقوى
                    // قليلًا من الصيغة الملحقة
                    if (
                        variantIndex >
                        0
                    ) {

                        variantWeight *=
                            0.92;

                    }


                    variantResult.hits.forEach(
                        function (
                            hit
                        ) {

                            rawHits.push({

                                hit:
                                    hit,

                                source:
                                    variantIndex ===
                                    0

                                        ? "term"

                                        : "variant",

                                searchTerm:
                                    cleanVariant,

                                queryWeight:
                                    variantWeight

                            });

                        }
                    );

                }
            );

        }
    );


    // ==================================
    // الكلمات الوظيفية لا تبحث منفردة
    // ==================================

    void functionTokens;


    // ==================================
    // دمج النتائج حسب الفقرة
    // ==================================

    const mergedResults =
        new Map();


    rawHits.forEach(
        function (
            rawItem
        ) {

            if (
                !rawItem ||
                !rawItem.hit ||
                !rawItem.hit.document
            ) {

                return;

            }


            const document =
                rawItem.hit.document;


            const paragraphIndex =
                Number(
                    document.paragraphIndex
                );


            if (
                Number.isNaN(
                    paragraphIndex
                )
            ) {

                return;

            }


            const oramaScore =
                Number(
                    rawItem.hit.score ||
                    0
                );


            const weightedOramaScore =
                oramaScore *
                Number(
                    rawItem.queryWeight ||
                    1
                );


            const existing =
                mergedResults.get(
                    paragraphIndex
                );


            if (
                !existing
            ) {

                mergedResults.set(
                    paragraphIndex,
                    {

                        document:
                            document,

                        oramaScore:
                            oramaScore,

                        weightedOramaScore:
                            weightedOramaScore,

                        sources:
                            [
                                rawItem.source
                            ]

                    }
                );

            }
            else {

                existing.oramaScore =
                    Math.max(
                        existing.oramaScore,
                        oramaScore
                    );


                existing.weightedOramaScore =
                    Math.max(
                        existing.weightedOramaScore,
                        weightedOramaScore
                    );


                if (
                    !existing.sources.includes(
                        rawItem.source
                    )
                ) {

                    existing.sources.push(
                        rawItem.source
                    );

                }

            }

        }
    );


    // ==================================
    // العناوين
    // ==================================

    const headings =
        Array.isArray(
            structureData.headings
        )
            ? structureData.headings
                .filter(
                    function (
                        heading
                    ) {

                        return (
                            heading &&
                            typeof heading.index !==
                                "undefined" &&
                            String(
                                heading.text ||
                                ""
                            ).trim()
                        );

                    }
                )
                .slice()
                .sort(
                    function (
                        a,
                        b
                    ) {

                        return (
                            Number(
                                a.index
                            ) -
                            Number(
                                b.index
                            )
                        );

                    }
                )
            : [];


    // ==================================
    // الحصول على أقرب عنوان
    // ==================================

    function getNearestHeading(
        paragraphIndex
    ) {

        let nearest =
            null;


        for (
            let i =
                headings.length - 1;

            i >= 0;

            i--
        ) {

            const heading =
                headings[i];


            if (
                Number(
                    heading.index
                ) <=
                Number(
                    paragraphIndex
                )
            ) {

                nearest =
                    heading;

                break;

            }

        }


        return nearest;

    }


    // ==================================
    // ترجيح عام بحسب نوع السؤال
    // ==================================

    function getProfileScore(
        text,
        heading
    ) {

        const normalizedText =
            normalizeSearchText(
                text
            );


        const normalizedHeading =
            normalizeSearchText(
                heading
            );


        let score =
            0;


        if (
            retrievalProfile ===
            "definition"
        ) {

            if (
                /تعريف|المقصود|المراد|يعني|يقصد|هو ان|هي ان/
                    .test(
                        normalizedText
                    )
            ) {

                score +=
                    5;

            }


            if (
                /تعريف|المقصود|المراد/
                    .test(
                        normalizedHeading
                    )
            ) {

                score +=
                    6;

            }

        }
        else if (
            retrievalProfile ===
            "effect"
        ) {

            if (
                /اثر|تاثير|نتيجة|ينتج|يترتب|انعكاس/
                    .test(
                        normalizedText
                    )
            ) {

                score +=
                    5;

            }


            if (
                /اثر|تاثير|نتيجة/
                    .test(
                        normalizedHeading
                    )
            ) {

                score +=
                    6;

            }

        }
        else if (
            retrievalProfile ===
            "comparison"
        ) {

            if (
                /علاق[ةه]|فرق|فروق|مقارنة|يقارن|تمييز|خلاف/
                    .test(
                        normalizedText
                    )
            ) {

                score +=
                    5;

            }


            if (
                /علاق[ةه]|فرق|فروق|مقارنة|يقارن|تمييز|خلاف/
                    .test(
                        normalizedHeading
                    )
            ) {

                score +=
                    7;

            }

        }
        else if (
            retrievalProfile ===
            "causes"
        ) {

            if (
                /سبب|اسباب|علة|علل|بسبب|لان/
                    .test(
                        normalizedText
                    )
            ) {

                score +=
                    5;

            }

        }
        else if (
            retrievalProfile ===
            "location"
        ) {

            if (
                /الفصل|المبحث|المطلب|الباب|الصفحة|موضع/
                    .test(
                        normalizedText
                    )
            ) {

                score +=
                    4;

            }


            if (
                /الفصل|المبحث|المطلب|الباب/
                    .test(
                        normalizedHeading
                    )
            ) {

                score +=
                    7;

            }

        }


        return score;

    }


    // ==================================
    // حساب قرب مفاهيم السؤال
    // باستخدام variants من المحلل
    // ==================================

    function calculateQueryProximity(
        normalizedText
    ) {

        const positions =
            [];


        contentTokens.forEach(
            function (
                token
            ) {

                const termItem =
                    weightedTerms.find(
                        function (
                            item
                        ) {

                            return (
                                item &&
                                item.term ===
                                token
                            );

                        }
                    );


                const variants =
                    termItem &&
                    Array.isArray(
                        termItem.variants
                    ) &&
                    termItem.variants.length >
                        0

                        ? termItem.variants

                        : [
                            token
                        ];


                let bestPosition =
                    -1;


                variants.forEach(
                    function (
                        variant
                    ) {

                        const normalizedVariant =
                            normalizeSearchText(
                                variant
                            );


                        const position =
                            normalizedText.indexOf(
                                normalizedVariant
                            );


                        if (
                            position >=
                            0 &&
                            (
                                bestPosition <
                                    0 ||
                                position <
                                    bestPosition
                            )
                        ) {

                            bestPosition =
                                position;

                        }

                    }
                );


                if (
                    bestPosition >=
                    0
                ) {

                    positions.push(
                        bestPosition
                    );

                }

            }
        );


        if (
            positions.length <
            2
        ) {

            return 0;

        }


        positions.sort(
            function (
                a,
                b
            ) {

                return (
                    a -
                    b
                );

            }
        );


        const span =
            positions[
                positions.length - 1
            ] -
            positions[0] +
            1;


        return (
            20 /
            Math.max(
                span,
                1
            )
        );

    }


    // ==================================
    // النتائج النهائية
    // ==================================

    const finalResults =
        [];


    mergedResults.forEach(
        function (
            mergedItem
        ) {

            const paragraph =
                mergedItem.document;


            if (
                !paragraph
            ) {

                return;

            }


            const originalText =
                String(
                    paragraph.text ||
                    ""
                )
                    .trim();


            if (
                !originalText
            ) {

                return;

            }


            const normalizedText =
                normalizeSearchText(
                    originalText
                );


            const nearestHeading =
                getNearestHeading(
                    paragraph.paragraphIndex
                );


            const headingText =
                nearestHeading
                    ? String(
                        nearestHeading.text ||
                        ""
                    ).trim()
                    : "";


            const normalizedHeading =
                normalizeSearchText(
                    headingText
                );


            // ==================================
            // المصطلحات المطابقة
            // ==================================

            const paragraphMatchedTerms =
                weightedTerms.filter(
                    function (
                        termItem
                    ) {

                        if (
                            !termItem ||
                            !termItem.term
                        ) {

                            return false;

                        }


                        const variants =
                            Array.isArray(
                                termItem.variants
                            ) &&
                            termItem.variants.length >
                                0

                                ? termItem.variants

                                : [
                                    termItem.term
                                ];


                        return variants.some(
                            function (
                                variant
                            ) {

                                const normalizedVariant =
                                    normalizeSearchText(
                                        variant
                                    );


                                return (
                                    normalizedText.includes(
                                        normalizedVariant
                                    ) ||
                                    normalizedHeading.includes(
                                        normalizedVariant
                                    )
                                );

                            }
                        );

                    }
                );


            // ==================================
            // التغطية الموزونة
            // ==================================

            let weightedMatchedTotal =
                0;


            let weightedMatchedValue =
                0;


            weightedTerms.forEach(
                function (
                    termItem
                ) {

                    if (
                        !termItem ||
                        !termItem.term
                    ) {

                        return;

                    }


                    const weight =
                        Number(
                            termItem.weight ||
                            0
                        );


                    weightedMatchedTotal +=
                        weight;


                    const variants =
                        Array.isArray(
                            termItem.variants
                        ) &&
                        termItem.variants.length >
                            0

                            ? termItem.variants

                            : [
                                termItem.term
                            ];


                    const isMatched =
                        variants.some(
                            function (
                                variant
                            ) {

                                const normalizedVariant =
                                    normalizeSearchText(
                                        variant
                                    );


                                return (
                                    normalizedText.includes(
                                        normalizedVariant
                                    ) ||
                                    normalizedHeading.includes(
                                        normalizedVariant
                                    )
                                );

                            }
                        );


                    if (
                        isMatched
                    ) {

                        weightedMatchedValue +=
                            weight;

                    }

                }
            );


            const weightedCoverage =
                weightedMatchedTotal >
                0

                    ? weightedMatchedValue /
                      weightedMatchedTotal

                    : 0;


            // ==================================
            // المفاهيم المباشرة
            // ==================================

            const directMatchedContentTokens =
                contentTokens.filter(
                    function (
                        token
                    ) {

                        const termItem =
                            weightedTerms.find(
                                function (
                                    item
                                ) {

                                    return (
                                        item &&
                                        item.term ===
                                        token
                                    );

                                }
                            );


                        const variants =
                            termItem &&
                            Array.isArray(
                                termItem.variants
                            ) &&
                            termItem.variants.length >
                                0

                                ? termItem.variants

                                : [
                                    token
                                ];


                        return variants.some(
                            function (
                                variant
                            ) {

                                const normalizedVariant =
                                    normalizeSearchText(
                                        variant
                                    );


                                return (
                                    normalizedText.includes(
                                        normalizedVariant
                                    ) ||
                                    normalizedHeading.includes(
                                        normalizedVariant
                                    )
                                );

                            }
                        );

                    }
                );


            const directContentCoverage =
                contentTokens.length >
                0

                    ? directMatchedContentTokens.length /
                      contentTokens.length

                    : 0;


            // ==================================
            // تغطية المفاهيم
            // ==================================

            const conceptCoverage =
                directContentCoverage;


            // ==================================
            // عائلات الفقرة
            // ==================================

            const paragraphTokens =
                tokenizeDocumentText(
                    originalText
                );


            const paragraphFamilies =
                [];


            paragraphTokens.forEach(
                function (
                    token
                ) {

                    const family =
                        getConservativeFamilyKey(
                            token,
                            null
                        );


                    if (
                        family &&
                        !paragraphFamilies.includes(
                            family
                        )
                    ) {

                        paragraphFamilies.push(
                            family
                        );

                    }

                }
            );


            const paragraphMatchedFamilies =
                queryFamilies.filter(
                    function (
                        family
                    ) {

                        return paragraphFamilies.includes(
                            family
                        );

                    }
                );


            const familyCoverage =
                queryFamilies.length >
                0

                    ? paragraphMatchedFamilies.length /
                      queryFamilies.length

                    : 0;


            // ==================================
            // العبارة الكاملة
            // ==================================

            const exactPhrase =
                normalizedText.includes(
                    searchTerm
                );


            // ==================================
            // العبارات المركبة المطابقة
            // ==================================

            let matchedPhraseCount =
                0;


            let matchedPhraseWeight =
                0;


            phraseCandidates.forEach(
                function (
                    phraseItem
                ) {

                    if (
                        !phraseItem ||
                        !phraseItem.phrase
                    ) {

                        return;

                    }


                    const phrase =
                        normalizeSearchText(
                            phraseItem.phrase
                        );


                    if (
                        normalizedText.includes(
                            phrase
                        ) ||
                        normalizedHeading.includes(
                            phrase
                        )
                    ) {

                        matchedPhraseCount +=
                            1;


                        matchedPhraseWeight +=
                            Number(
                                phraseItem.weight ||
                                1
                            );

                    }

                }
            );


            // ==================================
            // العنوان
            // ==================================

            let headingScore =
                0;


            contentTokens.forEach(
                function (
                    token
                ) {

                    const termItem =
                        weightedTerms.find(
                            function (
                                item
                            ) {

                                return (
                                    item &&
                                    item.term ===
                                    token
                                );

                            }
                        );


                    const variants =
                        termItem &&
                        Array.isArray(
                            termItem.variants
                        ) &&
                        termItem.variants.length >
                            0

                            ? termItem.variants

                            : [
                                token
                            ];


                    const matched =
                        variants.some(
                            function (
                                variant
                            ) {

                                return normalizedHeading.includes(
                                    normalizeSearchText(
                                        variant
                                    )
                                );

                            }
                        );


                    if (
                        matched
                    ) {

                        headingScore +=
                            4;

                    }

                }
            );


            if (
                normalizedHeading.includes(
                    searchTerm
                )
            ) {

                headingScore +=
                    18;

            }


            const headingMatchedContentTokens =
                contentTokens.filter(
                    function (
                        token
                    ) {

                        const termItem =
                            weightedTerms.find(
                                function (
                                    item
                                ) {

                                    return (
                                        item &&
                                        item.term ===
                                        token
                                    );

                                }
                            );


                        const variants =
                            termItem &&
                            Array.isArray(
                                termItem.variants
                            ) &&
                            termItem.variants.length >
                                0

                                ? termItem.variants

                                : [
                                    token
                                ];


                        return variants.some(
                            function (
                                variant
                            ) {

                                return normalizedHeading.includes(
                                    normalizeSearchText(
                                        variant
                                    )
                                );

                            }
                        );

                    }
                ).length;


            const headingCoverage =
                contentTokens.length >
                0

                    ? headingMatchedContentTokens /
                      contentTokens.length

                    : 0;


            // ==================================
            // القرب
            // ==================================

            const proximityScore =
                calculateQueryProximity(
                    normalizedText
                );


            // ==================================
            // نوع السؤال
            // ==================================

            const profileScore =
                getProfileScore(
                    originalText,
                    headingText
                );


            // ==================================
            // الدرجة الأساسية
            // ==================================

            let score =
                Number(
                    mergedItem.weightedOramaScore ||
                    0
                );


            // ==================================
            // تغطية المصطلحات
            // ==================================

            score +=
                weightedCoverage *
                24;


            // ==================================
            // تغطية المفاهيم
            // ==================================

            score +=
                conceptCoverage *
                18;


            // ==================================
            // اجتماع جميع المفاهيم
            // ==================================

            if (
                contentTokens.length >=
                    2 &&
                conceptCoverage >=
                    1
            ) {

                score +=
                    20;

            }
            else if (
                directMatchedContentTokens.length >=
                2
            ) {

                score +=
                    10;

            }


            // ==================================
            // العائلات
            // ==================================

            score +=
                familyCoverage *
                14;


            // ==================================
            // العبارات المركبة
            // ==================================

            score +=
                Math.min(
                    matchedPhraseWeight *
                    10,
                    20
                );


            // ==================================
            // العنوان
            // ==================================

            score +=
                headingScore;


            // ==================================
            // تغطية العنوان
            // ==================================

            score +=
                headingCoverage *
                8;


            // ==================================
            // القرب
            // ==================================

            score +=
                Math.min(
                    proximityScore,
                    8
                );


            // ==================================
            // نوع السؤال
            // ==================================

            score +=
                profileScore;


            // ==================================
            // العبارة الكاملة
            // ==================================

            if (
                exactPhrase
            ) {

                score +=
                    15;

            }


            // ==================================
            // اجتماع أكثر من عائلة
            // ==================================

            if (
                paragraphMatchedFamilies.length >=
                2
            ) {

                score +=
                    8;

            }


            if (
                directMatchedContentTokens.length >=
                2
            ) {

                score +=
                    5;

            }


            // ==================================
            // فلترة عامة جدًا
            // ==================================

            if (
                weightedCoverage ===
                    0 &&
                directContentCoverage ===
                    0 &&
                headingCoverage ===
                    0
            ) {

                return;

            }


            // ==================================
            // نوع التطابق
            // ==================================

            let matchType =
                "word";


            if (
                exactPhrase
            ) {

                matchType =
                    "exact";

            }
            else if (
                matchedPhraseCount >
                0
            ) {

                matchType =
                    "phrase";

            }
            else if (
                headingScore >=
                8
            ) {

                matchType =
                    "heading";

            }
            else if (
                paragraphMatchedFamilies.length >
                0
            ) {

                matchType =
                    "family";

            }


            // ==================================
            // بناء السياق
            // ==================================

            let context =
                originalText;


            let firstContentPosition =
                -1;


            let firstContentToken =
                "";


            for (
                let i = 0;

                i <
                    contentTokens.length;

                i++
            ) {

                const token =
                    contentTokens[i];


                const termItem =
                    weightedTerms.find(
                        function (
                            item
                        ) {

                            return (
                                item &&
                                item.term ===
                                token
                            );

                        }
                    );


                const variants =
                    termItem &&
                    Array.isArray(
                        termItem.variants
                    ) &&
                    termItem.variants.length >
                        0

                        ? termItem.variants

                        : [
                            token
                        ];


                for (
                    let j = 0;

                    j <
                        variants.length;

                    j++
                ) {

                    const normalizedVariant =
                        normalizeSearchText(
                            variants[j]
                        );


                    const position =
                        normalizedText.indexOf(
                            normalizedVariant
                        );


                    if (
                        position >=
                        0 &&
                        (
                            firstContentPosition <
                                0 ||
                            position <
                                firstContentPosition
                        )
                    ) {

                        firstContentPosition =
                            position;

                        firstContentToken =
                            normalizedVariant;

                    }

                }

            }


            if (
                firstContentPosition >=
                    0 &&
                firstContentToken
            ) {

                const start =
                    Math.max(
                        0,
                        firstContentPosition -
                        140
                    );


                const end =
                    Math.min(
                        normalizedText.length,
                        firstContentPosition +
                        firstContentToken.length +
                        360
                    );


                context =
                    normalizedText.substring(
                        start,
                        end
                    );

            }


            // ==================================
            // حفظ النتيجة
            // ==================================

            finalResults.push({

                paragraphIndex:
                    paragraph.paragraphIndex,

                paragraphId:
                    paragraph.paragraphId,

                text:
                    originalText,

                context:
                    context,

                heading:
                    headingText,

                headingLevel:
                    nearestHeading
                        ? String(
                            nearestHeading.style ||
                            ""
                        )
                        : "",

                score:
                    score,

                oramaScore:
                    Number(
                        mergedItem.oramaScore ||
                        0
                    ),

                weightedOramaScore:
                    Number(
                        mergedItem.weightedOramaScore ||
                        0
                    ),

                profile:
                    retrievalProfile,

                profileScore:
                    profileScore,

                relationScore:
                    0,

                conceptCoverage:
                    conceptCoverage,

                comparisonCoverage:
                    weightedCoverage,

                matchType:
                    matchType,

                matchedTerms:
                    paragraphMatchedTerms.map(
                        function (
                            item
                        ) {

                            return item.term;

                        }
                    ),

                matchedFamilies:
                    paragraphMatchedFamilies,

                matchedFamilyCount:
                    paragraphMatchedFamilies.length,

                matchedPhraseCount:
                    matchedPhraseCount,

                matchedPhraseWeight:
                    matchedPhraseWeight,

                queryCoverage:
                    directContentCoverage,

                weightedCoverage:
                    weightedCoverage,

                familyCoverage:
                    familyCoverage,

                headingScore:
                    headingScore,

                headingCoverage:
                    headingCoverage,

                proximityScore:
                    proximityScore,

                exactPhrase:
                    exactPhrase,

                searchSources:
                    Array.isArray(
                        mergedItem.sources
                    )
                        ? mergedItem.sources
                        : []

            });

        }
    );


    // ==================================
    // ترتيب النتائج
    // ==================================

    finalResults.sort(
        function (
            a,
            b
        ) {

            if (
                Number(
                    b.score
                ) !==
                Number(
                    a.score
                )
            ) {

                return (
                    Number(
                        b.score
                    ) -
                    Number(
                        a.score
                    )
                );

            }


            if (
                Number(
                    b.conceptCoverage
                ) !==
                Number(
                    a.conceptCoverage
                )
            ) {

                return (
                    Number(
                        b.conceptCoverage
                    ) -
                    Number(
                        a.conceptCoverage
                    )
                );

            }


            if (
                Number(
                    b.weightedCoverage
                ) !==
                Number(
                    a.weightedCoverage
                )
            ) {

                return (
                    Number(
                        b.weightedCoverage
                    ) -
                    Number(
                        a.weightedCoverage
                    )
                );

            }


            if (
                Number(
                    b.familyCoverage
                ) !==
                Number(
                    a.familyCoverage
                )
            ) {

                return (
                    Number(
                        b.familyCoverage
                    ) -
                    Number(
                        a.familyCoverage
                    )
                );

            }


            if (
                Number(
                    b.headingScore
                ) !==
                Number(
                    a.headingScore
                )
            ) {

                return (
                    Number(
                        b.headingScore
                    ) -
                    Number(
                        a.headingScore
                    )
                );

            }


            return (
                Number(
                    a.paragraphIndex
                ) -
                Number(
                    b.paragraphIndex
                )
            );

        }
    );


    // ==================================
    // عدد المرشحين
    // ==================================

    const candidateLimit =
        typeof settings.candidateLimit ===
            "number"

                ? settings.candidateLimit

                : 50;


    if (
        finalResults.length >
        candidateLimit
    ) {

        finalResults.splice(
            candidateLimit
        );

    }


    // ==================================
    // عدد الظهورات
    // ==================================

    let indexedOccurrences =
        0;


    const allParagraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    const allDocumentText =
        allParagraphs
            .map(
                function (
                    paragraph
                ) {

                    return (
                        paragraph &&
                        paragraph.text
                            ? paragraph.text
                            : ""
                    );

                }
            )
            .join(
                "\n"
            );


    const allDocumentTokens =
        tokenizeDocumentText(
            allDocumentText
        );


    const occurrenceFamilySet =
        new Set(
            queryFamilies
        );


    allDocumentTokens.forEach(
        function (
            token
        ) {

            const family =
                getConservativeFamilyKey(
                    token,
                    null
                );


            if (
                occurrenceFamilySet.has(
                    family
                )
            ) {

                indexedOccurrences +=
                    1;

            }

        }
    );


    // ==================================
    // النتيجة النهائية
    // ==================================

    return {

        query:
            searchTerm,

        profile:
            retrievalProfile,

        count:
            finalResults.length,

        results:
            finalResults,

        matchedTerms:
            weightedTerms
                .filter(
                    function (
                        item
                    ) {

                        return (
                            item &&
                            item.term
                        );

                    }
                )
                .map(
                    function (
                        item
                    ) {

                        return item.term;

                    }
                ),

        matchedFamilies:
            queryFamilies,

        phraseCandidates:
            phraseCandidates,

        queryAnalysis:
            queryAnalysis,

        totalQueryTerms:
            queryTokens.length,

        contentTermCount:
            contentTokens.length,

        functionTermCount:
            functionTokens.length,

        indexTokenCount:
            documentItem.indexTokenCount ||
            0,

        indexUniqueTerms:
            documentItem.indexUniqueTerms ||
            0,

        indexUniqueFamilies:
            documentItem.indexUniqueFamilies ||
            0,

        indexedOccurrences:
            indexedOccurrences

    };

}
// ======================================
// Console Test Bridge
// إتاحة البحث للاختبار من Console
// ======================================

window.testIndexedSearch =
    async function (
        query
    ) {

        if (!currentDocument) {

            console.warn(
                "لا يوجد مستند نشط."
            );

            return null;

        }

        try {

            const result =
                await searchIndexedDocument(
                    currentDocument.id,
                    query,
                    {
                        maxResults:
                            10
                    }
                );

            console.log(
                "======================================"
            );

            console.log(
                "اختبار البحث:",
                query
            );

            console.log(
                "عدد النتائج:",
                result.count
            );

            console.log(
                "الكلمات:",
                result.matchedTerms
            );

            console.log(
                "العائلات:",
                result.matchedFamilies
            );

            console.log(
                "النتائج:",
                result.results
            );

            console.log(
                "======================================"
            );

            return result;

        }
        catch (error) {

            console.error(
                "فشل اختبار البحث:",
                error
            );

            return null;

        }

    };

// =====================================================
// Orama Test
// للاختبار من Console
// =====================================================

window.testOramaSearch =
    async function (
        query
    ) {

        try {

            if (!currentDocument) {

                console.warn(
                    "لا يوجد مستند نشط."
                );

                return [];

            }


            const results =
                await searchOramaDocument(
                    currentDocument,
                    query
                );


            console.log(
                "======================================"
            );


            console.log(
                "اختبار Orama"
            );


            console.log(
                "السؤال:",
                query
            );


            console.log(
                "عدد النتائج:",
                results.length
            );


            results
                .slice(
                    0,
                    20
                )
                .forEach(
                    function (
                        result,
                        index
                    ) {

                        console.log(
                            "#" +
                            (
                                index + 1
                            ),
                            {

                                score:
                                    result.score,

                                matchType:
                                    result.matchType,

                                heading:
                                    result.heading,

                                paragraphIndex:
                                    result.paragraphIndex,

                                matchedTerms:
                                    result.matchedTerms,

                                matchedFamilies:
                                    result.matchedFamilies,

                                text:
                                    String(
                                        result.text ||
                                        ""
                                    ).substring(
                                        0,
                                        300
                                    )

                            }
                        );

                    }
                );


            console.log(
                "======================================"
            );


            return results;

        }
        catch (error) {

            console.error(
                "فشل اختبار Orama:",
                error
            );


            return [];

        }

    };


// =====================================================
// Orama Initial Status
// =====================================================

console.log(
    "تم تحميل طبقة Orama الجديدة بنجاح."
);
// =====================================================
// Research Tools
// PART 3
// سياق الاسترجاع + الإحالات + تصنيف السؤال
// =====================================================


// =====================================================
// Detect Retrieval Profile
// تحديد نوع الاسترجاع بحسب السؤال
// =====================================================

function getRetrievalProfile(
    query
) {

    const text =
        normalizeSearchText(
            query
        );


    const profile = {

        type:
            "general",

        maxResults:
            8,

        maxChars:
            8000

    };


    // ==================================
    // تعريف / معنى
    // ==================================

    if (
        /ماهوا|ماهو|ماهى|ماهي|ما هي|المقصود|معنى|تعريف|يقصد ب|المراد ب/
            .test(
                text
            )
    ) {

        profile.type =
            "definition";

        profile.maxResults =
            5;

        profile.maxChars =
            6000;

        return profile;

    }


    // ==================================
    // أثر / نتائج / تأثير
    // ==================================

    if (
        /اثر|تاثير|نتائج|ينتج عن|يترتب على|انعكاس/
            .test(
                text
            )
    ) {

        profile.type =
            "effect";

        profile.maxResults =
            8;

        profile.maxChars =
            9000;

        return profile;

    }


    // ==================================
    // مقارنة
    // ==================================

    if (
        /الفرق|الفروق|مقارنة|يقارن|ما الفرق|التمييز بين|يفترق/
            .test(
                text
            )
    ) {

        profile.type =
            "comparison";

        profile.maxResults =
            10;

        profile.maxChars =
            10000;

        return profile;

    }


    // ==================================
    // أسباب / علل
    // ==================================

    if (
        /لماذا|سبب|اسباب|علة|علل|لان|بسبب/
            .test(
                text
            )
    ) {

        profile.type =
            "causes";

        profile.maxResults =
            8;

        profile.maxChars =
            9000;

        return profile;

    }


    // ==================================
    // موضع / فصل / مبحث / مطلب
    // ==================================

    if (
        /اين|موضع|موضعه|الفصل|المبحث|المطلب|الصفحة/
            .test(
                text
            )
    ) {

        profile.type =
            "location";

        profile.maxResults =
            6;

        profile.maxChars =
            6000;

        return profile;

    }


    return profile;

}


// =====================================================
// Get Retrieval Limits
// =====================================================

function getRetrievalLimits(
    providerName,
    modelName
) {

    const providerValue =
        String(
            providerName || ""
        ).toLowerCase();


    if (
        providerValue ===
        "groq"
    ) {

        return {

            maxResults:
                4,

            maxChars:
                3500

        };

    }


    if (
        providerValue ===
        "openrouter"
    ) {

        return {

            maxResults:
                5,

            maxChars:
                5000

        };

    }


    if (
        providerValue ===
        "gemini"
    ) {

        return {

            maxResults:
                6,

            maxChars:
                6000

        };

    }


    if (
        providerValue ===
        "openai"
    ) {

        return {

            maxResults:
                6,

            maxChars:
                6000

        };

    }


    return {

        maxResults:
            4,

        maxChars:
            3500

    };

}


// =====================================================
// Estimate Token Count
// =====================================================

function estimateTokenCount(
    text
) {

    return Math.ceil(
        String(
            text || ""
        ).length / 4
    );

}


// =====================================================
// Common Text Length
// تقدير التشابه بين مقطعين
// =====================================================

function getCommonTextLength(
    textA,
    textB
) {

    const a =
        String(
            textA || ""
        );

    const b =
        String(
            textB || ""
        );


    if (
        !a ||
        !b
    ) {

        return 0;

    }


    let best =
        0;


    const minLength =
        Math.min(
            a.length,
            b.length
        );


    const maxWindow =
        Math.min(
            minLength,
            300
        );


    for (
        let length = maxWindow;
        length >= 20;
        length -= 10
    ) {

        let found =
            false;


        for (
            let i = 0;
            i + length <= a.length;
            i += 10
        ) {

            const part =
                a.substring(
                    i,
                    i + length
                );


            if (
                b.includes(
                    part
                )
            ) {

                best =
                    length;

                found =
                    true;

                break;

            }

        }


        if (found) {

            break;

        }

    }


    return best;

}


// =====================================================
// Build Retrieval Context
// تحويل نتائج البحث إلى سياق فعلي للذكاء الاصطناعي
// معالجة نتائج العناوين إلى فقرات المحتوى التابعة لها
// مع الحفاظ على الإحالات والانتقال إلى Word
// =====================================================

// =====================================================
// Build Retrieval Context
// تحويل نتائج البحث إلى سياق ذكي للذكاء الاصطناعي
//
// المبدأ:
//
// 1) نتائج Orama هي أساس الاسترجاع.
// 2) قوة العنوان تدخل في الوزن.
// 3) إذا كان هناك عنوان شديد الصلة بالسؤال:
//    يصبح العنوان مرساة للقسم.
// 4) تُضم أفضل الفقرات الواقعة تحت العنوان.
// 5) العنوان وحده ليس دليلًا.
// 6) الفقرات التابعة هي المادة التي تُرسل للـAI.
// 7) إذا لم يوجد عنوان قوي، نعود إلى الاسترجاع العام.
//
// ملاحظة:
// الدالة أصبحت async لأنها قد تحتاج بنية المستند
// لاستخراج الفقرات الواقعة تحت العنوان.
// =====================================================

async function buildRetrievalContext(
    searchResult,
    options
) {

    // ==================================
    // الإعدادات
    // ==================================

    const settings =
        options || {};


    const maxResults =
        typeof settings.maxResults ===
            "number"

                ? settings.maxResults

                : 8;


    const maxChars =
        typeof settings.maxChars ===
            "number"

                ? settings.maxChars

                : 8000;


    const includeNeighbors =
        settings.includeNeighbors !==
        false;


    // ==================================
    // التحقق
    // ==================================

    if (
        !searchResult ||
        !Array.isArray(
            searchResult.results
        )
    ) {

        return {

            query:
                searchResult &&
                searchResult.query
                    ? searchResult.query
                    : "",

            count:
                0,

            selectedCount:
                0,

            totalOccurrences:
                0,

            contexts:
                [],

            text:
                ""

        };

    }


    // ==================================
    // نتائج البحث الأصلية
    // ==================================

    const results =
        searchResult.results
            .filter(
                function (
                    result
                ) {

                    return (
                        result &&
                        typeof result.text ===
                            "string" &&
                        result.text.trim()
                    );

                }
            )
            .slice()
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        Number(
                            b.score ||
                            0
                        ) -
                        Number(
                            a.score ||
                            0
                        )
                    );

                }
            );


    if (
        results.length ===
        0
    ) {

        return {

            query:
                searchResult.query ||
                "",

            count:
                0,

            selectedCount:
                0,

            totalOccurrences:
                Number(
                    searchResult.indexedOccurrences ||
                    0
                ),

            contexts:
                [],

            text:
                ""

        };

    }


    // ==================================
    // تحليل الاستعلام
    // ==================================

    const queryAnalysis =
        searchResult.queryAnalysis ||
        {};


    const contentTokens =
        Array.isArray(
            queryAnalysis.contentTokens
        )
            ? queryAnalysis.contentTokens
            : [];


    const queryFamilies =
        Array.isArray(
            queryAnalysis.families
        )
            ? queryAnalysis.families
            : [];


    const weightedTerms =
        Array.isArray(
            queryAnalysis.weightedTerms
        )
            ? queryAnalysis.weightedTerms
            : [];


    const searchTerm =
        normalizeSearchText(
            searchResult.query ||
            ""
        );


    // ==================================
    // بنية المستند
    // ==================================

    let structureData =
        null;


    if (
        currentDocument
    ) {

        try {

            structureData =
                await ensureDocumentStructure(
                    currentDocument
                );

        }
        catch (
            error
        ) {

            console.warn(
                "تعذر تحميل بنية المستند لبناء سياق العناوين:",
                error
            );

        }

    }


    // ==================================
    // العناوين الفعلية للمستند
    //
    // نستخدم structureData.headings
    // وليس heading الموجود في النتيجة فقط.
    // ==================================

    const headings =
        structureData &&
        Array.isArray(
            structureData.headings
        )

            ? structureData.headings
                .filter(
                    function (
                        heading
                    ) {

                        return (
                            heading &&
                            typeof heading.index !==
                                "undefined" &&
                            String(
                                heading.text ||
                                ""
                            ).trim()
                        );

                    }
                )
                .slice()
                .sort(
                    function (
                        a,
                        b
                    ) {

                        return (
                            Number(
                                a.index
                            ) -
                            Number(
                                b.index
                            )
                        );

                    }
                )

            : [];


    // ==================================
    // جدول الفقرات
    // ==================================

    const paragraphs =
        structureData &&
        Array.isArray(
            structureData.paragraphs
        )

            ? structureData.paragraphs

            : [];


    const paragraphMap =
        new Map();


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                !paragraph ||
                typeof paragraph.index ===
                    "undefined"
            ) {

                return;

            }


            paragraphMap.set(
                Number(
                    paragraph.index
                ),
                paragraph
            );

        }
    );


    // ==================================
    // حساب قوة مطابقة العنوان
    //
    // هذه ليست نسبة مئوية.
    // إنها درجة ترجيح تضاف إلى النتيجة.
    // ==================================

    function calculateHeadingAnchorScore(
        heading
    ) {

        if (
            !heading
        ) {

            return 0;

        }


        const headingText =
            normalizeSearchText(
                heading.text ||
                ""
            );


        if (
            !headingText
        ) {

            return 0;

        }


        let anchorScore =
            0;


        // ==================================
        // 1) تطابق المفاهيم الأساسية
        // ==================================

        let matchedContentCount =
            0;


        contentTokens.forEach(
            function (
                token
            ) {

                const normalizedToken =
                    normalizeSearchText(
                        token
                    );


                if (
                    !normalizedToken
                ) {

                    return;

                }


                const termItem =
                    weightedTerms.find(
                        function (
                            item
                        ) {

                            return (
                                item &&
                                item.term ===
                                token
                            );

                        }
                    );


                const variants =
                    termItem &&
                    Array.isArray(
                        termItem.variants
                    ) &&
                    termItem.variants.length >
                        0

                        ? termItem.variants

                        : [
                            token
                        ];


                const matched =
                    variants.some(
                        function (
                            variant
                        ) {

                            return headingText.includes(
                                normalizeSearchText(
                                    variant
                                )
                            );

                        }
                    );


                if (
                    matched
                ) {

                    matchedContentCount +=
                        1;

                    anchorScore +=
                        18;

                }

            }
        );


        // ==================================
        // 2) العائلات
        // ==================================

        const headingTokens =
            tokenizeDocumentText(
                headingText
            );


        const headingFamilies =
            [];


        headingTokens.forEach(
            function (
                token
            ) {

                const family =
                    getConservativeFamilyKey(
                        token,
                        null
                    );


                if (
                    family &&
                    !headingFamilies.includes(
                        family
                    )
                ) {

                    headingFamilies.push(
                        family
                    );

                }

            }
        );


        let matchedFamilyCount =
            0;


        queryFamilies.forEach(
            function (
                family
            ) {

                if (
                    headingFamilies.includes(
                        family
                    )
                ) {

                    matchedFamilyCount +=
                        1;

                    anchorScore +=
                        10;

                }

            }
        );


        // ==================================
        // 3) العبارة الكاملة
        // ==================================

        if (
            searchTerm &&
            headingText ===
            searchTerm
        ) {

            anchorScore +=
                60;

        }
        else if (
            searchTerm &&
            headingText.includes(
                searchTerm
            )
        ) {

            anchorScore +=
                40;

        }


        // ==================================
        // 4) نوع السؤال
        // ==================================

        const profile =
            searchResult.profile ||
            "general";


        if (
            profile ===
            "definition" &&
            /تعريف|المقصود|تعريفه|تعريفها|معنى|ماهية/
                .test(
                    headingText
                )
        ) {

            anchorScore +=
                18;

        }


        if (
            profile ===
            "effect" &&
            /اثر|أثر|تاثير|تأثير|نتيجة|انعكاس/
                .test(
                    headingText
                )
        ) {

            anchorScore +=
                18;

        }


        if (
            profile ===
            "comparison" &&
            /فرق|فروق|مقارنة|علاقة|العلاقة|التمييز|بين/
                .test(
                    headingText
                )
        ) {

            anchorScore +=
                18;

        }


        if (
            profile ===
            "causes" &&
            /سبب|اسباب|أسباب|علة|علل/
                .test(
                    headingText
                )
        ) {

            anchorScore +=
                18;

        }


        // ==================================
        // 5) اجتماع أكثر من مفهوم
        // ==================================

        if (
            matchedContentCount >=
            2
        ) {

            anchorScore +=
                20;

        }


        if (
            matchedFamilyCount >=
            2
        ) {

            anchorScore +=
                12;

        }


        return anchorScore;

    }


    // ==================================
    // حساب درجات جميع العناوين
    // ==================================

    const headingAnchors =
        [];


    headings.forEach(
        function (
            heading,
            index
        ) {

            const anchorScore =
                calculateHeadingAnchorScore(
                    heading
                );


            if (
                anchorScore <=
                0
            ) {

                return;

            }


            const headingStart =
                Number(
                    heading.index
                );


            const nextHeading =
                headings[
                    index + 1
                ];


            const headingEnd =
                nextHeading
                    ? Number(
                        nextHeading.index
                    ) - 1

                    : paragraphs.length >
                        0
                        ? Number(
                            paragraphs[
                                paragraphs.length - 1
                            ].index
                        )
                        : headingStart;


            headingAnchors.push({

                heading:
                    heading,

                anchorScore:
                    anchorScore,

                startIndex:
                    headingStart,

                endIndex:
                    headingEnd

            });

        }
    );


    // ==================================
    // ترتيب المراسي
    // ==================================

    headingAnchors.sort(
        function (
            a,
            b
        ) {

            return (
                b.anchorScore -
                a.anchorScore
            );

        }
    );


    // ==================================
    // المراسي القوية
    //
    // ليست نسبة.
    // نستخدم حدًا وزنيًا.
    // ==================================

    const strongAnchors =
        headingAnchors.filter(
            function (
                anchor
            ) {

                return (
                    Number(
                        anchor.anchorScore
                    ) >=
                    35
                );

            }
        );


    // ==================================
    // الحصول على القسم التابع للعنوان
    // ==================================

    function getSectionParagraphs(
        anchor
    ) {

        if (
            !anchor
        ) {

            return [];

        }


        const section =
            [];


        paragraphs.forEach(
            function (
                paragraph
            ) {

                if (
                    !paragraph ||
                    typeof paragraph.index ===
                        "undefined"
                ) {

                    return;

                }


                const index =
                    Number(
                        paragraph.index
                    );


                // لا ندخل عنوان القسم نفسه
                // بوصفه دليلاً مستقلًا.
                if (
                    index <=
                    anchor.startIndex
                ) {

                    return;

                }


                if (
                    index >
                    anchor.endIndex
                ) {

                    return;

                }


                const text =
                    String(
                        paragraph.text ||
                        ""
                    ).trim();


                if (
                    !text
                ) {

                    return;

                }


                section.push(
                    paragraph
                );

            }
        );


        return section;

    }


    // ==================================
    // تحويل الفقرة إلى نتيجة قابلة للاختيار
    // ==================================

    function getResultForParagraph(
        paragraph
    ) {

        if (
            !paragraph
        ) {

            return null;

        }


        const paragraphIndex =
            Number(
                paragraph.index
            );


        // ----------------------------------
        // إذا كانت الفقرة موجودة أصلًا
        // في نتائج Orama نستخدم نتيجتها.
        // ----------------------------------

        const existing =
            results.find(
                function (
                    result
                ) {

                    return (
                        Number(
                            result.paragraphIndex
                        ) ===
                        paragraphIndex
                    );

                }
            );


        if (
            existing
        ) {

            return {

                ...existing

            };

        }


        // ----------------------------------
        // فقرة جديدة جلبناها بسبب عنوان قوي
        // ----------------------------------

        const text =
            String(
                paragraph.text ||
                ""
            ).trim();


        if (
            !text
        ) {

            return null;

        }


        const normalizedText =
            normalizeSearchText(
                text
            );


        let conceptCoverage =
            0;


        let matchedConcepts =
            0;


        contentTokens.forEach(
            function (
                token
            ) {

                const termItem =
                    weightedTerms.find(
                        function (
                            item
                        ) {

                            return (
                                item &&
                                item.term ===
                                token
                            );

                        }
                    );


                const variants =
                    termItem &&
                    Array.isArray(
                        termItem.variants
                    ) &&
                    termItem.variants.length >
                        0

                        ? termItem.variants

                        : [
                            token
                        ];


                const matched =
                    variants.some(
                        function (
                            variant
                        ) {

                            return normalizedText.includes(
                                normalizeSearchText(
                                    variant
                                )
                            );

                        }
                    );


                if (
                    matched
                ) {

                    matchedConcepts +=
                        1;

                }

            }
        );


        if (
            contentTokens.length >
            0
        ) {

            conceptCoverage =
                matchedConcepts /
                contentTokens.length;

        }


        let baseScore =
            conceptCoverage *
            20;


        return {

            paragraphIndex:
                paragraphIndex,

            paragraphId:
                paragraph.id,

            text:
                text,

            context:
                text,

            heading:
                "",

            headingLevel:
                paragraph.style ||
                "",

            score:
                baseScore,

            oramaScore:
                0,

            weightedOramaScore:
                0,

            profile:
                searchResult.profile ||
                "general",

            profileScore:
                0,

            relationScore:
                0,

            conceptCoverage:
                conceptCoverage,

            comparisonCoverage:
                conceptCoverage,

            matchType:
                "section",

            matchedTerms:
                [],

            matchedFamilies:
                [],

            matchedFamilyCount:
                0,

            matchedPhraseCount:
                0,

            matchedPhraseWeight:
                0,

            queryCoverage:
                conceptCoverage,

            weightedCoverage:
                conceptCoverage,

            familyCoverage:
                0,

            headingScore:
                0,

            headingCoverage:
                0,

            proximityScore:
                0,

            exactPhrase:
                false,

            searchSources:
                [
                    "heading-section"
                ]

        };

    }


    // ==================================
    // بناء المرشحين النهائيات
    // ==================================

    const candidatePool =
        [];


    // ==================================
    // أولًا:
    // الفقرات الناتجة من Orama
    // ==================================

    results.forEach(
        function (
            result
        ) {

            candidatePool.push({

                result:
                    result,

                sectionAnchor:
                    null,

                anchorScore:
                    0

            });

        }
    );


    // ==================================
    // ثانيًا:
    // الفقرات التابعة للعناوين القوية
    // ==================================

    strongAnchors.forEach(
        function (
            anchor
        ) {

            const sectionParagraphs =
                getSectionParagraphs(
                    anchor
                );


            sectionParagraphs.forEach(
                function (
                    paragraph
                ) {

                    const result =
                        getResultForParagraph(
                            paragraph
                        );


                    if (
                        !result
                    ) {

                        return;

                    }


                    candidatePool.push({

                        result:
                            result,

                        sectionAnchor:
                            anchor,

                        anchorScore:
                            anchor.anchorScore

                    });

                }
            );

        }
    );


    // ==================================
    // دمج المرشحين حسب الفقرة
    // ==================================

    const merged =
        new Map();


    candidatePool.forEach(
        function (
            item
        ) {

            if (
                !item ||
                !item.result
            ) {

                return;

            }


            const result =
                item.result;


            const paragraphIndex =
                Number(
                    result.paragraphIndex
                );


            if (
                Number.isNaN(
                    paragraphIndex
                )
            ) {

                return;

            }


            const existing =
                merged.get(
                    paragraphIndex
                );


            const weightedScore =
                Number(
                    result.score ||
                    0
                ) +
                Number(
                    item.anchorScore ||
                    0
                );


            if (
                !existing
            ) {

                merged.set(
                    paragraphIndex,
                    {

                        result:
                            result,

                        score:
                            weightedScore,

                        anchorScore:
                            Number(
                                item.anchorScore ||
                                0
                            ),

                        sectionAnchor:
                            item.sectionAnchor

                    }
                );

            }
            else {

                if (
                    weightedScore >
                    existing.score
                ) {

                    existing.score =
                        weightedScore;

                    existing.result =
                        result;

                    existing.anchorScore =
                        Number(
                            item.anchorScore ||
                            0
                        );

                    existing.sectionAnchor =
                        item.sectionAnchor;

                }

            }

        }
    );


    // ==================================
    // تحويل الخريطة إلى مصفوفة
    // ==================================

    const candidates =
        Array.from(
            merged.values()
        )
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        Number(
                            b.score
                        ) -
                        Number(
                            a.score
                        )
                    );

                }
            );


    // ==================================
    // الاختيار
    //
    // نمنع التكرار الشديد.
    // لكن نحافظ على المقاطع التابعة
    // للقسم القوي.
    // ==================================

    const selected =
        [];


    const selectedIndexes =
        new Set();


    // ==================================
    // معرفة القسم التابع للنتيجة
    // ==================================

    function getAnchorForParagraph(
        paragraphIndex
    ) {

        for (
            let i =
                strongAnchors.length - 1;

            i >= 0;

            i--
        ) {

            const anchor =
                strongAnchors[i];


            if (
                Number(
                    paragraphIndex
                ) >
                    Number(
                        anchor.startIndex
                    ) &&
                Number(
                    paragraphIndex
                ) <=
                    Number(
                        anchor.endIndex
                    )
            ) {

                return anchor;

            }

        }


        return null;

    }


    // ==================================
    // الاختيار المرحلي
    //
    // نعطي أولوية للمقاطع التي حصلت
    // على وزن عنوان قوي.
    // ==================================

    for (
        let i =
            0;

        i <
            candidates.length;

        i++
    ) {

        const candidateEntry =
            candidates[i];


        const candidate =
            candidateEntry.result;


        if (
            !candidate
        ) {

            continue;

        }


        const paragraphIndex =
            Number(
                candidate.paragraphIndex
            );


        if (
            selectedIndexes.has(
                paragraphIndex
            )
        ) {

            continue;

        }


        const candidateText =
            String(
                candidate.context ||
                candidate.text ||
                ""
            )
                .replace(
                    /\s+/g,
                    " "
                )
                .trim();


        if (
            !candidateText
        ) {

            continue;

        }


        // ==================================
        // منع التشابه الشديد
        // ==================================

        let similar =
            false;


        for (
            let j =
                0;

            j <
                selected.length;

            j++
        ) {

            const selectedText =
                String(
                    selected[j].result.context ||
                    selected[j].result.text ||
                    ""
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            if (
                !selectedText
            ) {

                continue;

            }


            const shorterLength =
                Math.min(
                    candidateText.length,
                    selectedText.length
                );


            const overlap =
                shorterLength >
                0
                    ? getCommonTextLength(
                        selectedText,
                        candidateText
                    ) /
                      shorterLength

                    : 0;


            if (
                overlap >=
                0.75
            ) {

                similar =
                    true;

                break;

            }

        }


        if (
            similar
        ) {

            continue;

        }


        // ==================================
        // المرساة الحالية
        // ==================================

        const anchor =
            candidateEntry.sectionAnchor ||
            getAnchorForParagraph(
                paragraphIndex
            );


        selected.push({

            result:
                candidate,

            score:
                candidateEntry.score,

            anchor:
                anchor

        });


        selectedIndexes.add(
            paragraphIndex
        );


        if (
            selected.length >=
            maxResults
        ) {

            break;

        }

    }


    // ==================================
    // إذا كان لدينا قسم قوي:
    //
    // نحاول ضمان عدم إرسال عنوانه فقط.
    // إذا كانت النتيجة المختارة مجرد عنوان
    // أو قصيرة جدًا، نضيف أفضل فقرة تالية
    // من القسم نفسه.
    // ==================================

    strongAnchors.forEach(
        function (
            anchor
        ) {

            const selectedFromAnchor =
                selected.filter(
                    function (
                        item
                    ) {

                        const paragraphIndex =
                            Number(
                                item.result.paragraphIndex
                            );


                        return (
                            paragraphIndex >
                                anchor.startIndex &&
                            paragraphIndex <=
                                anchor.endIndex
                        );

                    }
                );


            if (
                selectedFromAnchor.length >
                0
            ) {

                return;

            }


            const candidatesFromAnchor =
                candidates
                    .filter(
                        function (
                            item
                        ) {

                            if (
                                !item ||
                                !item.result
                            ) {

                                return false;

                            }


                            const paragraphIndex =
                                Number(
                                    item.result.paragraphIndex
                                );


                            return (
                                paragraphIndex >
                                    anchor.startIndex &&
                                paragraphIndex <=
                                    anchor.endIndex
                            );

                        }
                    )
                    .sort(
                        function (
                            a,
                            b
                        ) {

                            return (
                                Number(
                                    b.score
                                ) -
                                Number(
                                    a.score
                                )
                            );

                        }
                    );


            for (
                let i =
                    0;

                i <
                    candidatesFromAnchor.length;

                i++
            ) {

                const candidate =
                    candidatesFromAnchor[i];


                const paragraphIndex =
                    Number(
                        candidate.result.paragraphIndex
                    );


                if (
                    selectedIndexes.has(
                        paragraphIndex
                    )
                ) {

                    continue;

                }


                selected.push({

                    result:
                        candidate.result,

                    score:
                        candidate.score,

                    anchor:
                        anchor

                });


                selectedIndexes.add(
                    paragraphIndex
                );


                break;

            }

        }
    );


    // ==================================
    // ترتيب المختارات
    //
    // المقطع التابع لعنوان قوي يحتفظ
    // بقوة عنوانه في الترتيب.
    // ==================================

    selected.sort(
        function (
            a,
            b
        ) {

            return (
                Number(
                    b.score
                ) -
                Number(
                    a.score
                )
            );

        }
    );


    // ==================================
    // بناء السياق النهائي
    // ==================================

    const contexts =
        [];


    let totalChars =
        0;


    selected.forEach(
        function (
            selectedItem,
            index
        ) {

            if (
                totalChars >=
                maxChars
            ) {

                return;

            }


            const result =
                selectedItem.result;


            const main =
                String(
                    result.context ||
                    result.text ||
                    ""
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            if (
                !main
            ) {

                return;

            }


            let previous =
                includeNeighbors
                    ? String(
                        result.previousParagraphText ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim()

                    : "";


            let next =
                includeNeighbors
                    ? String(
                        result.nextParagraphText ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim()

                    : "";


            if (
                selectedIndexes.has(
                    Number(
                        result.paragraphIndex
                    ) - 1
                )
            ) {

                previous =
                    "";

            }


            if (
                selectedIndexes.has(
                    Number(
                        result.paragraphIndex
                    ) + 1
                )
            ) {

                next =
                    "";

            }


            // ==================================
            // المساحة المتاحة
            // ==================================

            const available =
                Math.max(
                    300,
                    maxChars -
                    totalChars -
                    300
                );


            let context =
                main.substring(
                    0,
                    available
                );


            // ==================================
            // السابق
            // ==================================

            if (
                previous &&
                context.length <
                    available -
                    150
            ) {

                context =
                    previous.substring(
                        Math.max(
                            0,
                            previous.length -
                            Math.max(
                                120,
                                available -
                                context.length -
                                1
                            )
                        )
                    ) +
                    " " +
                    context;

            }


            // ==================================
            // التالي
            // ==================================

            if (
                next &&
                context.length <
                    available -
                    150
            ) {

                context +=
                    " " +
                    next.substring(
                        0,
                        Math.max(
                            120,
                            available -
                            context.length -
                            1
                        )
                    );

            }


            // ==================================
            // العنوان المرجعي
            // ==================================

            let sectionHeading =
                result.heading ||
                "";


            if (
                selectedItem.anchor &&
                selectedItem.anchor.heading
            ) {

                sectionHeading =
                    selectedItem.anchor.heading.text ||
                    sectionHeading;

            }


            // ==================================
            // إنشاء السياق
            // ==================================

            const item = {

                rank:
                    index + 1,

                paragraphIndex:
                    result.paragraphIndex,

                paragraphId:
                    result.paragraphId,

                heading:
                    sectionHeading,

                headingScore:
                    Number(
                        result.headingScore ||
                        0
                    ),

                sectionAnchorScore:
                    Number(
                        selectedItem.score -
                        Number(
                            result.score ||
                            0
                        )
                    ),

                matchType:
                    selectedItem.anchor
                        ? "section"
                        : (
                            result.matchType ||
                            "orama"
                        ),

                previousParagraph:
                    previous,

                mainParagraph:
                    main,

                nextParagraph:
                    next,

                context:
                    context,

                score:
                    Number(
                        selectedItem.score ||
                        result.score ||
                        0
                    )

            };


            contexts.push(
                item
            );


            totalChars +=
                context.length;

        }
    );


    // ==================================
    // النص المرسل إلى الـAI
    // ==================================

    const text =
        contexts
            .map(
                function (
                    item
                ) {

                    let block =
                        "[مقطع " +
                        item.rank +
                        "]\n";


                    if (
                        item.heading
                    ) {

                        block +=
                            "العنوان: " +
                            item.heading +
                            "\n";

                    }


                    if (
                        item.previousParagraph
                    ) {

                        block +=
                            "السياق السابق: " +
                            item.previousParagraph +
                            "\n";

                    }


                    block +=
                        "المقطع المطابق: " +
                        item.mainParagraph;


                    if (
                        item.nextParagraph
                    ) {

                        block +=
                            "\nالسياق التالي: " +
                            item.nextParagraph;

                    }


                    return block;

                }
            )
            .join(
                "\n\n---\n\n"
            );


    // ==================================
    // النتيجة
    // ==================================

    return {

        query:
            searchResult.query ||
            "",

        count:
            Number(
                searchResult.count ||
                results.length
            ),

        selectedCount:
            contexts.length,

        totalOccurrences:
            Number(
                searchResult.indexedOccurrences ||
                0
            ),

        contexts:
            contexts,

        text:
            text

    };

}


// =====================================================
// Build AI Document Context
//
// الوظيفة:
// 1) استقبال سؤال المستخدم.
// 2) تشغيل محرك الاسترجاع العام.
// 3) بناء سياق مناسب للذكاء الاصطناعي.
// 4) حفظ مصادر الإحالات.
// 5) إعادة بيانات الاسترجاع والتحليل مع السياق.
//
// مبدأ مهم:
// المحرك يسترجع الأدلة.
// الذكاء الاصطناعي يفهم السؤال ويصوغ الإجابة.
// =====================================================

async function buildAIDocumentContext(
    query
) {

    // ==================================
    // التحقق من وجود مستند
    // ==================================

    if (
        !currentDocument
    ) {

        currentCitationSources =
            [];


        return {

            found:
                false,

            query:
                query,

            profile:
                "general",

            resultCount:
                0,

            selectedCount:
                0,

            totalOccurrences:
                0,

            matchedTerms:
                [],

            matchedFamilies:
                [],

            contexts:
                [],

            text:
                "",

            queryAnalysis:
                null

        };

    }


    try {

        // ==================================
        // تحديد نوع السؤال
        // ==================================

        const retrievalProfile =
            getRetrievalProfile(
                query
            );


        const profileType =
            retrievalProfile &&
            retrievalProfile.type
                ? retrievalProfile.type
                : "general";


        // ==================================
        // البحث في الفهرس
        // ==================================

        const searchResult =
            await searchIndexedDocument(
                currentDocument.id,
                query,
                {

                    profile:
                        profileType,

                    // نطلب عددًا واسعًا من المرشحين
                    // ثم تقتطع buildRetrievalContext
                    // العدد المناسب للسياق النهائي.
                    candidateLimit:
                        50

                }
            );


        // ==================================
        // التحقق من نتيجة البحث
        // ==================================

        if (
            !searchResult ||
            !Array.isArray(
                searchResult.results
            ) ||
            searchResult.results.length ===
                0
        ) {

            currentCitationSources =
                [];


            return {

                found:
                    false,

                query:
                    query,

                profile:
                    profileType,

                resultCount:
                    0,

                selectedCount:
                    0,

                totalOccurrences:
                    0,

                matchedTerms:
                    searchResult &&
                    Array.isArray(
                        searchResult.matchedTerms
                    )
                        ? searchResult.matchedTerms
                        : [],

                matchedFamilies:
                    searchResult &&
                    Array.isArray(
                        searchResult.matchedFamilies
                    )
                        ? searchResult.matchedFamilies
                        : [],

                contexts:
                    [],

                text:
                    "",

                queryAnalysis:
                    searchResult &&
                    searchResult.queryAnalysis
                        ? searchResult.queryAnalysis
                        : null

            };

        }


        // ==================================
        // إعدادات المزود والنموذج
        // ==================================

        const settings =
            getSavedSettings();


        const provider =
            settings &&
            settings.provider
                ? settings.provider
                : "";


        const model =
            settings &&
            settings.model
                ? settings.model
                : "";


        // ==================================
        // حدود الاسترجاع
        // ==================================

        const retrievalLimits =
            getRetrievalLimits(
                provider,
                model
            );


        let maxResults =
            Math.min(
                Number(
                    retrievalLimits.maxResults ||
                    6
                ),
                Number(
                    retrievalProfile.maxResults ||
                    6
                )
            );


        let maxChars =
            Math.min(
                Number(
                    retrievalLimits.maxChars ||
                    6000
                ),
                Number(
                    retrievalProfile.maxChars ||
                    6000
                )
            );


        // ==================================
        // حماية الحدود
        // ==================================

        if (
            !Number.isFinite(
                maxResults
            ) ||
            maxResults <=
                0
        ) {

            maxResults =
                6;

        }


        if (
            !Number.isFinite(
                maxChars
            ) ||
            maxChars <=
                0
        ) {

            maxChars =
                6000;

        }


        // ==================================
        // المقارنة تحتاج مساحة إضافية
        //
        // هذه ليست قاعدة للبحث نفسه.
        // فقط للسياق النهائي المرسل للـAI.
        // ==================================

        if (
            profileType ===
            "comparison"
        ) {

            maxResults =
                Math.min(
                    maxResults + 1,
                    6
                );


            maxChars =
                Math.min(
                    maxChars + 1000,
                    6500
                );

        }


        // ==================================
        // بناء السياق النهائي
        // ==================================

        const retrieval =
            await buildRetrievalContext(
                searchResult,
                {

                    maxResults:
                        maxResults,

                    maxChars:
                        maxChars,

                    includeNeighbors:
                        true

                }
            );


        // ==================================
        // التحقق من السياق
        // ==================================

        if (
            !retrieval ||
            typeof retrieval.text !==
                "string" ||
            !retrieval.text.trim()
        ) {

            currentCitationSources =
                [];


            return {

                found:
                    false,

                query:
                    query,

                profile:
                    profileType,

                resultCount:
                    Number(
                        searchResult.count ||
                        0
                    ),

                selectedCount:
                    0,

                totalOccurrences:
                    Number(
                        searchResult.indexedOccurrences ||
                        0
                    ),

                matchedTerms:
                    Array.isArray(
                        searchResult.matchedTerms
                    )
                        ? searchResult.matchedTerms
                        : [],

                matchedFamilies:
                    Array.isArray(
                        searchResult.matchedFamilies
                    )
                        ? searchResult.matchedFamilies
                        : [],

                contexts:
                    [],

                text:
                    "",

                queryAnalysis:
                    searchResult.queryAnalysis ||
                    null

            };

        }


        // ==================================
        // مصادر الإحالات
        //
        // نأخذها من السياق النهائي نفسه
        // حتى يكون ما يظهر للـAI
        // متوافقًا مع ما سنعرضه كمصدر.
        // ==================================

        const retrievedContexts =
            Array.isArray(
                retrieval.contexts
            )
                ? retrieval.contexts
                : [];


        currentCitationSources =
            retrievedContexts.map(
                function (
                    item,
                    index
                ) {

                    const paragraphIndex =
                        Number(
                            item.paragraphIndex
                        );


                    const paragraphId =
                        item.paragraphId !==
                            undefined &&
                        item.paragraphId !==
                            null

                            ? String(
                                item.paragraphId
                            )

                            : String(
                                paragraphIndex
                            );


                    const mainParagraph =
                        String(
                            item.mainParagraph ||
                            ""
                        ).trim();


                    const contextText =
                        String(
                            item.context ||
                            mainParagraph ||
                            ""
                        ).trim();


                    return {

                        // ----------------------------------
                        // رقم الإحالة
                        // ----------------------------------

                        citationIndex:
                            index + 1,

                        rank:
                            Number(
                                item.rank ||
                                index + 1
                            ),

                        paragraphIndex:
                            Number.isFinite(
                                paragraphIndex
                            )
                                ? paragraphIndex
                                : null,

                        paragraphId:
                            paragraphId,

                        heading:
                            String(
                                item.heading ||
                                ""
                            ).trim(),

                        mainParagraph:
                            mainParagraph,

                        text:
                            contextText,

                        score:
                            Number(
                                item.score ||
                                0
                            ),

                        matchType:
                            item.matchType ||
                            "word"

                    };

                }
            );


        // ==================================
        // النتيجة النهائية
        // ==================================

        return {

            found:
                true,

            query:
                query,

            profile:
                profileType,

            // ----------------------------------
            // معلومات الاسترجاع
            // ----------------------------------

            resultCount:
                Number(
                    searchResult.count ||
                    0
                ),

            selectedCount:
                Number(
                    retrievedContexts.length
                ),

            totalOccurrences:
                Number(
                    searchResult.indexedOccurrences ||
                    retrieval.totalOccurrences ||
                    0
                ),

            // ----------------------------------
            // الكلمات والعائلات
            // ----------------------------------

            matchedTerms:
                Array.isArray(
                    searchResult.matchedTerms
                )
                    ? searchResult.matchedTerms
                    : [],

            matchedFamilies:
                Array.isArray(
                    searchResult.matchedFamilies
                )
                    ? searchResult.matchedFamilies
                    : [],

            // ----------------------------------
            // تحليل السؤال
            // ----------------------------------

            queryAnalysis:
                searchResult.queryAnalysis ||
                null,

            // ----------------------------------
            // النتائج الخام
            //
            // مفيدة للطبقات اللاحقة،
            // لكنها لا تُرسل كلها للـAI.
            // ----------------------------------

            results:
                Array.isArray(
                    searchResult.results
                )
                    ? searchResult.results
                    : [],

            // ----------------------------------
            // السياق النهائي
            // ----------------------------------

            contexts:
                retrievedContexts,

            text:
                String(
                    retrieval.text ||
                    ""
                ),

            // ----------------------------------
            // الإحالات
            // ----------------------------------

            citations:
                currentCitationSources

        };

    }
    catch (
        error
    ) {

        currentCitationSources =
            [];


        console.error(
            "❌ خطأ buildAIDocumentContext:",
            error
        );


        console.error(
            "رسالة الخطأ:",
            error &&
            error.message
                ? error.message
                : error
        );


        console.error(
            "Stack:",
            error &&
            error.stack
                ? error.stack
                : "لا يوجد Stack"
        );


        return {

            found:
                false,

            query:
                query,

            profile:
                "general",

            resultCount:
                0,

            selectedCount:
                0,

            totalOccurrences:
                0,

            matchedTerms:
                [],

            matchedFamilies:
                [],

            contexts:
                [],

            text:
                "",

            queryAnalysis:
                null,

            error:
                error &&
                error.message
                    ? error.message
                    : String(
                        error
                    )

        };

    }

}


// =====================================================
// Temporary Retrieval Test
// اختبار السياق الجديد
// =====================================================

window.testRetrieval =
    async function (
        query
    ) {

        try {

            if (!currentDocument) {

                console.warn(
                    "لا يوجد مستند نشط."
                );

                return null;

            }


            const context =
                await buildAIDocumentContext(
                    query
                );


            console.log(
                "======================================"
            );


            console.log(
                "اختبار استرجاع Orama"
            );


            console.log(
                "السؤال:",
                query
            );


            console.log(
                "النوع:",
                context.profile
            );


            console.log(
                "عدد النتائج:",
                context.resultCount
            );


            console.log(
                "المقاطع المختارة:",
                context.selectedCount
            );


            console.log(
                "العائلات:",
                context.matchedFamilies
            );


            console.log(
                "المصادر:",
                currentCitationSources
            );


            console.log(
                "السياق:",
                context.text
            );


            console.log(
                "======================================"
            );


            return context;

        }
        catch (error) {

            console.error(
                "فشل اختبار الاسترجاع:",
                error
            );


            return null;

        }

    };


// =====================================================
// تحديث إحصاءات Orama في المستند
// =====================================================

async function updateOramaDocumentStats(
    documentItem
) {

    if (!documentItem) {

        return;

    }


    try {

        const structureData =
            await ensureDocumentStructure(
                documentItem
            );


        if (!structureData) {

            return;

        }


        const paragraphs =
            Array.isArray(
                structureData.paragraphs
            )
                ? structureData.paragraphs
                : [];


        let tokenCount =
            0;


        const terms =
            new Set();


        const families =
            new Set();


        paragraphs.forEach(
            function (
                paragraph
            ) {

                const tokens =
                    tokenizeDocumentText(
                        paragraph &&
                        paragraph.text
                            ? paragraph.text
                            : ""
                    );


                tokens.forEach(
                    function (
                        token
                    ) {

                        tokenCount +=
                            1;


                        terms.add(
                            token
                        );


                        const family =
                            getConservativeFamilyKey(
                                token,
                                null
                            );


                        if (family) {

                            families.add(
                                family
                            );

                        }

                    }
                );

            }
        );


        documentItem.indexTokenCount =
            tokenCount;


        documentItem.indexUniqueTerms =
            terms.size;


        documentItem.indexUniqueFamilies =
            families.size;


        documentItem.oramaSchemaVersion =
            ORAMA_SCHEMA_VERSION;


        documentItem.indexStatus =
            "indexed";


        documentItem.indexUpdatedAt =
            new Date().toISOString();


        saveDocuments();

    }
    catch (error) {

        console.warn(
            "تعذر تحديث إحصاءات Orama:",
            error
        );

    }

}


// =====================================================
// Build Orama Index For Current Document
// واجهة موحدة للفهرسة الجديدة
// =====================================================

async function buildOramaIndexForDocument(
    documentItem
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    updateDocumentIndexStatus(
        documentItem,
        "indexing"
    );


    try {

        const structureData =
            await ensureDocumentStructure(
                documentItem
            );


        if (!structureData) {

            throw new Error(
                "تعذر الحصول على بنية المستند."
            );

        }


        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


        await updateOramaDocumentStats(
            documentItem
        );


        return oramaRetrievalDb;

    }
    catch (error) {

        updateDocumentIndexStatus(
            documentItem,
            "error"
        );


        throw error;

    }

}


// =====================================================
// تحديث الفهرس عند تغيير المستند الحالي
// =====================================================

async function ensureOramaDocumentReady(
    documentItem
) {

    if (!documentItem) {

        return null;

    }


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (!structureData) {

        return null;

    }


    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    if (
        documentItem.oramaSchemaVersion !==
        ORAMA_SCHEMA_VERSION
    ) {

        await updateOramaDocumentStats(
            documentItem
        );

    }


    return db;

}
// =====================================================
// Research Tools
// PART 4
// الإحالات + Word + عرض المحادثة
// =====================================================





// =====================================================
// FORMAT AI MESSAGE
// Markdown + Interactive Citations
// =====================================================

function formatAIMessage(
    text,
    citationSources,
    fixedCitationGroupId
) {

    if (
        !text
    ) {

        return "";

    }


    const sources =
        Array.isArray(
            citationSources
        )
            ? citationSources
            : [];


    // ==================================
    // إنشاء هوية ثابتة لمجموعة الإحالات
    // الخاصة بهذه الرسالة
    // ==================================

    const citationGroupId =
        fixedCitationGroupId ||
        (
            "citation-" +
            Date.now() +
            "-" +
            Math.random()
                .toString(
                    36
                )
                .substring(
                    2,
                    10
                )
        );


    // ==================================
    // حفظ مصادر هذه الرسالة
    // ==================================

    citationRegistry.set(
        citationGroupId,
        sources.map(
            function (
                source
            ) {

                return {

                    rank:
                        Number(
                            source.rank ||
                            0
                        ),

                    paragraphIndex:
                        source.paragraphIndex !==
                            undefined
                            ? Number(
                                source.paragraphIndex
                            )
                            : null,

                    paragraphId:
                        source.paragraphId ||
                        "",

                    heading:
                        source.heading ||
                        "",

                    mainParagraph:
                        source.mainParagraph ||
                        source.text ||
                        "",

                    text:
                        source.text ||
                        source.mainParagraph ||
                        ""

                };

            }
        )
    );


    try {

        let html =
            marked.parse(
                String(
                    text
                ),
                {

                    breaks:
                        true,

                    gfm:
                        true

                }
            );


        // ==================================
        // الإحالات المجمعة
        // ==================================

        html =
            html.replace(
                /\[(مقطع\s*[0-9٠-٩]+(?:\s*(?:،|,|و)\s*(?:مقطع\s*)?[0-9٠-٩]+)*)\]/g,
                function (
                    match,
                    citationBody
                ) {

                    const normalized =
                        String(
                            citationBody
                        )
                            .replace(
                                /[٠-٩]/g,
                                function (
                                    digit
                                ) {

                                    return String(
                                        "٠١٢٣٤٥٦٧٨٩".indexOf(
                                            digit
                                        )
                                    );

                                }
                            );


                    const rankMatches =
                        normalized.match(
                            /\d+/g
                        );


                    if (
                        !Array.isArray(
                            rankMatches
                        )
                    ) {

                        return match;

                    }


                    const ranks =
                        [];


                    rankMatches.forEach(
                        function (
                            value
                        ) {

                            const rank =
                                Number(
                                    value
                                );


                            if (
                                !Number.isFinite(
                                    rank
                                ) ||
                                rank <=
                                    0
                            ) {

                                return;

                            }


                            if (
                                !ranks.includes(
                                    rank
                                )
                            ) {

                                ranks.push(
                                    rank
                                );

                            }

                        }
                    );


                    if (
                        ranks.length ===
                        0
                    ) {

                        return match;

                    }


                    // ==================================
                    // كل إحالة زر مستقل
                    // لكن جميعها تحمل نفس المجموعة
                    // ==================================

                    return ranks
                        .map(
                            function (
                                rank
                            ) {

                                const source =
                                    sources.find(
                                        function (
                                            item
                                        ) {

                                            return (
                                                Number(
                                                    item.rank
                                                ) ===
                                                rank
                                            );

                                        }
                                    );


                                if (
                                    !source
                                ) {

                                    return (
                                        "[مقطع " +
                                        rank +
                                        "]"
                                    );

                                }


                                return (

                                    '<button ' +
                                    'type="button" ' +
                                    'class="document-citation" ' +
                                    'data-citation-group="' +
                                    citationGroupId +
                                    '" ' +
                                    'data-citation-rank="' +
                                    rank +
                                    '" ' +
                                    'title="الانتقال إلى المقطع ' +
                                    rank +
                                    '">' +

                                    "[مقطع " +
                                    rank +
                                    "]" +

                                    "</button>"

                                );

                            }
                        )
                        .join(
                            " "
                        );

                }
            );


        return html;

    }
    catch (
        error
    ) {

        console.warn(
            "فشل تنسيق رسالة الذكاء الاصطناعي:",
            error
        );


        return String(
            text
        ).replace(
            /\n/g,
            "<br>"
        );

    }

}


// =====================================================
// Citation Click Handler
//
// كل زر إحالة يحمل:
// 1) رقم المقطع
// 2) مجموعة الإحالات الخاصة برسالته
//
// لذلك يمكن أن تكون هناك:
// [مقطع 1] في رسالة أولى
// [مقطع 1] في رسالة ثانية
//
// ولكل واحدة مصدر مختلف.
// =====================================================

if (chatArea) {

    chatArea.addEventListener(
        "click",
        function (
            event
        ) {

            const citation =
                event.target.closest(
                    ".document-citation"
                );


            if (
                !citation
            ) {

                return;

            }


            event.preventDefault();

            event.stopPropagation();


            // ==================================
            // مجموعة إحالات هذه الرسالة
            // ==================================

            const groupId =
                citation.getAttribute(
                    "data-citation-group"
                );


            // ==================================
            // رقم المقطع
            // ==================================

            const rank =
                Number(
                    citation.getAttribute(
                        "data-citation-rank"
                    )
                );


            if (
                !groupId ||
                !Number.isFinite(
                    rank
                )
            ) {

                console.warn(
                    "بيانات الإحالة غير صالحة:",
                    {
                        groupId:
                            groupId,

                        rank:
                            rank
                    }
                );

                return;

            }


            // ==================================
            // استرجاع مصادر الرسالة نفسها
            // ==================================

            const sources =
                citationRegistry.get(
                    groupId
                );


            if (
                !Array.isArray(
                    sources
                )
            ) {

                console.warn(
                    "لم يتم العثور على مجموعة الإحالات:",
                    groupId
                );

                return;

            }


            // ==================================
            // البحث عن المقطع داخل
            // مجموعة الرسالة نفسها
            // ==================================

            const source =
                sources.find(
                    function (
                        item
                    ) {

                        return (
                            item &&
                            Number(
                                item.rank
                            ) ===
                            rank
                        );

                    }
                );


            if (
                !source
            ) {

                console.warn(
                    "لم يتم العثور على مصدر الإحالة:",
                    {
                        groupId:
                            groupId,

                        rank:
                            rank
                    }
                );

                return;

            }


            // ==================================
            // توافق مع openCitationInWord(rank)
            //
            // لا نغيّر الدالة التي ثبت أنها تعمل.
            // نجعل currentCitationSources مؤقتًا
            // يساوي مصادر الرسالة التي نقرنا عليها.
            // ==================================

            currentCitationSources =
                sources;


            // ==================================
            // الانتقال إلى المقطع
            // ==================================

            openCitationInWord(
                rank
            );

        }
    );

}


// =====================================================
// Find Citation Source
// =====================================================

function getCitationSource(
    rank
) {

    if (
        !Array.isArray(
            currentCitationSources
        )
    ) {

        return null;

    }


    return (
        currentCitationSources.find(
            function (
                item
            ) {

                return (
                    Number(
                        item.rank
                    ) ===
                    Number(
                        rank
                    )
                );

            }
        ) ||
        null
    );

}


// =====================================================
// Normalize Citation Search Text
// =====================================================

function prepareCitationSearchText(
    text,
    maxLength
) {

    let value =
        String(
            text ||
            ""
        )
            .replace(
                /\s+/g,
                " "
            )
            .trim();


    if (!value) {

        return "";

    }


    const limit =
        typeof maxLength ===
            "number"
                ? maxLength
                : 180;


    if (
        value.length >
        limit
    ) {

        value =
            value
                .substring(
                    0,
                    limit
                )
                .trim();

    }


    return value;

}


// =====================================================
// Search Citation In Word
// البحث عن النص الأصلي داخل Word
// =====================================================

async function searchCitationInWord(
    body,
    searchText
) {

    if (
        !body ||
        !searchText
    ) {

        return [];

    }


    const results =
        body.search(
            searchText,
            {

                matchCase:
                    false,

                matchWholeWord:
                    false,

                matchWildcards:
                    false,

                ignorePunct:
                    true,

                ignoreSpace:
                    true

            }
        );


    results.load(
        "items"
    );


    await results.context.sync();


    return results.items || [];

}





// =====================================================
// Open Citation In Word
// =====================================================

async function openCitationInWord(
    rank
) {

    const source =
        Array.isArray(
            currentCitationSources
        )
            ? currentCitationSources.find(
                function (
                    item
                ) {

                    return (
                        Number(
                            item.rank
                        ) ===
                        Number(
                            rank
                        )
                    );

                }
            )
            : null;


    if (
        !source
    ) {

        console.warn(
            "لم يتم العثور على مصدر الإحالة:",
            rank
        );

        return;

    }


    if (
        typeof Word ===
        "undefined"
    ) {

        console.warn(
            "Word غير متاح."
        );

        return;

    }


    try {

        await Word.run(
            async function (
                context
            ) {

                const body =
                    context.document.body;


                // ==================================
                // 1) الأولوية لموضع الفقرة
                // ==================================

                const paragraphIndex =
                    Number(
                        source.paragraphIndex
                    );


                if (
                    Number.isFinite(
                        paragraphIndex
                    )
                ) {

                    const paragraphs =
                        body.paragraphs;


                    paragraphs.load(
                        "items/text"
                    );


                    await context.sync();


                    const items =
                        paragraphs.items;


                    const directIndexes =
                        [];


                    directIndexes.push(
                        paragraphIndex
                    );


                    if (
                        paragraphIndex >
                        0
                    ) {

                        directIndexes.push(
                            paragraphIndex - 1
                        );

                    }


                    directIndexes.push(
                        paragraphIndex + 1
                    );


                    for (
                        let i =
                            0;

                        i <
                            directIndexes.length;

                        i++
                    ) {

                        const index =
                            directIndexes[i];


                        if (
                            index <
                            0 ||
                            index >=
                            items.length
                        ) {

                            continue;

                        }


                        const paragraph =
                            items[index];


                        const paragraphText =
                            String(
                                paragraph.text ||
                                ""
                            ).trim();


                        if (
                            !paragraphText
                        ) {

                            continue;

                        }


                        const sourceText =
                            String(
                                source.mainParagraph ||
                                source.text ||
                                ""
                            ).trim();


                        if (
                            !sourceText
                        ) {

                            continue;

                        }


                        const normalizedSource =
                            normalizeSearchText(
                                sourceText
                            );


                        const normalizedParagraph =
                            normalizeSearchText(
                                paragraphText
                            );


                        const sourcePrefix =
                            normalizedSource.substring(
                                0,
                                Math.min(
                                    120,
                                    normalizedSource.length
                                )
                            );


                        const paragraphPrefix =
                            normalizedParagraph.substring(
                                0,
                                Math.min(
                                    120,
                                    normalizedParagraph.length
                                )
                            );


                        if (
                            (
                                sourcePrefix &&
                                paragraphPrefix &&
                                (
                                    normalizedParagraph.includes(
                                        sourcePrefix
                                    ) ||
                                    normalizedSource.includes(
                                        paragraphPrefix
                                    )
                                )
                            ) ||
                            index ===
                            paragraphIndex
                        ) {

                            paragraph.select(
                                "Select"
                            );


                            await context.sync();


                            return;

                        }

                    }

                }


                // ==================================
                // 2) البحث بالنص
                // ==================================

                let searchText =
                    String(
                        source.mainParagraph ||
                        source.text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim();


                if (
                    !searchText
                ) {

                    throw new Error(
                        "لا يوجد نص صالح للمقطع."
                    );

                }


                const originalSearchText =
                    searchText;


                if (
                    searchText.length >
                    220
                ) {

                    searchText =
                        searchText.substring(
                            0,
                            220
                        ).trim();

                }


                let results =
                    body.search(
                        searchText,
                        {

                            matchCase:
                                false,

                            matchWholeWord:
                                false,

                            matchWildcards:
                                false,

                            ignorePunct:
                                true,

                            ignoreSpace:
                                true

                        }
                    );


                results.load(
                    "items"
                );


                await context.sync();


                if (
                    results.items.length >
                    0
                ) {

                    results.items[0].select(
                        "Select"
                    );


                    await context.sync();


                    return;

                }


                // ==================================
                // 3) مقتطف أقصر
                // ==================================

                let fallback =
                    originalSearchText
                        .substring(
                            0,
                            Math.min(
                                100,
                                originalSearchText.length
                            )
                        )
                        .trim();


                if (
                    fallback.length >
                    20
                ) {

                    results =
                        body.search(
                            fallback,
                            {

                                matchCase:
                                    false,

                                matchWholeWord:
                                    false,

                                matchWildcards:
                                    false,

                                ignorePunct:
                                    true,

                                ignoreSpace:
                                    true

                            }
                        );


                    results.load(
                        "items"
                    );


                    await context.sync();


                    if (
                        results.items.length >
                        0
                    ) {

                        results.items[0].select(
                            "Select"
                        );


                        await context.sync();


                        return;

                    }

                }


                // ==================================
                // 4) النص المنظف
                // ==================================

                const normalized =
                    normalizeSearchText(
                        originalSearchText
                    );


                fallback =
                    normalized.substring(
                        0,
                        Math.min(
                            100,
                            normalized.length
                        )
                    ).trim();


                if (
                    fallback
                ) {

                    results =
                        body.search(
                            fallback,
                            {

                                matchCase:
                                    false,

                                matchWholeWord:
                                    false,

                                matchWildcards:
                                    false,

                                ignorePunct:
                                    true,

                                ignoreSpace:
                                    true

                            }
                        );


                    results.load(
                        "items"
                    );


                    await context.sync();


                    if (
                        results.items.length >
                        0
                    ) {

                        results.items[0].select(
                            "Select"
                        );


                        await context.sync();


                        return;

                    }

                }


                throw new Error(
                    "لم يتم العثور على نص المقطع في المستند."
                );

            }
        );

    }
    catch (
        error
    ) {

        console.error(
            "تعذر الانتقال إلى المقطع:",
            error
        );

    }

}


// =====================================================
// Render Chat
// عرض المحادثة
// =====================================================

function renderChat() {

    if (!chatArea) {

        return;

    }


    chatArea.innerHTML =
        "";


    if (!currentChat) {

        chatArea.innerHTML = `
            <div class="welcome">

                <div class="ai-symbol">
                    ✦
                </div>

                <h2>
                    مرحبًا بك
                </h2>

                <p>
                    ابدأ محادثة جديدة
                </p>

            </div>
        `;

        return;

    }


    if (
        !Array.isArray(
            currentChat.messages
        )
    ) {

        currentChat.messages =
            [];

    }


    currentChat.messages.forEach(
        function (
            msg
        ) {

            if (!msg) {

                return;

            }


            const div =
                document.createElement(
                    "div"
                );


            div.className =
                "message " +
                (
                    msg.role ===
                    "user"
                        ? "user-message"
                        : "ai-message"
                );


            // ==================================
            // رسالة المستخدم
            // ==================================

            if (
                msg.role ===
                "user"
            ) {

                div.textContent =
                    msg.text ||
                    "";

            }


            // ==================================
            // رسالة AI
            // ==================================

            else {

                div.innerHTML =
                    formatAIMessage(
                        msg.text ||
                        "",

                        Array.isArray(
                            msg.citationSources
                        )
                            ? msg.citationSources
                            : []
                    );

            }


            chatArea.appendChild(
                div
            );

        }
    );


    chatArea.scrollTop =
        chatArea.scrollHeight;

}


// =====================================================
// Save Chat
// =====================================================

function saveChat(
    chat
) {

    if (!chat) {

        return;

    }


    if (
        !Array.isArray(
            chats
        )
    ) {

        chats =
            [];

    }


    const existingIndex =
        chats.findIndex(
            function (
                item
            ) {

                return (
                    item &&
                    item.id ===
                        chat.id
                );

            }
        );


    if (
        existingIndex ===
        -1
    ) {

        chats.unshift(
            chat
        );

    }
    else {

        chats[
            existingIndex
        ] =
            chat;

    }


    saveChats();

}



// =====================================================
// Clone Citation Sources
//
// يحفظ جميع البيانات اللازمة للإحالة:
// - rank
// - paragraphIndex
// - paragraphId
// - heading
// - mainParagraph
// - text
//
// حتى تبقى الإحالة صالحة بعد إغلاق الدردشة وفتحها.
// =====================================================

function cloneCitationSources(
    sources
) {

    if (
        !Array.isArray(
            sources
        )
    ) {

        return [];

    }


    return sources.map(
        function (
            source
        ) {

            if (
                !source
            ) {

                return null;

            }


            return {

                rank:
                    Number(
                        source.rank ||
                        0
                    ),

                paragraphIndex:
                    source.paragraphIndex !==
                        undefined
                        ? Number(
                            source.paragraphIndex
                        )
                        : null,

                paragraphId:
                    source.paragraphId !==
                        undefined &&
                    source.paragraphId !==
                        null
                        ? String(
                            source.paragraphId
                        )
                        : "",

                heading:
                    String(
                        source.heading ||
                        ""
                    ),

                mainParagraph:
                    String(
                        source.mainParagraph ||
                        source.text ||
                        ""
                    ),

                text:
                    String(
                        source.text ||
                        source.mainParagraph ||
                        ""
                    ),

                score:
                    Number(
                        source.score ||
                        0
                    ),

                matchType:
                    source.matchType ||
                    "word"

            };

        }
    ).filter(
        function (
            source
        ) {

            return (
                source &&
                Number.isFinite(
                    source.rank
                ) &&
                source.rank >
                    0
            );

        }
    );

}


// =====================================================
// Add AI Message
// =====================================================

function addAIMessage(
    chat,
    text,
    citationSources
) {

    if (!chat) {

        return;

    }


    if (
        !Array.isArray(
            chat.messages
        )
    ) {

        chat.messages =
            [];

    }


    chat.messages.push({

        role:
            "ai",

        text:
            String(
                text ||
                ""
            ),

        citationSources:
            cloneCitationSources(
                citationSources
            )

    });


    chat.updatedAt =
        new Date().toISOString();


    saveChat(
        chat
    );

}


// =====================================================
// Add User Message
// =====================================================

function addUserMessage(
    chat,
    text
) {

    if (!chat) {

        return;

    }


    if (
        !Array.isArray(
            chat.messages
        )
    ) {

        chat.messages =
            [];

    }


    chat.messages.push({

        role:
            "user",

        text:
            String(
                text ||
                ""
            )

    });


    chat.updatedAt =
        new Date().toISOString();


    saveChat(
        chat
    );

}


// =====================================================
// Citation Diagnostics
// =====================================================

window.testCitation =
    async function (
        rank
    ) {

        try {

            const source =
                getCitationSource(
                    rank
                );


            if (!source) {

                console.warn(
                    "لا يوجد مصدر للإحالة:",
                    rank
                );

                return null;

            }


            console.log(
                "مصدر الإحالة:",
                source
            );


            await openCitationInWord(
                rank
            );


            return source;

        }
        catch (error) {

            console.error(
                "فشل اختبار الإحالة:",
                error
            );


            return null;

        }

    };


// =====================================================
// Citation Sources Diagnostics
// =====================================================

window.testCitationSources =
    function () {

        console.log(
            "مصادر الإحالات الحالية:",
            currentCitationSources
        );


        return (
            Array.isArray(
                currentCitationSources
            )
                ? currentCitationSources
                : []
        );

    };
    // =====================================================
// Research Tools
// PART 5
// المحادثات + المشاريع + الإعدادات
// =====================================================


// =====================================================
// Projects Storage
// =====================================================

function saveProjects() {

    localStorage.setItem(
        "WORD_AI_PROJECTS",
        JSON.stringify(
            projects
        )
    );

}


// =====================================================
// Chats Storage
// =====================================================

let chats = [];

// =====================================================
// Citation Registry
//
// كل مجموعة إحالات مرتبطة برسالة معينة.
// لا نعتمد على currentCitationSources
// عند إعادة عرض الرسائل القديمة.
// =====================================================

const citationRegistry =
    new Map();


let currentChat = null;

try {

    chats =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_CHATS"
            )
        ) || [];

}
catch (error) {

    console.warn(
        "تعذر تحميل المحادثات:",
        error
    );

    chats = [];

}


// =====================================================
// Save Chats
// =====================================================

function saveChats() {

    localStorage.setItem(
        "WORD_AI_CHATS",
        JSON.stringify(
            chats
        )
    );

}


// =====================================================
// Normalize Chats
// الحفاظ على بنية المحادثات القديمة
// =====================================================

chats =
    chats
        .filter(
            function (
                chat
            ) {

                return (
                    chat &&
                    typeof chat ===
                        "object"
                );

            }
        )
        .map(
            function (
                chat
            ) {

                const now =
                    new Date().toISOString();


                return {

                    id:
                        chat.id ||
                        Date.now(),

                    title:
                        chat.title ||
                        "محادثة جديدة",

                    messages:
                        Array.isArray(
                            chat.messages
                        )
                            ? chat.messages
                            : [],

                    isTemporary:
                        Boolean(
                            chat.isTemporary
                        ),

                    projectId:
                        chat.projectId ||
                        null,

                    createdAt:
                        chat.createdAt ||
                        now,

                    updatedAt:
                        chat.updatedAt ||
                        now

                };

            }
        );


// =====================================================
// Normalize Chat Messages
// =====================================================

chats.forEach(
    function (
        chat
    ) {

        chat.messages =
            chat.messages.map(
                function (
                    message
                ) {

                    if (
                        !message ||
                        typeof message !==
                            "object"
                    ) {

                        return null;

                    }


                    return {

                        role:
                            message.role ===
                            "ai"
                                ? "ai"
                                : "user",

                        text:
                            String(
                                message.text ||
                                ""
                            ),

                        citationSources:
                            Array.isArray(
                                message.citationSources
                            )
                                ? cloneCitationSources(
                                    message.citationSources
                                )
                                : []

                    };

                }
            )
            .filter(
                function (
                    message
                ) {

                    return message !==
                        null;

                }
            );

    }
);


// =====================================================
// Current Chat Helpers
// =====================================================

function createNewChat() {

    currentChat = {

        id:
            Date.now(),

        title:
            "محادثة جديدة",

        messages:
            [],

        isTemporary:
            true,

        projectId:
            null,

        createdAt:
            new Date().toISOString(),

        updatedAt:
            new Date().toISOString()

    };


    currentCitationSources =
        [];


    if (input) {

        input.value =
            "";

        input.style.height =
            "auto";

    }


    renderChat();

}


// =====================================================
// Finalize Temporary Chat
// تحويل المحادثة المؤقتة إلى محفوظة
// =====================================================

function finalizeCurrentChat(
    firstMessage
) {

    if (!currentChat) {

        return null;

    }


    if (
        !currentChat.isTemporary
    ) {

        return currentChat;

    }


    const text =
        String(
            firstMessage ||
            ""
        ).trim();


    currentChat.isTemporary =
        false;


    const chatProject =
        currentChat.projectId
            ? projects.find(
                function (
                    project
                ) {

                    return (
                        project &&
                        String(
                            project.id
                        ) ===
                        String(
                            currentChat.projectId
                        )
                    );

                }
            )
            : null;


    currentChat.title =
        text !== ""
            ? text.substring(
                0,
                30
            )
            : "محادثة جديدة";


    currentChat.updatedAt =
        new Date().toISOString();


    const exists =
        chats.some(
            function (
                chat
            ) {

                return (
                    chat &&
                    chat.id ===
                        currentChat.id
                );

            }
        );


    if (!exists) {

        chats.unshift(
            currentChat
        );

    }


    // ==================================
    // ربط المحادثة بالمشروع
    // ==================================

    if (
        chatProject
    ) {

        if (
            !Array.isArray(
                chatProject.chatIds
            )
        ) {

            chatProject.chatIds =
                [];

        }


        if (
            !chatProject.chatIds.includes(
                currentChat.id
            )
        ) {

            chatProject.chatIds.push(
                currentChat.id
            );


            chatProject.updatedAt =
                new Date().toISOString();


            saveProjects();

        }

    }


    saveChats();


    return currentChat;

}


// =====================================================
// Project Helpers
// =====================================================

function getCurrentProject() {

    return currentProject || null;

}



// =====================================================
// Create Project
// =====================================================

function createProject(
    name
) {

    const projectName =
        String(
            name ||
            ""
        ).trim();


    if (!projectName) {

        return null;

    }


    const now =
        new Date().toISOString();


    const project = {

        id:
            Date.now(),

        name:
            projectName,

        createdAt:
            now,

        updatedAt:
            now,

        documents:
            [],

        references:
            [],

        chatIds:
            [],

        settings: {

            citationStyle:
                "",

            notes:
                ""

        }

    };


    projects.unshift(
        project
    );


    saveProjects();


    renderProjects();




    return project;

}



// =====================================================
// AI Settings
// =====================================================

function getSavedSettings() {

    try {

        return (
            JSON.parse(
                localStorage.getItem(
                    "AI_SETTINGS"
                )
            ) || {}
        );

    }
    catch (error) {

        console.warn(
            "تعذر قراءة إعدادات الذكاء الاصطناعي:",
            error
        );


        return {};

    }

}


// =====================================================
// Save AI Settings
// =====================================================

function saveAISettings(
    data
) {

    const settings = {

        provider:
            String(
                data &&
                data.provider ||
                "openrouter"
            ),

        key:
            String(
                data &&
                data.key ||
                ""
            ),

        model:
            String(
                data &&
                data.model ||
                ""
            )

    };


    localStorage.setItem(
        "AI_SETTINGS",
        JSON.stringify(
            settings
        )
    );


    return settings;

}


// =====================================================
// Apply Saved Settings To UI
// =====================================================

function loadSettings() {

    const data =
        getSavedSettings();


    if (provider) {

        provider.value =
            data.provider ||
            "openrouter";

    }


    if (apiKey) {

        apiKey.value =
            data.key ||
            "";

    }


    if (modelSelect) {

        const savedModel =
            data.model ||
            "";


        modelSelect.innerHTML =
            "";


        if (
            savedModel !==
            ""
        ) {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                savedModel;


            option.textContent =
                savedModel;


            modelSelect.appendChild(
                option
            );


            modelSelect.value =
                savedModel;

        }
        else {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                "";


            option.textContent =
                "أدخل المفتاح ثم حدّث النماذج";


            modelSelect.appendChild(
                option
            );

        }

    }


    // ==================================
    // مزامنة الحالة الداخلية
    // ==================================

    AppState.aiSettings = {

        provider:
            data.provider ||
            "openrouter",

        key:
            data.key ||
            "",

        model:
            data.model ||
            ""

    };


    updateProviderInfo();

}


// =====================================================
// Get Active AI Settings
// =====================================================

function getActiveAISettings() {

    const saved =
        getSavedSettings();


    const active = {

        provider:
            (
                saved.provider ||
                (
                    provider
                        ? provider.value
                        : "openrouter"
                ) ||
                "openrouter"
            ).toLowerCase(),

        key:
            saved.key ||
            (
                apiKey
                    ? apiKey.value.trim()
                    : ""
            ) ||
            "",

        model:
            saved.model ||
            (
                modelSelect
                    ? modelSelect.value.trim()
                    : ""
            ) ||
            ""

    };


    AppState.aiSettings =
        active;


    return active;

}


// =====================================================
// Provider Info
// =====================================================

function updateProviderInfo() {

    if (
        !providerInfo ||
        !provider
    ) {

        return;

    }


    const value =
        provider.value;


    if (
        value ===
        "openrouter"
    ) {

        providerInfo.innerHTML =
            "OpenRouter: سيتم جلب النماذج المجانية المتاحة من حسابك.";

        return;

    }


    if (
        value ===
        "gemini"
    ) {

        providerInfo.innerHTML =
            "Gemini: سيتم جلب النماذج التي تدعم generateContent.";

        return;

    }


    if (
        value ===
        "groq"
    ) {

        providerInfo.innerHTML =
            "Groq: سيتم جلب النماذج المتاحة من حسابك.";

        return;

    }


    if (
        value ===
        "openai"
    ) {

        providerInfo.innerHTML =
            "OpenAI: سيتم جلب النماذج المتاحة من حسابك.";

        return;

    }


    providerInfo.innerHTML =
        "سيتم تحديد رابط الاتصال حسب مزود الذكاء الاصطناعي.";

}


// =====================================================
// Settings Button
// =====================================================

if (settingsBtn) {

    settingsBtn.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (projectsPopup) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.add(
                    "open"
                );

            }


            loadSettings();

        };

}


// =====================================================
// Close Settings
// =====================================================

if (closeSettings) {

    closeSettings.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }

        };

}


// =====================================================
// Show / Hide API Key
// =====================================================

if (
    showKey &&
    apiKey
) {

    showKey.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (
                apiKey.type ===
                "password"
            ) {

                apiKey.type =
                    "text";


                showKey.innerHTML =
                    "🙈";

            }
            else {

                apiKey.type =
                    "password";


                showKey.innerHTML =
                    "👁";

            }

        };

}


// =====================================================
// Provider Change
// =====================================================

if (provider) {

    provider.onchange =
        function () {

            updateProviderInfo();


            if (modelSelect) {

                modelSelect.innerHTML = `
                    <option value="">
                        أدخل المفتاح ثم حدّث النماذج
                    </option>
                `;

            }


            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "";

            }

        };

}


// =====================================================
// Save Settings Button
// =====================================================

if (saveSettings) {

    saveSettings.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            const settings = {

                provider:
                    provider
                        ? provider.value
                        : "openrouter",

                key:
                    apiKey
                        ? apiKey.value.trim()
                        : "",

                model:
                    modelSelect
                        ? modelSelect.value.trim()
                        : ""

            };


            if (!settings.key) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ يرجى إدخال مفتاح API.";

                }


                return;

            }


            if (!settings.model) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ يرجى تحديد نموذج الذكاء الاصطناعي.";

                }


                return;

            }


            saveAISettings(
                settings
            );


            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "✓ تم حفظ إعدادات الذكاء الاصطناعي";

            }

        };

}


// =====================================================
// Initialize Stored State
// =====================================================

saveProjects();

saveChats();

loadSettings();


// =====================================================
// State Diagnostics
// =====================================================

window.testAppState =
    function () {

        const state = {

            currentProject:
                currentProject,

            currentDocument:
                currentDocument,

            currentChat:
                currentChat,

            chatsCount:
                Array.isArray(
                    chats
                )
                    ? chats.length
                    : 0,

            projectsCount:
                Array.isArray(
                    projects
                )
                    ? projects.length
                    : 0,

            citationSources:
                Array.isArray(
                    currentCitationSources
                )
                    ? currentCitationSources.length
                    : 0,

            aiSettings:
                getActiveAISettings()

        };


        console.log(
            "Research Tools AppState:",
            state
        );


        return state;

    };
    // =====================================================
// Research Tools
// PART 6
// AI Providers + Models + Connection + Ask AI
// =====================================================


// =====================================================
// Gemini Model Normalization
// =====================================================

function normalizeGeminiModel(
    model
) {

    return String(
        model || ""
    ).replace(
        /^models\//,
        ""
    );

}


// =====================================================
// JSON
// =====================================================

async function readJSON(
    response
) {

    try {

        return await response.json();

    }
    catch (error) {

        return {};

    }

}


// =====================================================
// API Error
// =====================================================

function getAPIError(
    result,
    fallback
) {

    if (!result) {

        return fallback;

    }


    if (result.error) {

        if (
            typeof result.error ===
            "string"
        ) {

            return result.error;

        }


        if (
            result.error.message
        ) {

            let message =
                result.error.message;


            if (
                result.error.code
            ) {

                message +=
                    " | Code: " +
                    result.error.code;

            }


            return message;

        }

    }


    if (
        result.message
    ) {

        return String(
            result.message
        );

    }


    return fallback;

}


// =====================================================
// OpenAI / OpenRouter Answer
// =====================================================

function extractOpenAIStyleAnswer(
    result,
    providerName
) {

    if (
        result &&
        result.choices &&
        result.choices.length >
            0
    ) {

        const choice =
            result.choices[0];


        if (
            choice.message &&
            typeof choice.message.content ===
                "string"
        ) {

            return choice.message.content;

        }

    }


    throw new Error(
        "لم يصل رد صالح من " +
        providerName +
        "."
    );

}


// =====================================================
// Gemini Answer
// =====================================================

function extractGeminiAnswer(
    result
) {

    if (
        result &&
        result.candidates &&
        result.candidates.length >
            0
    ) {

        const candidate =
            result.candidates[0];


        if (
            candidate.content &&
            Array.isArray(
                candidate.content.parts
            )
        ) {

            const answerParts =
                candidate.content.parts
                    .filter(
                        function (
                            part
                        ) {

                            if (
                                !part ||
                                typeof part.text !==
                                    "string"
                            ) {

                                return false;

                            }


                            // ==================================
                            // عدم عرض أجزاء التفكير
                            // ==================================

                            if (
                                part.thought ===
                                true
                            ) {

                                return false;

                            }


                            return true;

                        }
                    )
                    .map(
                        function (
                            part
                        ) {

                            return part.text;

                        }
                    );


            if (
                answerParts.length >
                    0
            ) {

                return answerParts
                    .join(
                        "\n"
                    )
                    .trim();

            }

        }

    }


    throw new Error(
        "لم يصل رد صالح من Gemini."
    );

}


// =====================================================
// Populate Models
// =====================================================

function populateModels(
    models
) {

    if (!modelSelect) {

        return;

    }


    const saved =
        getSavedSettings();


    const savedModel =
        saved.model ||
        "";


    modelSelect.innerHTML =
        "";


    if (
        !Array.isArray(
            models
        ) ||
        models.length ===
            0
    ) {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            "";


        option.textContent =
            "لا توجد نماذج متاحة";


        modelSelect.appendChild(
            option
        );


        return;

    }


    models.forEach(
        function (
            item
        ) {

            if (
                !item ||
                !item.id
            ) {

                return;

            }


            const option =
                document.createElement(
                    "option"
                );


            option.value =
                item.id;


            option.textContent =
                item.name ||
                item.id;


            modelSelect.appendChild(
                option
            );

        }
    );


    if (savedModel) {

        const exists =
            Array.from(
                modelSelect.options
            ).some(
                function (
                    option
                ) {

                    return (
                        option.value ===
                        savedModel
                    );

                }
            );


        if (exists) {

            modelSelect.value =
                savedModel;

        }

    }

}


// =====================================================
// Load Groq Models
// =====================================================

async function loadGroqModels() {

    const key =
        apiKey
            ? apiKey.value.trim()
            : "";


    if (!key) {

        throw new Error(
            "يرجى إدخال مفتاح Groq أولاً."
        );

    }


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "⏳ جاري تحميل نماذج Groq...";

    }


    const response =
        await fetch(
            "https://api.groq.com/openai/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Authorization":
                        "Bearer " +
                        key,

                    "Content-Type":
                        "application/json"

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Groq."
            )
        );

    }


    if (
        !result.data ||
        !Array.isArray(
            result.data
        )
    ) {

        throw new Error(
            "لم تصل قائمة نماذج Groq."
        );

    }


    const models =
        result.data
            .filter(
                function (
                    item
                ) {

                    return (
                        item &&
                        item.id &&
                        item.active !==
                            false
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return String(
                        a.id
                    ).localeCompare(
                        String(
                            b.id
                        )
                    );

                }
            )
            .map(
                function (
                    item
                ) {

                    return {

                        id:
                            item.id,

                        name:
                            item.id

                    };

                }
            );


    populateModels(
        models
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج Groq: " +
            models.length;

    }

}


// =====================================================
// Load OpenRouter Models
// النماذج المجانية
// =====================================================

async function loadOpenRouterModels() {

    const key =
        apiKey
            ? apiKey.value.trim()
            : "";


    if (!key) {

        throw new Error(
            "يرجى إدخال مفتاح OpenRouter أولاً."
        );

    }


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "⏳ جاري تحميل نماذج OpenRouter المجانية...";

    }


    const response =
        await fetch(
            "https://openrouter.ai/api/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Authorization":
                        "Bearer " +
                        key,

                    "Content-Type":
                        "application/json",

                    "HTTP-Referer":
                        window.location.href,

                    "X-Title":
                        "Research Tools"

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenRouter."
            )
        );

    }


    if (
        !result.data ||
        !Array.isArray(
            result.data
        )
    ) {

        throw new Error(
            "لم تصل قائمة النماذج من OpenRouter."
        );

    }


    const freeModels =
        result.data.filter(
            function (
                item
            ) {

                return (
                    item &&
                    item.id &&
                    String(
                        item.id
                    ).endsWith(
                        ":free"
                    )
                );

            }
        );


    const models =
        freeModels
            .map(
                function (
                    item
                ) {

                    return {

                        id:
                            item.id,

                        name:
                            (
                                item.name ||
                                item.id
                            ) +
                            " (مجاني)"

                    };

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return String(
                        a.name
                    ).localeCompare(
                        String(
                            b.name
                        ),
                        "ar"
                    );

                }
            );


    populateModels(
        models
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث النماذج المجانية: " +
            models.length;

    }

}


// =====================================================
// Load OpenAI Models
// =====================================================

async function loadOpenAIModels() {

    const key =
        apiKey
            ? apiKey.value.trim()
            : "";


    if (!key) {

        throw new Error(
            "يرجى إدخال مفتاح OpenAI أولاً."
        );

    }


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "⏳ جاري تحميل نماذج OpenAI...";

    }


    const response =
        await fetch(
            "https://api.openai.com/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Authorization":
                        "Bearer " +
                        key,

                    "Content-Type":
                        "application/json"

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenAI."
            )
        );

    }


    if (
        !result.data ||
        !Array.isArray(
            result.data
        )
    ) {

        throw new Error(
            "لم تصل قائمة نماذج OpenAI."
        );

    }


    const models =
        result.data
            .filter(
                function (
                    item
                ) {

                    if (
                        !item ||
                        !item.id
                    ) {

                        return false;

                    }


                    const id =
                        String(
                            item.id
                        ).toLowerCase();


                    return (
                        id.startsWith(
                            "gpt-"
                        ) ||
                        id.startsWith(
                            "o1"
                        ) ||
                        id.startsWith(
                            "o3"
                        ) ||
                        id.startsWith(
                            "o4"
                        )
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return String(
                        a.id
                    ).localeCompare(
                        String(
                            b.id
                        )
                    );

                }
            )
            .map(
                function (
                    item
                ) {

                    return {

                        id:
                            item.id,

                        name:
                            item.id

                    };

                }
            );


    populateModels(
        models
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج OpenAI: " +
            models.length;

    }

}


// =====================================================
// Load Gemini Models
// =====================================================

async function loadGeminiModels() {

    const key =
        apiKey
            ? apiKey.value.trim()
            : "";


    if (!key) {

        throw new Error(
            "يرجى إدخال مفتاح Gemini أولاً."
        );

    }


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "⏳ جاري تحميل نماذج Gemini...";

    }


    const url =
        "https://generativelanguage.googleapis.com/v1beta/models?key=" +
        encodeURIComponent(
            key
        );


    const response =
        await fetch(
            url,
            {

                method:
                    "GET"

            }
        );


    const result =
        await readJSON(
            response
        );


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Gemini."
            )
        );

    }


    if (
        !result.models ||
        !Array.isArray(
            result.models
        )
    ) {

        throw new Error(
            "لم تصل قائمة نماذج Gemini."
        );

    }


    const models =
        result.models
            .filter(
                function (
                    item
                ) {

                    return (
                        item &&
                        item.name &&
                        Array.isArray(
                            item.supportedGenerationMethods
                        ) &&
                        item.supportedGenerationMethods.includes(
                            "generateContent"
                        )
                    );

                }
            )
            .map(
                function (
                    item
                ) {

                    const cleanId =
                        normalizeGeminiModel(
                            item.name
                        );


                    return {

                        id:
                            cleanId,

                        name:
                            item.displayName
                                ? item.displayName +
                                  " — " +
                                  cleanId
                                : cleanId

                    };

                }
            );


    populateModels(
        models
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج Gemini: " +
            models.length;

    }

}


// =====================================================
// Load Models
// =====================================================

async function loadModels() {

    const selectedProvider =
        provider
            ? provider.value
            : "openrouter";


    if (
        selectedProvider ===
        "openrouter"
    ) {

        await loadOpenRouterModels();

        return;

    }


    if (
        selectedProvider ===
        "gemini"
    ) {

        await loadGeminiModels();

        return;

    }


    if (
        selectedProvider ===
        "groq"
    ) {

        await loadGroqModels();

        return;

    }


    if (
        selectedProvider ===
        "openai"
    ) {

        await loadOpenAIModels();

        return;

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف."
    );

}


// =====================================================
// Refresh Models Button
// =====================================================

if (refreshModels) {

    refreshModels.onclick =
        async function (e) {

            e.preventDefault();
            e.stopPropagation();


            refreshModels.disabled =
                true;


            try {

                await loadModels();

            }
            catch (error) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ " +
                        (
                            error.message ||
                            "تعذر تحديث النماذج"
                        );

                }

            }
            finally {

                refreshModels.disabled =
                    false;

            }

        };

}


// =====================================================
// Test AI Connection
// =====================================================

async function testAIConnection() {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "يرجى إدخال مفتاح API أولاً."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "يرجى تحديد نموذج الذكاء الاصطناعي أولاً."
        );

    }


    // =================================================
    // OpenRouter
    // =================================================

    if (
        data.provider ===
        "openrouter"
    ) {

        const response =
            await fetch(
                "https://openrouter.ai/api/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key,

                        "Content-Type":
                            "application/json",

                        "HTTP-Referer":
                            window.location.href,

                        "X-Title":
                            "Research Tools"

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages: [

                                {

                                    role:
                                        "user",

                                    content:
                                        "أجب بكلمة واحدة فقط: متصل"

                                }

                            ],

                            max_tokens:
                                10

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenRouter."
                )
            );

        }


        return "✓ تم الاتصال بـ OpenRouter بنجاح";

    }


    // =================================================
    // OpenAI
    // =================================================

    if (
        data.provider ===
        "openai"
    ) {

        const response =
            await fetch(
                "https://api.openai.com/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key,

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages: [

                                {

                                    role:
                                        "user",

                                    content:
                                        "أجب بكلمة واحدة فقط: متصل"

                                }

                            ],

                            max_tokens:
                                10

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenAI."
                )
            );

        }


        return "✓ تم الاتصال بـ OpenAI بنجاح";

    }


    // =================================================
    // Gemini
    // =================================================

    if (
        data.provider ===
        "gemini"
    ) {

        const cleanModel =
            normalizeGeminiModel(
                data.model
            );


        const url =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(
                cleanModel
            ) +
            ":generateContent?key=" +
            encodeURIComponent(
                data.key
            );


        const response =
            await fetch(
                url,
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            contents: [

                                {

                                    parts: [

                                        {

                                            text:
                                                "أجب بكلمة واحدة فقط: متصل"

                                        }

                                    ]

                                }

                            ]

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Gemini."
                )
            );

        }


        return "✓ تم الاتصال بـ Gemini بنجاح";

    }


    // =================================================
    // Groq
    // =================================================

    if (
        data.provider ===
        "groq"
    ) {

        const response =
            await fetch(
                "https://api.groq.com/openai/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key,

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages: [

                                {

                                    role:
                                        "user",

                                    content:
                                        "أجب بكلمة واحدة فقط: متصل"

                                }

                            ],

                            max_tokens:
                                10

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Groq."
                )
            );

        }


        return "✓ تم الاتصال بـ Groq بنجاح";

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف."
    );

}


// =====================================================
// Test Connection Button
// =====================================================

if (testConnection) {

    testConnection.onclick =
        async function (e) {

            e.preventDefault();
            e.stopPropagation();


            testConnection.disabled =
                true;


            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⏳ جاري اختبار الاتصال...";

            }


            try {

                const message =
                    await testAIConnection();


                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        message;

                }

            }
            catch (error) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ " +
                        (
                            error.message ||
                            "تعذر الاتصال"
                        );

                }

            }
            finally {

                testConnection.disabled =
                    false;

            }

        };

}


// =====================================================
// Build AI Conversation Messages
// تجهيز السؤال والسياق قبل إرساله للمزود
// =====================================================

async function buildAIConversationMessages(
    text
) {

    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const conversationMessages =
        [];


    // ==================================
    // مهم:
    // تعريف documentContext تم قبل استخدامه
    // لإصلاح المشكلة الموجودة في النسخة القديمة
    // ==================================

    const historyLimit =
        documentContext &&
        documentContext.found
            ? 2
            : 4;


    if (
        currentChat &&
        Array.isArray(
            currentChat.messages
        )
    ) {

        const previousMessages =
            currentChat.messages
                .slice(
                    -historyLimit
                );


        previousMessages.forEach(
            function (
                msg
            ) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                let messageText =
                    String(
                        msg.text
                    ).trim();


                const maxHistoryChars =
                    documentContext &&
                    documentContext.found
                        ? 1000
                        : 1500;


                if (
                    messageText.length >
                    maxHistoryChars
                ) {

                    messageText =
                        messageText.substring(
                            0,
                            maxHistoryChars
                        ) +
                        "…";

                }


                conversationMessages.push({

                    role:
                        msg.role ===
                        "ai"
                            ? "assistant"
                            : "user",

                    content:
                        messageText

                });

            }
        );

    }


    // ==================================
    // السؤال الحالي
    // ==================================

    let userContent =
        text;


    // ==================================
    // وجود مستند
    // ==================================

    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",
                "",

                "=== سؤال المستخدم ===",
                text,

                "",

                "=== بيانات المستند ===",
                "اسم المستند: " +
                    (
                        currentDocument
                            ? currentDocument.name
                            : ""
                    ),

                "نوع الاسترجاع: " +
                    documentContext.profile,

                "العائلات المطابقة: " +
                    (
                        documentContext.matchedFamilies &&
                        documentContext.matchedFamilies.length >
                            0
                            ? documentContext.matchedFamilies.join(
                                "، "
                            )
                            : "لا توجد"
                    ),

                "",

                "=== المادة المستخرجة من المستند ===",
                documentContext.text,

                "",

                "=== قواعد الإجابة ===",

                "أجب عن سؤال المستخدم اعتمادًا على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",

                "استخرج الأفكار المرتبطة بالسؤال فقط.",

                "ادمج الأفكار المتشابهة في فكرة واحدة ولا تكررها بصيغ مختلفة.",

                "رتب الإجابة وفق محاور السؤال.",

                "إذا كان السؤال يتضمن أكثر من جانب، أجب عن جميع الجوانب التي تدعمها المقاطع.",

                "لا تضف معلومة أو حكمًا أو نسبة قول إلى المستند غير موجودة في المقاطع المستخرجة.",

                "إذا لم تتوفر في المقاطع إجابة عن جزء من السؤال، صرّح بذلك بوضوح.",

                "لا تستخدم المعرفة العامة لسد نقص المستند إلا إذا طلب المستخدم ذلك صراحة.",

                "حافظ على العربية والأسلوب الأكاديمي الواضح.",

                "لا تبدأ باعتذار أو تمهيد عام غير ضروري.",

                "لا تعيد صياغة سؤال المستخدم.",

                "اجعل الإجابة مترابطة ومعتدلة الطول.",

                "لا تحول كل مقطع إلى فقرة مستقلة إذا كانت المقاطع تخدم الفكرة نفسها.",

                "ضع الإحالة [مقطع رقم] بعد الفكرة التي يدعمها ذلك المقطع.",

                "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع إحالاتها في نهاية الفكرة.",

                "لا تكرر الإحالة نفسها دون فائدة.",

                "لا تخترع أرقام مقاطع.",

                "لا تنسب فكرة إلى مقطع لا يدعمها.",

                "استبعد أي مقطع مستخرج لا يجيب مباشرة عن سؤال المستخدم.",

                "وجود كلمات السؤال داخل المقطع لا يعني أن المقطع صالح للإجابة.",

                "قدّم خلاصة تركيبية للمقاطع بدل تلخيص كل مقطع على حدة."

            ].join(
                "\n"
            );

    }


    conversationMessages.push({

        role:
            "user",

        content:
            userContent

    });


    return {

        documentContext:
            documentContext,

        messages:
            conversationMessages,

        userContent:
            userContent

    };

}


// =====================================================
// Ask AI
// طلب غير متدفق
// =====================================================

async function askAI(
    text
) {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
        );

    }


    const aiContext =
        await buildAIConversationMessages(
            text
        );


    const conversationMessages =
        aiContext.messages;


    // =================================================
    // OpenRouter
    // =================================================

    if (
        data.provider ===
        "openrouter"
    ) {

        const response =
            await fetch(
                "https://openrouter.ai/api/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json",

                        "Authorization":
                            "Bearer " +
                            data.key,

                        "HTTP-Referer":
                            window.location.href,

                        "X-Title":
                            "Research Tools"

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages:
                                conversationMessages,

                            max_tokens:
                                8000,

                            temperature:
                                0.2

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenRouter."
                )
            );

        }


        return extractOpenAIStyleAnswer(
            result,
            "OpenRouter"
        );

    }


    // =================================================
    // OpenAI
    // =================================================

    if (
        data.provider ===
        "openai"
    ) {

        const response =
            await fetch(
                "https://api.openai.com/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json",

                        "Authorization":
                            "Bearer " +
                            data.key

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages:
                                conversationMessages,

                            max_tokens:
                                8000,

                            temperature:
                                0.2

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenAI."
                )
            );

        }


        return extractOpenAIStyleAnswer(
            result,
            "OpenAI"
        );

    }


    // =================================================
    // Gemini
    // =================================================

    if (
        data.provider ===
        "gemini"
    ) {

        const cleanModel =
            normalizeGeminiModel(
                data.model
            );


        const url =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(
                cleanModel
            ) +
            ":generateContent?key=" +
            encodeURIComponent(
                data.key
            );


        const contents =
            conversationMessages.map(
                function (
                    msg
                ) {

                    return {

                        role:
                            msg.role ===
                            "assistant"
                                ? "model"
                                : "user",

                        parts: [

                            {

                                text:
                                    msg.content

                            }

                        ]

                    };

                }
            );


        const response =
            await fetch(
                url,
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            contents:
                                contents

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Gemini."
                )
            );

        }


        return extractGeminiAnswer(
            result
        );

    }


    // =================================================
    // Groq
    // =================================================

    if (
        data.provider ===
        "groq"
    ) {

        const response =
            await fetch(
                "https://api.groq.com/openai/v1/chat/completions",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json",

                        "Authorization":
                            "Bearer " +
                            data.key

                    },

                    body:
                        JSON.stringify({

                            model:
                                data.model,

                            messages:
                                conversationMessages,

                            max_tokens:
                                3000,

                            temperature:
                                0.2

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Groq."
                )
            );

        }


        return extractOpenAIStyleAnswer(
            result,
            "Groq"
        );

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف: " +
        data.provider
    );

}

function escapeHTML(value) {

    return String(
        value ?? ""
    )
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
// =====================================================
// AI REFERENCE ANALYZER
// محلل المراجع بالذكاء الاصطناعي
// =====================================================

async function analyzeReferencesWithAI(materials) {

    if (
        !Array.isArray(materials) ||
        materials.length === 0
    ) {

        return {
            references: [],
            stats: {
                totalMaterials: 0,
                referenceCount: 0,
                multipleReferenceCount: 0,
                ibidCount: 0,
                internalReferenceCount: 0,
                hadithCount: 0,
                explanatoryCount: 0,
                mixedCount: 0,
                reviewCount: 0
            }
        };

    }


    const settings =
        getSavedSettings();


    const provider =
        String(
            settings?.provider ||
            "openrouter"
        )
            .trim()
            .toLowerCase();


    const key =
        String(
            settings?.key ||
            ""
        )
            .trim();


    const model =
        String(
            settings?.model ||
            ""
        )
            .trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
        );

    }


    // =====================================================
    // تجهيز المواد
    // =====================================================

    const payload =
        materials.map(
            function (
                material,
                index
            ) {

                return {

                    materialId:
                        String(
                            material?.id ||
                            `material-${index + 1}`
                        ),

                    source:
                        String(
                            material?.source ||
                            ""
                        ),

                    kind:
                        String(
                            material?.kind ||
                            ""
                        ),

                    noteNumber:
                        material?.noteNumber ??
                        null,

                    position:
                        material?.position ??
                        null,

                    text:
                        String(
                            material?.originalText ||
                            material?.cleanedText ||
                            material?.text ||
                            ""
                        )
                            .trim(),

                    context:
                        String(
                            material?.context ||
                            ""
                        )
                            .trim()

                };

            }
        );


    // =====================================================
    // خريطة المواد الأصلية
    //
    // نستخدمها لاستعادة source / noteNumber / position
    // إذا أغفلها الذكاء الاصطناعي في occurrence.
    // =====================================================

    const materialIndex =
        new Map();

    payload.forEach(
        function (
            material
        ) {

            materialIndex.set(
                String(
                    material.materialId
                ),
                material
            );

        }
    );


    // =====================================================
    // استعادة بيانات occurrence من المادة الأصلية
    //
    // لا نغير المرجع ولا locations ولا variants.
    // =====================================================

    function hydrateAIResult(
        aiResult
    ) {

        if (
            !aiResult ||
            !Array.isArray(
                aiResult.references
            )
        ) {

            return aiResult;

        }


        aiResult.references.forEach(
            function (
                reference
            ) {

                if (
                    !Array.isArray(
                        reference.occurrences
                    )
                ) {

                    reference.occurrences =
                        [];

                    return;

                }


                reference.occurrences.forEach(
                    function (
                        occurrence
                    ) {

                        const material =
                            materialIndex.get(
                                String(
                                    occurrence?.materialId ||
                                    ""
                                )
                            );


                        if (!material) {
                            return;
                        }


                        // source
                        if (
                            !occurrence.source &&
                            material.source
                        ) {

                            occurrence.source =
                                material.source;

                        }


                        // noteNumber
                        if (
                            occurrence.noteNumber ===
                                null ||
                            occurrence.noteNumber ===
                                undefined
                        ) {

                            occurrence.noteNumber =
                                material.noteNumber;

                        }


                        // position
                        if (
                            occurrence.position ===
                                null ||
                            occurrence.position ===
                                undefined
                        ) {

                            occurrence.position =
                                material.position;

                        }

                    }
                );

            }
        );


        return aiResult;

    }


    // =====================================================
    // تعليمات الذكاء الاصطناعي
    // =====================================================

    const systemPrompt = `

أنت محلل مراجع أكاديمي متخصص في البحوث العربية
والدراسات الإسلامية والفقه وأصول الفقه.

ستستلم مواد كاملة مرتبة بحسب ظهورها في مستند Word.

مهمتك تحديد هوية المراجع الواردة في هذه المواد وبناء قائمة
نهائية للمراجع.

المطلوب منك أساسًا:

- تصنيف المواد.
- فصل المراجع المتعددة داخل المادة الواحدة.
- حل "المصدر نفسه" اعتمادًا على السياق السابق.
- تحديد المؤلف والعنوان عندما تسمح المادة والسياق بذلك.
- تحديد المراجع التي تعود إلى الكتاب نفسه.
- عدم دمج الكتب المختلفة للمؤلف نفسه.

=====================================================
قواعد الهوية
=====================================================

المؤلف وحده لا يحدد هوية المرجع.

إذا كان المؤلف نفسه وله كتب مختلفة،
فكل كتاب مرجع مستقل.

مثال:

الغزالي - المنخول
الغزالي - أساس القياس
الغزالي - الوسيط في المذهب

ثلاثة مراجع مستقلة.

وكذلك:

أبو مؤنس - الثوابت والمتغيرات في التشريع الإسلامي
أبو مؤنس - التعليل بالحكمة وأثره في قواعد الفقه وأصوله

مرجعان مستقلان.

يمكن دمج العنوان المختصر مع العنوان الكامل
إذا كان واضحًا أنهما الكتاب نفسه.

مثال:

البحر المحيط
البحر المحيط في أصول الفقه

يجوز دمجهما إذا كان السياق واضحًا.

أما اختلاف العنوان اختلافًا جوهريًا فيعني كتابًا مستقلًا.

=====================================================
المراجع المتعددة — قاعدة الفصل الأساسية
=====================================================

المادة الواحدة قد تحتوي على مرجعين أو ثلاثة أو عشرة مراجع
مستقلة، خصوصًا الحواشي الطويلة.

يجب تحليل المادة كاملة واكتشاف كل مرجع مستقل فيها.

لا تجعل الحاشية الواحدة مرجعًا واحدًا لمجرد أنها مادة واحدة.

مثال:

"مجلة المجمع الفقهي عدد (4) ج2 ص2333-
والعدد (5) ج3 ص2265، دراسة شرعية لأهم العقود المالية المستحدثة
د/ محمد الأمين مصطفى الشنقيطي. 3- زكاة الحقوق المعنوية،
أ. د. عبد الحميد محمود البعلي، المدخل الفقهي العام،
للشيخ/ مصطفى الزرقا ص267-270، المعاملات المالية المعاصرة
في الفقه الإسلامي، د/محمد عثمان شبير، البيوع الشائعة
وأثر ضبط البيوع على شرعيتها د محمد سعيد رمضان البوطي،
مجلة الشريعة والقانون – الإمارات، 2/15."

هذه ليست مرجعًا واحدًا.

بل تحتوي على المراجع التالية:

1. مجلة المجمع الفقهي — عدد (4) — ج2 — ص2333
2. مجلة المجمع الفقهي — عدد (5) — ج3 — ص2265
3. محمد الأمين مصطفى الشنقيطي — دراسة شرعية لأهم العقود المالية المستحدثة
4. عبد الحميد محمود البعلي — زكاة الحقوق المعنوية
5. مصطفى الزرقا — المدخل الفقهي العام — ص267-270
6. محمد عثمان شبير — المعاملات المالية المعاصرة في الفقه الإسلامي
7. محمد سعيد رمضان البوطي — البيوع الشائعة وأثر ضبط البيوع على شرعيتها
   — مجلة الشريعة والقانون – الإمارات — 2/15

يجب إنشاء سجل مستقل لكل مرجع مستقل.

لا تدمج مرجعين مختلفين إذا اختلف المؤلف أو عنوان الكتاب.
كل مؤلف + عنوان كتاب هو هوية مستقلة.
ولا يجوز نقل location أو variant أو occurrence من مرجع إلى مرجع آخر.

=====================================================
كيف تكتشف بداية مرجع جديد؟
=====================================================

ابحث داخل المادة عن انتقال واضح من مرجع إلى آخر، مثل:

- ظهور مؤلف جديد.
- ظهور عنوان كتاب أو بحث جديد.
- ظهور عنوان مجلة أو دورية جديدة.
- ظهور بيانات مجلد أو صفحة مرتبطة بعنوان جديد.
- ظهور رقم مثل "3-" أو "4-" يدل على بداية مرجع جديد.
- ظهور صيغة مؤلف جديدة مثل:
  د/
  أ. د.
  للشيخ/
  ابن ...
  أبو ...
- انتهاء عنوان مرجع وبدء عنوان آخر.

لا تعتمد على الفاصلة وحدها.

افهم تركيب المادة وسياقها.

=====================================================
مثال مهم جدًا
=====================================================

إذا كانت المادة:

"الشيرازي، اللمع 37، السمعاني، قواطع الأدلة، 1/193"

فالنتيجة يجب أن تكون سجلين مستقلين:

السجل الأول:
author = الشيرازي
title = اللمع
location = ص 37

السجل الثاني:
author = السمعاني
title = قواطع الأدلة
volume = 1
page = 193

ممنوع وضعهما في سجل واحد.

=====================================================
قاعدة الحاشية الطويلة
=====================================================

إذا كانت الحاشية تحتوي على عدة مراجع متتابعة،
فلا تعتبرها مرجعًا واحدًا.

كل مرجع مستقل يحصل على:

- author
- title
- locations
- variants
- occurrences

ويمكن أن تشترك المراجع المختلفة في نفس materialId،
لأنها وردت في الحاشية الأصلية نفسها.

=====================================================
قاعدة variants
=====================================================

variants تستخدم فقط للصيغ المختلفة للمرجع نفسه.

لا تستخدم variants لجمع كتب مختلفة.

مثال صحيح:

الزركشي، المنثور 2/362
المنثور، 2/362
انظر: الزركشي، المنثور، 2/362

هذه variants لمرجع واحد.

أما:

الزركشي، المنثور 2/362
الزركشي، البحر المحيط 8/85

فهذان مرجعان مستقلان.

=====================================================
قاعدة حاسمة
=====================================================

عدد المراجع داخل المادة لا يهم.

قد تكون المادة:

مرجعًا واحدًا،
أو مرجعين،
أو خمسة،
أو عشرة.

المطلوب دائمًا هو فصل كل مرجع مستقل إلى سجل مستقل.

لا تسمح لبنية الحاشية الواحدة بأن تجعل عدة مراجع
في سجل واحد.

لا تكرر نفس صيغة الظهور أكثر من مرة.
اعتبر اختلاف المسافات وعلامات الترقيم البسيطة تكرارًا لنفس الصيغة.

=====================================================
قاعدة occurrences
=====================================================

إذا احتوت مادة واحدة على سبعة مراجع:

كل مرجع يحصل على occurrence واحد.

لا يصبح occurrence = 7 لأي مرجع.

نفس materialId يمكن أن يظهر في occurrence لعدة مراجع
لأنها وردت داخل المادة الأصلية نفسها.

=====================================================
المصدر نفسه
=====================================================

المصدر نفسه والمرجع نفسه ليسا كتابًا جديدًا.

اربطهما بالمرجع السابق المناسب.

إذا تعذر الربط بدرجة كافية من الثقة،
لا تخترع المرجع.

=====================================================
الإحالات الداخلية
=====================================================

مثل:

ص 65
ص 65-85
ص 25-26، ص 45-46
في هذه الدراسة ص 67

إذا كانت إحالة إلى صفحات الدراسة الحالية،
فهي ليست مرجعًا خارجيًا.

=====================================================
الأحاديث
=====================================================

أخرجه البخاري
رواه الترمذي
رواه البيهقي

تصنف تخريج حديث.

=====================================================
الشرح
=====================================================

الشرح وحده لا يدخل قائمة المراجع.

أما إذا احتوى الشرح على مرجع،
فاستخرج المرجع الموجود داخله.

=====================================================
المعلومات
=====================================================

لا تخترع مؤلفًا أو عنوانًا أو ناشرًا أو سنة.

إذا كانت المعلومة غير موجودة فاتركها فارغة.

احتفظ بالنص الأصلي داخل variants.

=====================================================
locations
=====================================================

اجمع جميع مواضع المرجع نفسه.

احتفظ بنطاق الصفحات كما ورد،
ولا تحوّل:

460-464

إلى:

464

ولا تحوّل:

96-101

إلى:

101.

قاعدة locations وvariants:

كل location يجب أن ينتمي إلى المرجع نفسه فقط، وينقل الجزء والصفحة أو نطاق الصفحة كما وردا في المادة الأصلية دون تغيير.

إذا ورد:
11/388-389
فيسجل:
volume = 11
pageRange = "388-389"

ولا يجوز تحويله إلى:
page = "388-389"
ولا إلى:
page = "389"

وإذا ورد:
15/927
فهو للمرجع الذي ورد معه فقط.

لا تنقل location أو variant من مادة إلى مرجع آخر.

variants يجب أن تكون النصوص الأصلية الفعلية التي تشير إلى المرجع نفسه، ولا تُنشئ صيغة جديدة من عندك.

=====================================================
occurrences
=====================================================

كل ظهور للمراجع داخل مادة أصلية يمثل occurrence واحدًا.

وجود أكثر من صفحة داخل المادة نفسها لا يعني أكثر من occurrence.

إذا احتوت مادة واحدة على مرجعين مستقلين،
يحصل كل منهما على occurrence واحد.

confidence رقم بين 0 و1، ويجب أن يعكس درجة الثقة الحقيقية في هوية المرجع.
للمراجع الواضحة جدًا استخدم 0.95 إلى 1.00، ولا تستخدم 0 إذا كانت هوية المرجع واضحة.

إذا أخطأت المادة في اسم المؤلف، مثل:
"النووي، العرف حجيته وأثره..."
بينما المرجع المعروف في نفس المادة هو الزبيري،
فلا تنقل الخطأ إلى هوية المرجع؛ احتفظ بالنص في variants،
واستخدم هوية المرجع التي تدل عليها بقية المادة، مع needsReview عند الحاجة.

كل occurrence يجب أن يحمل موضع الاستشهاد الخاص بهذا الظهور فقط.

إذا كان المرجع في هذه الحاشية:
النووي، المجموع 11/417

فيكون occurrence:
volume = "11"
page = "417"
pageRange = ""

ولا تستخدم أول location للمرجع بدل موضع هذا occurrence.

إذا ورد:
460-464
فيكون:
pageRange = "460-464"

إذا ورد:
11/388-389
فيكون:
volume = "11"
pageRange = "388-389"

الموضع في occurrence يجب أن يخص هذا الظهور وحده، ولا يجوز نقله من ظهور آخر.

=====================================================
النتيجة
=====================================================

أعد JSON فقط:

{
  "stats": {
    "totalMaterials": 0,
    "referenceCount": 0,
    "multipleReferenceCount": 0,
    "ibidCount": 0,
    "internalReferenceCount": 0,
    "hadithCount": 0,
    "explanatoryCount": 0,
    "mixedCount": 0,
    "reviewCount": 0
  },
  "references": [
    {
      "id": "reference-1",
      "type": "book",
      "author": "",
      "title": "",
      "edition": "",
      "editor": "",
      "publisher": "",
      "city": "",
      "year": "",
      "locations": [
        {
          "volume": "",
          "page": "",
          "pageRange": ""
        }
      ],
      "variants": [],
      "occurrences": [
        {
          "materialId": "",
          "source": "",
          "noteNumber": null,
          "volume": "",
          "page": "",
          "pageRange": ""
        }
      ],
      "notes": "",
      "confidence": 0.0,
      "needsReview": false
    }
  ]
}

قيم type:

book
article
journal
thesis
website
hadith
other

حافظ على ترتيب أول ظهور للمراجع.

لا تضع مرجعين مختلفين في سجل واحد.
لا تدمج كتابين مختلفين للمؤلف نفسه.
`;


    const userPrompt =
        [
            "حلل جميع المواد التالية.",
            "",
            "ابنِ قائمة المراجع النهائية.",
            "",
            JSON.stringify(
                payload,
                null,
                2
            )
        ].join("\n");


    let result;


    // =====================================================
    // Gemini
    // =====================================================

    if (
        provider === "gemini"
    ) {

        const response =
            await fetch(
                `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
                    normalizeGeminiModel(model)
                )}:generateContent?key=${encodeURIComponent(
                    key
                )}`,
                {
                    method:
                        "POST",

                    headers:
                        {
                            "Content-Type":
                                "application/json"
                        },

                    body:
                        JSON.stringify(
                            {
                                systemInstruction:
                                    {
                                        parts:
                                            [
                                                {
                                                    text:
                                                        systemPrompt
                                                }
                                            ]
                                    },

                                contents:
                                    [
                                        {
                                            role:
                                                "user",

                                            parts:
                                                [
                                                    {
                                                        text:
                                                            userPrompt
                                                    }
                                                ]
                                        }
                                    ],

                                generationConfig:
                                    {
                                        temperature:
                                            0.1,

                                        responseMimeType:
                                            "application/json"
                                    }
                            }
                        )
                }
            );


        result =
            await readJSON(
                response
            );


        if (
            !response.ok
        ) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Gemini أثناء تحليل المراجع."
                )
            );

        }


        const answer =
            extractGeminiAnswer(
                result
            );


        if (
            !answer ||
            !String(answer).trim()
        ) {

            throw new Error(
                "أعاد Gemini استجابة فارغة."
            );

        }


        return hydrateAIResult(
            parseUnifiedReferenceAIResult(
                answer
            )
        );

    }


    // =====================================================
    // OpenAI / OpenRouter / Groq
    // =====================================================

    const endpoint =
        provider === "openai"

            ? "https://api.openai.com/v1/chat/completions"

            : provider === "groq"

                ? "https://api.groq.com/openai/v1/chat/completions"

                : "https://openrouter.ai/api/v1/chat/completions";


    const headers = {

        "Content-Type":
            "application/json",

        "Authorization":
            "Bearer " + key

    };


    if (
        provider === "openrouter"
    ) {

        headers["HTTP-Referer"] =
            window.location.href;

        headers["X-Title"] =
            "Research Tools";

    }


    const response =
        await fetch(
            endpoint,
            {
                method:
                    "POST",

                headers:
                    headers,

                body:
                    JSON.stringify(
                        {
                            model:
                                model,

                            messages:
                                [
                                    {
                                        role:
                                            "system",

                                        content:
                                            systemPrompt
                                    },

                                    {
                                        role:
                                            "user",

                                        content:
                                            userPrompt
                                    }
                                ],

                            temperature:
                                0.1,

                            response_format:
                                {
                                    type:
                                        "json_object"
                                },

                            max_tokens:
                                provider === "groq"
                                    ? 16000
                                    : 24000
                        }
                    )
            }
        );


    result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                `فشل الاتصال بـ ${provider} أثناء تحليل المراجع.`
            )
        );

    }


    const answer =
        extractOpenAIStyleAnswer(
            result,
            provider
        );


    if (
        !answer ||
        !String(answer).trim()
    ) {

        throw new Error(
            `أعاد ${provider} استجابة فارغة.`
        );

    }


    return hydrateAIResult(
        parseUnifiedReferenceAIResult(
            answer
        )
    );

}

function normalizeUnifiedReferenceResult(
    parsedResult,
    materials
) {

    const result =
        parsedResult &&
        typeof parsedResult === "object" &&
        !Array.isArray(parsedResult)
            ? parsedResult
            : {};

    const references =
        Array.isArray(result.references)
            ? result.references
            : [];

    const statsSource =
        result.stats &&
        typeof result.stats === "object" &&
        !Array.isArray(result.stats)
            ? result.stats
            : {};

    function text(value) {
        return String(value ?? "").trim();
    }

    function normalizeTitle(value) {

        return text(value)
            .replace(/^الشيخ\/\s*/i, "")
            .replace(/^للشيخ\/\s*/i, "")
            .replace(/\s+/g, " ")
            .replace(/[.,،:؛]+$/g, "")
            .trim();

    }

    function normalizeAuthor(value) {

        return text(value)
            .replace(/\s+/g, " ")
            .trim();

    }

    function titleKey(value) {

        let t =
            normalizeTitle(value)
                .toLowerCase();

        /*
         * العناوين المختصرة التي نريد اعتبارها
         * هي والعنوان الكامل كتابًا واحدًا.
         *
         * لا نستخدم المؤلف وحده.
         */

        t = t
            .replace(
                /\s+في\s+شرح\s+/g,
                " "
            )
            .replace(
                /\s+شرح\s+/g,
                " "
            )
            .replace(
                /\s+أصول\s+الفقه$/g,
                ""
            )
            .trim();

        return t;
    }

    function sameAuthor(a, b) {

        const x =
            normalizeAuthor(a);

        const y =
            normalizeAuthor(b);

        if (!x || !y) {
            return false;
        }

        return x === y;
    }

    function titlesMatch(a, b) {

        const x =
            titleKey(a);

        const y =
            titleKey(b);

        if (!x || !y) {
            return false;
        }

        if (x === y) {
            return true;
        }

        /*
         * أحد العنوانين قد يكون نسخة مختصرة من الآخر.
         */
        return (
            x.includes(y) ||
            y.includes(x)
        );

    }

    function normalizeConfidence(value) {

        let n =
            Number(value);

        if (!Number.isFinite(n)) {
            return 0;
        }

        if (n > 1 && n <= 100) {
            n /= 100;
        }

        return Math.max(
            0,
            Math.min(1, n)
        );

    }

    function normalizeLocation(location) {

        if (
            !location ||
            typeof location !== "object"
        ) {
            return null;
        }

        return {
            volume:
                text(location.volume),

            page:
                text(location.page),

            pageRange:
                text(location.pageRange)
        };

    }

    function normalizeOccurrence(occurrence) {

        if (
            !occurrence ||
            typeof occurrence !== "object"
        ) {
            return null;
        }

        return {
            materialId:
                text(occurrence.materialId),

            source:
                text(occurrence.source),

            noteNumber:
                occurrence.noteNumber ?? null
        };

    }

    function addUnique(array, value) {

        const v = text(value);

        if (!v) {
            return;
        }

        if (!array.includes(v)) {
            array.push(v);
        }

    }

    const normalized = [];

    references.forEach(
        function (reference, index) {

            if (
                !reference ||
                typeof reference !== "object"
            ) {
                return;
            }

            const current = {

                id:
                    text(
                        reference.id ||
                        `reference-${index + 1}`
                    ),

                type:
                    text(
                        reference.type ||
                        "book"
                    ).toLowerCase(),

                author:
                    normalizeAuthor(
                        reference.author
                    ),

                title:
                    normalizeTitle(
                        reference.title
                    ),

                edition:
                    text(reference.edition),

                editor:
                    text(reference.editor),

                publisher:
                    text(reference.publisher),

                city:
                    text(reference.city),

                year:
                    text(reference.year),

                locations:
                    Array.isArray(reference.locations)
                        ? reference.locations
                            .map(normalizeLocation)
                            .filter(Boolean)
                        : [],

                variants:
                    Array.isArray(reference.variants)
                        ? reference.variants
                            .map(text)
                            .filter(Boolean)
                        : [],

                occurrences:
                    Array.isArray(reference.occurrences)
                        ? reference.occurrences
                            .map(normalizeOccurrence)
                            .filter(Boolean)
                        : [],

                notes:
                    text(reference.notes),

                confidence:
                    normalizeConfidence(
                        reference.confidence
                    ),

                needsReview:
                    Boolean(
                        reference.needsReview
                    )
            };

            /*
             * البحث عن مرجع موجود لنفس الكتاب.
             *
             * المؤلف + العنوان هما أساس الهوية.
             * المؤلف وحده لا يكفي.
             */
            let existing = null;

            for (
                let i = 0;
                i < normalized.length;
                i++
            ) {

                const candidate =
                    normalized[i];

                /*
                 * إذا كان كلا المرجعين يملكان مؤلفًا:
                 * يجب تطابق المؤلف.
                 *
                 * إذا غاب المؤلف من أحدهما:
                 * يسمح بالدمج فقط إذا كان العنوان واضحًا جدًا.
                 */
                const authorOK =
                    current.author &&
                    candidate.author
                        ? sameAuthor(
                            current.author,
                            candidate.author
                        )
                        : true;

                if (
                    !authorOK
                ) {
                    continue;
                }

                if (
                    titlesMatch(
                        current.title,
                        candidate.title
                    )
                ) {

                    existing =
                        candidate;

                    break;

                }

            }

            /*
             * لا يوجد سجل سابق: أضف المرجع.
             */
            if (!existing) {

                normalized.push(
                    current
                );

                return;

            }

            /*
             * =================================================
             * دمج آمن
             * =================================================
             */

            /*
             * إذا كان العنوان الموجود أكمل،
             * نحافظ عليه.
             */
            if (
                current.title.length >
                existing.title.length
            ) {

                existing.title =
                    current.title;

            }

            if (
                !existing.author &&
                current.author
            ) {

                existing.author =
                    current.author;

            }

            if (
                !existing.edition &&
                current.edition
            ) {

                existing.edition =
                    current.edition;

            }

            if (
                !existing.editor &&
                current.editor
            ) {

                existing.editor =
                    current.editor;

            }

            if (
                !existing.publisher &&
                current.publisher
            ) {

                existing.publisher =
                    current.publisher;

            }

            if (
                !existing.city &&
                current.city
            ) {

                existing.city =
                    current.city;

            }

            if (
                !existing.year &&
                current.year
            ) {

                existing.year =
                    current.year;

            }

            /*
             * جمع المواقع.
             */
            current.locations.forEach(
                function (location) {

                    const exists =
                        existing.locations.some(
                            function (oldLocation) {

                                return (
                                    oldLocation.volume ===
                                        location.volume &&
                                    oldLocation.page ===
                                        location.page &&
                                    oldLocation.pageRange ===
                                        location.pageRange
                                );

                            }
                        );

                    if (!exists) {

                        existing.locations.push(
                            location
                        );

                    }

                }
            );

            /*
             * جمع صيغ الظهور.
             */
            current.variants.forEach(
                function (variant) {

                    addUnique(
                        existing.variants,
                        variant
                    );

                }
            );

            /*
             * جمع مرات الظهور.
             */
            current.occurrences.forEach(
                function (occurrence) {

                    const exists =
                        existing.occurrences.some(
                            function (oldOccurrence) {

                                return (
                                    oldOccurrence.materialId ===
                                        occurrence.materialId &&
                                    oldOccurrence.source ===
                                        occurrence.source &&
                                    String(
                                        oldOccurrence.noteNumber ?? ""
                                    ) ===
                                        String(
                                            occurrence.noteNumber ?? ""
                                        )
                                );

                            }
                        );

                    if (!exists) {

                        existing.occurrences.push(
                            occurrence
                        );

                    }

                }
            );

            /*
             * جمع الملاحظات.
             */
            if (
                current.notes &&
                current.notes !== existing.notes
            ) {

                if (
                    existing.notes
                ) {

                    existing.notes +=
                        " | " +
                        current.notes;

                }
                else {

                    existing.notes =
                        current.notes;

                }

            }

            /*
             * الاحتفاظ بأعلى ثقة.
             */
            existing.confidence =
                Math.max(
                    existing.confidence,
                    current.confidence
                );

            if (
                current.needsReview
            ) {

                existing.needsReview =
                    true;

            }

        }
    );

    /*
     * =====================================================
     * تنظيف نهائي
     * =====================================================
     */

    normalized.forEach(
        function (reference) {

            reference.locations =
                reference.locations.filter(
                    function (location, index, array) {

                        return (
                            array.findIndex(
                                function (item) {

                                    return (
                                        item.volume ===
                                            location.volume &&
                                        item.page ===
                                            location.page &&
                                        item.pageRange ===
                                            location.pageRange
                                    );

                                }
                            ) === index
                        );

                    }
                );

            reference.variants =
                reference.variants.filter(
                    function (variant, index, array) {

                        return (
                            array.indexOf(
                                variant
                            ) === index
                        );

                    }
                );

            reference.occurrences =
                reference.occurrences.filter(
                    function (occurrence, index, array) {

                        return (
                            array.findIndex(
                                function (item) {

                                    return (
                                        item.materialId ===
                                            occurrence.materialId &&
                                        item.source ===
                                            occurrence.source &&
                                        String(
                                            item.noteNumber ?? ""
                                        ) ===
                                            String(
                                                occurrence.noteNumber ?? ""
                                            )
                                    );

                                }
                            ) === index
                        );

                    }
                );

        }
    );

    normalized.stats = {

        totalMaterials:
            Number(
                statsSource.totalMaterials ??
                (
                    Array.isArray(materials)
                        ? materials.length
                        : 0
                )
            ),

        referenceCount:
            Number(
                statsSource.referenceCount ??
                0
            ),

        multipleReferenceCount:
            Number(
                statsSource.multipleReferenceCount ??
                0
            ),

        ibidCount:
            Number(
                statsSource.ibidCount ??
                0
            ),

        internalReferenceCount:
            Number(
                statsSource.internalReferenceCount ??
                0
            ),

        hadithCount:
            Number(
                statsSource.hadithCount ??
                0
            ),

        explanatoryCount:
            Number(
                statsSource.explanatoryCount ??
                0
            ),

        mixedCount:
            Number(
                statsSource.mixedCount ??
                0
            ),

        reviewCount:
            Number(
                statsSource.reviewCount ??
                0
            )
    };

    return normalized;
}

function parseUnifiedReferenceAIResult(answer) {

    let parsed;

    const raw =
        String(
            answer ?? ""
        ).trim();


    if (!raw) {
        throw new Error(
            "أعاد الذكاء الاصطناعي استجابة فارغة."
        );
    }


    try {

        parsed =
            JSON.parse(raw);

    }
    catch (error) {

        const cleaned =
            raw
                .replace(
                    /^\s*```(?:json)?\s*/i,
                    ""
                )
                .replace(
                    /\s*```\s*$/i,
                    ""
                )
                .trim();


        try {

            parsed =
                JSON.parse(cleaned);

        }
        catch (secondError) {

            const start =
                cleaned.indexOf("{");

            const end =
                cleaned.lastIndexOf("}");


            if (
                start === -1 ||
                end === -1 ||
                end <= start
            ) {

                throw new Error(
                    "تعذر قراءة JSON الصادر من الذكاء الاصطناعي."
                );

            }


            try {

                parsed =
                    JSON.parse(
                        cleaned.slice(
                            start,
                            end + 1
                        )
                    );

            }
            catch (thirdError) {

                console.error(
                    "النص الذي تعذر تحليله:",
                    answer
                );

                throw new Error(
                    "تعذر قراءة JSON الصادر من الذكاء الاصطناعي."
                );

            }

        }

    }


    // ---------------------------------------------
    // دعم البنية الصحيحة:
    // { references: [...], stats: {...} }
    // ---------------------------------------------

    if (
        !parsed ||
        typeof parsed !== "object"
    ) {

        throw new Error(
            "بنية نتيجة الذكاء الاصطناعي غير صحيحة."
        );

    }


    const references =
        Array.isArray(
            parsed.references
        )
            ? parsed.references
            : [];


    const stats =
        parsed.stats &&
        typeof parsed.stats === "object" &&
        !Array.isArray(
            parsed.stats
        )

            ? parsed.stats

            : {};


    console.log(
        "parseUnifiedReferenceAIResult:",
        {
            referencesCount:
                references.length,

            stats:
                stats,

            parsed:
                parsed
        }
    );


    return {

        references:
            references,

        stats:
            {

                totalMaterials:
                    Number(
                        stats.totalMaterials ||
                        0
                    ),

                referenceCount:
                    Number(
                        stats.referenceCount ||
                        0
                    ),

                multipleReferenceCount:
                    Number(
                        stats.multipleReferenceCount ||
                        0
                    ),

                ibidCount:
                    Number(
                        stats.ibidCount ||
                        0
                    ),

                internalReferenceCount:
                    Number(
                        stats.internalReferenceCount ||
                        0
                    ),

                hadithCount:
                    Number(
                        stats.hadithCount ||
                        0
                    ),

                explanatoryCount:
                    Number(
                        stats.explanatoryCount ||
                        0
                    ),

                mixedCount:
                    Number(
                        stats.mixedCount ||
                        0
                    ),

                reviewCount:
                    Number(
                        stats.reviewCount ||
                        0
                    )

            }

    };

}
// =====================================================
// PARSE REFERENCE AI JSON
// =====================================================

function parseReferenceAIJSON(
    text
) {

    let raw =
        String(
            text || ""
        ).trim();


    if (!raw) {

        return [];

    }


    // إزالة Markdown إن أعاده النموذج رغم التعليمات
    raw =
        raw.replace(
            /^```(?:json)?\s*/i,
            ""
        );

    raw =
        raw.replace(
            /\s*```$/i,
            ""
        );


    try {

        const parsed =
            JSON.parse(
                raw
            );


        if (
            !parsed ||
            !Array.isArray(
                parsed.items
            )
        ) {

            throw new Error(
                "صيغة JSON الخاصة بتحليل المراجع غير صحيحة."
            );

        }


        return parsed.items;

    }
    catch (
        error
    ) {

        console.error(
            "تعذر تحليل JSON الذي أعاده الذكاء الاصطناعي:",
            text
        );


        throw new Error(
            "أعاد الذكاء الاصطناعي نتيجة غير منظمة أثناء تحليل المراجع."
        );

    }

}

// =====================================================
// AI REFERENCE MERGER
// توحيد وتجميع المراجع باستخدام الذكاء الاصطناعي
// =====================================================

async function mergeReferencesWithAI(
    aiReferenceResults
) {

    if (
        !Array.isArray(aiReferenceResults) ||
        aiReferenceResults.length === 0
    ) {
        return [];
    }


    const settings =
        getSavedSettings();

    const provider =
        String(
            settings.provider || "openrouter"
        ).toLowerCase();

    const key =
        String(
            settings.key || ""
        ).trim();

    const model =
        String(
            settings.model || ""
        ).trim();


    if (!key) {
        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي."
        );
    }


    if (!model) {
        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي."
        );
    }


    const systemPrompt = `

أنت الآن في المرحلة الثانية من تحليل المراجع الأكاديمية.

لديك مجموعة من نتائج تحليل أولي لمواد موجودة في مستند عربي
أكاديمي.

مهمتك الوحيدة هي بناء قائمة موحدة للمراجع.

قواعد صارمة:

1. لا تعتبر كل مادة مرجعًا.
2. لا تكرر المرجع نفسه بسبب اختلاف:
   - الصفحة
   - الجزء
   - طريقة كتابة الاسم
   - وجود "انظر"
   - وجود "المصدر نفسه".
3. اجمع الإحالات المختلفة إلى الكتاب نفسه في سجل واحد.
4. احتفظ بجميع مواضع الاستشهاد بالمرجع.
5. احتفظ بكل الصيغ الأصلية التي استخدمها الباحث.
6. "المصدر نفسه" و"المرجع نفسه" إحالة إلى مرجع سابق، وليسا مرجعًا جديدًا.
7. اربط "المصدر نفسه" بالمرجع السابق المناسب ضمن ترتيب المواد.
8. إذا احتوت مادة واحدة على عدة مراجع، أنشئ كل مرجع كسجل مستقل.
9. لا تدمج كتابين لمجرد تشابه العنوان.
10. لا تخترع اسم مؤلف أو عنوانًا غير موجود في البيانات.
11. إذا كان المؤلف غير مذكور، أبقِه فارغًا.
12. إذا وجدت احتمال خطأ في البيانات الأصلية، لا تصححه من نفسك؛ احتفظ به وأضف ملاحظة.
13. الإحالات الداخلية إلى صفحات الدراسة ليست مراجع.
14. تخريج الحديث ليس كتابًا، ولكنه يحتفظ كسجل حديث مستقل.
15. الشرح الذي لا يحتوي مرجعًا لا يدخل قائمة المراجع.
16. إذا كان الدمج غير مؤكد، لا تدمج السجلين وضع needsReview = true.

أعد JSON فقط بهذه الصيغة:

{
  "references": [
    {
      "id": "",
      "type": "book",
      "author": "",
      "title": "",
      "edition": "",
      "editor": "",
      "publisher": "",
      "city": "",
      "year": "",
      "locations": [
        {
          "volume": "",
          "page": "",
          "pageRange": ""
        }
      ],
      "variants": [],
      "occurrences": [],
      "notes": "",
      "confidence": 0.0,
      "needsReview": false
    }
  ]
}

في occurrences احتفظ بمعرّف المادة الأصلية
materialId، ومصدرها، ورقم الحاشية إن وجد.

في variants احتفظ بالنصوص الأصلية المختلفة التي تشير إلى المرجع نفسه.

لا تحذف المعلومات لمجرد أنها غير مكتملة.
`;


    // =====================================================
    // تقسيم النتائج إلى دفعات
    // =====================================================

    const batchSize = 20;

    const batches = [];

    for (
        let i = 0;
        i < aiReferenceResults.length;
        i += batchSize
    ) {

        batches.push(
            aiReferenceResults.slice(
                i,
                i + batchSize
            )
        );

    }


    let mergedReferences = [];


    // =====================================================
    // إرسال الدفعات
    // =====================================================

    for (
        let batchIndex = 0;
        batchIndex < batches.length;
        batchIndex++
    ) {

        const batch =
            batches[batchIndex];


        const userPrompt =
            [
                `هذه الدفعة رقم ${batchIndex + 1} من ${batches.length}.`,
                "",
                "وحّد المراجع التالية:",
                "",
                JSON.stringify(
                    batch,
                    null,
                    2
                )
            ].join("\n");


        let result;


        // =================================================
        // Gemini
        // =================================================

        if (
            provider === "gemini"
        ) {

            const response =
                await fetch(
                    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(normalizeGeminiModel(model))}:generateContent?key=${encodeURIComponent(key)}`,
                    {

                        method: "POST",

                        headers: {
                            "Content-Type":
                                "application/json"
                        },

                        body: JSON.stringify({

                            systemInstruction: {

                                parts: [
                                    {
                                        text:
                                            systemPrompt
                                    }
                                ]

                            },

                            contents: [

                                {
                                    role:
                                        "user",

                                    parts: [
                                        {
                                            text:
                                                userPrompt
                                        }
                                    ]

                                }

                            ],

                            generationConfig: {

                                temperature:
                                    0.1

                            }

                        })

                    }
                );


            result =
                await readJSON(
                    response
                );


            if (
                !response.ok
            ) {

                throw new Error(
                    getAPIError(
                        result,
                        "فشل تحليل دمج المراجع."
                    )
                );

            }


            const answer =
                extractGeminiAnswer(
                    result
                );


            mergedReferences.push(
                ...parseReferenceMergeJSON(
                    answer
                )
            );

        }

        // =================================================
        // OpenAI / OpenRouter / Groq
        // =================================================

        else {

            const endpoint =
                provider === "openai"

                    ? "https://api.openai.com/v1/chat/completions"

                    : provider === "groq"

                        ? "https://api.groq.com/openai/v1/chat/completions"

                        : "https://openrouter.ai/api/v1/chat/completions";


            const headers = {

                "Content-Type":
                    "application/json",

                "Authorization":
                    "Bearer " + key

            };


            if (
                provider ===
                "openrouter"
            ) {

                headers[
                    "HTTP-Referer"
                ] =
                    window.location.href;

                headers[
                    "X-Title"
                ] =
                    "Research Tools";

            }


            const response =
                await fetch(
                    endpoint,
                    {

                        method: "POST",

                        headers: headers,

                        body: JSON.stringify({

                            model:
                                model,

                            messages: [

                                {
                                    role:
                                        "system",

                                    content:
                                        systemPrompt
                                },

                                {
                                    role:
                                        "user",

                                    content:
                                        userPrompt
                                }

                            ],

                            temperature:
                                0.1,

                            max_tokens:
                                provider ===
                                "groq"
                                    ? 7000
                                    : 10000

                        })

                    }
                );


            result =
                await readJSON(
                    response
                );


            if (
                !response.ok
            ) {

                throw new Error(
                    getAPIError(
                        result,
                        "فشل تحليل دمج المراجع."
                    )
                );

            }


            const answer =
                extractOpenAIStyleAnswer(
                    result,
                    provider
                );


            mergedReferences.push(
                ...parseReferenceMergeJSON(
                    answer
                )
            );

        }

    }


    return mergedReferences;

}

// =====================================================
// PARSE REFERENCE MERGE JSON
// =====================================================

function parseReferenceMergeJSON(
    text
) {

    let raw =
        String(
            text || ""
        ).trim();


    if (!raw) {
        return [];
    }


    raw =
        raw.replace(
            /^```(?:json)?\s*/i,
            ""
        );


    raw =
        raw.replace(
            /\s*```$/i,
            ""
        );


    try {

        const parsed =
            JSON.parse(
                raw
            );


        if (
            !parsed ||
            !Array.isArray(
                parsed.references
            )
        ) {

            throw new Error(
                "صيغة دمج المراجع غير صحيحة."
            );

        }


        return parsed.references;

    }
    catch (
        error
    ) {

        console.error(
            "نتيجة دمج المراجع غير صالحة:",
            text
        );


        throw new Error(
            "أعاد الذكاء الاصطناعي نتيجة غير منظمة أثناء توحيد المراجع."
        );

    }

}
// =====================================================
// Ask AI Diagnostics
// =====================================================

window.testAskAI =
    async function (
        text
    ) {

        try {

            const answer =
                await askAI(
                    text
                );


            console.log(
                "إجابة AI:",
                answer
            );


            return answer;

        }
        catch (error) {

            console.error(
                "فشل askAI:",
                error
            );


            return null;

        }

    };
    
// =====================================================
// Research Tools
// PART 7
// Streaming Providers
// Groq + Gemini + OpenRouter + OpenAI
// =====================================================


// =====================================================
// Build Streaming History
// بناء تاريخ المحادثة بدون تكرار السؤال الحالي
// =====================================================

function buildStreamingHistory(
    documentContext
) {

    const messages =
        [];


    const historyLimit =
        documentContext &&
        documentContext.found
            ? 2
            : 4;


    if (
        !currentChat ||
        !Array.isArray(
            currentChat.messages
        )
    ) {

        return messages;

    }


    // ==================================
    // الرسالة الحالية موجودة بالفعل
    // في currentChat بعد sendMessage()
    // لذلك نستبعد آخر رسالة
    // ==================================

    const messagesWithoutCurrent =
        currentChat.messages.slice(
            0,
            -1
        );


    const previousMessages =
        messagesWithoutCurrent.slice(
            -historyLimit
        );


    previousMessages.forEach(
        function (
            msg
        ) {

            if (
                !msg ||
                !msg.text
            ) {

                return;

            }


            let messageText =
                String(
                    msg.text
                ).trim();


            const maxHistoryChars =
                documentContext &&
                documentContext.found
                    ? 1000
                    : 1500;


            if (
                messageText.length >
                maxHistoryChars
            ) {

                messageText =
                    messageText.substring(
                        0,
                        maxHistoryChars
                    ) +
                    "…";

            }


            messages.push({

                role:
                    msg.role ===
                    "ai"
                        ? "assistant"
                        : "user",

                content:
                    messageText

            });

        }
    );


    return messages;

}



// =====================================================
// Build Streaming Context
//
// المسؤول عن:
// 1) استدعاء محرك الاسترجاع.
// 2) تجهيز تاريخ المحادثة.
// 3) تجهيز السؤال + المادة المستخرجة.
// 4) تعريف الإحالات المتاحة فعليًا للذكاء الاصطناعي.
// 5) منع اختراع إحالات غير موجودة.
//
// جميع مزودي الذكاء الاصطناعي يعتمدون على هذه الدالة.
// =====================================================

// =====================================================
// Build Streaming Context
//
// إذا كان هناك مستند نشط:
//     → يستخرج سياق المستند
//     → يرسل المادة إلى الذكاء الاصطناعي
//
// إذا لم يوجد مستند نشط:
//     → تبقى الدردشة عامة
//     → لا يتم فرض أي سياق مستندي
// =====================================================

async function buildStreamingContext(
    text
) {

    // ==================================
    // استرجاع سياق المستند النشط
    //
    // buildAIDocumentContext() نفسها
    // تتعامل مع حالة عدم وجود currentDocument.
    // ==================================

    const documentContext =
        await buildAIDocumentContext(
            text
        );


    // ==================================
    // تاريخ المحادثة
    // ==================================

    const history =
        buildStreamingHistory(
            documentContext
        );


    // ==================================
    // السؤال الأصلي
    // ==================================

    let userContent =
        String(
            text ||
            ""
        );


    // ==================================
    // إذا وجد مستند نشط وسياق صالح
    // ==================================

    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",

                "",

                "=== سؤال المستخدم ===",

                text,

                "",

                "=== بيانات المستند ===",

                "اسم المستند: " +
                    (
                        currentDocument
                            ? currentDocument.name
                            : ""
                    ),

                "نوع الاسترجاع: " +
                    (
                        documentContext.profile ||
                        "general"
                    ),

                "العائلات المطابقة: " +
                    (
                        Array.isArray(
                            documentContext.matchedFamilies
                        ) &&
                        documentContext.matchedFamilies.length >
                            0

                            ? documentContext
                                .matchedFamilies
                                .join(
                                    "، "
                                )

                            : "لا توجد"
                    ),

                "",

                "=== المادة المستخرجة من المستند ===",

                documentContext.text,

                "",

                "=== قواعد الإجابة ===",

                "أجب عن سؤال المستخدم اعتمادًا على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",

                "استخرج الأفكار المرتبطة بالسؤال فقط.",

                "ادمج الأفكار المتشابهة في فكرة واحدة.",

                "رتب الإجابة وفق محاور السؤال.",

                "إذا كان السؤال يتضمن أكثر من جانب، أجب عن الجوانب التي تدعمها المادة.",

                "لا تضف معلومة أو حكمًا أو نسبة قول إلى المستند غير موجودة في المادة المستخرجة.",

                "إذا لم تكف المادة المستخرجة للإجابة عن جزء من السؤال، صرّح بذلك بوضوح.",

                "لا تستخدم المعرفة العامة لسد نقص المستند إلا إذا طلب المستخدم ذلك صراحة.",

                "حافظ على لغة السؤال ولغة المستند.",

                "لا تبدأ باعتذار أو تمهيد غير ضروري.",

                "لا تعيد صياغة سؤال المستخدم.",

                "ادمج المقاطع التي تحمل الفكرة نفسها.",

                "لا تحول كل مقطع مستخرج إلى فقرة مستقلة.",

                "ضع الإحالة [مقطع رقم] بعد الفكرة التي يدعمها المقطع.",

                "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع الإحالات.",

                "لا تخترع أرقام مقاطع.",

                "لا تنسب فكرة إلى مقطع لا يدعمها.",

                "وجود كلمات السؤال داخل المقطع لا يعني أن المقطع صالح للإجابة.",

                "قدّم خلاصة تركيبية للمادة المستخرجة."

            ].join(
                "\n"
            );

    }


    // ==================================
    // لا يوجد مستند نشط
    //
    // لا نضيف أي تعليمات بحثية.
    // يبقى السؤال كما كتبه المستخدم.
    // ==================================

    else {

        userContent =
            String(
                text ||
                ""
            );

    }


    // ==================================
    // السؤال الحالي
    // ==================================

    history.push({

        role:
            "user",

        content:
            userContent

    });


    return {

        documentContext:
            documentContext,

        messages:
            history,

        userContent:
            userContent

    };

}


// =====================================================
// Process OpenAI-Compatible SSE
// Groq + OpenRouter + OpenAI
// =====================================================

async function processOpenAICompatibleStream(
    response,
    onChunk,
    providerName
) {

    if (!response.body) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من " +
            providerName +
            "."
        );

    }


    const reader =
        response.body.getReader();


    const decoder =
        new TextDecoder(
            "utf-8"
        );


    let buffer =
        "";


    let fullAnswer =
        "";


    function processSSELine(
        line
    ) {

        const cleanLine =
            String(
                line ||
                ""
            ).trim();


        if (
            !cleanLine ||
            !cleanLine.startsWith(
                "data:"
            )
        ) {

            return;

        }


        const dataText =
            cleanLine.substring(
                5
            ).trim();


        if (
            !dataText ||
            dataText ===
                "[DONE]"
        ) {

            return;

        }


        let parsed;


        try {

            parsed =
                JSON.parse(
                    dataText
                );

        }
        catch (
            error
        ) {

            return;

        }


        const delta =
            parsed &&
            parsed.choices &&
            parsed.choices[0] &&
            parsed.choices[0].delta
                ? parsed.choices[0].delta.content
                : "";


        if (
            typeof delta !==
                "string" ||
            !delta
        ) {

            return;

        }


        fullAnswer +=
            delta;


        if (
            typeof onChunk ===
            "function"
        ) {

            onChunk(
                delta,
                fullAnswer
            );

        }

    }


    while (true) {

        const streamResult =
            await reader.read();


        if (
            streamResult.done
        ) {

            break;

        }


        buffer +=
            decoder.decode(
                streamResult.value,
                {
                    stream:
                        true
                }
            );


        buffer =
            buffer.replace(
                /\r\n/g,
                "\n"
            );


        buffer =
            buffer.replace(
                /\r/g,
                "\n"
            );


        let newlineIndex =
            buffer.indexOf(
                "\n"
            );


        while (
            newlineIndex !==
            -1
        ) {

            const line =
                buffer.substring(
                    0,
                    newlineIndex
                );


            buffer =
                buffer.substring(
                    newlineIndex + 1
                );


            processSSELine(
                line
            );


            newlineIndex =
                buffer.indexOf(
                    "\n"
                );

        }

    }


    if (
        buffer.trim()
    ) {

        processSSELine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من " +
            providerName +
            " عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// Stream Groq AI
// =====================================================

async function streamGroqAI(
    text,
    onChunk
) {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح Groq من الإعدادات."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج Groq."
        );

    }


    const streamingContext =
        await buildStreamingContext(
            text
        );


    const response =
        await fetch(
            "https://api.groq.com/openai/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "Authorization":
                        "Bearer " +
                        data.key

                },

                body:
                    JSON.stringify({

                        model:
                            data.model,

                        messages:
                            streamingContext.messages,

                        max_tokens:
                            3000,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (!response.ok) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Groq."
            )
        );

    }


    AppState.streaming.active =
        true;

    AppState.streaming.provider =
        "groq";

    AppState.streaming.text =
        "";


    try {

        const answer =
            await processOpenAICompatibleStream(
                response,
                function (
                    delta,
                    fullText
                ) {

                    AppState.streaming.text =
                        fullText;


                    if (
                        typeof onChunk ===
                        "function"
                    ) {

                        onChunk(
                            delta,
                            fullText
                        );

                    }

                },
                "Groq"
            );


        return answer;

    }
    finally {

        AppState.streaming.active =
            false;

        AppState.streaming.provider =
            "";

    }

}


// =====================================================
// Stream OpenRouter AI
// =====================================================

async function streamOpenRouterAI(
    text,
    onChunk
) {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenRouter من الإعدادات."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج OpenRouter."
        );

    }


    const streamingContext =
        await buildStreamingContext(
            text
        );


    const response =
        await fetch(
            "https://openrouter.ai/api/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "Authorization":
                        "Bearer " +
                        data.key,

                    "HTTP-Referer":
                        window.location.href,

                    "X-Title":
                        "Research Tools"

                },

                body:
                    JSON.stringify({

                        model:
                            data.model,

                        messages:
                            streamingContext.messages,

                        max_tokens:
                            3000,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (!response.ok) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenRouter."
            )
        );

    }


    AppState.streaming.active =
        true;

    AppState.streaming.provider =
        "openrouter";

    AppState.streaming.text =
        "";


    try {

        const answer =
            await processOpenAICompatibleStream(
                response,
                function (
                    delta,
                    fullText
                ) {

                    AppState.streaming.text =
                        fullText;


                    if (
                        typeof onChunk ===
                        "function"
                    ) {

                        onChunk(
                            delta,
                            fullText
                        );

                    }

                },
                "OpenRouter"
            );


        return answer;

    }
    finally {

        AppState.streaming.active =
            false;

        AppState.streaming.provider =
            "";

    }

}


// =====================================================
// Stream OpenAI AI
// =====================================================

async function streamOpenAI(
    text,
    onChunk
) {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenAI من الإعدادات."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج OpenAI."
        );

    }


    const streamingContext =
        await buildStreamingContext(
            text
        );


    const response =
        await fetch(
            "https://api.openai.com/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "Authorization":
                        "Bearer " +
                        data.key

                },

                body:
                    JSON.stringify({

                        model:
                            data.model,

                        messages:
                            streamingContext.messages,

                        max_tokens:
                            3000,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (!response.ok) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenAI."
            )
        );

    }


    AppState.streaming.active =
        true;

    AppState.streaming.provider =
        "openai";

    AppState.streaming.text =
        "";


    try {

        const answer =
            await processOpenAICompatibleStream(
                response,
                function (
                    delta,
                    fullText
                ) {

                    AppState.streaming.text =
                        fullText;


                    if (
                        typeof onChunk ===
                        "function"
                    ) {

                        onChunk(
                            delta,
                            fullText
                        );

                    }

                },
                "OpenAI"
            );


        return answer;

    }
    finally {

        AppState.streaming.active =
            false;

        AppState.streaming.provider =
            "";

    }

}


// =====================================================
// Stream Gemini AI
// =====================================================

async function streamGeminiAI(
    text,
    onChunk
) {

    const data =
        getActiveAISettings();


    if (!data.key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح Gemini من الإعدادات."
        );

    }


    if (!data.model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج Gemini من الإعدادات."
        );

    }


    const streamingContext =
        await buildStreamingContext(
            text
        );


    
    // ==================================
    // تعليمات النظام
    //
    // تختلف حسب وجود سياق مستندي فعلي.
    // ==================================

    const hasDocumentContext =
        Boolean(
            streamingContext &&
            streamingContext.documentContext &&
            streamingContext.documentContext.found
        );


    let systemInstruction;


    if (
        hasDocumentContext
    ) {

        // ==================================
        // الوضع البحثي
        // ==================================

        systemInstruction = [

            "أنت مساعد بحث أكاديمي يعمل على مستندات Word.",

            "المادة المستخرجة المرفقة في رسالة المستخدم هي المصدر الأساسي للإجابة.",

            "أجب عن السؤال اعتمادًا على المادة المستخرجة.",

            "استخرج الأفكار المرتبطة بالسؤال فقط.",

            "ادمج الأفكار المتشابهة ولا تكررها.",

            "رتب الإجابة وفق محاور السؤال.",

            "لا تحول كل مقطع مستخرج إلى فقرة مستقلة.",

            "لا تضف إلى المستند معلومة أو حكمًا أو نسبة قول غير موجودة في المادة المستخرجة.",

            "إذا لم تكف المادة المستخرجة للإجابة عن جزء من السؤال، صرّح بذلك بوضوح.",

            "لا تستخدم المعرفة العامة لسد نقص المستند إلا إذا طلب المستخدم ذلك صراحة.",

            "ضع الإحالات بعد الأفكار التي يدعمها المستند بصيغة [مقطع X].",

            "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع إحالاتها.",

            "لا تخترع أرقام المقاطع.",

            "لا تنسب فكرة إلى مقطع لا يدعمها.",

            "وجود كلمة من كلمات السؤال داخل مقطع لا يعني أن المقطع صالح للإجابة.",

            "قدّم إجابة تركيبية مترابطة.",

            "حافظ على لغة السؤال ولغة المستند.",

            "لا تبدأ باعتذار أو تمهيد غير ضروري.",

            "لا تعيد صياغة سؤال المستخدم."

        ].join(
            "\n"
        );

    }
    else {

        // ==================================
        // الوضع العام
        //
        // لا يوجد مستند مفعّل.
        // الذكاء الاصطناعي حر في الإجابة.
        // ==================================

        systemInstruction = [

            "أنت مساعد ذكاء اصطناعي عام داخل أداة تعمل في Microsoft Word.",

            "أجب عن سؤال المستخدم مباشرة.",

            "استخدم معرفتك العامة وقدراتك في الاستدلال والتحليل.",

            "لا تفترض وجود مستند أو مادة مستخرجة.",

            "لا تطلب من المستخدم تزويدك بنص مستند إلا إذا طلب المستخدم منك صراحة العمل على مستند غير متاح لك.",

            "قدّم إجابة واضحة ودقيقة ومناسبة للسؤال.",

            "حافظ على لغة المستخدم.",

            "لا تبدأ باعتذار أو تمهيد غير ضروري.",

            "لا تعيد صياغة سؤال المستخدم."

        ].join(
            "\n"
        );

    }


    // ==================================
    // تاريخ المحادثة
    // ==================================

    const conversationMessages =
        [];


    const previousMessages =
        streamingContext.messages.slice(
            0,
            -1
        );


    previousMessages.forEach(
        function (
            msg
        ) {

            conversationMessages.push({

                role:
                    msg.role ===
                    "assistant"
                        ? "model"
                        : "user",

                parts: [

                    {

                        text:
                            msg.content

                    }

                ]

            });

        }
    );


    // ==================================
    // السؤال الحالي
    // ==================================

    conversationMessages.push({

        role:
            "user",

        parts: [

            {

                text:
                    streamingContext.userContent

            }

        ]

    });


    const cleanModel =
        normalizeGeminiModel(
            data.model
        );


    const url =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        encodeURIComponent(
            cleanModel
        ) +
        ":streamGenerateContent?alt=sse";


    const response =
        await fetch(
            url,
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "x-goog-api-key":
                        data.key

                },

                body:
                    JSON.stringify({

                        systemInstruction: {

                            parts: [

                                {

                                    text:
                                        systemInstruction

                                }

                            ]

                        },

                        contents:
                            conversationMessages,

                        generationConfig: {

                            temperature:
                                0.2

                        }

                    })

            }
        );


    if (!response.ok) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Gemini."
            )
        );

    }


    if (!response.body) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من Gemini."
        );

    }


    AppState.streaming.active =
        true;

    AppState.streaming.provider =
        "gemini";

    AppState.streaming.text =
        "";


    const reader =
        response.body.getReader();


    const decoder =
        new TextDecoder(
            "utf-8"
        );


    let buffer =
        "";


    let fullAnswer =
        "";


    function processGeminiSSELine(
        line
    ) {

        const cleanLine =
            String(
                line ||
                ""
            ).trim();


        if (
            !cleanLine ||
            !cleanLine.startsWith(
                "data:"
            )
        ) {

            return;

        }


        const dataText =
            cleanLine
                .substring(
                    5
                )
                .trim();


        if (
            !dataText ||
            dataText ===
                "[DONE]"
        ) {

            return;

        }


        let parsed;


        try {

            parsed =
                JSON.parse(
                    dataText
                );

        }
        catch (
            error
        ) {

            return;

        }


        if (
            !parsed ||
            !Array.isArray(
                parsed.candidates
            ) ||
            !parsed.candidates[0]
        ) {

            return;

        }


        const candidate =
            parsed.candidates[0];


        if (
            !candidate.content ||
            !Array.isArray(
                candidate.content.parts
            )
        ) {

            return;

        }


        candidate.content.parts.forEach(
            function (
                part
            ) {

                if (
                    !part ||
                    typeof part.text !==
                        "string"
                ) {

                    return;

                }


                // ==================================
                // تجاهل أجزاء التفكير
                // ==================================

                if (
                    part.thought ===
                    true
                ) {

                    return;

                }


                const delta =
                    part.text;


                if (!delta) {

                    return;

                }


                fullAnswer +=
                    delta;


                AppState.streaming.text =
                    fullAnswer;


                if (
                    typeof onChunk ===
                    "function"
                ) {

                    onChunk(
                        delta,
                        fullAnswer
                    );

                }

            }
        );

    }


    try {

        while (true) {

            const streamResult =
                await reader.read();


            if (
                streamResult.done
            ) {

                break;

            }


            buffer +=
                decoder.decode(
                    streamResult.value,
                    {
                        stream:
                            true
                    }
                );


            buffer =
                buffer.replace(
                    /\r\n/g,
                    "\n"
                );


            buffer =
                buffer.replace(
                    /\r/g,
                    "\n"
                );


            let newlineIndex =
                buffer.indexOf(
                    "\n"
                );


            while (
                newlineIndex !==
                -1
            ) {

                const line =
                    buffer.substring(
                        0,
                        newlineIndex
                    );


                buffer =
                    buffer.substring(
                        newlineIndex + 1
                    );


                processGeminiSSELine(
                    line
                );


                newlineIndex =
                    buffer.indexOf(
                        "\n"
                    );

            }

        }


        if (
            buffer.trim()
        ) {

            processGeminiSSELine(
                buffer
            );

        }


        if (
            !fullAnswer.trim()
        ) {

            throw new Error(
                "لم يصل نص من Gemini عبر البث المتدفق."
            );

        }


        return fullAnswer.trim();

    }
    finally {

        AppState.streaming.active =
            false;

        AppState.streaming.provider =
            "";

    }

}


// =====================================================
// Stream AI
// واجهة موحدة
// =====================================================

async function streamAI(
    text,
    onChunk
) {

    const settings =
        getActiveAISettings();


    const selectedProvider =
        String(
            settings.provider ||
            "openrouter"
        ).toLowerCase();


    if (
        selectedProvider ===
        "groq"
    ) {

        return await streamGroqAI(
            text,
            onChunk
        );

    }


    if (
        selectedProvider ===
        "gemini"
    ) {

        return await streamGeminiAI(
            text,
            onChunk
        );

    }


    if (
        selectedProvider ===
        "openrouter"
    ) {

        return await streamOpenRouterAI(
            text,
            onChunk
        );

    }


    if (
        selectedProvider ===
        "openai"
    ) {

        return await streamOpenAI(
            text,
            onChunk
        );

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف: " +
        selectedProvider
    );

}


// =====================================================
// Streaming Test
// =====================================================

window.testStreamingAI =
    async function (
        text
    ) {

        try {

            const answer =
                await streamAI(
                    text,
                    function (
                        delta,
                        fullText
                    ) {

                        console.log(
                            delta
                        );

                    }
                );


            console.log(
                "الإجابة النهائية:",
                answer
            );


            return answer;

        }
        catch (error) {

            console.error(
                "فشل اختبار Streaming:",
                error
            );


            return null;

        }

    };
    
// =====================================================
// Research Tools
// PART 8
// Send Message + Streaming UI + Save AI Response
// =====================================================



// =====================================================
// Render Streaming Text
//
// لكل رسالة أثناء البث مجموعة إحالات ثابتة واحدة.
// لا يتم إنشاء citationGroupId جديد مع كل Chunk.
// =====================================================

function renderStreamingText(
    loadingElement,
    text
) {

    if (
        !loadingElement
    ) {

        return;

    }


    const value =
        String(
            text ||
            ""
        );


    if (
        value ===
        ""
    ) {

        loadingElement.innerHTML =
            "⏳ جاري التفكير...";

    }
    else {

        // ==================================
        // إنشاء هوية ثابتة لهذه الرسالة
        // أثناء البث
        // ==================================

        if (
            !loadingElement.dataset.citationGroupId
        ) {

            loadingElement.dataset.citationGroupId =
                "stream-citation-" +
                Date.now() +
                "-" +
                Math.random()
                    .toString(
                        36
                    )
                    .substring(
                        2,
                        10
                    );

        }


        loadingElement.innerHTML =
            formatAIMessage(
                value,
                currentCitationSources,
                loadingElement.dataset.citationGroupId
            );

    }


    if (
        chatArea
    ) {

        chatArea.scrollTop =
            chatArea.scrollHeight;

    }

}


// =====================================================
// Send Message
// إرسال الرسالة
// =====================================================

async function sendMessage() {

    if (!input) {

        return;

    }


    const text =
        input.value.trim();


    if (
        text ===
        ""
    ) {

        return;

    }


    // =================================================
    // إنشاء محادثة مؤقتة عند الحاجة
    // =================================================

    if (!currentChat) {

        currentChat = {

            id:
                Date.now(),

            title:
                "محادثة جديدة",

            messages:
                [],

            isTemporary:
                true,

            projectId:
                null,

            createdAt:
                new Date().toISOString(),

            updatedAt:
                new Date().toISOString()

        };

    }


    // =================================================
    // تثبيت المحادثة
    // =================================================

    if (
        currentChat.isTemporary
    ) {

        finalizeCurrentChat(
            text
        );

    }


    // =================================================
    // إضافة رسالة المستخدم
    // =================================================

    addUserMessage(
        currentChat,
        text
    );


    // =================================================
    // تحديث الواجهة
    // =================================================

    renderChat();

   

    renderSidebarChats();

    renderRecentChats();


    // =================================================
    // تنظيف صندوق الإدخال
    // =================================================

    input.value =
        "";

    input.style.height =
        "auto";


    // =================================================
    // فقاعة الإجابة
    // =================================================

    const loading =
        document.createElement(
            "div"
        );


    loading.className =
        "message ai-message";


    loading.innerHTML =
        "⏳ جاري التفكير...";


    if (chatArea) {

        chatArea.appendChild(
            loading
        );


        chatArea.scrollTop =
            chatArea.scrollHeight;

    }


    // =================================================
    // حالة العرض المتدفق
    // =================================================

    let pendingRenderText =
        "";


    let renderTimer =
        null;


    let requestFinished =
        false;


    // =================================================
    // تحديث الواجهة
    // =================================================

    function renderPendingText() {

        if (!loading) {

            return;

        }


        renderStreamingText(
            loading,
            pendingRenderText
        );


        renderTimer =
            null;

    }


    // =================================================
    // جدولة تحديث الواجهة
    // =================================================

    function scheduleRender(
        fullText
    ) {

        pendingRenderText =
            String(
                fullText ||
                ""
            );


        if (
            renderTimer !==
            null
        ) {

            return;

        }


        renderTimer =
            setTimeout(
                function () {

                    renderPendingText();

                },
                60
            );

    }


    // =================================================
    // بدء البث
    // =================================================

    try {

        const settings =
            getActiveAISettings();


        if (
            !settings.key.trim()
        ) {

            throw new Error(
                "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
            );

        }


        if (
            !settings.model.trim()
        ) {

            throw new Error(
                "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
            );

        }


        const answer =
            await streamAI(
                text,
                function (
                    delta,
                    fullText
                ) {

                    scheduleRender(
                        fullText
                    );

                }
            );


        requestFinished =
            true;


        // =================================================
        // إلغاء أي تحديث مؤجل
        // =================================================

        if (
            renderTimer !==
            null
        ) {

            clearTimeout(
                renderTimer
            );


            renderTimer =
                null;

        }


        // =================================================
        // عرض الرد النهائي
        // =================================================

        pendingRenderText =
            String(
                answer ||
                ""
            );


        renderPendingText();


        // =================================================
        // إزالة فقاعة التحميل
        // =================================================

        if (
            loading &&
            loading.parentNode
        ) {

            loading.remove();

        }


        // =================================================
        // حفظ مصادر الإحالات
        // =================================================

        const savedSources =
            cloneCitationSources(
                currentCitationSources
            );


        // =================================================
        // حفظ إجابة AI
        // =================================================

        addAIMessage(
            currentChat,
            answer,
            savedSources
        );


        // =================================================
        // تحديث الواجهة
        // =================================================

        renderChat();

        

        renderSidebarChats();

        renderRecentChats();


    }
    catch (error) {

        requestFinished =
            true;


        if (
            renderTimer !==
            null
        ) {

            clearTimeout(
                renderTimer
            );


            renderTimer =
                null;

        }


        if (
            loading &&
            loading.parentNode
        ) {

            loading.remove();

        }


        const errorMessage =
            (
                error &&
                error.message
            )
                ? error.message
                : "حدث خطأ غير معروف";


        // =================================================
        // حفظ الخطأ كرسالة AI
        // =================================================

        addAIMessage(
            currentChat,
            "خطأ: " +
            errorMessage,
            []
        );


        renderChat();

        

        renderSidebarChats();

        renderRecentChats();


        console.error(
            "فشل إرسال الرسالة:",
            error
        );

    }
    finally {

        if (
            requestFinished
        ) {

            AppState.streaming.active =
                false;

        }

    }

}


// =====================================================
// Send Button
// =====================================================

if (sendBtn) {

    sendBtn.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();

            sendMessage();

        };

}

if (referencesBtn && referencesWorkspace) {

    referencesBtn.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();

            referencesWorkspace.style.display =
                "block";


            if (chatArea) {

                chatArea.style.display =
                    "none";

            }


            if (documentTitle) {

                documentTitle.style.display =
                    "none";

            }


            if (inputArea) {

                inputArea.style.display =
                    "none";

            }

        };

}

// =====================================================
// Close References
// =====================================================

if (closeReferences) {

    closeReferences.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (referencesWorkspace) {

                referencesWorkspace.style.display =
                    "none";

            }


            if (chatArea) {

                chatArea.style.display =
                    "";

            }


            if (documentTitle) {

                documentTitle.style.display =
                    "";

            }


            if (inputArea) {

                inputArea.style.display =
                    "";

            }

        };

}

// =====================================================
// References Source
// المستند الحالي
// =====================================================

if (referencesContent) {

    let latestUnifiedReferences = [];

    let latestBibliographyComparison = null;

    function loadReferenceStyle() {

        try {

            const saved =
                JSON.parse(
                    localStorage.getItem(
                        "REFERENCE_STYLE"
                    ) || "{}"
                );

            referenceStyle.order =
                saved.order ||
                "appearance";

            referenceStyle.format =
                saved.format ||
                "author-title";

        }
        catch (error) {

            console.warn(
                "تعذر تحميل نمط المراجع:",
                error
            );

        }

    }

    loadReferenceStyle();

    referencesContent
        .querySelectorAll(
            ".references-source-option"
        )
        .forEach(
            function (button) {

                button.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();

                        const source =
                            button.getAttribute(
                                "data-reference-source"
                            );

                        // =====================================================
                        // لا نعالج إلا المستند الحالي
                        // =====================================================

                        if (
                            source !==
                            "current-document"
                        ) {

                            return;

                        }

                        referencesSourceDocument =
                            null;

                        // =====================================================
                        // واجهة اختيار المستند
                        // =====================================================

                        referencesSourceWorkspace.innerHTML =
                            `
                            <div class="references-selected-source">

                                <div class="references-selected-icon">

                                    <svg xmlns="http://www.w3.org/2000/svg"
                                         viewBox="0 0 24 24"
                                         fill="none"
                                         stroke="#000000"
                                         stroke-width="1"
                                         stroke-linecap="round"
                                         stroke-linejoin="round"
                                         aria-hidden="true">

                                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>

                                        <polyline points="14 2 14 8 20 8"/>

                                        <line x1="8" y1="13" x2="16" y2="13"/>

                                        <line x1="8" y1="17" x2="16" y2="17"/>

                                        <line x1="8" y1="9" x2="10" y2="9"/>

                                    </svg>

                                </div>

                                <div class="references-selected-info">

                                    <div class="references-selected-title">
                                        المستند المفتوح حاليًا في Word
                                    </div>

                                    <div class="references-selected-name">
                                        سيتم تحليل المستند النشط
                                    </div>

                                </div>



                            </div>

                            <div class="reference-style-options">

                                <div class="reference-style-title">
                                    نمط قائمة المراجع والحواشي
                                </div>

                                <div class="reference-style-row">

                                    <label>
                                        الترتيب
                                    </label>

                                    <select id="reference-order-select">

                                        <option value="appearance">
                                            حسب ترتيب الظهور
                                        </option>

                                        <option value="author">
                                            حسب المؤلف
                                        </option>

                                        <option value="title">
                                            حسب عنوان الكتاب
                                        </option>

                                    </select>

                                </div>

                                <div class="reference-style-row">

                                    <label>
                                        الصيغة
                                    </label>

                                    <select id="reference-format-select">

                                        <option value="author-title">
                                            المؤلف، العنوان
                                        </option>

                                        <option value="title-author">
                                            العنوان، المؤلف
                                        </option>

                                    </select>

                                </div>

                            </div>

                            <button
                                type="button"
                                id="analyze-references-btn"
                                class="analyze-references-btn">

                                تحليل المراجع

                            </button>

                            <button
                                type="button"
                                id="compare-references-btn"
                                class="compare-references-btn">

                                مقارنة قائمة المراجع

                            </button>
                            <button
                                type="button"
                                id="build-bibliography-btn"
                                class="build-bibliography-btn">

                                إنشاء/تحديث قائمة المراجع

                            </button>
                            `;


                        const referenceOrderSelect =
                            document.getElementById(
                                "reference-order-select"
                            );

                        const referenceFormatSelect =
                            document.getElementById(
                                "reference-format-select"
                            );

                        if (referenceOrderSelect) {

                            referenceOrderSelect.value =
                                referenceStyle.order;

                            referenceOrderSelect.onchange =
                                function () {

                                    referenceStyle.order =
                                        this.value;

                                    localStorage.setItem(
                                        "REFERENCE_STYLE",
                                        JSON.stringify(
                                            referenceStyle
                                        )
                                    );

                                };

                        }

                        if (referenceFormatSelect) {

                            referenceFormatSelect.value =
                                referenceStyle.format;

                            referenceFormatSelect.onchange =
                                function () {

                                    referenceStyle.format =
                                        this.value;

                                    localStorage.setItem(
                                        "REFERENCE_STYLE",
                                        JSON.stringify(
                                            referenceStyle
                                        )
                                    );

                                };

                        }

                        const analyzeReferencesBtn =
                            document.getElementById(
                                "analyze-references-btn"
                            );

                        if (
                            !analyzeReferencesBtn
                        ) {

                            return;

                        }

                        // =====================================================
                        // منع ربط الحدث أكثر من مرة
                        // =====================================================

                        analyzeReferencesBtn.onclick =
                            null;

                        // =====================================================
                        // دالة حماية HTML محلية ومضمونة
                        // لا تعتمد على escapeHTML خارجية
                        // =====================================================

                        function escapeReferenceHTML(
                            value
                        ) {

                            return String(
                                value ?? ""
                            )
                                .replace(
                                    /&/g,
                                    "&amp;"
                                )
                                .replace(
                                    /</g,
                                    "&lt;"
                                )
                                .replace(
                                    />/g,
                                    "&gt;"
                                )
                                .replace(
                                    /"/g,
                                    "&quot;"
                                )
                                .replace(
                                    /'/g,
                                    "&#039;"
                                );

                        }

                        // =====================================================
                        // تنفيذ التحليل
                        // =====================================================

                        analyzeReferencesBtn.onclick =
                            async function (e) {

                                e.preventDefault();
                                e.stopPropagation();

                                analyzeReferencesBtn.disabled =
                                    true;

                                analyzeReferencesBtn.textContent =
                                    "جارٍ قراءة المستند...";

                                try {

                                    // =================================================
                                    // 1. قراءة المستند
                                    // =================================================

                                    const referenceSources =
                                        await readReferenceSources();

                                    if (
                                        !referenceSources ||
                                        typeof referenceSources !==
                                        "object"
                                    ) {

                                        throw new Error(
                                            "تعذر الحصول على بيانات المستند."
                                        );

                                    }

                                    const mainText =
                                        String(
                                            referenceSources.mainText ||
                                            ""
                                        );

                                    if (
                                        !mainText.trim()
                                    ) {

                                        throw new Error(
                                            "تعذر الحصول على نص المستند المفتوح."
                                        );

                                    }

                                    const footnotes =
                                        Array.isArray(
                                            referenceSources.footnotes
                                        )
                                            ? referenceSources.footnotes
                                            : [];

                                    const endnotes =
                                        Array.isArray(
                                            referenceSources.endnotes
                                        )
                                            ? referenceSources.endnotes
                                            : [];

                                    // =================================================
                                    // 2. تجهيز المواد الخام
                                    // =================================================

                                    analyzeReferencesBtn.textContent =
                                        "جارٍ تجهيز مواد المراجع...";

                                    const processedReferences =
                                        processReferenceSources(
                                            referenceSources
                                        );

                                    if (
                                        !Array.isArray(
                                            processedReferences
                                        ) ||
                                        processedReferences.length === 0
                                    ) {

                                        throw new Error(
                                            "لم يتم العثور على مواد قابلة للتحليل."
                                        );

                                    }

                                    console.log(
                                        "المواد المرسلة إلى الذكاء الاصطناعي:",
                                        processedReferences
                                    );

                                    // =================================================
                                    // 3. طلب واحد فقط إلى الذكاء الاصطناعي
                                    //
                                    // الدالة الجديدة تقوم بكل شيء:
                                    // التصنيف
                                    // استخراج المراجع
                                    // فصل المتعدد
                                    // حل المصدر نفسه
                                    // دمج التكرارات
                                    // جمع المواقع
                                    // جمع الصيغ
                                    // بناء السجلات النهائية
                                    // =================================================

                                    analyzeReferencesBtn.textContent =
                                        "جارٍ تحليل وتوحيد المراجع بالذكاء الاصطناعي...";

                                    const aiAnalysisResult =
                                        await analyzeReferencesWithAI(
                                            processedReferences
                                        );

                                    if (
                                        !aiAnalysisResult ||
                                        typeof aiAnalysisResult !==
                                        "object"
                                    ) {

                                        throw new Error(
                                            "لم تصل نتيجة منظمة من الذكاء الاصطناعي."
                                        );

                                    }

                                    console.log(
                                        "النتيجة الكاملة من الذكاء الاصطناعي:",
                                        aiAnalysisResult
                                    );

                                    // =================================================
                                    // 4. استخراج الإحصاءات
                                    // =================================================

                                    const aiStats =
                                        aiAnalysisResult.stats &&
                                        typeof aiAnalysisResult.stats ===
                                        "object"

                                            ? aiAnalysisResult.stats

                                            : {};

                                    // =================================================
                                    // 5. استخراج المراجع النهائية
                                    // =================================================

                                    const finalReferenceResults =
                                        Array.isArray(aiAnalysisResult.references)
                                            ? mergeEquivalentReferences(
                                                aiAnalysisResult.references
                                            )
                                            : [];

                                            latestUnifiedReferences =
                                                finalReferenceResults;

                                            const unifiedFootnoteMap =
                                                buildUnifiedFootnoteMap(
                                                    latestUnifiedReferences
                                                );

                                            console.log(
                                                "خريطة الحواشي والمراجع الموحدة:",
                                                unifiedFootnoteMap
                                            );

                                            const footnoteSuggestions =
                                                await buildFootnoteSuggestions(
                                                    unifiedFootnoteMap
                                                );

                                            console.log(
                                                "توثيق الحواشي المقترح:",
                                                footnoteSuggestions
                                            );

                                            const footnoteSuggestionHTML =
                                                footnoteSuggestions
                                                    .filter(function (item) {
                                                        return item.reference;
                                                    })
                                                    .map(function (item, index) {

                                                        return `
                                                            <div class="footnote-suggestion-item">

                                                                <span class="footnote-suggestion-number">
                                                                    ${index + 1}
                                                                </span>

                                                                <div class="footnote-suggestion-content">

                                                                    <div class="footnote-suggestion-original">
                                                                        ${escapeReferenceHTML(
                                                                            item.originalText
                                                                        )}
                                                                    </div>

                                                                    <div class="footnote-suggestion-arrow">
                                                                        ←
                                                                    </div>

                                                                    <div class="footnote-suggestion-new">
                                                                        ${escapeReferenceHTML(
                                                                            item.suggestedText
                                                                        )}
                                                                    </div>

                                                                </div>

                                                            </div>
                                                        `;

                                                    })
                                                    .join("");

                                            referencesSourceWorkspace.insertAdjacentHTML(
                                                "beforeend",
                                                `
                                                <div class="footnote-suggestions-result">

                                                    <div class="footnote-suggestions-title">
                                                        توثيق الحواشي المقترح
                                                        <strong>
                                                            ${
                                                                footnoteSuggestions.filter(
                                                                    function (item) {
                                                                        return item.reference;
                                                                    }
                                                                ).length
                                                            }
                                                        </strong>
                                                    </div>

                                                    <div class="footnote-suggestions-list">
                                                        ${footnoteSuggestionHTML}
                                                    </div>

                                                    ${
                                                        footnoteSuggestions.some(function (item) {
                                                            return item.reference;
                                                        })
                                                            ? `
                                                                <button
                                                                    type="button"
                                                                    id="apply-footnote-suggestions-btn"
                                                                    class="apply-footnote-suggestions-btn">
                                                                    تطبيق التوثيق على الحواشي
                                                                </button>
                                                            `
                                                            : ""
                                                    }

                                                </div>
                                                `
                                            );

                                            

                                    if (
                                        !Array.isArray(
                                            finalReferenceResults
                                        )
                                    ) {

                                        throw new Error(
                                            "قائمة المراجع النهائية غير صالحة."
                                        );

                                    }

                                    console.log(
                                        "المراجع النهائية الموحدة:",
                                        finalReferenceResults
                                    );

                                    // =================================================
                                    // 6. الإحصاءات
                                    // =================================================

                                    const referenceCount =
                                        Number(
                                            aiStats.referenceCount ||
                                            0
                                        );

                                    const multipleReferenceCount =
                                        Number(
                                            aiStats.multipleReferenceCount ||
                                            0
                                        );

                                    const ibidCount =
                                        Number(
                                            aiStats.ibidCount ||
                                            0
                                        );

                                    const internalReferenceCount =
                                        Number(
                                            aiStats.internalReferenceCount ||
                                            0
                                        );

                                    const hadithCount =
                                        Number(
                                            aiStats.hadithCount ||
                                            0
                                        );

                                    const explanatoryCount =
                                        Number(
                                            aiStats.explanatoryCount ||
                                            0
                                        );

                                    const mixedCount =
                                        Number(
                                            aiStats.mixedCount ||
                                            0
                                        );

                                    const reviewCount =
                                        Number(
                                            aiStats.reviewCount ||
                                            0
                                        );

                                    // =================================================
                                    // 7. إجمالي مرات الظهور
                                    // =================================================

                                    const totalOccurrences =
                                        finalReferenceResults.reduce(
                                            function (
                                                total,
                                                reference
                                            ) {

                                                const occurrences =
                                                    Array.isArray(
                                                        reference?.occurrences
                                                    )
                                                        ? reference.occurrences.length
                                                        : 0;

                                                return (
                                                    total +
                                                    occurrences
                                                );

                                            },
                                            0
                                        );

                                    // =================================================
                                    // 8. إزالة النتيجة السابقة
                                    // =================================================

                                    referencesSourceWorkspace
                                        .querySelectorAll(
                                            ".references-analysis-result"
                                        )
                                        .forEach(
                                            function (
                                                element
                                            ) {

                                                element.remove();

                                            }
                                        );

                                    // =================================================
                                    // 9. بناء HTML للمراجع النهائية
                                    // =================================================

                                    const finalReferenceHTML =
                                        finalReferenceResults
                                            .map(
                                                function (
                                                    reference,
                                                    index
                                                ) {

                                                    const type =
                                                        String(
                                                            reference?.type ||
                                                            "book"
                                                        )
                                                            .trim()
                                                            .toLowerCase();

                                                    const typeLabel =
                                                        type === "hadith"
                                                            ? "تخريج حديث"
                                                            : type === "journal"
                                                                ? "مجلة"
                                                                : type === "article"
                                                                    ? "بحث / مقال"
                                                                    : type === "thesis"
                                                                        ? "رسالة علمية"
                                                                        : type === "website"
                                                                            ? "موقع إلكتروني"
                                                                            : "كتاب";

                                                    const author =
                                                        escapeReferenceHTML(
                                                            reference?.author ||
                                                            "مؤلف غير محدد"
                                                        );

                                                    const title =
                                                        escapeReferenceHTML(
                                                            reference?.title ||
                                                            "عنوان غير محدد"
                                                        );

                                                    const locations =
                                                        Array.isArray(
                                                            reference?.locations
                                                        )
                                                            ? reference.locations
                                                            : [];

                                                    const variants =
                                                        Array.isArray(
                                                            reference?.variants
                                                        )
                                                            ? reference.variants
                                                            : [];

                                                    const occurrences =
                                                        Array.isArray(
                                                            reference?.occurrences
                                                        )
                                                            ? reference.occurrences
                                                            : [];

                                                    const confidenceValue =
                                                        Number(
                                                            reference?.confidence ?? 0
                                                        );

                                                    const confidence =
                                                        Math.max(
                                                            0,
                                                            Math.min(
                                                                100,
                                                                Math.round(
                                                                    confidenceValue <= 1
                                                                        ? confidenceValue * 100
                                                                        : confidenceValue
                                                                )
                                                            )
                                                        );

                                                    const locationParts =
                                                        locations
                                                            .map(
                                                                function (
                                                                    location
                                                                ) {

                                                                    const volume =
                                                                        escapeReferenceHTML(
                                                                            location?.volume || ""
                                                                        );

                                                                    const page =
                                                                        escapeReferenceHTML(
                                                                            location?.page || ""
                                                                        );

                                                                    const pageRange =
                                                                        escapeReferenceHTML(
                                                                            location?.pageRange || ""
                                                                        );

                                                                    if (
                                                                        volume &&
                                                                        page
                                                                    ) {

                                                                        return `ج ${volume} / ص ${page}`;

                                                                    }

                                                                    if (
                                                                        pageRange
                                                                    ) {

                                                                        return `ص ${pageRange}`;

                                                                    }

                                                                    if (
                                                                        page
                                                                    ) {

                                                                        return `ص ${page}`;

                                                                    }

                                                                    return "";

                                                                }
                                                            )
                                                            .filter(
                                                                Boolean
                                                            );

                                                    const locationHTML =
                                                        locationParts.length
                                                            ? `
                                                                <div class="ai-reference-locations-line">
                                                                    ${locationParts
                                                                        .map(
                                                                            function (
                                                                                location
                                                                            ) {

                                                                                return `
                                                                                    <span class="ai-reference-location-chip">
                                                                                        ${location}
                                                                                    </span>
                                                                                `;

                                                                            }
                                                                        )
                                                                        .join("")}
                                                                </div>
                                                            `
                                                            : "";

                                                    const variantHTML =
                                                        variants.length
                                                            ? `
                                                                <details class="ai-reference-variants">
                                                                    <summary>
                                                                        صيغ الظهور
                                                                        <span>
                                                                            ${variants.length}
                                                                        </span>
                                                                    </summary>

                                                                    <div class="ai-reference-variants-list">

                                                                        ${variants
                                                                            .map(
                                                                                function (
                                                                                    variant
                                                                                ) {

                                                                                    return `
                                                                                        <div class="ai-reference-variant">
                                                                                            ${escapeReferenceHTML(
                                                                                                variant
                                                                                            )}
                                                                                        </div>
                                                                                    `;

                                                                                }
                                                                            )
                                                                            .join("")}

                                                                    </div>
                                                                </details>
                                                            `
                                                            : "";

                                                    const reviewHTML =
                                                        reference?.needsReview
                                                            ? `
                                                                <span class="ai-reference-review">
                                                                    يحتاج مراجعة
                                                                </span>
                                                            `
                                                            : "";

                                                    const referenceId =
                                                        escapeReferenceHTML(
                                                            reference?.id ||
                                                            `reference-${index + 1}`
                                                        );

                                                    return `
                                                        <div
                                                            class="ai-reference-row"
                                                            data-reference-id="${referenceId}">

                                                            <div class="ai-reference-index">
                                                                ${String(index + 1).padStart(2, "0")}
                                                            </div>

                                                            <div class="ai-reference-main">

                                                                <div class="ai-reference-heading">

                                                                    <span class="ai-reference-type">
                                                                        ${typeLabel}
                                                                    </span>

                                                                    <span class="ai-reference-author">
                                                                        ${author}
                                                                    </span>

                                                                </div>

                                                                <div class="ai-reference-title">
                                                                    ${title}
                                                                </div>

                                                                ${locationHTML}

                                                                <div class="ai-reference-meta">

                                                                    <span>
                                                                        ظهور
                                                                        <strong>
                                                                            ${occurrences.length}
                                                                        </strong>
                                                                    </span>

                                                                    <span class="ai-reference-separator">
                                                                        |
                                                                    </span>

                                                                    <span>
                                                                        ثقة
                                                                        <strong>
                                                                            ${confidence}%
                                                                        </strong>
                                                                    </span>

                                                                    ${reviewHTML}

                                                                </div>

                                                                ${variantHTML}

                                                            </div>

                                                        </div>
                                                    `;

                                                }
                                            )
                                            .join("");

                                    // =================================================
                                    // 10. عرض النتيجة
                                    // =================================================

                                    referencesSourceWorkspace.insertAdjacentHTML(
                                        "beforeend",
                                        `
                                        <div class="references-analysis-result">

                                            <div class="references-analysis-status success">

                                                ✓ تم تحليل وتوحيد المراجع بالذكاء الاصطناعي

                                            </div>

                                            <div class="references-analysis-info">

                                                <div>
                                                    المتن:
                                                    <strong>
                                                        ${mainText.length}
                                                    </strong>
                                                    حرف
                                                </div>

                                                <div>
                                                    الحواشي السفلية:
                                                    <strong>
                                                        ${footnotes.length}
                                                    </strong>
                                                </div>

                                                <div>
                                                    الحواشي الختامية:
                                                    <strong>
                                                        ${endnotes.length}
                                                    </strong>
                                                </div>

                                                <div>
                                                    المواد المرسلة للتحليل:
                                                    <strong>
                                                        ${processedReferences.length}
                                                    </strong>
                                                </div>

                                                <div>
                                                    المراجع:
                                                    <strong>
                                                        ${referenceCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    مواد تحتوي عدة مراجع:
                                                    <strong>
                                                        ${multipleReferenceCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    المصدر نفسه:
                                                    <strong>
                                                        ${ibidCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    إحالات داخلية:
                                                    <strong>
                                                        ${internalReferenceCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    تخريج أحاديث:
                                                    <strong>
                                                        ${hadithCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    شرح فقط:
                                                    <strong>
                                                        ${explanatoryCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    شرح + مرجع:
                                                    <strong>
                                                        ${mixedCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    تحتاج مراجعة:
                                                    <strong>
                                                        ${reviewCount}
                                                    </strong>
                                                </div>

                                                <div>
                                                    المراجع الموحدة:
                                                    <strong>
                                                        ${finalReferenceResults.length}
                                                    </strong>
                                                </div>

                                                <div>
                                                    مرات الاستشهاد:
                                                    <strong>
                                                        ${totalOccurrences}
                                                    </strong>
                                                </div>

                                            </div>

                                            <div class="references-final-results">

                                                <div class="references-final-results-title">

                                                    المراجع الموحدة
                                                    <strong>
                                                        ${finalReferenceResults.length}
                                                    </strong>

                                                </div>

                                                <div class="references-final-list">

                                                    ${finalReferenceHTML}

                                                </div>

                                            </div>

                                        </div>
                                        `
                                    );

                                    // =================================================
                                    // 11. إنهاء التحليل
                                    // =================================================

                                    analyzeReferencesBtn.textContent =
                                        "✓ تم التحليل والتوحيد";

                                    analyzeReferencesBtn.disabled =
                                        false;

                                    // =================================================
                                    // 12. السجل التشخيصي
                                    // =================================================

                                    console.log(
                                        "اكتمل تحليل وتوحيد المراجع:",
                                        {
                                            mainTextLength:
                                                mainText.length,

                                            footnotes:
                                                footnotes.length,

                                            endnotes:
                                                endnotes.length,

                                            processedMaterials:
                                                processedReferences.length,

                                            finalReferences:
                                                finalReferenceResults.length,

                                            referenceCount:
                                                referenceCount,

                                            multipleReferenceCount:
                                                multipleReferenceCount,

                                            ibidCount:
                                                ibidCount,

                                            internalReferenceCount:
                                                internalReferenceCount,

                                            hadithCount:
                                                hadithCount,

                                            explanatoryCount:
                                                explanatoryCount,

                                            mixedCount:
                                                mixedCount,

                                            reviewCount:
                                                reviewCount,

                                            totalOccurrences:
                                                totalOccurrences,

                                            references:
                                                finalReferenceResults
                                        }
                                    );

                                }
                                catch (
                                    error
                                ) {

                                    analyzeReferencesBtn.disabled =
                                        false;

                                    analyzeReferencesBtn.textContent =
                                        "تحليل المراجع";

                                    referencesSourceWorkspace
                                        .querySelectorAll(
                                            ".references-analysis-result"
                                        )
                                        .forEach(
                                            function (
                                                element
                                            ) {

                                                element.remove();

                                            }
                                        );

                                    const errorMessage =
                                        String(
                                            error?.message ||
                                            "حدث خطأ غير معروف."
                                        );

                                    const safeErrorMessage =
                                        escapeReferenceHTML(
                                            errorMessage
                                        );

                                    referencesSourceWorkspace.insertAdjacentHTML(
                                        "beforeend",
                                        `
                                        <div class="references-analysis-result">

                                            <div class="references-analysis-status error">

                                                ⚠ تعذر تحليل المستند

                                            </div>

                                            <div class="references-analysis-info">

                                                ${safeErrorMessage}

                                            </div>

                                        </div>
                                        `
                                    );

                                    console.error(
                                        "فشل تحليل وتوحيد المراجع:",
                                        error
                                    );

                                }

                                const compareReferencesBtn =
                                    document.getElementById(
                                        "compare-references-btn"
                                    );

                                if (compareReferencesBtn) {

                                    compareReferencesBtn.onclick =
                                        async function (e) {

                                            e.preventDefault();
                                            e.stopPropagation();

                                            compareReferencesBtn.disabled = true;
                                            compareReferencesBtn.textContent =
                                                "جارٍ مقارنة قائمة المراجع...";

                                            try {

                                                const bibliography =
                                                    await readBibliographyFromCurrentDocument();

                                                const comparison =
                                                    compareUnifiedReferencesWithBibliography(
                                                        latestUnifiedReferences,
                                                        bibliography
                                                    );

                                                    latestBibliographyComparison =
                                                        comparison;

                                                    const buildBibliographyBtn =
                                                        document.getElementById(
                                                            "build-bibliography-btn"
                                                        );

                                                    if (buildBibliographyBtn) {

                                                        buildBibliographyBtn.onclick =
                                                            async function (e) {

                                                                e.preventDefault();
                                                                e.stopPropagation();

                                                                if (
                                                                    !latestBibliographyComparison
                                                                ) {

                                                                    console.warn(
                                                                        "يجب إجراء المقارنة أولًا."
                                                                    );

                                                                    return;

                                                                }

                                                                const finalBibliography =
                                                                    buildFinalBibliography(
                                                                        latestUnifiedReferences,
                                                                        latestBibliographyComparison
                                                                    );

                                                                console.log(
                                                                    "قائمة المراجع النهائية المقترحة:",
                                                                    finalBibliography
                                                                );

                                                                const result =
                                                                    await writeFinalBibliographyToDocument(
                                                                        finalBibliography,
                                                                        latestBibliographyComparison
                                                                    );

                                                                if (result.created) {

                                                                    buildBibliographyBtn.textContent =
                                                                        `✓ تم إنشاء قائمة المراجع (${result.added})`;

                                                                }
                                                                else if (result.added > 0) {

                                                                    buildBibliographyBtn.textContent =
                                                                        `✓ تم تحديث قائمة المراجع (+${result.added})`;

                                                                }
                                                                else {

                                                                    buildBibliographyBtn.textContent =
                                                                        "✓ قائمة المراجع مكتملة";

                                                                }

                                                            };

                                                    }

                                                console.log(
                                                    "نتيجة مقارنة قائمة المراجع:",
                                                    comparison
                                                );

                                                referencesSourceWorkspace.insertAdjacentHTML(
                                                    "beforeend",
                                                    `
                                                    <div class="references-comparison-result">

                                                        <div class="references-comparison-title">
                                                            مقارنة قائمة المراجع
                                                        </div>

                                                        <div class="references-comparison-summary">

                                                            <div>
                                                                المراجع الموحدة:
                                                                <strong>
                                                                    ${comparison.totalUnifiedReferences}
                                                                </strong>
                                                            </div>

                                                            <div>
                                                                عناصر القائمة:
                                                                <strong>
                                                                    ${comparison.totalBibliographyEntries}
                                                                </strong>
                                                            </div>

                                                            <div>
                                                                المطابق:
                                                                <strong>
                                                                    ${comparison.matchedCount}
                                                                </strong>
                                                            </div>

                                                            <div>
                                                                غير موجود في القائمة:
                                                                <strong>
                                                                    ${comparison.missingFromBibliography.length}
                                                                </strong>
                                                            </div>

                                                            <div>
                                                                موجود دون استشهاد:
                                                                <strong>
                                                                    ${comparison.unusedBibliography.length}
                                                                </strong>
                                                            </div>

                                                        </div>

                                                        ${
                                                            comparison.missingFromBibliography.length
                                                                ? `
                                                                    <div class="references-comparison-section">

                                                                        <div class="references-comparison-section-title">
                                                                            المراجع المستشهد بها وغير الموجودة في القائمة
                                                                            <strong>
                                                                                ${comparison.missingFromBibliography.length}
                                                                            </strong>
                                                                        </div>

                                                                        <div class="references-comparison-list">

                                                                            ${comparison.missingFromBibliography
                                                                                .map(
                                                                                    function (
                                                                                        reference,
                                                                                        index
                                                                                    ) {

                                                                                        return `
                                                                                            <div class="references-comparison-item">

                                                                                                <span class="references-comparison-number">
                                                                                                    ${index + 1}
                                                                                                </span>

                                                                                                <span class="references-comparison-text">

                                                                                                    <strong>
                                                                                                        ${escapeReferenceHTML(
                                                                                                            reference?.author ||
                                                                                                            "مؤلف غير محدد"
                                                                                                        )}
                                                                                                    </strong>

                                                                                                    ${
                                                                                                        reference?.title
                                                                                                            ? `
                                                                                                                —
                                                                                                                ${escapeReferenceHTML(
                                                                                                                    reference.title
                                                                                                                )}
                                                                                                            `
                                                                                                            : ""
                                                                                                    }

                                                                                                </span>

                                                                                            </div>
                                                                                        `;

                                                                                    }
                                                                                )
                                                                                .join("")}

                                                                        </div>

                                                                    </div>
                                                                `
                                                                : ""
                                                        }

                                                        ${
                                                            comparison.unusedBibliography.length
                                                                ? `
                                                                    <div class="references-comparison-section">

                                                                        <div class="references-comparison-section-title">
                                                                            مراجع موجودة في القائمة ولم يُعثر لها على استشهاد
                                                                            <strong>
                                                                                ${comparison.unusedBibliography.length}
                                                                            </strong>
                                                                        </div>

                                                                        <div class="references-comparison-list">

                                                                            ${comparison.unusedBibliography
                                                                                .map(
                                                                                    function (
                                                                                        item,
                                                                                        index
                                                                                    ) {

                                                                                        return `
                                                                                            <div class="references-comparison-item">

                                                                                                <span class="references-comparison-number">
                                                                                                    ${index + 1}
                                                                                                </span>

                                                                                                <span class="references-comparison-text">
                                                                                                    ${escapeReferenceHTML(
                                                                                                        item.text
                                                                                                    )}
                                                                                                </span>

                                                                                            </div>
                                                                                        `;

                                                                                    }
                                                                                )
                                                                                .join("")}

                                                                        </div>

                                                                    </div>
                                                                `
                                                                : ""
                                                        }

                                                    </div>
                                                    `
                                                );

                                                compareReferencesBtn.textContent =
                                                    "✓ تمت المقارنة";

                                            }
                                            catch (error) {

                                                compareReferencesBtn.textContent =
                                                    "مقارنة قائمة المراجع";

                                                console.error(
                                                    "فشل مقارنة قائمة المراجع:",
                                                    error
                                                );

                                            }
                                            finally {

                                                compareReferencesBtn.disabled = false;

                                            }

                                        };

                                }

                            };

                    };

            }
        );

}

function formatFootnoteReference(
    reference,
    occurrence
) {

    if (!reference) {
        return "";
    }

    const base =
        formatReferenceForOutput(
            reference
        );

    let locationText = "";

    const volume =
        String(
            occurrence?.volume ||
            ""
        ).trim();

    const page =
        String(
            occurrence?.page ||
            ""
        ).trim();

    const pageRange =
        String(
            occurrence?.pageRange ||
            ""
        ).trim();

    if (volume && page) {

        locationText =
            `${volume}/${page}`;

    }
    else if (volume && pageRange) {

        locationText =
            `${volume}/${pageRange}`;

    }
    else if (pageRange) {

        locationText =
            `ص ${pageRange}`;

    }
    else if (page) {

        locationText =
            `ص ${page}`;

    }

    return locationText
        ? `${base}، ${locationText}`
        : base;
}

function buildUnifiedFootnoteMap(unifiedReferences) {

    const map = new Map();

    if (!Array.isArray(unifiedReferences)) {
        return map;
    }

    function normalizeSource(source) {

        const value =
            String(source || "")
                .trim()
                .toLowerCase();

        if (
            value === "footnote" ||
            value === "footnotes" ||
            value === "حاشية" ||
            value === "حاشية سفلية" ||
            value === "الحواشي السفلية"
        ) {
            return "footnote";
        }

        if (
            value === "endnote" ||
            value === "endnotes" ||
            value === "حاشية ختامية" ||
            value === "الحواشي الختامية"
        ) {
            return "endnote";
        }

        return value;
    }

    unifiedReferences.forEach(function (reference) {

        const occurrences =
            Array.isArray(reference?.occurrences)
                ? reference.occurrences
                : [];

        occurrences.forEach(function (occurrence) {

            const noteNumber =
                Number(
                    occurrence?.noteNumber
                );

            if (
                !Number.isFinite(noteNumber)
            ) {
                return;
            }

            const source =
                normalizeSource(
                    occurrence?.source
                );

            if (!source) {
                return;
            }

            const key =
                `${source}:${noteNumber}`;

            if (!map.has(key)) {
                map.set(key, []);
            }

            map.get(key).push({
                reference: reference,
                occurrence: occurrence
            });

        });

    });

    console.log(
        "مفاتيح خريطة الحواشي:",
        Array.from(map.keys())
    );

    return map;
}

async function buildFootnoteSuggestions(
    unifiedFootnoteMap
) {

    return await Word.run(
        async function (context) {

            const body =
                context.document.body;

            const footnotes =
                body.footnotes;

            const endnotes =
                body.endnotes;

            footnotes.load("items");
            endnotes.load("items");

            await context.sync();

            footnotes.items.forEach(function (note) {
                note.reference.load("text");
                note.body.load("text");
            });

            endnotes.items.forEach(function (note) {
                note.reference.load("text");
                note.body.load("text");
            });

            await context.sync();

            const suggestions = [];

            function addNotes(notes, source) {

                notes.forEach(function (note, index) {

                    const noteNumber =
                        index + 1;

                    const key =
                        `${source}:${noteNumber}`;

                    const matches =
                        unifiedFootnoteMap.get(key) || [];

                    /*
                     * الحاشية قد تحتوي عدة مراجع.
                     */
                    if (matches.length === 0) {

                        suggestions.push({
                            source: source,
                            noteNumber: noteNumber,
                            originalText:
                                String(
                                    note.body?.text || ""
                                ).trim(),
                            references: [],
                            suggestedTexts: []
                        });

                        return;
                    }

                    suggestions.push({

                        source: source,

                        noteNumber: noteNumber,

                        originalText:
                            String(
                                note.body?.text || ""
                            ).trim(),

                        references:
                            matches.map(function (item) {
                                return {
                                    reference:
                                        item.reference,
                                    occurrence:
                                        item.occurrence
                                };
                            }),

                        suggestedTexts:
                            matches.map(function (item) {

                                return formatFootnoteReference(
                                    item.reference,
                                    item.occurrence
                                );

                            })

                    });

                });

            }

            addNotes(
                footnotes.items,
                "footnote"
            );

            addNotes(
                endnotes.items,
                "endnote"
            );

            return suggestions;

        }
    );

}

function mergeEquivalentReferences(references) {

    const merged = [];

    function clean(value) {
        return String(value ?? "")
            .trim()
            .replace(/\s+/g, " ")
            .replace(/[.,،:؛]+$/g, "");
    }

    function titleKey(title) {
        return clean(title)
            .replace(/^الـ/, "ال")
            .toLowerCase();
    }

    references.forEach(function (reference) {

        const author = clean(reference.author);
        const title = clean(reference.title);

        let existing = merged.find(function (item) {

            if (
                author &&
                item.author &&
                author !== clean(item.author)
            ) {
                return false;
            }

            const a = titleKey(title);
            const b = titleKey(item.title);

            return (
                a === b ||
                a.includes(b) ||
                b.includes(a)
            );

        });

        if (!existing) {

            merged.push({
                ...reference,
                locations: Array.isArray(reference.locations)
                    ? [...reference.locations]
                    : [],
                variants: Array.isArray(reference.variants)
                    ? [...reference.variants]
                    : [],
                occurrences: Array.isArray(reference.occurrences)
                    ? [...reference.occurrences]
                    : []
            });

            return;
        }

        if (
            title.length >
            clean(existing.title).length
        ) {
            existing.title = title;
        }

        (reference.locations || []).forEach(function (location) {

            const exists =
                existing.locations.some(function (old) {

                    return (
                        String(old.volume ?? "") === String(location.volume ?? "") &&
                        String(old.page ?? "") === String(location.page ?? "") &&
                        String(old.pageRange ?? "") === String(location.pageRange ?? "")
                    );

                });

            if (!exists) {
                existing.locations.push(location);
            }

        });

        (reference.variants || []).forEach(function (variant) {

            const value = clean(variant);

            if (
                value &&
                !existing.variants.some(function (item) {
                    return clean(item) === value;
                })
            ) {
                existing.variants.push(value);
            }

        });

        (reference.occurrences || []).forEach(function (occurrence) {

            const exists =
                existing.occurrences.some(function (old) {

                    return (
                        old.materialId === occurrence.materialId &&
                        old.source === occurrence.source &&
                        String(old.noteNumber ?? "") ===
                        String(occurrence.noteNumber ?? "")
                    );

                });

            if (!exists) {
                existing.occurrences.push(occurrence);
            }

        });

        existing.confidence =
            Math.max(
                Number(existing.confidence ?? 0),
                Number(reference.confidence ?? 0)
            );

        existing.needsReview =
            Boolean(
                existing.needsReview ||
                reference.needsReview
            );

    });

    return merged;
}


// =====================================================
// Scope Elements
// =====================================================

const scopeBtn =
    document.getElementById("scope-btn");

const scopePanel =
    document.getElementById("scope-panel");

const scopeLabel =
    document.getElementById("scope-label");

const scopeStatus =
    document.getElementById("scope-status");

const scopeClear =
    document.getElementById("scope-clear");


function updateScopeStatus() {

    if (
        !scopeStatus ||
        !scopeLabel
    ) {

        return;

    }


    if (
        researchScope &&
        researchScope.name
    ) {

        scopeLabel.textContent =
            researchScope.name;


        scopeStatus.classList.add(
            "active"
        );

    }

    else {

        scopeLabel.textContent =
            "";

        scopeStatus.classList.remove(
            "active"
        );

    }

}

// =====================================================
// Clear Scope
// إلغاء نطاق البحث
// =====================================================

if (scopeClear) {

    scopeClear.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();

            researchScope = null;

            updateScopeStatus();

        };

}
// =====================================================
// Scope Button
// فتح وإغلاق لوحة نطاق البحث
// =====================================================

if (scopeBtn && scopePanel) {

    console.log(
        "SCOPE:",
        scopeBtn,
        scopePanel
    );

    scopeBtn.onclick =
        function (e) {

            console.log(
                "SCOPE CLICK"
            );

            e.preventDefault();
            e.stopPropagation();


            const isOpen =
                scopePanel.classList.contains(
                    "open"
                );


            if (isOpen) {

                scopePanel.classList.remove(
                    "open"
                );

                scopePanel.setAttribute(
                    "aria-hidden",
                    "true"
                );

                return;

            }


            scopePanel.classList.add(
                "open"
            );

            scopePanel.setAttribute(
                "aria-hidden",
                "false"
            );

        };


    // إغلاق اللوحة عند النقر خارجها
    document.addEventListener(
        "click",
        function (e) {

            if (
                !scopePanel.contains(e.target) &&
                !scopeBtn.contains(e.target)
            ) {

                scopePanel.classList.remove(
                    "open"
                );

                scopePanel.setAttribute(
                    "aria-hidden",
                    "true"
                );

            }

        }
    );

}

// =====================================================
// Scope Options
// التنقل داخل لوحة نطاق البحث
// =====================================================

if (
    scopePanel
) {

    const scopePanelContent =
        scopePanel.querySelector(
            ".scope-panel-content"
        );

    const scopePanelTitle =
        scopePanel.querySelector(
            ".scope-panel-title"
        );


    function showScopeHome() {

        scopePanelTitle.textContent =
            "نطاق البحث";


        scopePanelContent.innerHTML =
            `
            <button
                type="button"
                class="scope-option"
                data-scope-type="project">

                <span class="scope-option-icon">
                    ▱
                </span>

                <span class="scope-option-text">
                    المشاريع
                </span>

            </button>


            <button
                type="button"
                class="scope-option"
                data-scope-type="document">

                <span class="scope-option-icon">
                    ▤
                </span>

                <span class="scope-option-text">
                    المستندات
                </span>

            </button>


            <button
                type="button"
                class="scope-option"
                data-scope-type="reference">

                <span class="scope-option-icon">
                    ≡
                </span>

                <span class="scope-option-text">
                    المراجع
                </span>

            </button>


            <button
                type="button"
                class="scope-option"
                data-scope-type="library">

                <span class="scope-option-icon">
                    ▥
                </span>

                <span class="scope-option-text">
                    المكتبة
                </span>

            </button>
            `;


        bindScopeOptions();

    }


    function showScopeType(
        type
    ) {

        const titles = {

            project:
                "المشاريع",

            document:
                "المستندات",

            reference:
                "المراجع",

            library:
                "المكتبة"

        };


        const labels = {

            project:
                "مشروع",

            document:
                "مستند",

            reference:
                "مرجع",

            library:
                "مكتبة"

        };


        scopePanelTitle.innerHTML =
            `
            <button
                type="button"
                class="scope-back-btn"
                aria-label="رجوع">
                ←
            </button>

            <span>
                ${titles[type] || "نطاق البحث"}
            </span>
            `;


        scopePanelContent.innerHTML =
            `
            <button
                type="button"
                class="scope-mode-option"
                data-scope-mode="all">

                <span class="scope-mode-mark">
                    ◉
                </span>

                <span>
                    جميع ${labels[type] || ""}
                </span>

            </button>


            <button
                type="button"
                class="scope-mode-option"
                data-scope-mode="specific"
                data-scope-type="${type}">

                <span class="scope-mode-mark">
                    ⊙
                </span>

                <span>
                    ${labels[type] || ""} محدد
                </span>

            </button>
            `;


        const backButton =
            scopePanel.querySelector(
                ".scope-back-btn"
            );


        if (
            backButton
        ) {

            backButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();

                    showScopeHome();

                };

        }


        scopePanelContent
            .querySelectorAll(
                ".scope-mode-option"
            )
            .forEach(
                function (
                    button
                ) {

                    button.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();
                            e.stopPropagation();


                            const mode =
                                button.getAttribute(
                                    "data-scope-mode"
                                );


                            // =========================================
                            // جميع
                            // =========================================

                            if (
                                mode ===
                                "all"
                            ) {

                                const scopeNames = {

                                    project:
                                        "جميع المشاريع",

                                    document:
                                        "جميع المستندات",

                                    reference:
                                        "جميع المراجع",

                                    library:
                                        "جميع المكتبة"

                                };


                                researchScope = {

                                    type:
                                        type,

                                    scope:
                                        "all",

                                    id:
                                        null,

                                    name:
                                        scopeNames[type] ||
                                        "نطاق البحث"

                                };

                                updateScopeStatus();


                                scopePanel.classList.remove(
                                    "open"
                                );


                                scopePanel.setAttribute(
                                    "aria-hidden",
                                    "true"
                                );


                                console.log(
                                    "نطاق البحث الحالي:",
                                    researchScope
                                );


                                return;

                            }


                            // =========================================
                            // محدد
                            // =========================================

                            if (
                                mode ===
                                "specific"
                            ) {

                                if (
                                    type ===
                                    "project"
                                ) {

                                    showProjectScopePicker();

                                }

                                else if (
                                    type ===
                                    "document"
                                ) {

                                    showDocumentScopePicker();

                                }

                                else if (
                                    type ===
                                    "reference"
                                ) {

                                    showReferenceScopePicker();

                                }

                                else if (
                                    type ===
                                    "library"
                                ) {

                                    showLibraryScopePicker();

                                }

                            }

                        };

                }
            );

    }

    function showProjectScopePicker() {

        scopePanelTitle.innerHTML =
            `
            <button
                type="button"
                class="scope-back-btn"
                aria-label="رجوع">
                ←
            </button>

            <span>
                اختيار المشروع
            </span>
            `;


        scopePanelContent.innerHTML =
            "";


        const backButton =
            scopePanel.querySelector(
                ".scope-back-btn"
            );


        if (
            backButton
        ) {

            backButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();

                    showScopeType(
                        "project"
                    );

                };

        }


        if (
            !Array.isArray(
                projects
            ) ||
            projects.length ===
                0
        ) {

            scopePanelContent.innerHTML =
                `
                <div class="scope-empty">
                    لا توجد مشاريع
                </div>
                `;

            return;

        }


        projects.forEach(
            function (
                project
            ) {

                if (
                    !project
                ) {

                    return;

                }


                const button =
                    document.createElement(
                        "button"
                    );


                button.type =
                    "button";


                button.className =
                    "scope-project-item";


                button.innerHTML =
                    `
                    <span
                        class="scope-project-icon">
                        ▱
                    </span>

                    <span
                        class="scope-project-name">
                        ${project.name}
                    </span>
                    `;


                button.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        researchScope = {

                            type:
                                "project",

                            scope:
                                "specific",

                            id:
                                project.id,

                            name:
                                project.name

                        };

                        updateScopeStatus();


                        scopePanel.classList.remove(
                            "open"
                        );

                        scopePanel.setAttribute(
                            "aria-hidden",
                            "true"
                        );


                        console.log(
                            "نطاق البحث الحالي:",
                            researchScope
                        );

                    };


                scopePanelContent.appendChild(
                    button
                );

            }
        );

    }

    function showDocumentScopePicker() {

        scopePanelTitle.innerHTML =
            `
            <button
                type="button"
                class="scope-back-btn"
                aria-label="رجوع">
                ←
            </button>

            <span>
                اختيار المستند
            </span>
            `;


        scopePanelContent.innerHTML =
            "";


        const backButton =
            scopePanel.querySelector(
                ".scope-back-btn"
            );


        if (
            backButton
        ) {

            backButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();

                    showScopeType(
                        "document"
                    );

                };

        }


        if (
            !Array.isArray(
                documents
            ) ||
            documents.length ===
                0
        ) {

            scopePanelContent.innerHTML =
                `
                <div class="scope-empty">
                    لا توجد مستندات
                </div>
                `;

            return;

        }


        documents.forEach(
            function (
                documentItem
            ) {

                if (
                    !documentItem
                ) {

                    return;

                }


                const button =
                    document.createElement(
                        "button"
                    );


                button.type =
                    "button";


                button.className =
                    "scope-project-item";


                button.innerHTML =
                    `
                    <span
                        class="scope-project-icon">
                        ▤
                    </span>

                    <span
                        class="scope-project-name">
                        ${documentItem.name}
                    </span>
                    `;


                button.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        researchScope = {

                            type:
                                "document",

                            scope:
                                "specific",

                            id:
                                documentItem.id,

                            name:
                                documentItem.name

                        };

                        updateScopeStatus();


                        setCurrentDocument(
                            documentItem
                        );


                        scopePanel.classList.remove(
                            "open"
                        );

                        scopePanel.setAttribute(
                            "aria-hidden",
                            "true"
                        );


                        console.log(
                            "نطاق البحث الحالي:",
                            researchScope
                        );

                    };


                scopePanelContent.appendChild(
                    button
                );

            }
        );

    }

    function showReferenceScopePicker() {

        scopePanelTitle.innerHTML =
            `
            <button
                type="button"
                class="scope-back-btn"
                aria-label="رجوع">
                ←
            </button>

            <span>
                اختيار المرجع
            </span>
            `;


        scopePanelContent.innerHTML =
            "";


        const backButton =
            scopePanel.querySelector(
                ".scope-back-btn"
            );


        if (
            backButton
        ) {

            backButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();

                    showScopeType(
                        "reference"
                    );

                };

        }


        // المراجع غير مبنية بعد في النظام الحالي
        scopePanelContent.innerHTML =
            `
            <div class="scope-empty">
                لا توجد مراجع متاحة حاليًا
            </div>
            `;

    }

    function showLibraryScopePicker() {

        scopePanelTitle.innerHTML =
            `
            <button
                type="button"
                class="scope-back-btn"
                aria-label="رجوع">
                ←
            </button>

            <span>
                اختيار المكتبة
            </span>
            `;


        scopePanelContent.innerHTML =
            "";


        const backButton =
            scopePanel.querySelector(
                ".scope-back-btn"
            );


        if (
            backButton
        ) {

            backButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();

                    showScopeType(
                        "library"
                    );

                };

        }


        // المكتبات غير مبنية بعد في النظام الحالي
        scopePanelContent.innerHTML =
            `
            <div class="scope-empty">
                لا توجد مكتبات متاحة حاليًا
            </div>
            `;

    }


    function bindScopeOptions() {

        scopePanelContent
            .querySelectorAll(
                ".scope-option"
            )
            .forEach(
                function (
                    button
                ) {

                    button.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();
                            e.stopPropagation();


                            const type =
                                button.getAttribute(
                                    "data-scope-type"
                                );


                            showScopeType(
                                type
                            );

                        };

                }
            );

    }


    bindScopeOptions();

}
// =====================================================
// Keyboard
// =====================================================

if (input) {

    input.onkeydown =
        function (e) {

            if (
                e.key ===
                    "Enter" &&
                !e.shiftKey
            ) {

                e.preventDefault();


                sendMessage();

            }

        };

}


// =====================================================
// Input Auto Height
// تمدد صندوق الإدخال مع الكتابة
// =====================================================

if (input) {

    input.oninput =
        function () {

            input.style.height =
                "auto";


            const maxHeight =
                120;


            input.style.height =
                Math.min(
                    input.scrollHeight,
                    maxHeight
                ) +
                "px";


            input.style.overflowY =
                input.scrollHeight >
                maxHeight
                    ? "auto"
                    : "hidden";

        };

}


// =====================================================
// Send Message Test
// =====================================================

window.testSendMessage =
    async function (
        text
    ) {

        if (input) {

            input.value =
                String(
                    text ||
                    ""
                );

        }


        await sendMessage();

    };
    // =====================================================
// Research Tools
// PART 9
// Projects + Documents + Chats + Sidebar
// =====================================================


// =====================================================
// Project Icon
// =====================================================

const projectIcon = `
    <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="#000000"
        stroke-width="1"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true">

        <path
            d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z">
        </path>

    </svg>
`;


// =====================================================
// Chat Icon
// =====================================================

const chatIcon = `
    <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="#000000"
        stroke-width="1"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true">

        <path
            d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z">
        </path>

    </svg>
`;



// =====================================================
// RENDER PROJECTS
//
// المشروع هو الحاوية الأساسية:
//    المشروع
//      ├─ المستندات
//      ├─ المحادثات
//      └─ المراجع
//
// لا نحذف القائمة القديمة الآن.
// هذا النقل تدريجي، ولذلك تبقى الوظائف القديمة تعمل.
// =====================================================

function renderProjects() {

    if (
        !projectsList
    ) {

        return;

    }


    projectsList.innerHTML =
        "";


    if (
        !Array.isArray(
            projects
        ) ||
        projects.length ===
            0
    ) {

        projectsList.innerHTML =
            `
            <div class="empty-project">
                لا توجد مشاريع
            </div>
            `;

        return;

    }


    // =====================================================
    // مستندات المشروع
    // =====================================================

    function renderProjectDocuments(
        container,
        project
    ) {

        if (
            !container ||
            !project
        ) {

            return;

        }


        // =========================================
        // تنظيف المحتوى
        // =========================================

        container.innerHTML =
            "";


        // =========================================
        // صندوق إضافة مستند
        // =========================================

        const addDocumentBox =
            document.createElement(
                "div"
            );


        addDocumentBox.className =
            "project-add-document-box";


        const addDocumentButton =
            document.createElement(
                "button"
            );


        addDocumentButton.type =
            "button";


        addDocumentButton.className =
            "project-add-document";


        addDocumentButton.title =
            "إضافة مستند Word";


        addDocumentButton.innerHTML =
            `
            <span class="project-add-document-icon">
                +
            </span>

            <span class="project-add-document-text">
                إضافة مستند
            </span>
            `;


        addDocumentBox.appendChild(
            addDocumentButton
        );


        container.appendChild(
            addDocumentBox
        );


        // =========================================
        // إضافة مستند إلى هذا المشروع
        // =========================================

        addDocumentButton.onclick =
            function (
                e
            ) {

                e.preventDefault();

                e.stopPropagation();


                // تثبيت المشروع المستهدف فقط
                currentProject =
                    project;


                // أداة اختيار ملف Word الأصلية
                if (
                    !wordDocumentPicker
                ) {

                    console.warn(
                        "لم يتم العثور على أداة اختيار ملف Word."
                    );

                    return;

                }


                // السماح باختيار نفس الملف مرة أخرى
                wordDocumentPicker.value =
                    "";


                // فتح نافذة اختيار الملف
                wordDocumentPicker.click();

            };


        // =========================================
        // مستندات المشروع
        // =========================================

        const projectDocuments =
            getProjectDocuments(
                project.id
            );


        // =========================================
        // لا توجد مستندات
        // =========================================

        if (
            !Array.isArray(
                projectDocuments
            ) ||
            projectDocuments.length ===
                0
        ) {

            const empty =
                document.createElement(
                    "div"
                );


            empty.className =
                "empty-document";


            empty.textContent =
                "لا توجد مستندات";


            container.appendChild(
                empty
            );


            return;

        }


        // =========================================
        // رسم مستندات المشروع
        // =========================================

        projectDocuments.forEach(
            function (
                documentItem,
                index
            ) {

                if (
                    !documentItem
                ) {

                    return;

                }


                const item =
                    document.createElement(
                        "div"
                    );


                item.className =
                    "project-document-item";


                                // =========================================
                                // العنوان
                                // =========================================

                                const title =
                                    document.createElement(
                                        "span"
                                    );


                                title.className =
                                    "project-document-title";


                                const titleText =
                                    document.createElement(
                                        "span"
                                    );


                                titleText.className =
                                    "document-name";


                                titleText.textContent =
                                    documentItem.name;


                                title.appendChild(
                                    titleText
                                );


                                // =========================================
                                // حالة المستند
                                // =========================================

                                const status =
                                    document.createElement(
                                        "span"
                                    );


                                status.className =
                                    "project-document-status";


                                if (
                                    documentItem.indexStatus ===
                                    "indexed"
                                ) {

                                    status.textContent =
                                        "✓ مفهرس";


                                    if (
                                        documentItem.indexTokenCount
                                    ) {

                                        status.textContent +=
                                            " · " +
                                            documentItem.indexTokenCount +
                                            " كلمة";

                                    }


                                    if (
                                        documentItem.indexUniqueTerms
                                    ) {

                                        status.textContent +=
                                            " · " +
                                            documentItem.indexUniqueTerms +
                                            " فريدة";

                                    }


                                    if (
                                        documentItem.indexUniqueFamilies
                                    ) {

                                        status.textContent +=
                                            " · " +
                                            documentItem.indexUniqueFamilies +
                                            " عائلة";

                                    }

                                }
                                else if (
                                    documentItem.indexStatus ===
                                    "indexing"
                                ) {

                                    status.textContent =
                                        "جارٍ الفهرسة...";

                                }
                                else if (
                                    documentItem.indexStatus ===
                                    "error"
                                ) {

                                    status.textContent =
                                        "⚠ فشل الفهرسة";

                                }
                                else if (
                                    documentItem.readStatus ===
                                    "reading"
                                ) {

                                    status.textContent =
                                        "جارٍ القراءة...";

                                }
                                else if (
                                    documentItem.readStatus ===
                                    "read"
                                ) {

                                    status.textContent =
                                        "✓ تمت القراءة";

                                }
                                else {

                                    status.textContent =
                                        "جديد";

                                }

                // =========================================
                // قائمة المستند
                // =========================================

                const menuButton =
                    document.createElement(
                        "button"
                    );


                menuButton.type =
                    "button";


                menuButton.className =
                    "project-document-menu";


                menuButton.textContent =
                    "⋮";


                menuButton.title =
                    "خيارات المستند";


                const options =
                    document.createElement(
                        "div"
                    );


                options.className =
                    "project-document-options";


                options.innerHTML =
                    `
                    <div class="rename-project-document">
                        ✏ إعادة تسمية
                    </div>

                    <div class="move-project-document-up">
                        ↑ نقل إلى أعلى
                    </div>

                    <div class="move-project-document-down">
                        ↓ نقل إلى أسفل
                    </div>

                    <div class="delete-project-document">
                        🗑 حذف
                    </div>
                    `;


                // =========================================
                // إخفاء النقل لأعلى
                // =========================================

                if (
                    index ===
                    0
                ) {

                    const moveUp =
                        options.querySelector(
                            ".move-project-document-up"
                        );


                    if (
                        moveUp
                    ) {

                        moveUp.style.display =
                            "none";

                    }

                }


                // =========================================
                // إخفاء النقل لأسفل
                // =========================================

                if (
                    index ===
                    projectDocuments.length -
                        1
                ) {

                    const moveDown =
                        options.querySelector(
                            ".move-project-document-down"
                        );


                    if (
                        moveDown
                    ) {

                        moveDown.style.display =
                            "none";

                    }

                }


                // =========================================
                // فتح خيارات المستند
                // =========================================

                menuButton.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        document
                            .querySelectorAll(
                                ".project-document-options.open"
                            )
                            .forEach(
                                function (
                                    menu
                                ) {

                                    if (
                                        menu !==
                                        options
                                    ) {

                                        menu.classList.remove(
                                            "open"
                                        );

                                    }

                                }
                            );


                        options.classList.toggle(
                            "open"
                        );

                    };


                // =========================================
                // إعادة التسمية
                // =========================================

                const renameDocument =
                    options.querySelector(
                        ".rename-project-document"
                    );


                if (
                    renameDocument
                ) {

                    renameDocument.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();

                            e.stopPropagation();


                            options.classList.remove(
                                "open"
                            );


                            const oldName =
                                documentItem.name;


                            const edit =
                                document.createElement(
                                    "input"
                                );


                            edit.className =
                                "edit-project-document-title";


                            edit.value =
                                oldName;


                            title.replaceWith(
                                edit
                            );


                            edit.focus();


                            edit.setSelectionRange(
                                edit.value.length,
                                edit.value.length
                            );


                            edit.addEventListener(
                                "keydown",
                                function (
                                    event
                                ) {

                                    if (
                                        event.key ===
                                        "Enter"
                                    ) {

                                        event.preventDefault();


                                        const value =
                                            edit.value.trim();


                                        documentItem.name =
                                            value ||
                                            oldName;


                                        documentItem.updatedAt =
                                            new Date()
                                                .toISOString();


                                        saveDocuments();



                                        renderProjectDocuments(
                                            container,
                                            project
                                        );

                                    }
                                    else if (
                                        event.key ===
                                        "Escape"
                                    ) {

                                        event.preventDefault();


                                        renderProjectDocuments(
                                            container,
                                            project
                                        );

                                    }

                                }
                            );

                        };

                }


                // =========================================
                // نقل لأعلى
                // =========================================

                const moveUp =
                    options.querySelector(
                        ".move-project-document-up"
                    );


                if (
                    moveUp
                ) {

                    moveUp.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();

                            e.stopPropagation();


                            if (
                                index <=
                                0
                            ) {

                                return;

                            }


                            const previous =
                                projectDocuments[
                                    index - 1
                                ];


                            [
                                documentItem.order,
                                previous.order
                            ] =
                            [
                                previous.order,
                                documentItem.order
                            ];


                            saveDocuments();


                            renderProjectDocuments(
                                container,
                                project
                            );

                        };

                }


                // =========================================
                // نقل لأسفل
                // =========================================

                const moveDown =
                    options.querySelector(
                        ".move-project-document-down"
                    );


                if (
                    moveDown
                ) {

                    moveDown.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();

                            e.stopPropagation();


                            if (
                                index >=
                                projectDocuments.length -
                                    1
                            ) {

                                return;

                            }


                            const next =
                                projectDocuments[
                                    index + 1
                                ];


                            [
                                documentItem.order,
                                next.order
                            ] =
                            [
                                next.order,
                                documentItem.order
                            ];


                            saveDocuments();


                            renderProjectDocuments(
                                container,
                                project
                            );

                        };

                }


                // =========================================
                // حذف المستند
                // =========================================

                const deleteDocument =
                    options.querySelector(
                        ".delete-project-document"
                    );


                if (
                    deleteDocument
                ) {

                    deleteDocument.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();

                            e.stopPropagation();


                            options.classList.remove(
                                "open"
                            );


                            const oldConfirm =
                                document.querySelector(
                                    ".document-delete-confirm"
                                );


                            if (
                                oldConfirm
                            ) {

                                oldConfirm.remove();

                            }


                            const confirmBox =
                                document.createElement(
                                    "div"
                                );


                            confirmBox.className =
                                "document-delete-confirm";


                            confirmBox.innerHTML =
                                `
                                <div class="document-delete-dialog">

                                    <div class="document-delete-message">
                                        هل تريد حذف المستند؟
                                    </div>

                                    <div class="document-delete-name">
                                        ${documentItem.name}
                                    </div>

                                    <div class="document-delete-buttons">

                                        <button
                                            type="button"
                                            class="confirm-project-document-delete">
                                            حذف
                                        </button>

                                        <button
                                            type="button"
                                            class="cancel-project-document-delete">
                                            إلغاء
                                        </button>

                                    </div>

                                </div>
                                `;


                            document.body.appendChild(
                                confirmBox
                            );


                            const confirmDelete =
                                confirmBox.querySelector(
                                    ".confirm-project-document-delete"
                                );


                            if (
                                confirmDelete
                            ) {

                                confirmDelete.onclick =
                                    async function () {

                                        documents =
                                            documents.filter(
                                                function (
                                                    doc
                                                ) {

                                                    return (
                                                        String(
                                                            doc.id
                                                        ) !==
                                                        String(
                                                            documentItem.id
                                                        )
                                                    );

                                                }
                                            );


                                        if (
                                            Array.isArray(
                                                project.documents
                                            )
                                        ) {

                                            project.documents =
                                                project.documents.filter(
                                                    function (
                                                        id
                                                    ) {

                                                        return (
                                                            String(
                                                                id
                                                            ) !==
                                                            String(
                                                                documentItem.id
                                                            )
                                                        );

                                                    }
                                                );

                                        }


                                        project.updatedAt =
                                            new Date()
                                                .toISOString();


                                        try {

                                            await deleteWorkingWordFile(
                                                documentItem.storageId
                                            );

                                        }
                                        catch (
                                            error
                                        ) {

                                            console.warn(
                                                "تعذر حذف نسخة العمل:",
                                                error
                                            );

                                        }


                                        if (
                                            currentDocument &&
                                            String(
                                                currentDocument.id
                                            ) ===
                                            String(
                                                documentItem.id
                                            )
                                        ) {

                                            currentDocument =
                                                null;


                                            currentCitationSources =
                                                [];


                                            if (
                                                documentTitle
                                            ) {

                                                documentTitle.textContent =
                                                    "لا يوجد مستند مفتوح";

                                            }

                                        }


                                        oramaRetrievalDb =
                                            null;


                                        oramaRetrievalCacheKey =
                                            "";


                                        oramaRetrievalDocumentId =
                                            null;


                                        const remaining =
                                            getProjectDocuments(
                                                project.id
                                            );


                                        remaining.forEach(
                                            function (
                                                doc,
                                                newIndex
                                            ) {

                                                doc.order =
                                                    newIndex +
                                                    1;

                                            }
                                        );


                                        saveDocuments();

                                        saveProjects();


                                        confirmBox.remove();


                                        renderProjectDocuments(
                                            container,
                                            project
                                        );

                                    };

                            }


                            const cancelDelete =
                                confirmBox.querySelector(
                                    ".cancel-project-document-delete"
                                );


                            if (
                                cancelDelete
                            ) {

                                cancelDelete.onclick =
                                    function () {

                                        confirmBox.remove();

                                    };

                            }


                            confirmBox.onclick =
                                function (
                                    event
                                ) {

                                    if (
                                        event.target ===
                                        confirmBox
                                    ) {

                                        confirmBox.remove();

                                    }

                                };

                        };

                }


                // =========================================
                // بناء المستند
                // =========================================

                item.appendChild(
                    title
                );


                item.appendChild(
                    status
                );


                item.appendChild(
                    menuButton
                );


                item.appendChild(
                    options
                );


                container.appendChild(
                    item
                );

            }
        );

    }


    // =====================================================
    // محادثات المشروع
    // =====================================================

    function renderProjectChats(
        container,
        project
    ) {

        if (
            !container ||
            !project
        ) {

            return;

        }


        container.innerHTML =
            "";


        const projectChatIds =
            Array.isArray(
                project.chatIds
            )
                ? project.chatIds
                : [];


        const projectChats =
            chats.filter(
                function (
                    chat
                ) {

                    if (
                        !chat
                    ) {

                        return false;

                    }


                    if (
                        chat.projectId !==
                            undefined &&
                        chat.projectId !==
                            null
                    ) {

                        return (
                            String(
                                chat.projectId
                            ) ===
                            String(
                                project.id
                            )
                        );

                    }


                    return projectChatIds.some(
                        function (
                            id
                        ) {

                            return (
                                String(
                                    id
                                ) ===
                                String(
                                    chat.id
                                )
                            );

                        }
                    );

                }
            );


        if (
            projectChats.length ===
            0
        ) {

            container.innerHTML =
                `
                <div class="empty-chat">
                    لا توجد محادثات لهذا المشروع
                </div>
                `;

            return;

        }


        projectChats.forEach(
            function (
                chat
            ) {

                const item =
                    document.createElement(
                        "div"
                    );


                item.className =
                    "project-chat-item";


                item.innerHTML =
                    `
                    <span class="project-chat-title">
                        ${chatIcon}
                        ${chat.title || "محادثة جديدة"}
                    </span>
                    `;


                item.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        currentProject =
                            project;


                        currentChat =
                            chat;


                        renderChat();


                        if (
                            projectsPopup
                        ) {

                            projectsPopup.classList.remove(
                                "open"
                            );

                        }

                    };


                container.appendChild(
                    item
                );

            }
        );

    }


    // =====================================================
    // بناء المشاريع
    // =====================================================

    projects.forEach(
        function (
            project
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "project-item";


            item.innerHTML =
                `
                <div class="project-header">

                    <button
                        class="project-expand-toggle"
                        type="button"
                        aria-label="فتح المشروع"
                        aria-expanded="false">
                        ◂
                    </button>


                    <span class="project-title">
                        ${projectIcon}
                        ${project.name}
                    </span>


                    <button
                        class="project-menu"
                        type="button">
                        ⋮
                    </button>

                </div>


                <div
                    class="project-resources"
                    style="display:none;">

                    <div class="project-resource-section">

                        <div
                            class="project-resource-header">

                            <span>
                                ▤ المستندات
                            </span>

                            <span class="project-resource-arrow">
                                ›
                            </span>

                        </div>


                        <div
                            class="project-resource-content">
                        </div>

                    </div>


                    <div class="project-resource-section">

                        <div
                            class="project-resource-header">

                            <span>
                                ◯ المحادثات
                            </span>

                            <span class="project-resource-arrow">
                                ›
                            </span>

                        </div>


                        <div
                            class="project-resource-content">
                        </div>

                    </div>


                    <div class="project-resource-section">

                        <div
                            class="project-resource-header">

                            <span>
                                ≡ المراجع
                            </span>

                            <span class="project-resource-arrow">
                                ›
                            </span>

                        </div>


                        <div
                            class="project-resource-content">

                            <div class="empty-reference">
                                لا توجد مراجع لهذا المشروع حاليًا
                            </div>

                        </div>

                    </div>

                </div>


                <div class="project-options-menu">

                    <div class="rename-project">
                        ✏ إعادة تسمية
                    </div>

                    <div class="delete-project">
                        🗑 حذف
                    </div>

                </div>
                `;


            const projectExpandToggle =
                item.querySelector(
                    ".project-expand-toggle"
                );


            const projectResources =
                item.querySelector(
                    ".project-resources"
                );


            // =================================================
            // فتح / إغلاق المشروع
            // =================================================

            if (
                projectExpandToggle &&
                projectResources
            ) {

                projectExpandToggle.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        const open =
                            projectResources.style.display !==
                            "none";


                        if (
                            open
                        ) {

                            projectResources.style.display =
                                "none";


                            projectExpandToggle.textContent =
                                "◂";


                            projectExpandToggle.setAttribute(
                                "aria-expanded",
                                "false"
                            );

                        }
                        else {

                            projectResources.style.display =
                                "block";


                            projectExpandToggle.textContent =
                                "▾";


                            projectExpandToggle.setAttribute(
                                "aria-expanded",
                                "true"
                            );


                            // =====================================
                            // اختيار المشروع دون إعادة بناء القائمة
                            // =====================================

                            currentProject =
                                project;

                        }

                    };

            }


            // =================================================
            // أقسام المشروع
            // =================================================

            const sections =
                item.querySelectorAll(
                    ".project-resource-section"
                );


            sections.forEach(
                function (
                    section,
                    sectionIndex
                ) {

                    const header =
                        section.querySelector(
                            ".project-resource-header"
                        );


                    const content =
                        section.querySelector(
                            ".project-resource-content"
                        );


                    const arrow =
                        section.querySelector(
                            ".project-resource-arrow"
                        );


                    if (
                        !header ||
                        !content
                    ) {

                        return;

                    }


                    header.onclick =
                        function (
                            e
                        ) {

                            e.preventDefault();

                            e.stopPropagation();


                            const open =
                                header.classList.contains(
                                    "open"
                                );


                            if (
                                open
                            ) {

                                header.classList.remove(
                                    "open"
                                );


                                content.style.display =
                                    "none";


                                if (
                                    arrow
                                ) {

                                    arrow.textContent =
                                        "›";

                                }


                                return;

                            }


                            // =====================================
                            // إغلاق بقية أقسام المشروع
                            // =====================================

                            sections.forEach(
                                function (
                                    otherSection
                                ) {

                                    const otherHeader =
                                        otherSection.querySelector(
                                            ".project-resource-header"
                                        );


                                    const otherContent =
                                        otherSection.querySelector(
                                            ".project-resource-content"
                                        );


                                    const otherArrow =
                                        otherSection.querySelector(
                                            ".project-resource-arrow"
                                        );


                                    if (
                                        otherHeader
                                    ) {

                                        otherHeader.classList.remove(
                                            "open"
                                        );

                                    }


                                    if (
                                        otherContent
                                    ) {

                                        otherContent.style.display =
                                            "none";

                                    }


                                    if (
                                        otherArrow
                                    ) {

                                        otherArrow.textContent =
                                            "›";

                                    }

                                }
                            );


                            header.classList.add(
                                "open"
                            );


                            content.style.display =
                                "block";


                            if (
                                arrow
                            ) {

                                arrow.textContent =
                                    "⌄";

                            }


                            // =====================================
                            // مستندات
                            // =====================================

                            if (
                                sectionIndex ===
                                0
                            ) {

                                renderProjectDocuments(
                                    content,
                                    project
                                );

                            }


                            // =====================================
                            // محادثات
                            // =====================================

                            else if (
                                sectionIndex ===
                                1
                            ) {

                                renderProjectChats(
                                    content,
                                    project
                                );

                            }


                            // =====================================
                            // مراجع
                            // =====================================

                            else if (
                                sectionIndex ===
                                2
                            ) {

                                content.innerHTML =
                                    `
                                    <div class="empty-reference">
                                        لا توجد مراجع لهذا المشروع حاليًا
                                    </div>
                                    `;

                            }

                        };

                }
            );


            // =================================================
            // النقر على المشروع
            // =================================================

            item.onclick =
                function (
                    e
                ) {

                    if (
                        e.target.closest(
                            ".project-expand-toggle"
                        ) ||
                        e.target.closest(
                            ".project-menu"
                        ) ||
                        e.target.closest(
                            ".project-options-menu"
                        ) ||
                        e.target.closest(
                            ".project-resource-header"
                        ) ||
                        e.target.closest(
                            ".project-resource-content"
                        )
                    ) {

                        return;

                    }


                    e.stopPropagation();


                    currentProject =
                        project;


                    if (
                        projectsPopup
                    ) {

                        projectsPopup.classList.remove(
                            "open"
                        );

                    }

                };


            // =================================================
            // قائمة خيارات المشروع
            // =================================================

            const menu =
                item.querySelector(
                    ".project-menu"
                );


            const options =
                item.querySelector(
                    ".project-options-menu"
                );


            if (
                menu &&
                options
            ) {

                menu.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        document
                            .querySelectorAll(
                                ".project-options-menu.open"
                            )
                            .forEach(
                                function (
                                    openMenu
                                ) {

                                    if (
                                        openMenu !==
                                        options
                                    ) {

                                        openMenu.classList.remove(
                                            "open"
                                        );

                                    }

                                }
                            );


                        options.classList.add(
                            "open"
                        );


                        const rect =
                            menu.getBoundingClientRect();


                        const menuWidth =
                            140;


                        const menuHeight =
                            options.offsetHeight ||
                            80;


                        const margin =
                            8;


                        let left =
                            rect.left -
                            menuWidth -
                            margin;


                        let top =
                            rect.bottom +
                            margin;


                        if (
                            left <
                            margin
                        ) {

                            left =
                                rect.right +
                                margin;

                        }


                        if (
                            left +
                                menuWidth >
                            window.innerWidth -
                                margin
                        ) {

                            left =
                                window.innerWidth -
                                menuWidth -
                                margin;

                        }


                        if (
                            top +
                                menuHeight >
                            window.innerHeight -
                                margin
                        ) {

                            top =
                                rect.top -
                                menuHeight -
                                margin;

                        }


                        if (
                            top <
                            margin
                        ) {

                            top =
                                margin;

                        }


                        options.style.position =
                            "fixed";


                        options.style.left =
                            left +
                            "px";


                        options.style.top =
                            top +
                            "px";


                        options.style.right =
                            "auto";


                        options.style.bottom =
                            "auto";


                        options.style.zIndex =
                            "999999";

                    };

            }


            // =================================================
            // إعادة التسمية
            // =================================================

            const renameProject =
                options
                    ? options.querySelector(
                        ".rename-project"
                    )
                    : null;


            if (
                renameProject
            ) {

                renameProject.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const title =
                            item.querySelector(
                                ".project-title"
                            );


                        if (
                            !title
                        ) {

                            return;

                        }


                        const oldName =
                            project.name;


                        title.innerHTML =
                            `
                            <input
                                class="edit-project-title"
                                value="${oldName}">
                            `;


                        const edit =
                            title.querySelector(
                                ".edit-project-title"
                            );


                        if (
                            !edit
                        ) {

                            return;

                        }


                        edit.focus();


                        edit.setSelectionRange(
                            edit.value.length,
                            edit.value.length
                        );


                        edit.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();


                                    project.name =
                                        edit.value.trim() ||
                                        oldName;


                                    project.updatedAt =
                                        new Date()
                                            .toISOString();


                                    saveProjects();


                                    renderProjects();


                                }
                                else if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();


                                    renderProjects();

                                }

                            };

                    };

            }


            // =================================================
            // حذف المشروع
            // =================================================

            const deleteProject =
                options
                    ? options.querySelector(
                        ".delete-project"
                    )
                    : null;


            if (
                deleteProject
            ) {

                deleteProject.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "project-delete-confirm";


                        confirmBox.innerHTML =
                            `
                            <div class="confirm-dialog">

                                <p>
                                    هل تريد حذف المشروع:
                                    <br>

                                    <strong>
                                        ${project.name}
                                    </strong>

                                    ؟
                                </p>

                                <button
                                    class="confirm-project-delete"
                                    type="button">
                                    حذف
                                </button>

                                <button
                                    class="cancel-project-delete"
                                    type="button">
                                    إلغاء
                                </button>

                            </div>
                            `;


                        document.body.appendChild(
                            confirmBox
                        );


                        const confirmDelete =
                            confirmBox.querySelector(
                                ".confirm-project-delete"
                            );


                        if (
                            confirmDelete
                        ) {

                            confirmDelete.onclick =
                                async function () {

                                    const projectDocumentIds =
                                        Array.isArray(
                                            project.documents
                                        )
                                            ? project.documents
                                            : [];


                                    // =====================================
                                    // حذف ملفات المستندات
                                    // =====================================

                                    for (
                                        let i =
                                            0;

                                        i <
                                            projectDocumentIds.length;

                                        i++
                                    ) {

                                        const doc =
                                            documents.find(
                                                function (
                                                    d
                                                ) {

                                                    return (
                                                        d &&
                                                        String(
                                                            d.id
                                                        ) ===
                                                        String(
                                                            projectDocumentIds[
                                                                i
                                                            ]
                                                        )
                                                    );

                                                }
                                            );


                                        if (
                                            doc
                                        ) {

                                            try {

                                                await deleteWorkingWordFile(
                                                    doc.storageId
                                                );

                                            }
                                            catch (
                                                error
                                            ) {

                                                console.warn(
                                                    "تعذر حذف نسخة العمل:",
                                                    error
                                                );

                                            }

                                        }

                                    }


                                    // =====================================
                                    // حذف مستندات المشروع
                                    // =====================================

                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return !projectDocumentIds.some(
                                                    function (
                                                        id
                                                    ) {

                                                        return (
                                                            String(
                                                                id
                                                            ) ===
                                                            String(
                                                                doc.id
                                                            )
                                                        );

                                                    }
                                                );

                                            }
                                        );


                                    // =====================================
                                    // حذف محادثات المشروع
                                    // =====================================

                                    chats =
                                        chats.filter(
                                            function (
                                                chat
                                            ) {

                                                return !(
                                                    chat &&
                                                    String(
                                                        chat.projectId
                                                    ) ===
                                                    String(
                                                        project.id
                                                    )
                                                );

                                            }
                                        );


                                    // =====================================
                                    // حذف المشروع
                                    // =====================================

                                    projects =
                                        projects.filter(
                                            function (
                                                p
                                            ) {

                                                return (
                                                    String(
                                                        p.id
                                                    ) !==
                                                    String(
                                                        project.id
                                                    )
                                                );

                                            }
                                        );


                                    // =====================================
                                    // تصفير الحالة الحالية
                                    // =====================================

                                    if (
                                        currentProject &&
                                        String(
                                            currentProject.id
                                        ) ===
                                        String(
                                            project.id
                                        )
                                    ) {

                                        currentProject =
                                            null;


                                        currentDocument =
                                            null;


                                        currentChat =
                                            null;


                                        currentCitationSources =
                                            [];


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


                                    // =====================================
                                    // إعادة ضبط Orama
                                    // =====================================

                                    oramaRetrievalDb =
                                        null;


                                    oramaRetrievalCacheKey =
                                        "";


                                    oramaRetrievalDocumentId =
                                        null;


                                    // =====================================
                                    // حفظ
                                    // =====================================

                                    saveDocuments();

                                    saveChats();

                                    saveProjects();


                                    // =====================================
                                    // إعادة الرسم
                                    // =====================================

                                    renderProjects();

                                    

                                    

                                    renderSidebarChats();

                                    renderRecentChats();


                                    confirmBox.remove();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-project-delete"
                            );


                        if (
                            cancelDelete
                        ) {

                            cancelDelete.onclick =
                                function () {

                                    confirmBox.remove();

                                };

                        }

                    };

            }


            projectsList.appendChild(
                item
            );

        }
    );

}

// =====================================================
// Projects Button
// =====================================================

if (projectsBtn) {

    projectsBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (!projectsPopup)
                return;


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            projectsPopup.classList.toggle(
                "open"
            );


            renderProjects();

        };

}


// =====================================================
// New Project Button
// =====================================================

if (newProjectBtn) {

    newProjectBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (!projectsPopup)
                return;


            const oldBox =
                document.querySelector(
                    ".project-create-box"
                );


            if (oldBox) {

                oldBox.remove();

            }


            const box =
                document.createElement(
                    "div"
                );


            box.className =
                "project-create-box";


            box.innerHTML = `

                <div class="rename-dialog">

                    <input
                        class="new-project-name"
                        placeholder="اسم المشروع">

                    <button
                        class="save-project"
                        type="button">
                        حفظ
                    </button>

                    <button
                        class="cancel-project"
                        type="button">
                        إلغاء
                    </button>

                </div>

            `;


            document.body.appendChild(
                box
            );


            const buttonRect =
                newProjectBtn.getBoundingClientRect();


            const boxHeight =
                120;


            const screenMargin =
                12;


            let left =
                buttonRect.left;


            let top =
                buttonRect.bottom +
                8;


            const actualBoxWidth =
                box.offsetWidth ||
                240;


            if (
                left +
                    actualBoxWidth >
                window.innerWidth -
                    screenMargin
            ) {

                left =
                    window.innerWidth -
                    actualBoxWidth -
                    screenMargin;

            }


            if (
                left <
                screenMargin
            ) {

                left =
                    screenMargin;

            }


            if (
                top +
                    boxHeight >
                window.innerHeight -
                    screenMargin
            ) {

                top =
                    buttonRect.top -
                    boxHeight -
                    8;

            }


            box.style.position =
                "fixed";


            box.style.left =
                left +
                "px";


            box.style.top =
                top +
                "px";


            box.style.zIndex =
                "999999";


            const inputProject =
                box.querySelector(
                    ".new-project-name"
                );


            if (inputProject) {

                inputProject.focus();

            }


            const saveProject =
                box.querySelector(
                    ".save-project"
                );


            if (saveProject) {

                saveProject.onclick =
                    function () {

                        const name =
                            inputProject
                                ? inputProject.value.trim()
                                : "";


                        if (
                            name !==
                            ""
                        ) {

                            createProject(
                                name
                            );

                        }


                        box.remove();

                    };

            }


            const cancelProject =
                box.querySelector(
                    ".cancel-project"
                );


            if (cancelProject) {

                cancelProject.onclick =
                    function () {

                        box.remove();

                    };

            }


            box.onclick =
                function (event) {

                    event.stopPropagation();

                };

        };

}





// =====================================================
// Render Sidebar Chats
// =====================================================

function renderSidebarChats() {

    const list =
        document.getElementById(
            "new-chat-list"
        );


    if (!list)
        return;


    list.innerHTML =
        "";


    if (
        chats.length ===
        0
    ) {

        list.innerHTML = `
            <div class="empty-chat">
                لا توجد محادثات
            </div>
        `;

        return;

    }


    chats
        .slice(
            0,
            8
        )
        .forEach(
            function (
                chat
            ) {

                const item =
                    document.createElement(
                        "div"
                    );


                item.className =
                    "recent-chat-item";


                item.innerHTML = `
                    <span class="chat-title">
                        ${chatIcon}
                        ${chat.title}
                    </span>
                `;


                item.onclick =
                    function (e) {

                        e.stopPropagation();


                        currentChat =
                            chat;


                        currentCitationSources =
                            [];


                        renderChat();


                        if (projectsPopup) {

                            projectsPopup.classList.remove(
                                "open"
                            );

                        }


                        if (chatPopup) {

                            chatPopup.classList.remove(
                                "open"
                            );

                        }


                        if (searchPopup) {

                            searchPopup.classList.remove(
                                "open"
                            );

                        }

                    };


                list.appendChild(
                    item
                );

            }
        );

}




// =====================================================
// Recent Chats Popup
// =====================================================

function renderRecentChats() {

    if (!recentChatList)
        return;


    recentChatList.innerHTML =
        "";


    if (
        chats.length ===
        0
    ) {

        recentChatList.innerHTML =
            "<div class='empty-chat'>لا توجد محادثات</div>";

        return;

    }


    chats
        .slice(
            0,
            8
        )
        .forEach(
            function (
                chat
            ) {

                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "recent-chat-item";


                div.innerHTML = `
                    <span class="chat-title">
                        ${chatIcon}
                        ${chat.title}
                    </span>
                `;


                div.onclick =
                    function () {

                        currentChat =
                            chat;


                        renderChat();


                        if (chatPopup) {

                            chatPopup.classList.remove(
                                "open"
                            );

                        }

                    };


                recentChatList.appendChild(
                    div
                );

            }
        );

}


// =====================================================
// Chat Button
// =====================================================

if (chatBtn) {

    chatBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (!chatPopup)
                return;


            if (projectsPopup) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            if (
                chatPopup.classList.contains(
                    "open"
                )
            ) {

                chatPopup.classList.remove(
                    "open"
                );


                return;

            }


            


            chatPopup.classList.add(
                "open"
            );

        };

}


// =====================================================
// New Chat Button
// =====================================================

if (newChatBtn) {

    newChatBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (projectsPopup) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            createNewChat();

        };

}


// =====================================================
// Search Button
// البحث الحالي في أسماء المحادثات
// =====================================================

if (searchBtn) {

    searchBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (!searchPopup)
                return;


            if (projectsPopup) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            searchPopup.classList.toggle(
                "open"
            );


            if (
                searchPopup.classList.contains(
                    "open"
                ) &&
                searchInput
            ) {

                searchInput.focus();

            }

        };

}


// =====================================================
// Search Input
// البحث في عناوين المحادثات
// =====================================================

if (searchInput) {

    searchInput.oninput =
        function () {

            const keyword =
                searchInput.value
                    .trim()
                    .toLowerCase();


            if (!searchResults)
                return;


            searchResults.innerHTML =
                "";


            if (
                keyword ===
                ""
            ) {

                return;

            }


            chats.forEach(
                function (
                    chat
                ) {

                    const title =
                        String(
                            chat.title ||
                            ""
                        );


                    if (
                        title
                            .toLowerCase()
                            .includes(
                                keyword
                            )
                    ) {

                        const item =
                            document.createElement(
                                "div"
                            );


                        item.className =
                            "search-result-item";


                        item.innerHTML = `
                            <span class="chat-title">
                                ${chatIcon}
                                ${title}
                            </span>
                        `;


                        item.onclick =
                            function () {

                                currentChat =
                                    chat;


                                renderChat();


                                if (
                                    searchPopup
                                ) {

                                    searchPopup.classList.remove(
                                        "open"
                                    );

                                }


                                searchInput.value =
                                    "";


                                searchResults.innerHTML =
                                    "";

                            };


                        searchResults.appendChild(
                            item
                        );

                    }

                }
            );

        };

}
// ======================================
// Word AI Assistant
// PART 10 / 10
// Initialization + Diagnostics + Sidebar Pin
// ======================================


// =====================================================
// Test Current Document Index
// اختبار جاهزية مستند Orama الحالي
// =====================================================

window.testCurrentDocumentIndex =
    async function () {

    if (!currentDocument) {

        console.warn(
            "لا يوجد مستند نشط."
        );

        return null;

    }


    try {

        const structureData =
            await ensureDocumentStructure(
                currentDocument
            );


        if (!structureData) {

            console.warn(
                "لا توجد بنية محفوظة للمستند."
            );

            return null;

        }


        const db =
            await ensureOramaDocumentReady(
                currentDocument
            );


        if (!db) {

            console.warn(
                "تعذر تجهيز فهرس Orama."
            );

            return null;

        }


        const result =
            await searchIndexedDocument(
                currentDocument.id,
                "استصلاح",
                {
                    maxResults:
                        10
                }
            );


        console.log(
            "======================================"
        );

        console.log(
            "اختبار Orama للمستند الحالي"
        );

        console.log(
            "المستند:",
            currentDocument.name
        );

        console.log(
            "عدد الفقرات:",
            structureData.paragraphs
                ? structureData.paragraphs.length
                : 0
        );

        console.log(
            "عدد العناوين:",
            structureData.headings
                ? structureData.headings.length
                : 0
        );

        console.log(
            "عدد نتائج استصلاح:",
            result.count
        );

        console.log(
            "العائلات المطابقة:",
            result.matchedFamilies
        );

        console.log(
            "الكلمات المطابقة:",
            result.matchedTerms
        );

        console.log(
            "أول النتائج:",
            (
                result.results ||
                []
            ).slice(
                0,
                10
            )
        );

        console.log(
            "======================================"
        );


        return result;

    }
    catch (error) {

        console.error(
            "فشل اختبار المستند:",
            error
        );


        return null;

    }

}


// =====================================================
// Full Orama Diagnostic
// اختبار البحث في صيغ عائلة استصلاح
// =====================================================

window.testIstislahIndex =
    async function () {

        if (!currentDocument) {

            console.warn(
                "لا يوجد مستند نشط."
            );

            return null;

        }


        const testQueries = [

            "استصلاح",

            "الاستصلاح",

            "بالاستصلاح",

            "استصلاحيا"

        ];


        const results =
            [];


        console.log(
            "======================================"
        );

        console.log(
            "اختبار عائلة استصلاح مع Orama"
        );

        console.log(
            "المستند:",
            currentDocument.name
        );

        console.log(
            "======================================"
        );


        for (
            let i = 0;
            i < testQueries.length;
            i++
        ) {

            const query =
                testQueries[i];


            try {

                const result =
                    await searchIndexedDocument(
                        currentDocument.id,
                        query,
                        {
                            maxResults:
                                10
                        }
                    );


                const item = {

                    query:
                        query,

                    count:
                        result.count,

                    matchedTerms:
                        result.matchedTerms,

                    matchedFamilies:
                        result.matchedFamilies,

                    results:
                        (
                            result.results ||
                            []
                        ).slice(
                            0,
                            3
                        )

                };


                results.push(
                    item
                );


                console.log(
                    "الاستعلام:",
                    query
                );

                console.log(
                    "عدد النتائج:",
                    result.count
                );

                console.log(
                    "العائلات:",
                    result.matchedFamilies
                );

                console.log(
                    "أول النتائج:",
                    result.results.slice(
                        0,
                        3
                    )
                );

                console.log(
                    "--------------------------------------"
                );

            }
            catch (error) {

                console.error(
                    "فشل الاستعلام:",
                    query,
                    error
                );

            }

        }


        return results;

    };


// =====================================================
// Initialization
// التهيئة النهائية للتطبيق
// =====================================================


renderProjects();

renderSidebarChats();

renderRecentChats();

renderChat();

loadSettings();


// =====================================================
// Initial State
// =====================================================

if (currentProject) {

    renderProjects();

}


if (currentChat) {

    renderChat();

}


// =====================================================
// End Office.onReady
// الإغلاق الصحيح هنا
// =====================================================

});


// =====================================================
// Sidebar Pin
// خارج Office.onReady
// =====================================================

const sidebar =
    document.querySelector(
        ".sidebar"
    );


const pinSidebar =
    document.getElementById(
        "pin-sidebar"
    );


if (
    pinSidebar &&
    sidebar
) {

    let sidebarPinned =
        localStorage.getItem(
            "sidebarPinned"
        ) ===
        "true";


    // ==================================
    // استعادة حالة التثبيت
    // ==================================

    if (sidebarPinned) {

        sidebar.classList.add(
            "pinned"
        );


        pinSidebar.classList.add(
            "pinned"
        );


        document.body.classList.add(
            "sidebar-is-pinned"
        );

    }


    // ==================================
    // زر التثبيت
    // ==================================

    pinSidebar.addEventListener(
        "click",
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            sidebarPinned =
                !sidebarPinned;


            sidebar.classList.toggle(
                "pinned",
                sidebarPinned
            );


            pinSidebar.classList.toggle(
                "pinned",
                sidebarPinned
            );


            document.body.classList.toggle(
                "sidebar-is-pinned",
                sidebarPinned
            );


            localStorage.setItem(
                "sidebarPinned",
                sidebarPinned
                    ? "true"
                    : "false"
            );

        }
    );

}


// =====================================================
// Global Error Diagnostics
// تشخيص الأخطاء غير المعالجة
// =====================================================

window.addEventListener(
    "unhandledrejection",
    function (
        event
    ) {

        console.error(
            "Unhandled Promise Rejection:",
            event.reason
        );

    }
);


window.addEventListener(
    "error",
    function (
        event
    ) {

        console.error(
            "Global JavaScript Error:",
            event.error ||
            event.message
        );

    }
);


// =====================================================
// Final App Marker
// =====================================================

console.log(
    "✓ تم تحميل taskpane.js بالكامل."
);