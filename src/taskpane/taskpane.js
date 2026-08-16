// ======================================
// Word AI Assistant
// Main Application Controller
// PART 1 / 3
// التخزين + المستندات + الفهرسة + Orama
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

const documentsList =
    document.getElementById("documents-list");

const addDocumentBtn =
    document.getElementById("add-document-btn");

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

const expandedSidebar =
    document.getElementById("expanded-sidebar");

const sidebarToggleBtn =
    document.getElementById("sidebar-toggle-btn");

const expandedSidebarToggleSlot =
    document.getElementById(
        "expanded-sidebar-toggle-slot"
    );

const sidebarTogglePlaceholder =
    document.createComment(
        "sidebar-toggle-placeholder"
    );

if (
    sidebarToggleBtn &&
    sidebarToggleBtn.parentNode
) {

    sidebarToggleBtn.parentNode.insertBefore(
        sidebarTogglePlaceholder,
        sidebarToggleBtn
    );

}

const input =
    document.getElementById("user-input");

const sendBtn =
    document.getElementById("send-btn");

const chatArea =
    document.getElementById("chat-area");

const documentTitle =
    document.getElementById("document-title");


// ======================================
// Citation Click Handler
// التعامل مع النقر على إحالات المقاطع
// ======================================

if (chatArea) {

    chatArea.addEventListener(
        "click",
        function (event) {

            const citation =
                event.target.closest(
                    ".document-citation"
                );

            if (!citation) {

                return;

            }

            event.preventDefault();
            event.stopPropagation();

            const rank =
                Number(
                    citation.getAttribute(
                        "data-citation-rank"
                    )
                );

            if (
                !Number.isNaN(
                    rank
                )
            ) {

                openCitationInWord(
                    rank
                );

            }

        }
    );

}


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
catch (error) {

    projects = [];

}


projects =
    projects
        .filter(
            function (
                project
            ) {

                return (
                    project &&
                    typeof project ===
                        "object"
                );

            }
        )
        .map(
            function (
                project
            ) {

                const now =
                    new Date()
                        .toISOString();

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
                        typeof project.settings ===
                            "object"
                            ? project.settings
                            : {
                                citationStyle: "",
                                notes: ""
                            }

                };

            }
        );


let currentProject =
    null;

let currentDocument =
    null;


// ======================================
// Citation Sources
// مصادر الإحالات الحالية
// ======================================

let currentCitationSources =
    [];


// ======================================
// Orama Retrieval Cache
// ذاكرة محرك البحث
// ======================================

let oramaRetrievalDb =
    null;

let oramaRetrievalCacheKey =
    "";

let oramaRetrievalRecords =
    [];


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
            .filter(
                function (
                    documentItem
                ) {

                    return (
                        documentItem &&
                        typeof documentItem ===
                            "object"
                    );

                }
            )
            .map(
                function (
                    documentItem
                ) {

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

                }
            );

}
catch (error) {

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


saveDocuments();


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
// Version of indexing algorithm
// مهم: تغييره يجبر الفهرس القديم
// على إعادة البناء
// ======================================

const INDEX_SCHEMA_VERSION =
    5;


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
                            reader.result ||
                            ""
                        );


                    const commaIndex =
                        result.indexOf(
                            ","
                        );


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
                            commaIndex +
                            1
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
        new Date()
            .toISOString();


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
            typeof order ===
                "number"
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

    if (
        !documentItem
    ) {

        return;

    }


    documentItem.readStatus =
        status;


    documentItem.updatedAt =
        new Date()
            .toISOString();


    if (
        status ===
        "read"
    ) {

        documentItem.readAt =
            new Date()
                .toISOString();

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

    if (
        !documentItem
    ) {

        return;

    }


    documentItem.indexStatus =
        status;


    documentItem.updatedAt =
        new Date()
            .toISOString();


    saveDocuments();

}


// ======================================
// Get Project Documents
// ======================================

function getProjectDocuments(
    projectId
) {

    if (
        !projectId
    ) {

        return [];

    }


    return documents
        .filter(
            function (
                documentItem
            ) {

                return (
                    documentItem &&
                    documentItem.projectId ===
                        projectId
                );

            }
        )
        .sort(
            function (
                a,
                b
            ) {

                return (
                    a.order -
                    b.order
                );

            }
        );

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
                        text ||
                        ""
                    ),

                updatedAt:
                    new Date()
                        .toISOString()

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


// ======================================
// Normalize Search Text
// ======================================

function normalizeSearchText(
    text
) {

    return String(
        text ||
        ""
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

        // توحيد الهمزة
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
// Conservative Family Key
// مفتاح عائلة الكلمة
// ======================================

function getConservativeFamilyKey(
    word,
    surfaceSet
) {

    let w =
        normalizeSearchText(
            word
        );


    if (
        !w
    ) {

        return "";

    }


    w =
        w.replace(
            /[^\p{L}\p{N}]+/gu,
            ""
        );


    if (
        !w
    ) {

        return "";

    }


    if (
        w.length <=
        3
    ) {

        return w;

    }


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
// Legacy Document Index
// يبقى للتوافق والإحصاءات فقط
// وليس محرك البحث
// =====================================================

function buildDocumentIndex(
    documentId,
    text
) {

    if (
        !documentId
    ) {

        throw new Error(
            "معرّف المستند غير موجود."
        );

    }


    if (
        typeof text !==
            "string"
    ) {

        throw new Error(
            "نص المستند غير صالح للفهرسة."
        );

    }


    const tokens =
        tokenizeDocumentText(
            text
        );


    const terms =
        {};

    const families =
        {};


    const surfaceSet =
        new Set(
            tokens
        );


    const paragraphTexts =
        String(
            text ||
            ""
        )
            .split(
                /\r\n|\r|\n/
            );


    let globalTokenPosition =
        0;


    paragraphTexts.forEach(
        function (
            paragraphText,
            paragraphIndex
        ) {

            const normalizedParagraphText =
                normalizeSearchText(
                    paragraphText
                );


            const tokenMatches =
                Array.from(
                    normalizedParagraphText.matchAll(
                        /[\p{L}\p{N}]+/gu
                    )
                );


            tokenMatches.forEach(
                function (
                    match,
                    tokenIndex
                ) {

                    const surface =
                        match[0];


                    const charStart =
                        typeof match.index ===
                            "number"
                            ? match.index
                            : -1;


                    const charEnd =
                        charStart ===
                            -1
                            ? -1
                            : charStart +
                              surface.length;


                    if (
                        !surface
                    ) {

                        return;

                    }


                    const familyKey =
                        getConservativeFamilyKey(
                            surface,
                            surfaceSet
                        );


                    if (
                        !terms[surface]
                    ) {

                        terms[surface] = {

                            count:
                                0,

                            positions:
                                [],

                            occurrences:
                                [],

                            family:
                                familyKey ||
                                ""

                        };

                    }


                    terms[surface]
                        .count +=
                        1;


                    terms[surface]
                        .positions
                        .push(
                            globalTokenPosition
                        );


                    terms[surface]
                        .occurrences
                        .push({

                            paragraphIndex:
                                paragraphIndex,

                            tokenIndex:
                                tokenIndex,

                            globalIndex:
                                globalTokenPosition,

                            charStart:
                                charStart,

                            charEnd:
                                charEnd

                        });


                    if (
                        !familyKey
                    ) {

                        globalTokenPosition +=
                            1;

                        return;

                    }


                    if (
                        !families[
                            familyKey
                        ]
                    ) {

                        families[
                            familyKey
                        ] = {

                            count:
                                0,

                            positions:
                                [],

                            occurrences:
                                [],

                            words:
                                {},

                            uniqueWords:
                                0

                        };

                    }


                    families[
                        familyKey
                    ]
                        .count +=
                        1;


                    families[
                        familyKey
                    ]
                        .positions
                        .push(
                            globalTokenPosition
                        );


                    families[
                        familyKey
                    ]
                        .occurrences
                        .push({

                            paragraphIndex:
                                paragraphIndex,

                            tokenIndex:
                                tokenIndex,

                            globalIndex:
                                globalTokenPosition,

                            charStart:
                                charStart,

                            charEnd:
                                charEnd,

                            word:
                                surface

                        });


                    if (
                        !families[
                            familyKey
                        ].words[
                            surface
                        ]
                    ) {

                        families[
                            familyKey
                        ].words[
                            surface
                        ] =
                            0;


                        families[
                            familyKey
                        ]
                            .uniqueWords +=
                            1;

                    }


                    families[
                        familyKey
                    ]
                        .words[
                            surface
                        ] +=
                            1;


                    globalTokenPosition +=
                        1;

                }
            );

        }
    );


    return {

        documentId:
            String(
                documentId
            ),

        indexVersion:
            INDEX_SCHEMA_VERSION,

        tokenCount:
            tokens.length,

        uniqueTerms:
            Object.keys(
                terms
            ).length,

        uniqueFamilies:
            Object.keys(
                families
            ).length,

        terms:
            terms,

        families:
            families,

        updatedAt:
            new Date()
                .toISOString()

    };

}


// ======================================
// Save Document Index
// ======================================

async function saveDocumentIndex(
    documentId,
    indexData
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
                    DOCUMENT_INDEX_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_INDEX_STORE_NAME
                );


            const request =
                store.put(
                    indexData,
                    String(
                        documentId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        indexData
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
// Get Document Index
// ======================================

async function getDocumentIndex(
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
                    DOCUMENT_INDEX_STORE_NAME,
                    "readonly"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_INDEX_STORE_NAME
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
// Build And Save Document Index
// ======================================

async function rebuildDocumentIndex(
    documentId,
    text
) {

    if (
        !documentId
    ) {

        throw new Error(
            "معرّف المستند غير موجود."
        );

    }


    if (
        typeof text !==
            "string"
    ) {

        throw new Error(
            "نص المستند غير صالح للفهرسة."
        );

    }


    const indexData =
        buildDocumentIndex(
            documentId,
            text
        );


    await saveDocumentIndex(
        documentId,
        indexData
    );


    if (
        currentDocument &&
        String(
            currentDocument.id
        ) ===
        String(
            documentId
        )
    ) {

        currentDocument.indexStatus =
            "indexed";

        currentDocument.indexTokenCount =
            indexData.tokenCount;

        currentDocument.indexUniqueTerms =
            indexData.uniqueTerms;

        currentDocument.indexUniqueFamilies =
            indexData.uniqueFamilies;

        currentDocument.indexSchemaVersion =
            INDEX_SCHEMA_VERSION;

        currentDocument.indexUpdatedAt =
            indexData.updatedAt;

        saveDocuments();

    }


    // إبطال ذاكرة Orama
    oramaRetrievalDb =
        null;

    oramaRetrievalCacheKey =
        "";

    oramaRetrievalRecords =
        [];


    return indexData;

}


// ======================================
// Build Document Structure
// ======================================

async function buildDocumentStructure(
    documentItem
) {

    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    const file =
        await getWorkingWordFile(
            documentItem.storageId
        );


    if (
        !file
    ) {

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
                    new Date()
                        .toISOString()

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


    // إبطال ذاكرة Orama
    oramaRetrievalDb =
        null;

    oramaRetrievalCacheKey =
        "";

    oramaRetrievalRecords =
        [];


    return structure;

}


// ======================================
// Ensure Document Index
// ======================================
// للتوافق والإحصاءات فقط.
// لا يستخدم كمحرك بحث.
//
// ======================================

async function ensureDocumentIndex(
    documentItem
) {

    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    let index =
        await getDocumentIndex(
            documentItem.id
        );


    let validIndex =
        false;


    if (
        index &&
        index.indexVersion ===
            INDEX_SCHEMA_VERSION &&
        index.terms &&
        typeof index.terms ===
            "object" &&
        index.families &&
        typeof index.families ===
            "object"
    ) {

        validIndex =
            true;


        const termKeys =
            Object.keys(
                index.terms
            );


        for (
            let i = 0;
            i < termKeys.length;
            i++
        ) {

            const term =
                index.terms[
                    termKeys[i]
                ];


            if (
                !term ||
                !Array.isArray(
                    term.positions
                ) ||
                !Array.isArray(
                    term.occurrences
                )
            ) {

                validIndex =
                    false;

                break;

            }

        }


        if (
            validIndex
        ) {

            const familyKeys =
                Object.keys(
                    index.families
                );


            for (
                let i = 0;
                i < familyKeys.length;
                i++
            ) {

                const family =
                    index.families[
                        familyKeys[i]
                    ];


                if (
                    !family ||
                    !Array.isArray(
                        family.positions
                    ) ||
                    !Array.isArray(
                        family.occurrences
                    ) ||
                    !family.words ||
                    typeof family.words !==
                        "object"
                ) {

                    validIndex =
                        false;

                    break;

                }

            }

        }

    }


    if (
        validIndex
    ) {

        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            index.tokenCount ||
            0;


        documentItem.indexUniqueTerms =
            index.uniqueTerms ||
            0;


        documentItem.indexUniqueFamilies =
            index.uniqueFamilies ||
            0;


        documentItem.indexSchemaVersion =
            index.indexVersion;


        documentItem.indexUpdatedAt =
            index.updatedAt ||
            "";


        saveDocuments();


        return index;

    }


    const textData =
        await getDocumentText(
            documentItem.id
        );


    if (
        !textData ||
        typeof textData.text !==
            "string"
    ) {

        return null;

    }


    updateDocumentIndexStatus(
        documentItem,
        "indexing"
    );


    try {

        index =
            await rebuildDocumentIndex(
                documentItem.id,
                textData.text
            );


        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            index.tokenCount ||
            0;


        documentItem.indexUniqueTerms =
            index.uniqueTerms ||
            0;


        documentItem.indexUniqueFamilies =
            index.uniqueFamilies ||
            0;


        documentItem.indexSchemaVersion =
            INDEX_SCHEMA_VERSION;


        documentItem.indexUpdatedAt =
            index.updatedAt ||
            "";


        saveDocuments();


        return index;

    }
    catch (error) {

        updateDocumentIndexStatus(
            documentItem,
            "error"
        );


        throw error;

    }

}


// ======================================
// Orama Retrieval Index
// فهرس Orama الرئيسي
// ======================================

async function ensureOramaRetrievalIndex(
    documentItem,
    structureData
) {

    const {
        create,
        insertMultiple
    } =
        require(
            "@orama/orama"
        );


    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


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
                structureData.updatedAt ||
                ""
            ),

            String(
                paragraphs.length
            ),

            String(
                headings.length
            )

        ].join(
            "|"
        );


    if (
        oramaRetrievalDb &&
        oramaRetrievalCacheKey ===
            cacheKey &&
        Array.isArray(
            oramaRetrievalRecords
        ) &&
        oramaRetrievalRecords.length
    ) {

        return oramaRetrievalDb;

    }


    // ==================================
    // كل كلمات المستند
    // ==================================

    const allTokens =
        [];


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                paragraph &&
                paragraph.text
            ) {

                allTokens.push(
                    ...tokenizeDocumentText(
                        paragraph.text
                    )
                );

            }

        }
    );


    const surfaceSet =
        new Set(
            allTokens
        );


    // ==================================
    // مستوى العنوان
    // ==================================

    function getHeadingLevel(
        style
    ) {

        const match =
            String(
                style ||
                ""
            ).match(
                /Heading\s*([1-9])/i
            );


        return (
            match
                ? Number(
                    match[1]
                )
                : 9
        );

    }


    // ==================================
    // أقرب عنوان
    // ==================================

    function getNearestHeading(
        paragraphIndex
    ) {

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
                ) <
                Number(
                    paragraphIndex
                )
            ) {

                return heading;

            }

        }


        return null;

    }


    // ==================================
    // أقرب عنوان رئيسي
    // Heading 1–3
    // ==================================

    function getMajorHeading(
        paragraphIndex
    ) {

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
                ) <
                Number(
                    paragraphIndex
                ) &&
                getHeadingLevel(
                    heading.style
                ) <=
                    3
            ) {

                return heading;

            }

        }


        return null;

    }


    // ==================================
    // نهاية القسم
    // يتوقف عند عنوان من نفس المستوى
    // أو مستوى أعلى
    // ==================================

    function getSectionEnd(
        headingIndex,
        level
    ) {

        let endIndex =
            paragraphs.length -
            1;


        for (
            let i = 0;
            i < headings.length;
            i++
        ) {

            const heading =
                headings[i];


            const index =
                Number(
                    heading.index
                );


            if (
                index <=
                Number(
                    headingIndex
                )
            ) {

                continue;

            }


            if (
                getHeadingLevel(
                    heading.style
                ) <=
                    level
            ) {

                endIndex =
                    index -
                    1;

                break;

            }

        }


        return endIndex;

    }


    // ==================================
    // نص العائلات
    // ==================================

    function buildFamilyText(
        text
    ) {

        const tokens =
            tokenizeDocumentText(
                text
            );


        const familyTokens =
            [];


        tokens.forEach(
            function (
                token
            ) {

                const family =
                    getConservativeFamilyKey(
                        token,
                        surfaceSet
                    );


                if (
                    family
                ) {

                    familyTokens.push(
                        family
                    );

                }

            }
        );


        return familyTokens.join(
            " "
        );

    }


    // ==================================
    // السجلات
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


            if (
                !text
            ) {

                return;

            }


            const paragraphIndex =
                Number(
                    paragraph.index
                );


            const nearestHeading =
                getNearestHeading(
                    paragraphIndex
                );


            const majorHeading =
                getMajorHeading(
                    paragraphIndex
                );


            const majorHeadingLevel =
                majorHeading
                    ? getHeadingLevel(
                        majorHeading.style
                    )
                    : 9;


            const sectionEnd =
                majorHeading
                    ? getSectionEnd(
                        majorHeading.index,
                        majorHeadingLevel
                    )
                    : paragraphs.length -
                        1;


            records.push({

                id:
                    "p-" +
                    String(
                        paragraphIndex
                    ),

                paragraphIndex:
                    paragraphIndex,

                text:
                    text,

                heading:
                    nearestHeading
                        ? String(
                            nearestHeading.text ||
                            ""
                        ).trim()
                        : "",

                headingIndex:
                    nearestHeading
                        ? Number(
                            nearestHeading.index
                        )
                        : -1,

                headingLevel:
                    nearestHeading
                        ? getHeadingLevel(
                            nearestHeading.style
                        )
                        : 9,

                majorHeading:
                    majorHeading
                        ? String(
                            majorHeading.text ||
                            ""
                        ).trim()
                        : "",

                majorHeadingIndex:
                    majorHeading
                        ? Number(
                            majorHeading.index
                        )
                        : -1,

                majorHeadingLevel:
                    majorHeadingLevel,

                sectionEnd:
                    sectionEnd,

                familyText:
                    buildFamilyText(
                        text
                    ),

                isHeading:
                    Boolean(
                        nearestHeading &&
                        Number(
                            nearestHeading.index
                        ) ===
                        paragraphIndex
                    )

            });

        }
    );


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

                text:
                    "string",

                heading:
                    "string",

                headingIndex:
                    "number",

                headingLevel:
                    "number",

                majorHeading:
                    "string",

                majorHeadingIndex:
                    "number",

                majorHeadingLevel:
                    "number",

                sectionEnd:
                    "number",

                familyText:
                    "string",

                isHeading:
                    "boolean"

            },

            language:
                "arabic"

        });


    // ==================================
    // الإدخال في دفعة واحدة
    // ==================================

    if (
        records.length >
        0
    ) {

        insertMultiple(
            db,
            records,
            500
        );

    }


    // ==================================
    // حفظ الذاكرة
    // ==================================

    oramaRetrievalDb =
        db;


    oramaRetrievalCacheKey =
        cacheKey;


    oramaRetrievalRecords =
        records;


    console.log(
        "تم بناء فهرس Orama:",
        records.length,
        "فقرة"
    );


    return db;

}


// =====================================================
// Search Indexed Document
// البحث الجديد بالكامل باستخدام Orama
// =====================================================

async function searchIndexedDocument(
    documentId,
    query,
    options
) {

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

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            totalQueryTerms:
                0,

            indexedOccurrences:
                0

        };

    }


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


    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    const {
        search
    } =
        require(
            "@orama/orama"
        );


    // ==================================
    // كلمات السؤال
    // ==================================

    const queryTokens =
        getSearchQueryTokens(
            searchTerm
        );


    // ==================================
    // العائلات المطلوبة
    // ==================================

    const surfaceSet =
        new Set();


    oramaRetrievalRecords.forEach(
        function (
            record
        ) {

            tokenizeDocumentText(
                record.text
            ).forEach(
                function (
                    token
                ) {

                    surfaceSet.add(
                        token
                    );

                }
            );

        }
    );


    const queryFamilies =
        [];


    queryTokens.forEach(
        function (
            token
        ) {

            const family =
                getConservativeFamilyKey(
                    token,
                    surfaceSet
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


    // ==================================
    // إنشاء استعلامات Orama
    // ==================================
    // لا نرسل السؤال الطويل فقط.
    // نستخدم أيضًا الكلمات المهمة والعائلات.
    // ==================================

    const searchTerms =
        [
            searchTerm,
            ...queryTokens,
            ...queryFamilies
        ]
            .filter(
                function (
                    value,
                    index,
                    array
                ) {

                    return (
                        value &&
                        array.indexOf(
                            value
                        ) ===
                        index
                    );

                }
            );


    const resultMap =
        new Map();


    searchTerms.forEach(
        function (
            term
        ) {

            let result;


            try {

                result =
                    search(
                        db,
                        {

                            term:
                                term,

                            properties: [

                                "text",
                                "heading",
                                "majorHeading",
                                "familyText"

                            ],

                            boost: {

                                heading:
                                    8,

                                majorHeading:
                                    7,

                                familyText:
                                    3,

                                text:
                                    1

                            },

                            tolerance:
                                2,

                            limit:
                                50

                        }
                    );

            }
            catch (error) {

                console.warn(
                    "فشل استعلام Orama:",
                    term,
                    error
                );

                return;

            }


            if (
                !result ||
                !Array.isArray(
                    result.hits
                )
            ) {

                return;

            }


            result.hits.forEach(
                function (
                    hit
                ) {

                    if (
                        !hit ||
                        !hit.document
                    ) {

                        return;

                    }


                    const index =
                        Number(
                            hit.document
                                .paragraphIndex
                        );


                    if (
                        !resultMap.has(
                            index
                        )
                    ) {

                        resultMap.set(
                            index,
                            {

                                document:
                                    hit.document,

                                score:
                                    0,

                                bestScore:
                                    0,

                                matchedTerms:
                                    new Set()

                            }
                        );

                    }


                    const entry =
                        resultMap.get(
                            index
                        );


                    const hitScore =
                        Number(
                            hit.score ||
                            0
                        );


                    entry.score +=
                        hitScore;


                    entry.bestScore =
                        Math.max(
                            entry.bestScore,
                            hitScore
                        );


                    entry.matchedTerms.add(
                        term
                    );

                }
            );

        }
    );


    if (
        resultMap.size ===
        0
    ) {

        return {

            query:
                searchTerm,

            count:
                0,

            results:
                [],

            matchedTerms:
                queryTokens,

            matchedFamilies:
                queryFamilies,

            totalQueryTerms:
                queryTokens.length,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // خريطة الفقرات
    // ==================================

    const paragraphMap =
        new Map();


    oramaRetrievalRecords.forEach(
        function (
            record
        ) {

            paragraphMap.set(
                Number(
                    record.paragraphIndex
                ),
                record
            );

        }
    );


    // ==================================
    // قوة العنوان
    // ==================================

    function calculateHeadingScore(
        heading
    ) {

        const normalizedHeading =
            normalizeSearchText(
                heading
            );


        if (
            !normalizedHeading
        ) {

            return 0;

        }


        let score =
            0;


        // ==================================
        // تطابق العبارة نفسها
        // ==================================

        if (
            normalizedHeading.includes(
                searchTerm
            )
        ) {

            score +=
                60;

        }


        // ==================================
        // العائلات
        // ==================================

        const headingTokens =
            getSearchQueryTokens(
                normalizedHeading
            );


        const headingFamilies =
            new Set();


        headingTokens.forEach(
            function (
                token
            ) {

                const family =
                    getConservativeFamilyKey(
                        token,
                        surfaceSet
                    );


                if (
                    family
                ) {

                    headingFamilies.add(
                        family
                    );

                }

            }
        );


        let matchedFamilies =
            0;


        queryFamilies.forEach(
            function (
                family
            ) {

                if (
                    headingFamilies.has(
                        family
                    )
                ) {

                    matchedFamilies +=
                        1;

                }

            }
        );


        if (
            queryFamilies.length >
            0
        ) {

            score +=
                (
                    matchedFamilies /
                    queryFamilies.length
                ) *
                35;

        }


        // ==================================
        // أولوية المطلب/المبحث
        // ==================================

        if (
            /المطلب|المبحث|الفصل|الباب/
                .test(
                    normalizedHeading
                )
        ) {

            score +=
                8;

        }


        // ==================================
        // العناوين الجزئية
        // ==================================

        if (
            /^اولا|^أولا|^ثانيا|^ثالثا|^رابعا|^خامسا/
                .test(
                    normalizedHeading
                )
        ) {

            score -=
                3;

        }


        return score;

    }


    // ==================================
    // اكتشاف الأقسام
    // ==================================

    const sectionMap =
        new Map();


    resultMap.forEach(
        function (
            entry
        ) {

            const record =
                entry.document;


            const sectionIndex =
                Number(
                    record.majorHeadingIndex
                );


            if (
                sectionIndex <
                0
            ) {

                return;

            }


            const score =
                calculateHeadingScore(
                    record.majorHeading
                );


            const old =
                sectionMap.get(
                    sectionIndex
                );


            if (
                !old ||
                score >
                    old.score
            ) {

                sectionMap.set(
                    sectionIndex,
                    {

                        score:
                            score,

                        heading:
                            record.majorHeading,

                        headingIndex:
                            record.majorHeadingIndex,

                        sectionEnd:
                            record.sectionEnd,

                        level:
                            record.majorHeadingLevel

                    }
                );

            }

        }
    );


    // ==================================
    // الأقسام ذات الصلة
    // ==================================

    const prioritySections =
        Array.from(
            sectionMap.entries()
        )
            .filter(
                function (
                    entry
                ) {

                    return (
                        entry[1].score >=
                        12
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        b[1].score -
                        a[1].score
                    );

                }
            )
            .slice(
                0,
                3
            );


    const prioritySectionMap =
        new Map(
            prioritySections
        );


    // ==================================
    // إنشاء النتائج الأساسية
    // ==================================

    const rankedMap =
        new Map();


    resultMap.forEach(
        function (
            entry,
            paragraphIndex
        ) {

            const record =
                entry.document;


            let score =
                entry.score;


            const normalizedText =
                normalizeSearchText(
                    record.text
                );


            const exactPhrase =
                normalizedText.includes(
                    searchTerm
                );


            const headingScore =
                calculateHeadingScore(
                    record.heading
                );


            const majorHeadingScore =
                calculateHeadingScore(
                    record.majorHeading
                );


            if (
                exactPhrase
            ) {

                score +=
                    35;

            }


            score +=
                headingScore;


            score +=
                majorHeadingScore;


            // ==================================
            // تغطية كلمات السؤال
            // ==================================

            let matchedTermsCount =
                0;


            queryTokens.forEach(
                function (
                    token
                ) {

                    if (
                        normalizedText.includes(
                            token
                        )
                    ) {

                        matchedTermsCount +=
                            1;

                    }

                }
            );


            const coverage =
                queryTokens.length >
                0
                    ? matchedTermsCount /
                      queryTokens.length
                    : 0;


            score +=
                coverage *
                15;


            // ==================================
            // أولوية القسم
            // ==================================

            const sectionIndex =
                Number(
                    record.majorHeadingIndex
                );


            const prioritySection =
                prioritySectionMap.get(
                    sectionIndex
                );


            if (
                prioritySection
            ) {

                score +=
                    30 +
                    prioritySection.score;

            }


            // ==================================
            // تخفيض المقدمة والمنهج
            // عندما يوجد قسم حقيقي قوي
            // ==================================

            if (
                prioritySections.length >
                0
            ) {

                const sectionHeading =
                    normalizeSearchText(
                        record.majorHeading ||
                        ""
                    );


                if (
                    sectionHeading ===
                        "مقدمة" ||
                    sectionHeading.includes(
                        "مشكلة الدراسة"
                    ) ||
                    sectionHeading.includes(
                        "اهداف الدراسة"
                    ) ||
                    sectionHeading.includes(
                        "أهداف الدراسة"
                    ) ||
                    sectionHeading.includes(
                        "منهج الدراسة"
                    ) ||
                    sectionHeading.includes(
                        "اسئلة الدراسة"
                    ) ||
                    sectionHeading.includes(
                        "أسئلة الدراسة"
                    )
                ) {

                    score -=
                        30;

                }

            }


            // ==================================
            // نوع المطابقة
            // ==================================

            let matchType =
                "orama";


            if (
                majorHeadingScore >=
                12 ||
                headingScore >=
                12
            ) {

                matchType =
                    "heading";

            }
            else if (
                exactPhrase
            ) {

                matchType =
                    "exact";

            }
            else if (
                matchedTermsCount >
                0
            ) {

                matchType =
                    "term";

            }


            // ==================================
            // موضع النص
            // ==================================

            let context =
                record.text;


            const exactPosition =
                normalizedText.indexOf(
                    searchTerm
                );


            if (
                exactPosition >=
                0
            ) {

                const contextStart =
                    Math.max(
                        0,
                        exactPosition -
                        120
                    );


                const contextEnd =
                    Math.min(
                        record.text.length,
                        exactPosition +
                        searchTerm.length +
                        300
                    );


                context =
                    record.text.substring(
                        contextStart,
                        contextEnd
                    );

            }
            else {

                context =
                    record.text.substring(
                        0,
                        420
                    );

            }


            const previous =
                paragraphMap.get(
                    paragraphIndex -
                    1
                );


            const next =
                paragraphMap.get(
                    paragraphIndex +
                    1
                );


            rankedMap.set(
                paragraphIndex,
                {

                    paragraphIndex:
                        paragraphIndex,

                    paragraphId:
                        record.id,

                    text:
                        record.text,

                    context:
                        context,

                    heading:
                        record.heading ||
                        "",

                    headingLevel:
                        record.headingLevel,

                    majorHeading:
                        record.majorHeading ||
                        "",

                    majorHeadingIndex:
                        record.majorHeadingIndex,

                    sectionEnd:
                        record.sectionEnd,

                    previousParagraphText:
                        previous
                            ? previous.text
                            : "",

                    nextParagraphText:
                        next
                            ? next.text
                            : "",

                    score:
                        score,

                    matchType:
                        matchType,

                    matchedFamilies:
                        queryFamilies,

                    matchedFamilyCount:
                        queryFamilies.length,

                    familyOccurrencesInParagraph:
                        matchedTermsCount,

                    exactWordMatches:
                        matchedTermsCount,

                    totalQueryTerms:
                        queryTokens.length,

                    headingMatch:
                        (
                            headingScore >=
                            12 ||
                            majorHeadingScore >=
                            12
                        ),

                    sectionPriority:
                        Boolean(
                            prioritySection
                        )

                }

            );

        }
    );


    // ==================================
    // توسيع القسم الأقوى
    // ==================================
    // مهم جدًا:
    // لا نضيف 3 فقرات فقط.
    // نأخذ نطاق المطلب/المبحث حتى
    // قبل العنوان من المستوى نفسه.
    // ==================================

    prioritySections.forEach(
        function (
            entry
        ) {

            const sectionIndex =
                Number(
                    entry[0]
                );


            const section =
                entry[1];


            const sectionRecords =
                oramaRetrievalRecords
                    .filter(
                        function (
                            record
                        ) {

                            return (
                                Number(
                                    record.majorHeadingIndex
                                ) ===
                                sectionIndex &&
                                Number(
                                    record.paragraphIndex
                                ) <=
                                Number(
                                    section.sectionEnd
                                )
                            );

                        }
                    );


            sectionRecords.forEach(
                function (
                    record
                ) {

                    const index =
                        Number(
                            record.paragraphIndex
                        );


                    if (
                        rankedMap.has(
                            index
                        )
                    ) {

                        const existing =
                            rankedMap.get(
                                index
                            );


                        existing.score +=
                            10;


                        existing.sectionPriority =
                            true;


                        rankedMap.set(
                            index,
                            existing
                        );


                        return;

                    }


                    const previous =
                        paragraphMap.get(
                            index -
                            1
                        );


                    const next =
                        paragraphMap.get(
                            index +
                            1
                        );


                    rankedMap.set(
                        index,
                        {

                            paragraphIndex:
                                index,

                            paragraphId:
                                record.id,

                            text:
                                record.text,

                            context:
                                record.text,

                            heading:
                                record.heading ||
                                "",

                            headingLevel:
                                record.headingLevel,

                            majorHeading:
                                record.majorHeading ||
                                "",

                            majorHeadingIndex:
                                record.majorHeadingIndex,

                            sectionEnd:
                                record.sectionEnd,

                            previousParagraphText:
                                previous
                                    ? previous.text
                                    : "",

                            nextParagraphText:
                                next
                                    ? next.text
                                    : "",

                            score:
                                20 +
                                section.score,

                            matchType:
                                "section",

                            matchedFamilies:
                                queryFamilies,

                            matchedFamilyCount:
                                0,

                            familyOccurrencesInParagraph:
                                0,

                            exactWordMatches:
                                0,

                            totalQueryTerms:
                                queryTokens.length,

                            headingMatch:
                                true,

                            sectionPriority:
                                true

                        }

                    );

                }
            );

        }
    );


    // ==================================
    // تحويل إلى مصفوفة
    // ==================================

    const results =
        Array.from(
            rankedMap.values()
        );


    // ==================================
    // ترتيب نهائي
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
                    b.sectionPriority
                ) !==
                Boolean(
                    a.sectionPriority
                )
            ) {

                return (
                    b.sectionPriority
                        ? 1
                        : -1
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


            return (
                a.paragraphIndex -
                b.paragraphIndex
            );

        }
    );


    // ==================================
    // منع التكرار شبه التام
    // ==================================

    const selected =
        [];


    const MAX_TEXT_OVERLAP =
        0.75;


    results.forEach(
        function (
            candidate
        ) {

            if (
                selected.length >=
                30
            ) {

                return;

            }


            const candidateText =
                String(
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

                return;

            }


            let tooSimilar =
                false;


            for (
                let i = 0;
                i < selected.length;
                i++
            ) {

                const selectedText =
                    String(
                        selected[i].text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim();


                const shorterLength =
                    Math.min(
                        candidateText.length,
                        selectedText.length
                    );


                const commonLength =
                    getCommonTextLength(
                        candidateText,
                        selectedText
                    );


                const overlap =
                    shorterLength > 0
                        ? commonLength /
                          shorterLength
                        : 0;


                if (
                    overlap >=
                    MAX_TEXT_OVERLAP
                ) {

                    tooSimilar =
                        true;

                    break;

                }

            }


            if (
                !tooSimilar
            ) {

                selected.push(
                    candidate
                );

            }

        }
    );


    // ==================================
    // إجمالي الظهورات
    // ==================================

    const indexedOccurrences =
        selected.reduce(
            function (
                total,
                item
            ) {

                return (
                    total +
                    Number(
                        item.familyOccurrencesInParagraph ||
                        0
                    )
                );

            },
            0
        );


    return {

        query:
            searchTerm,

        count:
            selected.length,

        results:
            selected,

        matchedTerms:
            queryTokens,

        matchedFamilies:
            queryFamilies,

        totalQueryTerms:
            queryTokens.length,

        indexTokenCount:
            oramaRetrievalRecords.length,

        indexUniqueTerms:
            0,

        indexUniqueFamilies:
            queryFamilies.length,

        indexedOccurrences:
            indexedOccurrences

    };

}

// ======================================
// Word AI Assistant
// PART 2 / 3
// المشاريع + المستندات + المحادثات
// ======================================


// ======================================
// Project Icon
// ======================================

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


// ======================================
// Chat Icon
// ======================================

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


// ======================================
// Save Projects
// ======================================

function saveProjects() {

    localStorage.setItem(
        "WORD_AI_PROJECTS",
        JSON.stringify(
            projects
        )
    );

}


// ======================================
// Render Documents
// ======================================

function renderDocuments() {

    if (!documentsList)
        return;


    documentsList.innerHTML =
        "";


    if (!currentProject) {

        documentsList.innerHTML = `
            <div class="empty-document">
                اختر مشروعًا لعرض مستنداته
            </div>
        `;

        return;

    }


    const projectDocuments =
        getProjectDocuments(
            currentProject.id
        );


    if (
        projectDocuments.length ===
        0
    ) {

        documentsList.innerHTML = `
            <div class="empty-document">
                لا توجد مستندات
            </div>
        `;

        return;

    }


    projectDocuments.forEach(
        function (
            documentItem,
            index
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "document-item";


            if (
                currentDocument &&
                currentDocument.id ===
                    documentItem.id
            ) {

                item.classList.add(
                    "active-document"
                );

            }


            const title =
                document.createElement(
                    "span"
                );


            title.className =
                "document-title";


            title.textContent =
                documentItem.name;


            const status =
                document.createElement(
                    "span"
                );


            status.className =
                "document-read-status";


            if (
                documentItem.indexStatus ===
                "indexed"
            ) {

                status.textContent =
                    "✓ مفهرس · " +
                    documentItem.indexTokenCount +
                    " كلمة · " +
                    documentItem.indexUniqueTerms +
                    " فريدة";


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


            title.onclick =
                function (e) {

                    e.preventDefault();
                    e.stopPropagation();


                    setCurrentDocument(
                        documentItem
                    );

                };


            const menuButton =
                document.createElement(
                    "button"
                );


            menuButton.className =
                "document-menu";


            menuButton.type =
                "button";


            menuButton.title =
                "خيارات المستند";


            menuButton.textContent =
                "⋮";


            const options =
                document.createElement(
                    "div"
                );


            options.className =
                "document-options-menu";


            options.innerHTML = `

                <div class="rename-document">
                    ✏ إعادة تسمية
                </div>

                <div class="move-document-up">
                    ↑ نقل إلى أعلى
                </div>

                <div class="move-document-down">
                    ↓ نقل إلى أسفل
                </div>

                <div class="delete-document">
                    🗑 حذف
                </div>

            `;


            if (
                index ===
                0
            ) {

                const moveUp =
                    options.querySelector(
                        ".move-document-up"
                    );


                if (moveUp) {

                    moveUp.style.display =
                        "none";

                }

            }


            if (
                index ===
                projectDocuments.length -
                1
            ) {

                const moveDown =
                    options.querySelector(
                        ".move-document-down"
                    );


                if (moveDown) {

                    moveDown.style.display =
                        "none";

                }

            }


            menuButton.onclick =
                function (e) {

                    e.preventDefault();
                    e.stopPropagation();


                    document
                        .querySelectorAll(
                            ".document-options-menu.open"
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


            // ==================================
            // Rename
            // ==================================

            const renameButton =
                options.querySelector(
                    ".rename-document"
                );


            if (renameButton) {

                renameButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldName =
                            documentItem.name;


                        const inputRename =
                            document.createElement(
                                "input"
                            );


                        inputRename.className =
                            "edit-document-title";


                        inputRename.value =
                            oldName;


                        title.replaceWith(
                            inputRename
                        );


                        inputRename.focus();


                        inputRename.setSelectionRange(
                            inputRename.value.length,
                            inputRename.value.length
                        );


                        function finishRename(
                            saveChange
                        ) {

                            const newName =
                                inputRename.value.trim();


                            if (
                                saveChange &&
                                newName !==
                                    ""
                            ) {

                                documentItem.name =
                                    newName;


                                documentItem.updatedAt =
                                    new Date()
                                        .toISOString();


                                saveDocuments();

                            }
                            else {

                                documentItem.name =
                                    oldName;

                            }


                            renderDocuments();

                        }


                        inputRename.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();


                                    finishRename(
                                        true
                                    );

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();


                                    finishRename(
                                        false
                                    );

                                }

                            };

                    };

            }


            // ==================================
            // Move Up
            // ==================================

            const moveUpButton =
                options.querySelector(
                    ".move-document-up"
                );


            if (moveUpButton) {

                moveUpButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (
                            index <=
                            0
                        ) {

                            return;

                        }


                        const previousDocument =
                            projectDocuments[
                                index - 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            previousDocument.order;


                        previousDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        previousDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // ==================================
            // Move Down
            // ==================================

            const moveDownButton =
                options.querySelector(
                    ".move-document-down"
                );


            if (moveDownButton) {

                moveDownButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (
                            index >=
                            projectDocuments.length -
                            1
                        ) {

                            return;

                        }


                        const nextDocument =
                            projectDocuments[
                                index + 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            nextDocument.order;


                        nextDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        nextDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // ==================================
            // Delete
            // ==================================

            const deleteButton =
                options.querySelector(
                    ".delete-document"
                );


            if (deleteButton) {

                deleteButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldConfirm =
                            document.querySelector(
                                ".document-delete-confirm"
                            );


                        if (oldConfirm) {

                            oldConfirm.remove();

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "document-delete-confirm";


                        confirmBox.innerHTML = `

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
                                        class="confirm-document-delete">
                                        حذف
                                    </button>

                                    <button
                                        type="button"
                                        class="cancel-document-delete">
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
                                ".confirm-document-delete"
                            );


                        if (confirmDelete) {

                            confirmDelete.onclick =
                                async function () {

                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return (
                                                    doc.id !==
                                                    documentItem.id
                                                );

                                            }
                                        );


                                    if (
                                        currentProject &&
                                        Array.isArray(
                                            currentProject.documents
                                        )
                                    ) {

                                        currentProject.documents =
                                            currentProject.documents.filter(
                                                function (
                                                    id
                                                ) {

                                                    return (
                                                        id !==
                                                        documentItem.id
                                                    );

                                                }
                                            );


                                        currentProject.updatedAt =
                                            new Date()
                                                .toISOString();


                                        saveProjects();

                                    }


                                    try {

                                        await deleteWorkingWordFile(
                                            documentItem.storageId
                                        );

                                    }
                                    catch (
                                        storageError
                                    ) {

                                        console.warn(
                                            "تعذر حذف نسخة العمل:",
                                            storageError
                                        );

                                    }


                                    if (
                                        currentDocument &&
                                        currentDocument.id ===
                                            documentItem.id
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


                                    // إبطال ذاكرة Orama
                                    oramaRetrievalDb =
                                        null;

                                    oramaRetrievalCacheKey =
                                        "";

                                    oramaRetrievalRecords =
                                        [];


                                    const remaining =
                                        getProjectDocuments(
                                            currentProject.id
                                        );


                                    remaining.forEach(
                                        function (
                                            doc,
                                            newIndex
                                        ) {

                                            doc.order =
                                                newIndex +
                                                1;


                                            doc.updatedAt =
                                                new Date()
                                                    .toISOString();

                                        }
                                    );


                                    saveDocuments();


                                    confirmBox.remove();


                                    renderDocuments();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-document-delete"
                            );


                        if (cancelDelete) {

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


            documentsList.appendChild(
                item
            );

        }
    );

}


// ======================================
// Chat System
// ======================================

let chats = [];


try {

    chats =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_CHATS"
            )
        ) || [];

}
catch (error) {

    chats = [];

}


let currentChat =
    null;


function saveChats() {

    localStorage.setItem(
        "WORD_AI_CHATS",
        JSON.stringify(
            chats
        )
    );

}


// ======================================
// Initialize Sidebar Sections
// ======================================

function initializeSidebarSections() {

    const headers =
        document.querySelectorAll(
            ".section-title[data-target], .section-toggle[data-target]"
        );


    headers.forEach(
        function (
            header
        ) {

            const targetId =
                header.getAttribute(
                    "data-target"
                );


            if (
                !targetId
            ) {

                return;

            }


            const target =
                document.getElementById(
                    targetId
                );


            if (
                !target
            ) {

                return;

            }


            target.classList.remove(
                "open"
            );


            header.classList.remove(
                "open"
            );


            header.onclick =
                function (e) {

                    e.preventDefault();
                    e.stopPropagation();


                    const isOpen =
                        target.classList.contains(
                            "open"
                        );


                    if (
                        isOpen
                    ) {

                        target.classList.remove(
                            "open"
                        );


                        header.classList.remove(
                            "open"
                        );

                    }
                    else {

                        target.classList.add(
                            "open"
                        );


                        header.classList.add(
                            "open"
                        );

                    }

                };

        }
    );

}


// ======================================
// Render Projects
// ======================================

function renderProjects() {

    if (
        !projectsList
    ) {

        return;

    }


    projectsList.innerHTML =
        "";


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


            item.innerHTML = `

                <span class="project-title">
                    ${projectIcon}
                    ${project.name}
                </span>

                <button
                    class="project-menu"
                    type="button">
                    ⋮
                </button>

                <div class="project-options-menu">

                    <div class="rename-project">
                        ✏ إعادة تسمية
                    </div>

                    <div class="delete-project">
                        🗑 حذف
                    </div>

                </div>

            `;


            item.onclick =
                function (e) {

                    e.stopPropagation();


                    setCurrentProject(
                        project
                    );


                    if (
                        projectsPopup
                    ) {

                        projectsPopup.classList.remove(
                            "open"
                        );

                    }

                };


            const menu =
                item.querySelector(
                    ".project-menu"
                );


            if (menu) {

                menu.onclick =
                    function (e) {

                        e.stopPropagation();


                        document
                            .querySelectorAll(
                                ".project-options-menu"
                            )
                            .forEach(
                                function (
                                    m
                                ) {

                                    m.classList.remove(
                                        "open"
                                    );

                                }
                            );


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (
                            !options
                        ) {

                            return;

                        }


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
                            left + "px";


                        options.style.top =
                            top + "px";


                        options.style.right =
                            "auto";


                        options.style.bottom =
                            "auto";


                        options.style.zIndex =
                            "999999";

                    };

            }


            // ==================================
            // Rename Project
            // ==================================

            const renameProject =
                item.querySelector(
                    ".rename-project"
                );


            if (renameProject) {

                renameProject.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (
                            options
                        ) {

                            options.classList.remove(
                                "open"
                            );

                        }


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


                        title.innerHTML = `
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

                                    const value =
                                        edit.value.trim();


                                    project.name =
                                        value !==
                                            ""
                                            ? value
                                            : oldName;


                                    project.updatedAt =
                                        new Date()
                                            .toISOString();


                                    saveProjects();


                                    renderProjects();


                                    renderExpandedProjects();

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    project.name =
                                        oldName;


                                    renderProjects();

                                }

                            };

                    };

            }


            // ==================================
            // Delete Project
            // ==================================

            const deleteProject =
                item.querySelector(
                    ".delete-project"
                );


            if (deleteProject) {

                deleteProject.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (
                            options
                        ) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "project-delete-confirm";


                        confirmBox.innerHTML = `

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
                                function () {

                                    const projectDocumentIds =
                                        Array.isArray(
                                            project.documents
                                        )
                                            ? project.documents
                                            : [];


                                    projectDocumentIds
                                        .forEach(
                                            function (
                                                id
                                            ) {

                                                const doc =
                                                    documents.find(
                                                        function (
                                                            d
                                                        ) {

                                                            return (
                                                                d.id ===
                                                                id
                                                            );

                                                        }
                                                    );


                                                if (
                                                    doc
                                                ) {

                                                    deleteWorkingWordFile(
                                                        doc.storageId
                                                    )
                                                        .catch(
                                                            function () {}
                                                        );

                                                }

                                            }
                                        );


                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return (
                                                    !projectDocumentIds.includes(
                                                        doc.id
                                                    )
                                                );

                                            }
                                        );


                                    projects =
                                        projects.filter(
                                            function (
                                                p
                                            ) {

                                                return (
                                                    p.id !==
                                                    project.id
                                                );

                                            }
                                        );


                                    if (
                                        currentProject &&
                                        currentProject.id ===
                                            project.id
                                    ) {

                                        currentProject =
                                            null;


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


                                    // إبطال Orama
                                    oramaRetrievalDb =
                                        null;

                                    oramaRetrievalCacheKey =
                                        "";

                                    oramaRetrievalRecords =
                                        [];


                                    saveDocuments();

                                    saveProjects();

                                    renderProjects();

                                    renderExpandedProjects();

                                    renderDocuments();

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


// ======================================
// Projects Button
// ======================================

if (
    projectsBtn
) {

    projectsBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (
                !projectsPopup
            ) {

                return;

            }


            if (
                chatPopup
            ) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (
                searchPopup
            ) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (
                settingsWindow
            ) {

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


// ======================================
// Expanded Sidebar
// ======================================

if (
    sidebarToggleBtn &&
    expandedSidebar
) {

    sidebarToggleBtn.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (
                !expandedSidebar ||
                !expandedSidebarToggleSlot
            ) {

                return;

            }


            const isOpening =
                !expandedSidebar.classList.contains(
                    "open"
                );


            if (
                isOpening
            ) {

                expandedSidebar.classList.add(
                    "open"
                );


                document.body.classList.add(
                    "expanded-sidebar-open"
                );


                sidebarToggleBtn.title =
                    "إخفاء القائمة";


                sidebarToggleBtn.classList.add(
                    "sidebar-open"
                );


                expandedSidebarToggleSlot.appendChild(
                    sidebarToggleBtn
                );

            }
            else {

                expandedSidebar.classList.remove(
                    "open"
                );


                document.body.classList.remove(
                    "expanded-sidebar-open"
                );


                sidebarToggleBtn.title =
                    "إظهار القائمة";


                sidebarToggleBtn.classList.remove(
                    "sidebar-open"
                );


                if (
                    sidebarTogglePlaceholder.parentNode
                ) {

                    sidebarTogglePlaceholder
                        .parentNode
                        .insertBefore(
                            sidebarToggleBtn,
                            sidebarTogglePlaceholder.nextSibling
                        );

                }

            }

        };

}


// ======================================
// Render Projects in Expanded Sidebar
// ======================================

function renderExpandedProjects() {

    const list =
        document.getElementById(
            "expanded-projects-list"
        );


    if (
        !list
    ) {

        return;

    }


    list.innerHTML =
        "";


    projects.forEach(
        function (
            project
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "expanded-project-item";


            item.innerHTML = `
                <span>
                    ${projectIcon}
                    ${project.name}
                </span>
            `;


            item.onclick =
                function (e) {

                    e.stopPropagation();


                    setCurrentProject(
                        project
                    );

                };


            list.appendChild(
                item
            );

        }
    );

}


// ======================================
// Sidebar Recent Chats
// ======================================

function renderSidebarChats() {

    const list =
        document.getElementById(
            "new-chat-list"
        );


    if (
        !list
    ) {

        return;

    }


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


                        renderChat();


                        if (
                            projectsPopup
                        ) {

                            projectsPopup
                                .classList
                                .remove(
                                    "open"
                                );

                        }


                        if (
                            chatPopup
                        ) {

                            chatPopup
                                .classList
                                .remove(
                                    "open"
                                );

                        }


                        if (
                            searchPopup
                        ) {

                            searchPopup
                                .classList
                                .remove(
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


// ======================================
// Create New Chat
// ======================================

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
            currentProject
                ? currentProject.id
                : null

    };


    renderSidebarChats();


    renderRecentChats();


    if (
        input
    ) {

        input.value =
            "";


        input.style.height =
            "auto";

    }


    if (
        chatArea
    ) {

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

    }

}


// ======================================
// New Chat Button
// ======================================

if (
    newChatBtn
) {

    newChatBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (
                projectsPopup
            ) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (
                chatPopup
            ) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (
                searchPopup
            ) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (
                settingsWindow
            ) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            createNewChat();

        };

}


// ======================================
// Render Chat
// ======================================

function renderChat() {

    if (
        !chatArea
    ) {

        return;

    }


    chatArea.innerHTML =
        "";


    if (
        !currentChat
    ) {

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


    currentChat.messages.forEach(
        function (
            msg
        ) {

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


            if (
                msg.role ===
                "user"
            ) {

                div.textContent =
                    msg.text ||
                    "";

            }
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


// ======================================
// Render Chat List
// ======================================

function renderChatList() {

    const list =
        document.getElementById(
            "chat-list"
        );


    if (
        !list
    ) {

        return;

    }


    list.innerHTML =
        "";


    if (
        chats.length ===
        0
    ) {

        list.innerHTML =
            "<div class='empty-chat'>لا توجد محادثات</div>";


        return;

    }


    chats.forEach(
        function (
            chat
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "chat-history-item";


            item.innerHTML = `

                <span class="chat-title">
                    ${chatIcon}
                    ${chat.title}
                </span>

                <button
                    class="chat-menu"
                    type="button">
                    ⋮
                </button>

                <div class="chat-options-menu">

                    <div class="rename-chat">
                        ✏ إعادة تسمية
                    </div>

                    <div class="delete-chat">
                        🗑 حذف
                    </div>

                </div>

            `;


            const title =
                item.querySelector(
                    ".chat-title"
                );


            if (
                title
            ) {

                title.onclick =
                    function (e) {

                        e.stopPropagation();


                        currentChat =
                            chat;


                        renderChat();


                        if (
                            expandedSidebar
                        ) {

                            expandedSidebar.classList.remove(
                                "open"
                            );

                        }

                    };

            }


            const menu =
                item.querySelector(
                    ".chat-menu"
                );


            if (
                menu
            ) {

                menu.onclick =
                    function (e) {

                        e.stopPropagation();


                        document
                            .querySelectorAll(
                                ".chat-options-menu"
                            )
                            .forEach(
                                function (
                                    m
                                ) {

                                    m.classList.remove(
                                        "open"
                                    );

                                }
                            );


                        const options =
                            item.querySelector(
                                ".chat-options-menu"
                            );


                        if (
                            !options
                        ) {

                            return;

                        }


                        options.classList.add(
                            "open"
                        );


                        const rect =
                            menu.getBoundingClientRect();


                        options.style.position =
                            "fixed";


                        options.style.left =
                            Math.max(
                                8,
                                rect.left -
                                140 -
                                8
                            ) + "px";


                        options.style.top =
                            rect.bottom +
                            8 +
                            "px";


                        options.style.zIndex =
                            "999999";

                    };

            }


            // ==================================
            // Rename Chat
            // ==================================

            const renameBtn =
                item.querySelector(
                    ".rename-chat"
                );


            if (
                renameBtn
            ) {

                renameBtn.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".chat-options-menu"
                            );


                        if (
                            options
                        ) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const titleSpan =
                            item.querySelector(
                                ".chat-title"
                            );


                        if (
                            !titleSpan
                        ) {

                            return;

                        }


                        const oldName =
                            chat.title;


                        titleSpan.innerHTML = `
                            <input
                                class="edit-chat-title"
                                value="${oldName}">
                        `;


                        const editInput =
                            titleSpan.querySelector(
                                ".edit-chat-title"
                            );


                        if (
                            !editInput
                        ) {

                            return;

                        }


                        editInput.focus();


                        editInput.setSelectionRange(
                            editInput.value.length,
                            editInput.value.length
                        );


                        editInput.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    const value =
                                        editInput.value.trim();


                                    chat.title =
                                        value !==
                                            ""
                                            ? value
                                            : oldName;


                                    saveChats();


                                    renderChatList();


                                    renderSidebarChats();


                                    renderRecentChats();

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    chat.title =
                                        oldName;


                                    renderChatList();

                                }

                            };

                    };

            }


            // ==================================
            // Delete Chat
            // ==================================

            const deleteBtn =
                item.querySelector(
                    ".delete-chat"
                );


            if (
                deleteBtn
            ) {

                deleteBtn.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".chat-options-menu"
                            );


                        if (
                            options
                        ) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "delete-confirm";


                        confirmBox.innerHTML = `

                            <div class="confirm-dialog">

                                <p>
                                    هل تريد حذف المحادثة:
                                    <br>

                                    <strong>
                                        ${chat.title}
                                    </strong>

                                    ؟
                                </p>

                                <button
                                    class="confirm-delete"
                                    type="button">
                                    حذف
                                </button>

                                <button
                                    class="cancel-delete"
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
                                ".confirm-delete"
                            );


                        if (
                            confirmDelete
                        ) {

                            confirmDelete.onclick =
                                function () {

                                    chats =
                                        chats.filter(
                                            function (
                                                c
                                            ) {

                                                return (
                                                    c.id !==
                                                    chat.id
                                                );

                                            }
                                        );


                                    if (
                                        currentChat &&
                                        currentChat.id ===
                                            chat.id
                                    ) {

                                        currentChat =
                                            null;


                                        renderChat();

                                    }


                                    saveChats();


                                    renderChatList();


                                    renderSidebarChats();


                                    renderRecentChats();


                                    confirmBox.remove();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-delete"
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


            list.appendChild(
                item
            );

        }
    );

}


// ======================================
// Recent Chats Popup
// ======================================

function renderRecentChats() {

    if (
        !recentChatList
    ) {

        return;

    }


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


                        if (
                            chatPopup
                        ) {

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


// ======================================
// Chat Button
// ======================================

if (
    chatBtn
) {

    chatBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (
                !chatPopup
            ) {

                return;

            }


            if (
                projectsPopup
            ) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (
                searchPopup
            ) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (
                settingsWindow
            ) {

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


            renderChatList();


            chatPopup.classList.add(
                "open"
            );

        };

}


// ======================================
// Save Initial Projects/Chats State
// ======================================

saveProjects();
saveChats();

        // ======================================
        // New Project
        // ======================================

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
                                    name !== ""
                                ) {

                                    const now =
                                        new Date()
                                            .toISOString();


                                    const newProject = {

                                        id:
                                            Date.now(),

                                        name:
                                            name,

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
                                        newProject
                                    );


                                    saveProjects();


                                    renderProjects();


                                    renderExpandedProjects();

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
                        function (e) {

                            e.stopPropagation();

                        };

                };

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
            new Date()
                .toISOString();


        saveProjects();

    }

}


// ======================================
// Set Active Project
// ======================================

function setCurrentProject(
    project
) {

    if (!project) {

        currentProject =
            null;


        renderDocuments();


        return;

    }


    currentProject =
        project;


    renderDocuments();

}


// ======================================
// Read Current Working Document
// القراءة + الفهرسة + البنية
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


                    return body.text ||
                        "";

                }
            );


        await saveDocumentText(
            documentItem.id,
            text
        );


        updateDocumentReadStatus(
            documentItem,
            "read"
        );


        updateDocumentIndexStatus(
            documentItem,
            "indexing"
        );


        const indexData =
            await rebuildDocumentIndex(
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


        // ==============================
        // تثبيت بيانات الفهرس
        // ==============================

        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            indexData.tokenCount ||
            0;


        documentItem.indexUniqueTerms =
            indexData.uniqueTerms ||
            0;


        documentItem.indexUniqueFamilies =
            indexData.uniqueFamilies ||
            0;


        documentItem.indexSchemaVersion =
            INDEX_SCHEMA_VERSION;


        documentItem.indexUpdatedAt =
            indexData.updatedAt ||
            new Date()
                .toISOString();


        saveDocuments();


        // ==============================
        // بناء Orama فعليًا
        // ==============================

        oramaRetrievalDb =
            null;


        oramaRetrievalCacheKey =
            "";


        oramaRetrievalRecords =
            [];


        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


        console.log(
            "تمت قراءة المستند وبناء الفهرس بنجاح:",
            {
                documentId:
                    documentItem.id,

                tokenCount:
                    indexData.tokenCount,

                uniqueTerms:
                    indexData.uniqueTerms,

                uniqueFamilies:
                    indexData.uniqueFamilies,

                indexVersion:
                    indexData.indexVersion,

                oramaRecords:
                    oramaRetrievalRecords.length
            }
        );


        return text;

    }
    catch (error) {

        updateDocumentReadStatus(
            documentItem,
            "error"
        );


        updateDocumentIndexStatus(
            documentItem,
            "error"
        );


        console.error(
            "فشل قراءة/فهرسة المستند:",
            error
        );


        throw error;

    }

}


// ======================================
// Set Active Document
// ======================================

function setCurrentDocument(
    documentItem
) {

    if (!documentItem) {

        currentDocument =
            null;


        currentCitationSources =
            [];


        if (documentTitle) {

            documentTitle.textContent =
                "لا يوجد مستند مفتوح";

        }


        renderDocuments();


        return;

    }


    currentDocument =
        documentItem;


    currentCitationSources =
        [];


    if (documentTitle) {

        documentTitle.textContent =
            documentItem.name;

    }


    if (
        documentItem.readStatus ===
        "read"
    ) {

        ensureDocumentIndex(
            documentItem
        )
            .then(
                async function () {

                    const structure =
                        await ensureDocumentStructure(
                            documentItem
                        );


                    if (structure) {

                        await ensureOramaRetrievalIndex(
                            documentItem,
                            structure
                        );

                    }


                    renderDocuments();

                }
            )
            .catch(
                function (error) {

                    console.error(
                        "تعذر تحديث فهرس المستند:",
                        error
                    );


                    renderDocuments();

                }
            );


        renderDocuments();


        return;

    }


    readCurrentWordDocument(
        documentItem
    )
        .then(
            function (text) {

                console.log(
                    "محتوى نسخة العمل:",
                    text
                );


                renderDocuments();

            }
        )
        .catch(
            function (error) {

                console.error(
                    "تعذر قراءة نسخة العمل:",
                    error
                );


                renderDocuments();

            }
        );


    renderDocuments();

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
        new Date()
            .toISOString();

}


// ======================================
// Add Document
// ======================================

if (
    addDocumentBtn &&
    wordDocumentPicker
) {

    addDocumentBtn.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            if (!currentProject) {

                if (documentsList) {

                    documentsList.innerHTML = `
                        <div class="empty-document">
                            اختر مشروعًا أولًا لإضافة مستند
                        </div>
                    `;

                }


                return;

            }


            wordDocumentPicker.value =
                "";


            wordDocumentPicker.click();

        };


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


                renderDocuments();


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


                if (documentsList) {

                    documentsList.innerHTML = `
                        <div class="empty-document">
                            تعذر استيراد المستند
                        </div>
                    `;

                }

            }

        };

}


// ======================================
// Format AI Message
// ======================================

function formatAIMessage(
    text,
    citationSources
) {

    if (!text) {

        return "";

    }


    const sources =
        Array.isArray(
            citationSources
        )
            ? citationSources
            : currentCitationSources;


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


        html =
            html.replace(
                /\[مقطع\s*([0-9٠-٩\s،,]+)\]/g,
                function (
                    match,
                    ranksText
                ) {

                    const normalized =
                        String(
                            ranksText
                        )
                            .replace(
                                /٠/g,
                                "0"
                            )
                            .replace(
                                /١/g,
                                "1"
                            )
                            .replace(
                                /٢/g,
                                "2"
                            )
                            .replace(
                                /٣/g,
                                "3"
                            )
                            .replace(
                                /٤/g,
                                "4"
                            )
                            .replace(
                                /٥/g,
                                "5"
                            )
                            .replace(
                                /٦/g,
                                "6"
                            )
                            .replace(
                                /٧/g,
                                "7"
                            )
                            .replace(
                                /٨/g,
                                "8"
                            )
                            .replace(
                                /٩/g,
                                "9"
                            );


                    const ranks =
                        normalized
                            .split(
                                /[،,]+/
                            )
                            .map(
                                function (
                                    value
                                ) {

                                    return Number(
                                        value.trim()
                                    );

                                }
                            )
                            .filter(
                                function (
                                    value
                                ) {

                                    return (
                                        !Number.isNaN(
                                            value
                                        )
                                    );

                                }
                            );


                    if (
                        ranks.length ===
                        0
                    ) {

                        return match;

                    }


                    const citationButtons =
                        ranks.map(
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


                                if (!source) {

                                    return (
                                        "[مقطع " +
                                        rank +
                                        "]"
                                    );

                                }


                                return `
                                    <button
                                        type="button"
                                        class="document-citation"
                                        data-citation-rank="${rank}"
                                        title="الانتقال إلى المقطع ${rank}">
                                        [مقطع ${rank}]
                                    </button>
                                `;

                            }
                        );


                    return citationButtons.join(
                        " "
                    );

                }
            );


        return html;

    }
    catch (error) {

        console.error(
            "Markdown formatting error:",
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


// ======================================
// Open Citation In Word
// ======================================

async function openCitationInWord(
    rank
) {

    const source =
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
        );


    if (!source) {

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
            "Word API غير متاحة."
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


                if (!searchText) {

                    throw new Error(
                        "لا يوجد نص صالح للمقطع."
                    );

                }


                if (
                    searchText.length >
                    180
                ) {

                    searchText =
                        searchText.substring(
                            0,
                            180
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
                    results.items.length ===
                    0
                ) {

                    const fallbackText =
                        searchText.substring(
                            0,
                            80
                        );


                    const fallbackResults =
                        body.search(
                            fallbackText,
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


                    fallbackResults.load(
                        "items"
                    );


                    await context.sync();


                    if (
                        fallbackResults.items.length ===
                        0
                    ) {

                        throw new Error(
                            "لم يتم العثور على نص المقطع في المستند."
                        );

                    }


                    fallbackResults.items[0].select(
                        "Select"
                    );


                    await context.sync();


                    return;

                }


                results.items[0].select(
                    "Select"
                );


                await context.sync();

            }
        );

    }
    catch (error) {

        console.error(
            "تعذر الانتقال إلى المقطع:",
            error
        );

    }

}


// =====================================================
// AI SETTINGS
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

        return {};

    }

}


function saveAISettings(
    data
) {

    localStorage.setItem(
        "AI_SETTINGS",
        JSON.stringify(
            data
        )
    );

}


// ======================================
// Provider Info
// ======================================

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


// ======================================
// Load Settings
// ======================================

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


        const option =
            document.createElement(
                "option"
            );


        option.value =
            savedModel;


        option.textContent =
            savedModel ||
            "أدخل المفتاح ثم حدّث النماذج";


        modelSelect.appendChild(
            option
        );


        if (
            savedModel
        ) {

            modelSelect.value =
                savedModel;

        }

    }


    updateProviderInfo();

}


// ======================================
// Settings Button
// ======================================

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


// ======================================
// Provider Change
// ======================================

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


// ======================================
// Close Settings
// ======================================

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


// ======================================
// Show / Hide Key
// ======================================

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


// ======================================
// Save AI Settings
// ======================================

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


// ======================================
// Load Models
// ======================================

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


// ======================================
// Groq Models
// ======================================

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

                    Authorization:
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

                    return a.id.localeCompare(
                        b.id
                    );

                }
            );


    populateModels(
        models.map(
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
        )
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج Groq: " +
            models.length;

    }

}


// ======================================
// OpenRouter Models
// ======================================

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

                    Authorization:
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
        freeModels.map(
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
        );


    models.sort(
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


// ======================================
// OpenAI Models
// ======================================

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

                    Authorization:
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
        result.data.filter(
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
                    item.id.toLowerCase();


                return (
                    id.startsWith("gpt-") ||
                    id.startsWith("o1") ||
                    id.startsWith("o3") ||
                    id.startsWith("o4")
                );

            }
        );


    models.sort(
        function (
            a,
            b
        ) {

            return a.id.localeCompare(
                b.id
            );

        }
    );


    populateModels(
        models.map(
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
        )
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج OpenAI: " +
            models.length;

    }

}


// ======================================
// Gemini Models
// ======================================

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
        result.models.filter(
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
        );


    populateModels(
        models.map(
            function (
                item
            ) {

                const cleanId =
                    String(
                        item.name
                    ).replace(
                        /^models\//,
                        ""
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
        )
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث نماذج Gemini: " +
            models.length;

    }

}


// ======================================
// Populate Models
// ======================================

function populateModels(
    models
) {

    if (!modelSelect)
        return;


    const saved =
        getSavedSettings();


    const savedModel =
        saved.model ||
        "";


    modelSelect.innerHTML =
        "";


    if (
        !models ||
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

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                item.id;


            option.textContent =
                item.name;


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


// ======================================
// Test AI Connection
// ======================================

async function testAIConnection() {

    const data = {

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


    if (!data.key) {

        throw new Error(
            "يرجى إدخال مفتاح API أولاً."
        );

    }


    if (!data.model) {

        throw new Error(
            "يرجى تحديد نموذج الذكاء الاصطناعي أولاً."
        );

    }


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

                        Authorization:
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

                        Authorization:
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


    if (
        data.provider ===
        "gemini"
    ) {

        const model =
            normalizeGeminiModel(
                data.model
            );


        const url =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(
                model
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

                        Authorization:
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


// ======================================
// Test Connection Button
// ======================================

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
// Detect Retrieval Profile
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


    if (
        /ماهو|ماهى|ماهي|ما هي|المقصود|معنى|تعريف|يقصد ب|المراد ب/
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


    if (
        /لماذا|سبب|اسباب|علة|علل|لأن|لان|بسبب/
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
            providerName ||
            ""
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
            text ||
            ""
        ).length /
        4
    );

}


// =====================================================
// Build AI Document Context
// =====================================================

async function buildAIDocumentContext(
    query
) {

    if (!currentDocument) {

        currentCitationSources =
            [];


        return {

            found:
                false,

            query:
                query,

            profile:
                "general",

            text:
                ""

        };

    }


    try {

        const retrievalProfile =
            getRetrievalProfile(
                query
            );


        const searchResult =
            await searchIndexedDocument(
                currentDocument.id,
                query,
                {
                    profile:
                        retrievalProfile.type
                }
            );


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
                    retrievalProfile.type,

                text:
                    ""

            };

        }


        const settings =
            getSavedSettings();


        const retrievalLimits =
            getRetrievalLimits(
                settings.provider,
                settings.model
            );


        let maxResults =
            retrievalLimits.maxResults;


        let maxChars =
            retrievalLimits.maxChars;


        if (
            retrievalProfile.type ===
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


        const retrieval =
            buildRetrievalContext(
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


        if (
            !retrieval ||
            !retrieval.text
        ) {

            currentCitationSources =
                [];


            return {

                found:
                    false,

                query:
                    query,

                profile:
                    retrievalProfile.type,

                text:
                    ""

            };

        }


        currentCitationSources =
            Array.isArray(
                retrieval.contexts
            )
                ? retrieval.contexts.map(
                    function (
                        item
                    ) {

                        return {

                            rank:
                                item.rank,

                            paragraphIndex:
                                item.paragraphIndex,

                            heading:
                                item.heading ||
                                "",

                            mainParagraph:
                                item.mainParagraph ||
                                "",

                            text:
                                item.context ||
                                item.mainParagraph ||
                                ""

                        };

                    }
                )
                : [];


        return {

            found:
                true,

            query:
                query,

            profile:
                retrievalProfile.type,

            resultCount:
                Number(
                    searchResult.count ||
                    0
                ),

            selectedCount:
                Number(
                    retrieval.selectedCount ||
                    0
                ),

            totalOccurrences:
                Number(
                    retrieval.totalOccurrences ||
                    0
                ),

            matchedFamilies:
                Array.isArray(
                    searchResult.matchedFamilies
                )
                    ? searchResult.matchedFamilies
                    : [],

            matchedTerms:
                Array.isArray(
                    searchResult.matchedTerms
                )
                    ? searchResult.matchedTerms
                    : [],

            contexts:
                retrieval.contexts ||
                [],

            text:
                retrieval.text ||
                ""

        };

    }
    catch (error) {

        currentCitationSources =
            [];


        console.warn(
            "تعذر استرجاع سياق المستند:",
            error
        );


        return {

            found:
                false,

            query:
                query,

            profile:
                "general",

            text:
                ""

        };

    }

}


// =====================================================
// AI REQUEST
// =====================================================

async function askAI(
    text
) {

    const data =
        getSavedSettings();


    const selectedProvider =
        (
            data.provider ||
            "openrouter"
        ).toLowerCase();


    const key =
        data.key ||
        "";


    const model =
        data.model ||
        "";


    if (!key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
        );

    }


    if (!model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
        );

    }


    // مهم:
    // بناء سياق المستند قبل حساب historyLimit
    // حتى لا يحدث الوصول إلى const قبل تهيئته.

    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const conversationMessages =
        [];


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
            currentChat.messages.slice(
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


    let userContent =
        text;


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
                        documentContext.matchedFamilies.length
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
                "استبعد المقاطع التي لا تجيب مباشرة عن السؤال حتى لو احتوت كلمات البحث.",
                "إذا تكررت الفكرة نفسها في أكثر من مقطع، اذكرها مرة واحدة واجمع الإحالات.",
                "لا تذكر مشكلة الدراسة أو أهداف البحث أو منهجه أو أسئلته إلا إذا طلب المستخدم ذلك صراحة.",
                "ضع الإحالة [مقطع رقم] بعد الفكرة التي يدعمها المقطع.",
                "لا تخترع أرقام المقاطع أو الإحالات.",
                "حافظ على العربية والأسلوب الأكاديمي الواضح.",
                "لا تبدأ باعتذار أو تمهيد عام غير ضروري.",
                "لا تعيد صياغة سؤال المستخدم.",
                "قدّم خلاصة تركيبية مترابطة من المادة المستخرجة."

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


    if (
        selectedProvider ===
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

                        Authorization:
                            "Bearer " +
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

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


        if (
            !response.ok
        ) {

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


    if (
        selectedProvider ===
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

                        Authorization:
                            "Bearer " +
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

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


        if (
            !response.ok
        ) {

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


    if (
        selectedProvider ===
        "gemini"
    ) {

        const cleanModel =
            normalizeGeminiModel(
                model
            );


        const url =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(
                cleanModel
            ) +
            ":generateContent?key=" +
            encodeURIComponent(
                key
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


        if (
            !response.ok
        ) {

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


    if (
        selectedProvider ===
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

                        Authorization:
                            "Bearer " +
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

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


        if (
            !response.ok
        ) {

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
        selectedProvider
    );

}


// =====================================================
// Search box
// البحث الحالي = المحادثات
// =====================================================

if (searchBtn) {

    searchBtn.onclick =
        function (e) {

            e.stopPropagation();


            if (
                !searchPopup
            ) {

                return;

            }


            if (
                projectsPopup
            ) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (
                chatPopup
            ) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (
                settingsWindow
            ) {

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


// ======================================
// Search Input
// ======================================

if (searchInput) {

    searchInput.oninput =
        function () {

            const keyword =
                searchInput.value
                    .trim()
                    .toLowerCase();


            if (
                !searchResults
            ) {

                return;

            }


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

                    if (
                        chat.title &&
                        chat.title
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
                                ${chat.title}
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

                                    searchPopup
                                        .classList
                                        .remove(
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


// =====================================================
// PART 2 compatibility: helper for context
// =====================================================

async function getCurrentStructureForSearch() {

    if (
        !currentDocument
    ) {

        return null;

    }


    return await ensureDocumentStructure(
        currentDocument
    );

}


// =====================================================
// END MAIN FUNCTIONS
// =====================================================


// ======================================
// Initial Render
// ======================================

initializeSidebarSections();

renderProjects();

renderExpandedProjects();

renderDocuments();

renderChatList();

renderSidebarChats();

renderRecentChats();

renderChat();

loadSettings();


// ======================================
// End Office.onReady
// ======================================

});


// ======================================
// Sidebar Pin
// خارج Office.onReady
// ======================================

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


    pinSidebar.addEventListener(
        "click",
        function (e) {

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
