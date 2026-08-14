// ======================================
// Word AI Assistant
// Main Application Controller
// PART 1 / 3
// التخزين + المستندات + الفهرسة
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


let currentProject = null;
let currentDocument = null;


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
        JSON.stringify(documents)
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
// مهم: تغييره يجبر المستندات القديمة
// على إعادة الفهرسة بالخوارزمية الجديدة
// ======================================

const INDEX_SCHEMA_VERSION =
    4;


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
                        !db.objectStoreNames
                            .contains(
                                DOCUMENT_STORE_NAME
                            )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames
                            .contains(
                                DOCUMENT_INDEX_STORE_NAME
                            )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_INDEX_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames
                            .contains(
                                DOCUMENT_TEXT_STORE_NAME
                            )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_TEXT_STORE_NAME
                        );

                    }


                    if (
                        !db.objectStoreNames
                            .contains(
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

function fileToBase64(file) {

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
                        commaIndex === -1
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
            String(documentId),

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
        .filter(function (documentItem) {

            return (
                documentItem &&
                documentItem.projectId ===
                    projectId
            );

        })
        .sort(function (a, b) {

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


// ======================================
// Normalize Search Text
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
// تعريف واحد فقط
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
        matches || []
    );

}



// ======================================
// Conservative Family Key
// مفتاح عائلة الكلمة
// يعتمد على الأصل اللغوي المحافظ
// مع إزالة السوابق واللواحق المحددة فقط
// ======================================

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
        w.length <= 3
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
    // من الأطول إلى الأقصر
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
    // من الأكثر تحديدًا إلى الأقصر
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


    // ==================================
    // حماية الحد الأدنى
    // ==================================

    const MIN_ROOT_LENGTH = 3;


    // ==================================
    // إزالة السوابق
    // يمكن إزالة أكثر من سابقة
    // إذا كانت متتابعة بوضوح
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




// ======================================
// Build Document Index
// النسخة الوحيدة المعتمدة
// فهرسة الكلمات + العائلات + المواضع الذكية
// ======================================

function buildDocumentIndex(
    documentId,
    text
) {

    // ==================================
    // التحقق من البيانات
    // ==================================

    if (!documentId) {

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


    // ==================================
    // استخراج الكلمات الكلية
    // ==================================

    const tokens =
        tokenizeDocumentText(
            text
        );


    // ==================================
    // هياكل الفهرس
    // ==================================

    const terms = {};
    const families = {};


    // ==================================
    // الكلمات الموجودة فعليًا
    // ==================================

    const surfaceSet =
        new Set(
            tokens
        );


    // ==================================
    // تقسيم النص إلى فقرات
    // ==================================
    // نحافظ على ترتيب الفقرات نفسه
    // لاستخدامه لاحقًا في الاسترجاع
    // ==================================

    const paragraphTexts =
        String(text || "").split(
            /\r\n|\r|\n/
        );


    // ==================================
    // فهرسة الكلمات
    // ==================================

    let globalTokenPosition =
        0;


    paragraphTexts.forEach(
        function (
            paragraphText,
            paragraphIndex
        ) {

            const paragraphTokens =
                tokenizeDocumentText(
                    paragraphText
                );


            // ----------------------------------
            // كلمات الفقرة
            // ----------------------------------

            paragraphTokens.forEach(
                function (
                    token,
                    tokenIndex
                ) {

                    const surface =
                        normalizeSearchText(
                            token
                        );


                    if (!surface) {

                        return;

                    }


                    // ==================================
                    // تحديد العائلة
                    // ==================================

                    const familyKey =
                        getConservativeFamilyKey(
                            surface,
                            surfaceSet
                        );


                    // ==================================
                    // terms
                    // ==================================

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
                                familyKey || ""

                        };

                    }


                    terms[surface]
                        .count +=
                        1;


                    // الموضع العام القديم
                    terms[surface]
                        .positions
                        .push(
                            globalTokenPosition
                        );


                    // الموضع الذكي
                    terms[surface]
                        .occurrences
                        .push({

                            paragraphIndex:
                                paragraphIndex,

                            tokenIndex:
                                tokenIndex,

                            globalIndex:
                                globalTokenPosition

                        });


                    // ==================================
                    // لا توجد عائلة
                    // ==================================

                    if (!familyKey) {

                        globalTokenPosition +=
                            1;

                        return;

                    }


                    // ==================================
                    // إنشاء العائلة
                    // ==================================

                    if (
                        !families[familyKey]
                    ) {

                        families[familyKey] = {

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


                    // ==================================
                    // العدد الكلي
                    // ==================================

                    families[familyKey]
                        .count +=
                        1;


                    // ==================================
                    // الموضع العام
                    // ==================================

                    families[familyKey]
                        .positions
                        .push(
                            globalTokenPosition
                        );


                    // ==================================
                    // الموضع الذكي
                    // ==================================

                    families[familyKey]
                        .occurrences
                        .push({

                            paragraphIndex:
                                paragraphIndex,

                            tokenIndex:
                                tokenIndex,

                            globalIndex:
                                globalTokenPosition,

                            word:
                                surface

                        });


                    // ==================================
                    // اللفظ الأصلي
                    // ==================================

                    if (
                        !families[familyKey]
                            .words[surface]
                    ) {

                        families[familyKey]
                            .words[surface] =
                            0;


                        families[familyKey]
                            .uniqueWords +=
                            1;

                    }


                    families[familyKey]
                        .words[surface] +=
                        1;


                    // ==================================
                    // التالي
                    // ==================================

                    globalTokenPosition +=
                        1;

                }
            );

        }
    );


    // ==================================
    // النتيجة النهائية
    // ==================================

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
            new Date().toISOString()

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

    if (!documentId) {

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


    return indexData;

}


// ======================================
// Build Document Structure
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
// Ensure Document Index
// التأكد من أن الفهرس حديث وكامل
// وإعادة بنائه عند الحاجة
// ======================================

async function ensureDocumentIndex(
    documentItem
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    // ==================================
    // قراءة الفهرس الحالي
    // ==================================

    let index =
        await getDocumentIndex(
            documentItem.id
        );


    // ==================================
    // التحقق من أن الفهرس حديث
    // وكامل بالبنية الجديدة
    // ==================================

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


        // ----------------------------------
        // التحقق من بنية المصطلحات
        // ----------------------------------

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


        // ----------------------------------
        // التحقق من بنية العائلات
        // ----------------------------------

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


    // ==================================
    // الفهرس صالح
    // ==================================

    if (
        validIndex
    ) {

        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            index.tokenCount || 0;


        documentItem.indexUniqueTerms =
            index.uniqueTerms || 0;


        documentItem.indexUniqueFamilies =
            index.uniqueFamilies || 0;


        documentItem.indexSchemaVersion =
            index.indexVersion;


        documentItem.indexUpdatedAt =
            index.updatedAt || "";


        saveDocuments();


        return index;

    }


    // ==================================
    // الفهرس قديم أو ناقص
    // ==================================

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


    // ==================================
    // بدء إعادة الفهرسة
    // ==================================

    updateDocumentIndexStatus(
        documentItem,
        "indexing"
    );


    try {

        // ----------------------------------
        // بناء الفهرس الجديد
        // ----------------------------------

        index =
            await rebuildDocumentIndex(
                documentItem.id,
                textData.text
            );


        // ----------------------------------
        // التأكد من بنية المستند
        // ----------------------------------

        await ensureDocumentStructure(
            documentItem
        );


        // ----------------------------------
        // تثبيت بيانات المستند
        // ----------------------------------

        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            index.tokenCount || 0;


        documentItem.indexUniqueTerms =
            index.uniqueTerms || 0;


        documentItem.indexUniqueFamilies =
            index.uniqueFamilies || 0;


        documentItem.indexSchemaVersion =
            INDEX_SCHEMA_VERSION;


        documentItem.indexUpdatedAt =
            index.updatedAt || "";


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
// Read Current Working Document
// القراءة + الفهرسة + البنية في مسار واحد
// ======================================

async function readCurrentWordDocument(
    documentItem
) {

    if (!documentItem) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    // ==================================
    // حالة القراءة
    // ==================================

    updateDocumentReadStatus(
        documentItem,
        "reading"
    );


    try {

        // ==================================
        // التحقق من دعم Word API
        // ==================================

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


        // ==================================
        // استرجاع نسخة العمل
        // ==================================

        const file =
            await getWorkingWordFile(
                documentItem.storageId
            );


        if (!file) {

            throw new Error(
                "لم يتم العثور على نسخة العمل."
            );

        }


        // ==================================
        // تحويل الملف إلى Base64
        // ==================================

        const base64 =
            await fileToBase64(
                file
            );


        // ==================================
        // قراءة النص من مستند Word المخفي
        // ==================================

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


        // ==================================
        // حفظ النص الخام
        // ==================================

        await saveDocumentText(
            documentItem.id,
            text
        );


        // ==================================
        // تثبيت حالة القراءة
        // ==================================

        updateDocumentReadStatus(
            documentItem,
            "read"
        );


        // ==================================
        // بدء الفهرسة
        // ==================================

        updateDocumentIndexStatus(
            documentItem,
            "indexing"
        );


        // ==================================
        // بناء الفهرس الجديد مرة واحدة
        // ==================================

        const indexData =
            await rebuildDocumentIndex(
                documentItem.id,
                text
            );


        // ======================================
        // اختبار بحث عائلة استصلاح
        // ======================================

        const testQueries = [
            "استصلاح",
            "الاستصلاح",
            "بالاستصلاح",
            "استصلاحيا"
        ];

        console.log(
            "======================================"
        );

        console.log(
            "اختبار البحث في عائلة استصلاح"
        );

        console.log(
            "======================================"
        );

        for (
            let i = 0;
            i < testQueries.length;
            i++
        ) {

            const testQuery =
                testQueries[i];

            try {

                const testResult =
                    await searchIndexedDocument(
                        documentItem.id,
                        testQuery
                    );

                console.log(
                    "الاستعلام:",
                    testQuery
                );

                console.log(
                    "عدد النتائج:",
                    testResult.count
                );

                console.log(
                    "العائلات المطابقة:",
                    testResult.matchedFamilies
                );

                console.log(
                    "عدد الظهورات المفهرسة:",
                    testResult.indexedOccurrences
                );

                console.log(
                    "أول 3 نتائج:",
                    testResult.results.slice(
                        0,
                        3
                    )
                );

            }
            catch (error) {

                console.error(
                    "فشل الاستعلام:",
                    testQuery,
                    error
                );

            }

            console.log(
                "--------------------------------------"
            );

        }

        console.log(
            "انتهى اختبار بحث عائلة استصلاح."
        );

        console.log(
            "======================================"
        );

        // ==================================
        // بناء بنية المستند
        // ==================================

        const structureData =
            await buildDocumentStructure(
                documentItem
            );


        await saveDocumentStructure(
            documentItem.id,
            structureData
        );


        // ==================================
        // حفظ إحصاءات الفهرس في المستند
        // ==================================

        documentItem.indexTokenCount =
            indexData.tokenCount || 0;


        documentItem.indexUniqueTerms =
            indexData.uniqueTerms || 0;


        documentItem.indexUniqueFamilies =
            indexData.uniqueFamilies || 0;


        documentItem.indexSchemaVersion =
            INDEX_SCHEMA_VERSION;


        documentItem.indexUpdatedAt =
            indexData.updatedAt ||
            new Date().toISOString();


        // ==================================
        // تثبيت حالة الفهرسة
        // ==================================

        updateDocumentIndexStatus(
            documentItem,
            "indexed"
        );


        // ==================================
        // سجل تشخيصي
        // ==================================

        console.log(
            "تمت قراءة المستند وفهرسته بنجاح:",
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
                    indexData.indexVersion

            }
        );

        // ==================================
        // فحص عائلة استصلاح مباشرة
        // ==================================

        const istislahFamily =
            indexData.families &&
            indexData.families["استصلاح"];

        console.log(
            "======================================"
        );

        console.log(
            "عائلة استصلاح:",
            istislahFamily
        );

        if (istislahFamily) {

            console.log(
                "عدد أفراد العائلة:",
                istislahFamily.count
            );

            console.log(
                "الألفاظ:",
                istislahFamily.words
            );

            console.log(
                "عدد الألفاظ المختلفة:",
                istislahFamily.uniqueWords
            );

            console.log(
                "عدد occurrences:",
                istislahFamily.occurrences
                    ? istislahFamily.occurrences.length
                    : 0
            );

        }
        else {

            console.warn(
                "❌ لم توجد عائلة استصلاح في الفهرس."
            );

        }

        console.log(
            "======================================"
        );

        // ==================================
        // اختبار الفهرس بعد اكتمال الفهرسة
        // ==================================

        testCurrentDocumentIndex();


        return text;

    }
    catch (error) {

        // ==================================
        // فشل القراءة أو الفهرسة
        // ==================================

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


        if (documentTitle) {

            documentTitle.textContent =
                "لا يوجد مستند مفتوح";

        }


        renderDocuments();


        return;

    }


    currentDocument =
        documentItem;


    if (documentTitle) {

        documentTitle.textContent =
            documentItem.name;

    }


    // ==================================
    // إذا كان النص موجودًا مسبقًا
    // لا نعيد قراءة الملف من Word
    // بل نتحقق من نسخة الفهرس
    // ==================================

    if (
        documentItem.readStatus ===
        "read"
    ) {

        ensureDocumentIndex(
            documentItem
        )
        .then(
            function () {

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


    // ==================================
    // المستند جديد
    // ==================================

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
                    projectDocuments.length + 1;


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


            if (index === 0) {

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
                projectDocuments.length - 1
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
                            function (menu) {

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
                                newName !== ""
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


                        if (index <= 0)
                            return;


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
                            projectDocuments.length - 1
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


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


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
                                                newIndex + 1;


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
catch (e) {

    chats = [];

}


let currentChat = null;


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


            if (!targetId)
                return;


            const target =
                document.getElementById(
                    targetId
                );


            if (!target)
                return;


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


                    if (isOpen) {

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

    if (!projectsList)
        return;


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


                    if (projectsPopup) {

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


                        if (!options)
                            return;


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
                            left < margin
                        ) {

                            left =
                                rect.right +
                                margin;

                        }


                        if (
                            left + menuWidth >
                            window.innerWidth -
                            margin
                        ) {

                            left =
                                window.innerWidth -
                                menuWidth -
                                margin;

                        }


                        if (
                            top + menuHeight >
                            window.innerHeight -
                            margin
                        ) {

                            top =
                                rect.top -
                                menuHeight -
                                margin;

                        }


                        if (
                            top < margin
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


                        if (options) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const title =
                            item.querySelector(
                                ".project-title"
                            );


                        if (!title)
                            return;


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


                        if (!edit)
                            return;


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
                                        value !== ""
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


                        if (options) {

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


                        if (confirmDelete) {

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


                                                if (doc) {

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


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


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


                        if (cancelDelete) {

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
                buttonRect.bottom + 8;


            const actualBoxWidth =
                box.offsetWidth ||
                240;


            if (
                left + actualBoxWidth >
                window.innerWidth -
                screenMargin
            ) {

                left =
                    window.innerWidth -
                    actualBoxWidth -
                    screenMargin;

            }


            if (
                left < screenMargin
            ) {

                left =
                    screenMargin;

            }


            if (
                top + boxHeight >
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
                left + "px";

            box.style.top =
                top + "px";

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


            if (isOpening) {

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


    if (!list)
        return;


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


                        renderChat();


                        if (projectsPopup) {

                            projectsPopup
                                .classList
                                .remove(
                                    "open"
                                );

                        }


                        if (chatPopup) {

                            chatPopup
                                .classList
                                .remove(
                                    "open"
                                );

                        }


                        if (searchPopup) {

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


    if (input) {

        input.value =
            "";

        input.style.height =
            "auto";

    }


    if (chatArea) {

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


// ======================================
// Render Chat
// ======================================

function formatAIMessage(
    text
) {

    if (!text)
        return "";


    try {

        return marked.parse(
            text,
            {
                breaks: true,
                gfm: true
            }
        );

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


function renderChat() {

    if (!chatArea)
        return;


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
                    msg.role === "user"
                        ? "user-message"
                        : "ai-message"
                );


            if (
                msg.role ===
                "user"
            ) {

                div.textContent =
                    msg.text || "";

            }
            else {

                div.innerHTML =
                    formatAIMessage(
                        msg.text || ""
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


    if (!list)
        return;


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


            if (title) {

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


            if (menu) {

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


                        if (!options)
                            return;


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


            if (renameBtn) {

                renameBtn.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".chat-options-menu"
                            );


                        if (options) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const titleSpan =
                            item.querySelector(
                                ".chat-title"
                            );


                        if (!titleSpan)
                            return;


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


                        if (!editInput)
                            return;


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
                                        value !== ""
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


            if (deleteBtn) {

                deleteBtn.onclick =
                    function (e) {

                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".chat-options-menu"
                            );


                        if (options) {

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


                        if (confirmDelete) {

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


                        if (cancelDelete) {

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


// ======================================
// Chat Button
// ======================================

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
// Word AI Assistant
// PART 3 / 3
// AI + Search + Send + Initialization
// ======================================


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
    catch (e) {

        return {};

    }

}


// ======================================
// Save AI Settings
// ======================================

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
        "openai"
    ) {

        providerInfo.innerHTML =
            "OpenAI: سيتم جلب النماذج المتاحة من حسابك.";

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


        if (
            savedModel !== ""
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

            // ======================================
            // Temporary Index Test Bridge
            // ======================================

            window.testIstislahIndex = async function () {

                try {

                    if (!currentDocument) {

                        console.warn(
                            "لا يوجد مستند نشط."
                        );

                        return;

                    }

                    const result =
                        await searchIndexedDocument(
                            currentDocument.id,
                            "استصلاح"
                        );

                    console.log(
                        "======================================"
                    );

                    console.log(
                        "اختبار استصلاح"
                    );

                    console.log(
                        "المستند:",
                        currentDocument.name
                    );

                    console.log(
                        "عدد النتائج:",
                        result.count
                    );

                    console.log(
                        "العائلات المطابقة:",
                        result.matchedFamilies
                    );

                    console.log(
                        "عدد الظهورات:",
                        result.indexedOccurrences
                    );

                    console.log(
                        "أول النتائج:",
                        result.results.slice(
                            0,
                            10
                        )
                    );

                    console.log(
                        "======================================"
                    );

                }
                catch (error) {

                    console.error(
                        "فشل اختبار استصلاح:",
                        error
                    );

                }

            };

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
// Show/Hide Key
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
// Save Settings
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


// =====================================================
// AI MODELS
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
        "openai"
    ) {

        await loadOpenAIModels();

        return;

    }


    if (
        selectedProvider ===
        "gemini"
    ) {

        await loadGeminiModels();

        return;

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف."
    );

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

                if (
                    !item ||
                    !item.id ||
                    !item.pricing
                ) {

                    return false;

                }


                return (
                    String(
                        item.pricing.prompt
                    ) === "0" &&
                    String(
                        item.pricing.completion
                    ) === "0"
                );

            }
        );


    populateModels(
        freeModels.map(
            function (
                item
            ) {

                return {

                    id:
                        item.id,

                    name:
                        item.name ||
                        item.id

                };

            }
        )
    );


    if (settingsStatus) {

        settingsStatus.innerHTML =
            "✓ تم تحديث النماذج المجانية: " +
            freeModels.length;

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

                if (
                    !item ||
                    !item.name ||
                    !item.supportedGenerationMethods
                ) {

                    return false;

                }


                return (
                    item.supportedGenerationMethods
                        .includes(
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
        models.length === 0
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


// =====================================================
// AI CONNECTION TEST
// =====================================================

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


        return (
            "✓ تم الاتصال بـ OpenRouter بنجاح"
        );

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


        return (
            "✓ تم الاتصال بـ OpenAI بنجاح"
        );

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


        return (
            "✓ تم الاتصال بـ Gemini بنجاح"
        );

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


    const conversationMessages =
        [];


    if (
        currentChat &&
        Array.isArray(
            currentChat.messages
        )
    ) {

        currentChat.messages.forEach(
            function (
                msg
            ) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                conversationMessages.push({

                    role:
                        msg.role ===
                        "ai"
                            ? "assistant"
                            : "user",

                    content:
                        String(
                            msg.text
                        )

                });

            }
        );

    }


    conversationMessages.push({

        role:
            "user",

        content:
            text

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

                        "Authorization":
                            "Bearer " +
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

                            messages:
                                conversationMessages

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

                        "Authorization":
                            "Bearer " +
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

                            messages:
                                conversationMessages

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


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف: " +
        selectedProvider
    );

}


// ======================================
// Gemini Model Normalization
// ======================================

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


// ======================================
// JSON
// ======================================

async function readJSON(
    response
) {

    try {

        return await response.json();

    }
    catch (e) {

        return {};

    }

}


// ======================================
// API Error
// ======================================

function getAPIError(
    result,
    fallback
) {

    if (
        result &&
        result.error
    ) {

        if (
            typeof result.error ===
            "string"
        ) {

            return result.error;

        }


        if (
            result.error.message
        ) {

            return result.error.message;

        }

    }


    return fallback;

}


// ======================================
// OpenAI/OpenRouter Answer
// ======================================

function extractOpenAIStyleAnswer(
    result,
    providerName
) {

    if (
        result &&
        result.choices &&
        result.choices.length > 0
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


// ======================================
// Gemini Answer
// ======================================

function extractGeminiAnswer(
    result
) {

    if (
        result &&
        result.candidates &&
        result.candidates.length > 0
    ) {

        const candidate =
            result.candidates[0];


        if (
            candidate.content &&
            candidate.content.parts &&
            Array.isArray(
                candidate.content.parts
            )
        ) {

            const textParts =
                candidate.content.parts
                    .filter(
                        function (
                            part
                        ) {

                            return (
                                part &&
                                typeof part.text ===
                                "string"
                            );

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
                textParts.length >
                0
            ) {

                return textParts.join(
                    "\n"
                );

            }

        }

    }


    throw new Error(
        "لم يصل رد صالح من Gemini."
    );

}


// =====================================================
// Search Document Context
// البحث النصي القديم - للإبقاء على الوظيفة
// =====================================================

async function searchDocumentContext(
    documentId,
    query
) {

    const searchTerm =
        normalizeSearchText(
            query
        );


    if (!searchTerm) {

        return {

            query:
                "",

            count:
                0,

            results:
                []

        };

    }


    const structureData =
        await getDocumentStructure(
            documentId
        );


    if (!structureData) {

        throw new Error(
            "لا توجد بنية محفوظة لهذا المستند."
        );

    }


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
            : [];


    const results =
        [];


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                !paragraph ||
                !paragraph.text
            ) {

                return;

            }


            const originalText =
                String(
                    paragraph.text
                );


            const normalizedParagraph =
                normalizeSearchText(
                    originalText
                );


            if (
                !normalizedParagraph.includes(
                    searchTerm
                )
            ) {

                return;

            }


            let nearestHeading =
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
                    heading &&
                    heading.index <
                        paragraph.index
                ) {

                    nearestHeading =
                        heading;

                    break;

                }

            }


            const position =
                normalizedParagraph.indexOf(
                    searchTerm
                );


            const start =
                Math.max(
                    0,
                    position - 100
                );


            const end =
                Math.min(
                    originalText.length,
                    position +
                    searchTerm.length +
                    160
                );


            results.push({

                paragraphIndex:
                    paragraph.index,

                paragraphId:
                    paragraph.id,

                text:
                    originalText,

                context:
                    originalText.substring(
                        start,
                        end
                    ),

                heading:
                    nearestHeading
                        ? nearestHeading.text
                        : "",

                headingLevel:
                    nearestHeading
                        ? nearestHeading.style
                        : ""

            });

        }
    );


    return {

        query:
            searchTerm,

        count:
            results.length,

        results:
            results

    };

}


// =====================================================
// Search Indexed Document
// البحث الذكي بالعائلات والمواضع المباشرة
// =====================================================

async function searchIndexedDocument(
    documentId,
    query
) {

    // ==================================
    // تطبيع الاستعلام
    // ==================================

    const searchTerm =
        normalizeSearchText(
            query
        );

    if (!searchTerm) {

        return {

            query:
                "",

            count:
                0,

            results:
                []

        };

    }


    // ==================================
    // المستند
    // ==================================

    const documentItem =
        documents.find(
            function (doc) {

                return (
                    doc &&
                    String(doc.id) ===
                    String(documentId)
                );

            }
        ) || currentDocument;


    if (!documentItem) {

        throw new Error(
            "لم يتم العثور على المستند."
        );

    }


    // ==================================
    // التأكد من وجود فهرس حديث
    // ==================================

    const indexData =
        await ensureDocumentIndex(
            documentItem
        );


    if (!indexData) {

        throw new Error(
            "لا يوجد فهرس صالح لهذا المستند."
        );

    }


    const indexedTerms =
        indexData.terms || {};


    const indexedFamilies =
        indexData.families || {};


    // ==================================
    // كلمات الاستعلام
    // ==================================

    const queryTokens =
        tokenizeDocumentText(
            searchTerm
        );


    if (
        queryTokens.length ===
        0
    ) {

        return {

            query:
                searchTerm,

            count:
                0,

            results:
                []

        };

    }


    // ==================================
    // تحديد عائلات الاستعلام
    // ==================================

    const indexedSurfaceSet =
        new Set(
            Object.keys(
                indexedTerms
            )
        );


    const queryFamilyKeys =
        [];


    queryTokens.forEach(
        function (
            token
        ) {

            const familyKey =
                getConservativeFamilyKey(
                    token,
                    indexedSurfaceSet
                );


            if (
                familyKey &&
                !queryFamilyKeys.includes(
                    familyKey
                )
            ) {

                queryFamilyKeys.push(
                    familyKey
                );

            }

        }
    );


    // ==================================
    // العائلات الموجودة فعليًا
    // ==================================

    const matchedFamilies =
        queryFamilyKeys.filter(
            function (
                familyKey
            ) {

                return Boolean(
                    indexedFamilies[
                        familyKey
                    ]
                );

            }
        );


    // ==================================
    // الكلمات المطابقة حرفيًا
    // ==================================

    const matchedExactTerms =
        queryTokens.filter(
            function (
                token
            ) {

                return Boolean(
                    indexedTerms[
                        token
                    ]
                );

            }
        );


    // ==================================
    // لا توجد مطابقة
    // ==================================

    if (
        matchedFamilies.length === 0 &&
        matchedExactTerms.length === 0
    ) {

        return {

            query:
                searchTerm,

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            totalQueryTerms:
                queryTokens.length,

            indexTokenCount:
                indexData.tokenCount,

            indexUniqueTerms:
                indexData.uniqueTerms,

            indexUniqueFamilies:
                indexData.uniqueFamilies || 0,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // بنية المستند
    // ==================================

    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (!structureData) {

        throw new Error(
            "لا توجد بنية محفوظة لهذا المستند."
        );

    }


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
            : [];


    // ==================================
    // خريطة الفقرات
    // الوصول المباشر برقم الفقرة
    // ==================================

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
                    paragraph.index,
                    paragraph
                );

            }

        }
    );


    // ==================================
    // الفقرات المرشحة
    // ==================================

    const candidateParagraphs =
        new Map();


    // ==================================
    // إضافة مواضع العائلات
    // ==================================

    matchedFamilies.forEach(
        function (
            familyKey
        ) {

            const family =
                indexedFamilies[
                    familyKey
                ];


            if (!family) {

                return;

            }


            // ----------------------------------
            // النسخة الجديدة
            // occurrences
            // ----------------------------------

            if (
                Array.isArray(
                    family.occurrences
                )
            ) {

                family.occurrences.forEach(
                    function (
                        occurrence
                    ) {

                        if (
                            !occurrence
                        ) {

                            return;

                        }


                        const paragraphIndex =
                            occurrence.paragraphIndex;


                        if (
                            typeof paragraphIndex ===
                            "undefined"
                        ) {

                            return;

                        }


                        if (
                            !candidateParagraphs.has(
                                paragraphIndex
                            )
                        ) {

                            candidateParagraphs.set(
                                paragraphIndex,
                                new Map()
                            );

                        }


                        const paragraphMapEntry =
                            candidateParagraphs.get(
                                paragraphIndex
                            );


                        if (
                            !paragraphMapEntry.has(
                                familyKey
                            )
                        ) {

                            paragraphMapEntry.set(
                                familyKey,
                                []
                            );

                        }


                        paragraphMapEntry
                            .get(
                                familyKey
                            )
                            .push(
                                occurrence
                            );

                    }
                );

                return;

            }


            // ----------------------------------
            // توافق رجعي مع الفهرس القديم
            // ----------------------------------

            if (
                Array.isArray(
                    family.positions
                )
            ) {

                console.warn(
                    "العائلة تستخدم positions القديمة:",
                    familyKey
                );

            }

        }
    );


    // ==================================
    // إضافة الكلمات الحرفية
    // ==================================

    matchedExactTerms.forEach(
        function (
            term
        ) {

            const termData =
                indexedTerms[
                    term
                ];


            if (!termData) {

                return;

            }


            // ----------------------------------
            // النسخة الجديدة
            // ----------------------------------

            if (
                Array.isArray(
                    termData.occurrences
                )
            ) {

                termData.occurrences.forEach(
                    function (
                        occurrence
                    ) {

                        if (
                            !occurrence
                        ) {

                            return;

                        }


                        const paragraphIndex =
                            occurrence.paragraphIndex;


                        if (
                            typeof paragraphIndex ===
                            "undefined"
                        ) {

                            return;

                        }


                        if (
                            !candidateParagraphs.has(
                                paragraphIndex
                            )
                        ) {

                            candidateParagraphs.set(
                                paragraphIndex,
                                new Map()
                            );

                        }


                        const paragraphMapEntry =
                            candidateParagraphs.get(
                                paragraphIndex
                            );


                        if (
                            !paragraphMapEntry.has(
                                term
                            )
                        ) {

                            paragraphMapEntry.set(
                                term,
                                []
                            );

                        }


                        paragraphMapEntry
                            .get(
                                term
                            )
                            .push(
                                occurrence
                            );

                    }
                );

            }

        }
    );


    // ==================================
    // إنشاء النتائج
    // ==================================

    const results =
        [];


    candidateParagraphs.forEach(
        function (
            matchedEntries,
            paragraphIndex
        ) {

            const paragraph =
                paragraphMap.get(
                    paragraphIndex
                );


            if (
                !paragraph ||
                !paragraph.text
            ) {

                return;

            }


            const originalText =
                String(
                    paragraph.text
                );


            const normalizedText =
                normalizeSearchText(
                    originalText
                );


            // ==================================
            // عدد العائلات المطابقة
            // ==================================

            let matchedFamilyCount =
                0;


            matchedFamilies.forEach(
                function (
                    familyKey
                ) {

                    if (
                        matchedEntries.has(
                            familyKey
                        )
                    ) {

                        matchedFamilyCount +=
                            1;

                    }

                }
            );


            // ==================================
            // عدد الكلمات الحرفية المطابقة
            // ==================================

            let exactWordMatches =
                0;


            matchedExactTerms.forEach(
                function (
                    term
                ) {

                    if (
                        matchedEntries.has(
                            term
                        )
                    ) {

                        exactWordMatches +=
                            1;

                    }

                }
            );


            // ==================================
            // الدرجة الأساسية
            // ==================================

            let score =
                0;


            if (
                matchedFamilies.length >
                0
            ) {

                score +=
                    (
                        matchedFamilyCount /
                        matchedFamilies.length
                    ) * 5;

            }


            // ==================================
            // المطابقة الحرفية للعبارة
            // ==================================

            if (
                normalizedText.includes(
                    searchTerm
                )
            ) {

                score +=
                    6;

            }


            // ==================================
            // المطابقة الحرفية للكلمات
            // ==================================

            score +=
                exactWordMatches *
                1.25;


            // ==================================
            // العنوان الأقرب
            // ==================================

            let nearestHeading =
                null;


            for (
                let i =
                    headings.length - 1;
                i >= 0;
                i--
            ) {

                const heading =
                    headings[
                        i
                    ];


                if (
                    heading &&
                    heading.index <
                        paragraphIndex
                ) {

                    nearestHeading =
                        heading;

                    break;

                }

            }


            // ==================================
            // تحديد موضع المقتطف
            // ==================================

            let searchPosition =
                normalizedText.indexOf(
                    searchTerm
                );


            // ----------------------------------
            // إذا لم نجد العبارة حرفيًا
            // نبحث عن أحد ألفاظ العائلة
            // ----------------------------------

            if (
                searchPosition ===
                -1
            ) {

                const familyWords =
                    [];


                matchedFamilies.forEach(
                    function (
                        familyKey
                    ) {

                        const family =
                            indexedFamilies[
                                familyKey
                            ];


                        if (
                            !family ||
                            !family.words
                        ) {

                            return;

                        }


                        Object.keys(
                            family.words
                        )
                        .forEach(
                            function (
                                word
                            ) {

                                familyWords.push(
                                    word
                                );

                            }
                        );

                    }
                );


                for (
                    let i = 0;
                    i < familyWords.length;
                    i++
                ) {

                    const position =
                        normalizedText.indexOf(
                            familyWords[i]
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
            // المقتطف
            // ==================================

            let context =
                originalText;


            if (
                searchPosition !==
                -1
            ) {

                const start =
                    Math.max(
                        0,
                        searchPosition -
                        120
                    );


                const end =
                    Math.min(
                        originalText.length,
                        searchPosition +
                        searchTerm.length +
                        300
                    );


                context =
                    originalText.substring(
                        start,
                        end
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
            // نوع المطابقة
            // ==================================

            let matchType =
                "family";


            if (
                normalizedText.includes(
                    searchTerm
                )
            ) {

                matchType =
                    "exact";

            }
            else if (
                exactWordMatches > 0
            ) {

                matchType =
                    "word";

            }


            // ==================================
            // النتيجة
            // ==================================

            results.push({

                paragraphIndex:
                    paragraphIndex,

                paragraphId:
                    paragraph.id,

                text:
                    originalText,

                context:
                    context,

                heading:
                    nearestHeading
                        ? nearestHeading.text
                        : "",

                headingLevel:
                    nearestHeading
                        ? nearestHeading.style
                        : "",

                matchedTerms:
                    matchedExactTerms,

                matchedFamilies:
                    matchedFamilies,

                matchedFamilyCount:
                    matchedFamilyCount,

                exactWordMatches:
                    exactWordMatches,

                totalQueryTerms:
                    queryTokens.length,

                score:
                    score,

                matchType:
                    matchType

            });

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
    // عدد الظهورات المفهرسة
    // ==================================

    let indexedOccurrences =
        0;


    matchedFamilies.forEach(
        function (
            familyKey
        ) {

            const family =
                indexedFamilies[
                    familyKey
                ];


            if (family) {

                indexedOccurrences +=
                    Number(
                        family.count ||
                        0
                    );

            }

        }
    );


    // ==================================
    // النتيجة النهائية
    // ==================================

    return {

        query:
            searchTerm,

        count:
            results.length,

        results:
            results,

        matchedTerms:
            matchedExactTerms,

        matchedFamilies:
            matchedFamilies,

        totalQueryTerms:
            queryTokens.length,

        indexTokenCount:
            indexData.tokenCount,

        indexUniqueTerms:
            indexData.uniqueTerms,

        indexUniqueFamilies:
            indexData.uniqueFamilies ||
            0,

        indexedOccurrences:
            indexedOccurrences

    };

}


// =====================================================
// Temporary/Debug Test
// يبقى داخل التطبيق، لكنه لا يعتمد على window
// =====================================================

async function testCurrentDocumentIndex() {

    if (!currentDocument) {

        console.warn(
            "لا يوجد مستند نشط."
        );

        return;

    }


    try {

        const index =
            await getDocumentIndex(
                currentDocument.id
            );


        console.log(
            "الفهرس الحقيقي للمستند:",
            index
        );


        if (!index) {

            console.warn(
                "لا يوجد فهرس محفوظ لهذا المستند."
            );


            return;

        }


        console.log(
            "عدد الكلمات:",
            index.tokenCount
        );


        console.log(
            "عدد الكلمات الفريدة:",
            index.uniqueTerms
        );


        console.log(
            "عدد العائلات:",
            index.uniqueFamilies
        );


        console.log(
            "عائلة استصلاح:",
            index.families &&
            index.families["استصلاح"]
        );

        // ======================================
        // اختبار البحث في عائلة استصلاح
        // ======================================

        const testQueries = [
            "استصلاح",
            "الاستصلاح",
            "بالاستصلاح",
            "استصلاحيا"
        ];

        console.log(
            "======================================"
        );

        console.log(
            "اختبار صيغ عائلة استصلاح"
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
                        query
                    );

                console.log(
                    "الاستعلام:",
                    query
                );

                console.log(
                    "العائلة المطابقة:",
                    result.matchedFamilies
                );

                console.log(
                    "عدد الفقرات:",
                    result.count
                );

                console.log(
                    "عدد الظهورات:",
                    result.indexedOccurrences
                );

            }
            catch (error) {

                console.error(
                    "فشل البحث عن:",
                    query,
                    error
                );

            }

        }


    }
    catch (error) {

        console.error(
            "فشل اختبار الفهرس:",
            error
        );

    }

}


// =====================================================
// Search box
// ملاحظة: البحث الحالي في أسماء المحادثات
// ولا نغيّره في هذه المرحلة حتى لا نعبث بالواجهة
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
                )
            ) {

                if (searchInput) {

                    searchInput.focus();

                }

            }

        };

}


// ======================================
// Search Input
// البحث الحالي = المحادثات
// ======================================

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

                    if (
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
// Send Message
// =====================================================

async function sendMessage() {

    if (!input)
        return;


    const text =
        input.value.trim();


    if (
        text ===
        ""
    ) {

        return;

    }


    // ==================================
    // إنشاء محادثة
    // ==================================

    if (!currentChat) {

        currentChat = {

            id:
                Date.now(),

            title:
                text.substring(
                    0,
                    30
                ),

            messages:
                [],

            isTemporary:
                true,

            projectId:
                currentProject
                    ? currentProject.id
                    : null

        };

    }


    // ==================================
    // حفظ المحادثة المؤقتة
    // ==================================

    if (
        currentChat.isTemporary
    ) {

        currentChat.isTemporary =
            false;


        currentChat.projectId =
            currentProject
                ? currentProject.id
                : null;


        currentChat.title =
            text.substring(
                0,
                30
            );


        const alreadyExists =
            chats.some(
                function (
                    chat
                ) {

                    return (
                        chat.id ===
                        currentChat.id
                    );

                }
            );


        if (!alreadyExists) {

            chats.unshift(
                currentChat
            );

        }


        if (
            currentProject &&
            currentChat.projectId ===
                currentProject.id
        ) {

            if (
                !Array.isArray(
                    currentProject.chatIds
                )
            ) {

                currentProject.chatIds =
                    [];

            }


            if (
                !currentProject.chatIds.includes(
                    currentChat.id
                )
            ) {

                currentProject.chatIds.push(
                    currentChat.id
                );


                currentProject.updatedAt =
                    new Date()
                        .toISOString();


                saveProjects();

            }

        }


        saveChats();

    }


    // ==================================
    // User Message
    // ==================================

    currentChat.messages.push({

        role:
            "user",

        text:
            text

    });


    saveChats();


    renderChat();


    renderChatList();


    renderSidebarChats();


    renderRecentChats();


    input.value =
        "";


    input.style.height =
        "auto";


    // ==================================
    // Loading
    // ==================================

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


    // ==================================
    // AI
    // ==================================

    try {

        const answer =
            await askAI(
                text
            );


        if (loading) {

            loading.remove();

        }


        currentChat.messages.push({

            role:
                "ai",

            text:
                answer || ""

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();

    }
    catch (error) {

        if (loading) {

            loading.remove();

        }


        currentChat.messages.push({

            role:
                "ai",

            text:
                "خطأ: " +
                (
                    error.message ||
                    "حدث خطأ غير معروف"
                )

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();

    }

}


// ======================================
// Send Button
// ======================================

if (sendBtn) {

    sendBtn.onclick =
        function (e) {

            e.preventDefault();


            sendMessage();

        };

}


// ======================================
// Keyboard
// ======================================

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
// خارج Office.onReady كما كان
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