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
// =====================================================
// تثبيت موضع زر الشريط الجانبي الأصلي
// =====================================================

const sidebarToggleOriginalParent =
    sidebarToggleBtn
        ? sidebarToggleBtn.parentNode
        : null;
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

    if (!project) {

        currentProject =
            null;


        renderProjects();


        return;

    }


    currentProject =
        project;


    renderProjects();

}


// ======================================
// Add Document
// ======================================

if (
    addDocumentBtn &&
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
            currentProject
                ? currentProject.id
                : null,

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


    currentChat.projectId =
        currentProject
            ? currentProject.id
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


    renderExpandedProjects();


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
                currentProject
                    ? currentProject.id
                    : null,

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

    renderChatList();

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

        renderChatList();

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

        renderChatList();

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
// الحفاظ على سلوك صندوق الإدخال
// =====================================================

if (input) {

    input.oninput =
        function () {

            input.style.height =
                "auto";


            input.style.height =
                Math.min(
                    input.scrollHeight,
                    180
                ) +
                "px";

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
// Initialize Sidebar Sections
// =====================================================

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
        // زر إضافة مستند
        // =========================================

        addDocumentButton.onclick =
            function (
                e
            ) {

                e.preventDefault();
                e.stopPropagation();


                // هذا المستند سيضاف إلى هذا المشروع
                if (
                    !project
                ) {

                    return;

                }


                // جعل المشروع الحالي هو المشروع الذي ضغط المستخدم
                setCurrentProject(
                    project
                );


                // استخدام أداة اختيار ملف Word الموجودة أصلًا
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


                // فتح نافذة اختيار ملف Word
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
        //
        // صندوق الإضافة يبقى ظاهرًا
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
                // حالة التفعيل
                // =========================================

                const isActive =
                    Boolean(
                        currentDocument &&
                        String(
                            currentDocument.id
                        ) ===
                        String(
                            documentItem.id
                        )
                    );


                if (
                    isActive
                ) {

                    item.classList.add(
                        "active-document"
                    );

                }


                // =========================================
                // العنوان
                // =========================================

                const title =
                    document.createElement(
                        "span"
                    );


                title.className =
                    "project-document-title";


                const activeMark =
                    document.createElement(
                        "span"
                    );


                activeMark.className =
                    "document-active-mark";


                activeMark.textContent =
                    isActive
                        ? "✓ "
                        : "";


                const titleText =
                    document.createElement(
                        "span"
                    );


                titleText.className =
                    "document-name";


                titleText.textContent =
                    documentItem.name;


                title.appendChild(
                    activeMark
                );


                title.appendChild(
                    titleText
                );


                // =========================================
                // تفعيل / إلغاء التفعيل
                // =========================================

                title.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();

                        e.stopPropagation();


                        const sameDocument =
                            Boolean(
                                currentDocument &&
                                String(
                                    currentDocument.id
                                ) ===
                                String(
                                    documentItem.id
                                )
                            );


                        if (
                            sameDocument
                        ) {

                            setCurrentDocument(
                                null
                            );

                        }
                        else {

                            setCurrentDocument(
                                documentItem
                            );

                        }


                        renderProjectDocuments(
                            container,
                            project
                        );

                    };


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


                                        if (
                                            currentDocument &&
                                            String(
                                                currentDocument.id
                                            ) ===
                                            String(
                                                documentItem.id
                                            ) &&
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                documentItem.name;

                                        }


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
                                    ".project-document-delete-confirm"
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
                                "project-document-delete-confirm";


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
    // بناء المشروع
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


                        setCurrentProject(
                            project
                        );

                    }

                };


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

                    const addDocumentButton =
                        section.querySelector(
                            ".project-add-document"
                        );



                    // =================================================
                    // زر إضافة مستند
                    // =================================================

                    if (
                        addDocumentButton &&
                        sectionIndex === 0
                    ) {

                        addDocumentButton.onclick =
                            function (
                                e
                            ) {

                                e.preventDefault();

                                e.stopPropagation();


                                // هذا المشروع هو المشروع المستهدف
                                setCurrentProject(
                                    project
                                );


                                if (
                                    !wordDocumentPicker
                                ) {

                                    console.warn(
                                        "لم يتم العثور على أداة اختيار ملف Word."
                                    );

                                    return;

                                }


                                wordDocumentPicker.value =
                                    "";


                                wordDocumentPicker.click();

                            };

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


                                arrow.textContent =
                                    "›";


                                return;

                            }


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


                            arrow.textContent =
                                "⌄";


                            if (
                                sectionIndex ===
                                0
                            ) {

                                renderProjectDocuments(
                                    content,
                                    project
                                );

                            }
                            else if (
                                sectionIndex ===
                                1
                            ) {

                                renderProjectChats(
                                    content,
                                    project
                                );

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

                                openMenu.classList.remove(
                                    "open"
                                );

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


            // =================================================
            // إعادة التسمية
            // =================================================

            const renameProject =
                options.querySelector(
                    ".rename-project"
                );


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

                                renderExpandedProjects();

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


            // =================================================
            // حذف المشروع
            // =================================================

            const deleteProject =
                options.querySelector(
                    ".delete-project"
                );


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


                    confirmDelete.onclick =
                        async function () {

                            const projectDocumentIds =
                                Array.isArray(
                                    project.documents
                                )
                                    ? project.documents
                                    : [];


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


                            oramaRetrievalDb =
                                null;


                            oramaRetrievalCacheKey =
                                "";


                            oramaRetrievalDocumentId =
                                null;


                            saveDocuments();

                            saveChats();

                            saveProjects();


                            confirmBox.remove();


                            renderProjects();

                            renderExpandedProjects();

                

                            renderChatList();

                            renderSidebarChats();

                            renderRecentChats();

                        };


                    const cancelDelete =
                        confirmBox.querySelector(
                            ".cancel-project-delete"
                        );


                    cancelDelete.onclick =
                        function () {

                            confirmBox.remove();

                        };

                };


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
// Sidebar Toggle
// مستقل تمامًا عن المستندات
// =====================================================

if (
    sidebarToggleBtn &&
    expandedSidebar &&
    expandedSidebarToggleSlot
) {

    sidebarToggleBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();

            e.stopPropagation();


            const isOpening =
                !expandedSidebar.classList.contains(
                    "open"
                );


            // ==================================
            // فتح الشريط
            // ==================================

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


                // نقل الزر إلى مكانه داخل
                // الشريط الموسع
                if (
                    sidebarToggleBtn.parentNode !==
                    expandedSidebarToggleSlot
                ) {

                    expandedSidebarToggleSlot.appendChild(
                        sidebarToggleBtn
                    );

                }


                return;

            }


            // ==================================
            // إغلاق الشريط
            // ==================================

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


            // ==================================
            // إعادة الزر إلى موضعه الأصلي
            // ==================================

            if (
                sidebarToggleOriginalParent &&
                sidebarTogglePlaceholder
            ) {

                sidebarToggleOriginalParent.insertBefore(
                    sidebarToggleBtn,
                    sidebarTogglePlaceholder.nextSibling
                );

            }

        };

}


// =====================================================
// Render Projects In Expanded Sidebar
// =====================================================

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
// Render Chat List
// =====================================================

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
                            ) +
                            "px";


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
                                        value !==
                                            ""
                                            ? value
                                            : oldName;


                                    chat.updatedAt =
                                        new Date()
                                            .toISOString();


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


                                        currentCitationSources =
                                            [];


                                        renderChat();

                                    }


                                    if (
                                        currentProject &&
                                        Array.isArray(
                                            currentProject.chatIds
                                        )
                                    ) {

                                        currentProject.chatIds =
                                            currentProject.chatIds.filter(
                                                function (
                                                    id
                                                ) {

                                                    return (
                                                        id !==
                                                        chat.id
                                                    );

                                                }
                                            );


                                        currentProject.updatedAt =
                                            new Date()
                                                .toISOString();


                                        saveProjects();

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


            renderChatList();


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

initializeSidebarSections();

renderProjects();

renderExpandedProjects();



renderChatList();

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