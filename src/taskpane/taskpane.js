// ======================================
// Word AI Assistant
// Main Application Controller
// PART 1 / 4
// التخزين + المشاريع الأساسية + المستندات
// + قراءة المستند + بنية المستند
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
// Citation Sources
// مصادر الإحالات الحالية
// ======================================

let currentCitationSources = [];


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

    console.warn(
        "تعذر قراءة المشاريع المحفوظة:",
        error
    );

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
                    typeof project === "object"
                );

            }
        )
        .map(
            function (
                project
            ) {

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

                                citationStyle:
                                    "",

                                notes:
                                    ""

                            }

                };

            }
        );


// ======================================
// Current Project / Document
// ======================================

let currentProject =
    null;

let currentDocument =
    null;


// ======================================
// Orama Retrieval Cache
// ذاكرة محرك Orama
// ======================================

let oramaRetrievalDb =
    null;

let oramaRetrievalCacheKey =
    "";


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
                        typeof documentItem === "object"
                    );

                }
            )
            .map(
                function (
                    documentItem
                ) {

                    // ------------------------------
                    // إحصاءات الفهرسة
                    // ------------------------------

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


                    if (
                        typeof documentItem.indexSchemaVersion !==
                        "number"
                    ) {

                        documentItem.indexSchemaVersion =
                            0;

                    }


                    // ------------------------------
                    // حالات المستند
                    // ------------------------------

                    if (
                        !documentItem.readStatus
                    ) {

                        documentItem.readStatus =
                            "new";

                    }


                    if (
                        !documentItem.indexStatus
                    ) {

                        documentItem.indexStatus =
                            "new";

                    }


                    return documentItem;

                }
            );

}
catch (error) {

    console.warn(
        "تعذر قراءة المستندات المحفوظة:",
        error
    );

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
// IndexedDB
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
// Orama Version
// مهم: تغييرها يجبر Orama على إعادة
// بناء فهرس الذاكرة
// ======================================

const ORAMA_RETRIEVAL_VERSION =
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


                    // ------------------------------
                    // ملفات Word
                    // ------------------------------

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


                    // ------------------------------
                    // التخزين القديم للفهرس
                    // يبقى للتوافق مع البيانات السابقة
                    // ولا يمثل محرك البحث الجديد
                    // ------------------------------

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


                    // ------------------------------
                    // نصوص المستندات
                    // ------------------------------

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


                    // ------------------------------
                    // بنية المستند
                    // ------------------------------

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

                    const db =
                        request.result;


                    // عند ترقية قاعدة البيانات لاحقًا
                    // لا نريد أن تتسبب نسخة اتصال قديمة
                    // في مشاكل blocked
                    db.onversionchange =
                        function () {

                            db.close();

                        };


                    resolve(
                        db
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error ||
                        new Error(
                            "فشل فتح قاعدة بيانات المستندات."
                        )
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
                    String(
                        fileId
                    )
                );


            request.onsuccess =
                function () {

                    resolve(
                        String(
                            fileId
                        )
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error ||
                        new Error(
                            "فشل حفظ نسخة العمل."
                        )
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };


            transaction.onerror =
                function () {

                    reject(
                        transaction.error ||
                        new Error(
                            "فشل حفظ نسخة العمل."
                        )
                    );

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
                    String(
                        fileId
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
                        request.error ||
                        new Error(
                            "فشل قراءة نسخة العمل."
                        )
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
                    String(
                        fileId
                    )
                );


            request.onsuccess =
                function () {

                    resolve();

                };


            request.onerror =
                function () {

                    reject(
                        request.error ||
                        new Error(
                            "فشل حذف نسخة العمل."
                        )
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
            ORAMA_RETRIEVAL_VERSION,

        indexUpdatedAt:
            "",

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
        new Date().toISOString();


    if (
        status === "read"
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

    if (
        !documentItem
    ) {

        return;

    }


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
                    Number(
                        a.order || 0
                    ) -
                    Number(
                        b.order || 0
                    )
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
                        request.error ||
                        new Error(
                            "فشل حفظ نص المستند."
                        )
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
                        request.error ||
                        new Error(
                            "فشل قراءة نص المستند."
                        )
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
                        request.error ||
                        new Error(
                            "فشل حفظ بنية المستند."
                        )
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
                        request.error ||
                        new Error(
                            "فشل قراءة بنية المستند."
                        )
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
// Build Document Structure
// استخراج الفقرات والعناوين والجداول
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

                        return {

                            index:
                                index,

                            id:
                                String(
                                    index
                                ),

                            text:
                                String(
                                    paragraph.text ||
                                    ""
                                ).trim(),

                            style:
                                String(
                                    paragraph.styleBuiltIn ||
                                    ""
                                ),

                            tableNestingLevel:
                                Number(
                                    paragraph.tableNestingLevel ||
                                    0
                                )

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
                                Number(
                                    table.rowCount ||
                                    0
                                ),

                            columns:
                                Number(
                                    table.columnCount ||
                                    0
                                ),

                            style:
                                String(
                                    table.styleBuiltIn ||
                                    ""
                                )

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
// Ensure Document Structure
// ======================================

async function ensureDocumentStructure(
    documentItem
) {

    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


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
// القراءة + حفظ النص + بناء البنية
// ======================================

async function readCurrentWordDocument(
    documentItem
) {

    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم تحديد المستند."
        );

    }


    updateDocumentReadStatus(
        documentItem,
        "reading"
    );


    try {

        // ------------------------------
        // التحقق من دعم Word API
        // ------------------------------

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


        // ------------------------------
        // استرجاع نسخة العمل
        // ------------------------------

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


        // ------------------------------
        // تحويل الملف إلى Base64
        // ------------------------------

        const base64 =
            await fileToBase64(
                file
            );


        // ------------------------------
        // قراءة النص
        // ------------------------------

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


                    return String(
                        body.text ||
                        ""
                    );

                }
            );


        // ------------------------------
        // حفظ النص
        // ------------------------------

        await saveDocumentText(
            documentItem.id,
            text
        );


        // ------------------------------
        // تثبيت حالة القراءة
        // ------------------------------

        updateDocumentReadStatus(
            documentItem,
            "read"
        );


        // ------------------------------
        // بناء بنية المستند
        // ------------------------------

        const structureData =
            await buildDocumentStructure(
                documentItem
            );


        await saveDocumentStructure(
            documentItem.id,
            structureData
        );


        // ------------------------------
        // إحصاءات أولية
        // ------------------------------
        // هذه الإحصاءات مؤقتة حتى يبني Orama
        // فهرس الاسترجاع في الجزء التالي.

        documentItem.indexTokenCount =
            tokenizeDocumentText(
                text
            ).length;


        documentItem.indexUniqueTerms =
            new Set(
                tokenizeDocumentText(
                    text
                )
            ).size;


        documentItem.indexUniqueFamilies =
            0;


        documentItem.indexSchemaVersion =
            ORAMA_RETRIEVAL_VERSION;


        documentItem.indexUpdatedAt =
            new Date().toISOString();


        updateDocumentIndexStatus(
            documentItem,
            "indexed"
        );


        saveDocuments();


        console.log(
            "تمت قراءة المستند بنجاح:",
            {
                documentId:
                    documentItem.id,

                tokenCount:
                    documentItem.indexTokenCount,

                uniqueTerms:
                    documentItem.indexUniqueTerms,

                headings:
                    structureData.headingCount,

                paragraphs:
                    structureData.paragraphCount,

                tables:
                    structureData.tableCount

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
            "فشل قراءة المستند:",
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

    if (
        !documentItem
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


        if (
            typeof renderDocuments ===
            "function"
        ) {

            renderDocuments();

        }


        return;

    }


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


    // ------------------------------
    // إذا كانت النسخة قد قرئت
    // نتحقق من البنية فقط
    // ------------------------------

    if (
        documentItem.readStatus ===
        "read"
    ) {

        ensureDocumentStructure(
            documentItem
        )
            .then(
                function () {

                    if (
                        typeof renderDocuments ===
                        "function"
                    ) {

                        renderDocuments();

                    }

                }
            )
            .catch(
                function (
                    error
                ) {

                    console.error(
                        "تعذر تحديث بنية المستند:",
                        error
                    );


                    if (
                        typeof renderDocuments ===
                        "function"
                    ) {

                        renderDocuments();

                    }

                }
            );


        if (
            typeof renderDocuments ===
            "function"
        ) {

            renderDocuments();

        }


        return;

    }


    // ------------------------------
    // المستند جديد
    // ------------------------------

    readCurrentWordDocument(
        documentItem
    )
        .then(
            function () {

                if (
                    typeof renderDocuments ===
                    "function"
                ) {

                    renderDocuments();

                }

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


                if (
                    typeof renderDocuments ===
                    "function"
                ) {

                    renderDocuments();

                }

            }
        );


    if (
        typeof renderDocuments ===
        "function"
    ) {

        renderDocuments();

    }

}


// ======================================
// Update Document Timestamp
// ======================================

function touchDocument(
    documentItem
) {

    if (
        !documentItem
    ) {

        return;

    }


    documentItem.updatedAt =
        new Date().toISOString();


    saveDocuments();

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

    if (
        !project
    ) {

        currentProject =
            null;


        currentDocument =
            null;


        currentCitationSources =
            [];


        if (
            typeof renderDocuments ===
            "function"
        ) {

            renderDocuments();

        }


        return;

    }


    currentProject =
        project;


    if (
        currentDocument &&
        currentDocument.projectId !==
            project.id
    ) {

        currentDocument =
            null;

        currentCitationSources =
            [];

    }


    if (
        typeof renderDocuments ===
        "function"
    ) {

        renderDocuments();

    }

}


// ======================================
// Add Document
// ======================================

if (
    addDocumentBtn &&
    wordDocumentPicker
) {

    addDocumentBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (
                !currentProject
            ) {

                if (
                    documentsList
                ) {

                    documentsList.innerHTML =
                        `
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


                if (
                    !file
                ) {

                    return;

                }


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


                if (
                    !currentProject
                ) {

                    return;

                }


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


                if (
                    typeof renderDocuments ===
                    "function"
                ) {

                    renderDocuments();

                }


                console.log(
                    "تم استيراد مستند Word:",
                    {

                        name:
                            documentItem.name,

                        fileName:
                            documentItem.fileName,

                        storageId:
                            documentItem.storageId,

                        documentId:
                            documentItem.id

                    }
                );

            }
            catch (
                error
            ) {

                console.error(
                    "فشل استيراد مستند Word:",
                    error
                );


                if (
                    documentsList
                ) {

                    documentsList.innerHTML =
                        `
                            <div class="empty-document">
                                تعذر استيراد المستند
                            </div>
                        `;

                }

            }

        };

}


// ======================================
// PART 1 END
// ======================================

});
// ======================================
// Word AI Assistant
// PART 2 / 4
// محرك البحث الجديد Orama
// بناء الفهرس + البحث + العناوين
// + نطاقات المطالب + سياق الاسترجاع
// ======================================


// =====================================================
// Orama Loader
// تحميل Orama مرة واحدة
// =====================================================

function getOramaEngine() {

    if (
        typeof require !==
        "function"
    ) {

        throw new Error(
            "Webpack require غير متاح في هذه البيئة."
        );

    }


    return require(
        "@orama/orama"
    );

}


// =====================================================
// Get Orama Components
// =====================================================

function getOramaComponents() {

    try {

        return require(
            "@orama/orama/components"
        );

    }
    catch (error) {

        console.warn(
            "تعذر تحميل مكونات Orama:",
            error
        );

        return {};

    }

}


// =====================================================
// Normalize Orama Query
// تطبيع سؤال البحث قبل إرساله إلى Orama
// =====================================================

function normalizeOramaQuery(
    query
) {

    return normalizeSearchText(
        String(
            query ||
            ""
        )
    )
        .replace(
            /[،,؛;؟?!:()[\]{}"«»]/g,
            " "
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim();

}


// =====================================================
// Get Heading Level Number
// Heading1 => 1
// Heading2 => 2
// ...
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


    if (
        match
    ) {

        return Number(
            match[1]
        );

    }


    return 9;

}


// =====================================================
// Sort Headings
// =====================================================

function sortHeadings(
    headings
) {

    return (
        Array.isArray(
            headings
        )
            ? headings
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
            : []
    );

}


// =====================================================
// Find Heading For Paragraph
// العثور على العنوان الأقرب للفقرة
// =====================================================

function findNearestHeading(
    paragraphIndex,
    headings
) {

    let nearest =
        null;


    const sorted =
        sortHeadings(
            headings
        );


    for (
        let i =
            sorted.length - 1;

        i >= 0;

        i--
    ) {

        const heading =
            sorted[i];


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


// =====================================================
// Build Orama Retrieval Index
// الفهرس الرئيسي الجديد
// =====================================================

async function ensureOramaRetrievalIndex(
    documentItem,
    structureData
) {

    if (
        !documentItem
    ) {

        throw new Error(
            "لم يتم تحديد المستند لبناء فهرس Orama."
        );

    }


    if (
        !structureData
    ) {

        throw new Error(
            "لا توجد بنية للمستند لبناء فهرس Orama."
        );

    }


    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    const headings =
        sortHeadings(
            structureData.headings
        );


    // ==========================================
    // مفتاح ذاكرة الفهرس
    // ==========================================

    const cacheKey =
        [

            String(
                documentItem.id
            ),

            String(
                documentItem.indexUpdatedAt ||
                structureData.updatedAt ||
                ""
            ),

            String(
                ORAMA_RETRIEVAL_VERSION
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


    // ==========================================
    // استخدام الفهرس الموجود في الذاكرة
    // ==========================================

    if (
        oramaRetrievalDb &&
        oramaRetrievalCacheKey ===
            cacheKey
    ) {

        return oramaRetrievalDb;

    }


    const {
        create,
        insertMultiple
    } =
        getOramaEngine();


    // ==========================================
    // إنشاء قاعدة Orama
    // ==========================================

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
                    "string",

                isHeading:
                    "boolean"

            },

            language:
                "arabic"

        });


    // ==========================================
    // بناء السجلات
    // ==========================================

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
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            if (
                !text
            ) {

                return;

            }


            const paragraphIndex =
                Number(
                    paragraph.index
                );


            const heading =
                findNearestHeading(
                    paragraphIndex,
                    headings
                );


            const isHeading =
                Boolean(
                    heading &&
                    Number(
                        heading.index
                    ) ===
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

                text:
                    text,

                heading:
                    heading
                        ? String(
                            heading.text ||
                            ""
                        )
                            .replace(
                                /\s+/g,
                                " "
                            )
                            .trim()
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
                    isHeading

            });

        }
    );


    // ==========================================
    // إدخال البيانات في Orama
    // ==========================================

    if (
        records.length >
        0
    ) {

        await insertMultiple(
            db,
            records
        );

    }


    // ==========================================
    // حفظ الذاكرة
    // ==========================================

    oramaRetrievalDb =
        db;


    oramaRetrievalCacheKey =
        cacheKey;


    oramaDocumentDb =
        db;


    oramaDocumentId =
        String(
            documentItem.id
        );


    console.log(
        "تم بناء فهرس Orama للاسترجاع:",
        records.length,
        "فقرة"
    );


    console.log(
        "عدد العناوين:",
        headings.length
    );


    return db;

}


// =====================================================
// Get Orama Index For Document
// التأكد من وجود الفهرس الصحيح للمستند
// =====================================================

async function getOramaDocumentIndex(
    documentItem
) {

    if (
        !documentItem
    ) {

        return null;

    }


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData
    ) {

        return null;

    }


    return await ensureOramaRetrievalIndex(
        documentItem,
        structureData
    );

}


// =====================================================
// Calculate Heading Relevance
// حساب قوة العنوان بالنسبة للاستعلام
// =====================================================

function calculateHeadingRelevance(
    heading,
    queryTokens,
    normalizedQuery
) {

    if (
        !heading
    ) {

        return 0;

    }


    const headingText =
        normalizeSearchText(
            heading
        );


    if (
        !headingText
    ) {

        return 0;

    }


    const tokens =
        Array.isArray(
            queryTokens
        )
            ? queryTokens
            : [];


    let matched =
        0;


    tokens.forEach(
        function (
            token
        ) {

            if (
                token &&
                headingText.includes(
                    token
                )
            ) {

                matched +=
                    1;

            }

        }
    );


    const coverage =
        tokens.length >
        0
            ? matched /
              tokens.length
            : 0;


    let score =
        0;


    // ==========================================
    // تطابق العنوان كاملًا
    // ==========================================

    if (
        headingText ===
        normalizedQuery
    ) {

        score +=
            100;

    }


    // ==========================================
    // العنوان يحتوي على السؤال
    // ==========================================

    if (
        normalizedQuery &&
        headingText.includes(
            normalizedQuery
        )
    ) {

        score +=
            60;

    }


    // ==========================================
    // عدد الكلمات المطابقة
    // ==========================================

    score +=
        matched *
        15;


    // ==========================================
    // تغطية السؤال
    // ==========================================

    score +=
        coverage *
        35;


    return score;

}


// =====================================================
// Build Orama Search Query
// =====================================================

function buildOramaSearchQuery(
    query,
    queryTokens
) {

    const normalizedQuery =
        normalizeOramaQuery(
            query
        );


    const tokens =
        Array.isArray(
            queryTokens
        )
            ? queryTokens.filter(
                function (
                    token
                ) {

                    return (
                        token &&
                        token.length >=
                        3
                    );

                }
            )
            : [];


    // ==========================================
    // في حال وجود سؤال مركب
    // نبحث بالعبارة وكلماتها
    // ==========================================

    const terms =
        [];


    if (
        normalizedQuery
    ) {

        terms.push(
            normalizedQuery
        );

    }


    tokens.forEach(
        function (
            token
        ) {

            if (
                !terms.includes(
                    token
                )
            ) {

                terms.push(
                    token
                );

            }

        }
    );


    return terms;

}


// =====================================================
// Search Orama Document
// المحرك الفعلي الجديد
// =====================================================

async function searchOramaDocument(
    documentItem,
    query,
    options
) {

    if (
        !documentItem
    ) {

        return {

            query:
                String(
                    query ||
                    ""
                ),

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            profile:
                "general"

        };

    }


    const normalizedQuery =
        normalizeOramaQuery(
            query
        );


    if (
        !normalizedQuery
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

            profile:
                "general"

        };

    }


    const settings =
        options ||
        {};


    const profile =
        settings.profile ||
        "general";


    const maxResults =
        typeof settings.maxResults ===
            "number"
                ? Math.max(
                    1,
                    settings.maxResults
                )
                : 10;


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData
    ) {

        return {

            query:
                normalizedQuery,

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            profile:
                profile

        };

    }


    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    if (
        !db
    ) {

        return {

            query:
                normalizedQuery,

            count:
                0,

            results:
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            profile:
                profile

        };

    }


    const {
        search
    } =
        getOramaEngine();


    // ==========================================
    // كلمات الاستعلام
    // ==========================================

    const queryTokens =
        getSearchQueryTokens(
            normalizedQuery
        );


    const searchTerms =
        buildOramaSearchQuery(
            normalizedQuery,
            queryTokens
        );


    // ==========================================
    // تنفيذ البحث الأساسي
    // ==========================================

    let hits =
        [];


    // أولًا: البحث بالعبارة كاملة
    // ==========================================

    try {

        const fullResult =
            await search(
                db,
                {

                    term:
                        normalizedQuery,

                    properties:
                        [
                            "text",
                            "heading"
                        ],

                    boost:
                        {
                            heading:
                                6
                        },

                    tolerance:
                        2,

                    limit:
                        Math.max(
                            20,
                            maxResults *
                            4
                        )

                }
            );


        if (
            fullResult &&
            Array.isArray(
                fullResult.hits
            )
        ) {

            hits =
                fullResult.hits;

        }

    }
    catch (error) {

        console.warn(
            "فشل البحث بالعبارة الكاملة في Orama:",
            error
        );

    }


    // ==========================================
    // البحث الاحتياطي بكلمات الاستعلام
    // ==========================================

    if (
        searchTerms.length >
        1
    ) {

        for (
            let i = 0;
            i < searchTerms.length;
            i++
        ) {

            const term =
                searchTerms[i];


            if (
                !term ||
                term ===
                    normalizedQuery
            ) {

                continue;

            }


            try {

                const termResult =
                    await search(
                        db,
                        {

                            term:
                                term,

                            properties:
                                [
                                    "text",
                                    "heading"
                                ],

                            boost:
                                {
                                    heading:
                                        6
                                },

                            tolerance:
                                2,

                            limit:
                                Math.max(
                                    20,
                                    maxResults *
                                    3
                                )

                        }
                    );


                if (
                    termResult &&
                    Array.isArray(
                        termResult.hits
                    )
                ) {

                    hits =
                        hits.concat(
                            termResult.hits
                        );

                }

            }
            catch (
                error
            ) {

                console.warn(
                    "فشل البحث الجزئي في Orama:",
                    term,
                    error
                );

            }

        }

    }


    // ==========================================
    // إزالة التكرار
    // ==========================================

    const uniqueHits =
        new Map();


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


            const document =
                hit.document;


            const id =
                String(
                    document.id ||
                    (
                        "p-" +
                        document.paragraphIndex
                    )
                );


            const existing =
                uniqueHits.get(
                    id
                );


            if (
                !existing ||
                Number(
                    hit.score ||
                    0
                ) >
                Number(
                    existing.score ||
                    0
                )
            ) {

                uniqueHits.set(
                    id,
                    hit
                );

            }

        }
    );


    // ==========================================
    // إعداد بيانات العناوين
    // ==========================================

    const headings =
        sortHeadings(
            structureData.headings
        );


    // ==========================================
    // حساب قوة النتائج
    // ==========================================

    const scoredResults =
        [];


    uniqueHits.forEach(
        function (
            hit
        ) {

            const doc =
                hit.document;


            const paragraphIndex =
                Number(
                    doc.paragraphIndex
                );


            const heading =
                String(
                    doc.heading ||
                    ""
                )
                    .trim();


            const headingRelevance =
                calculateHeadingRelevance(
                    heading,
                    queryTokens,
                    normalizedQuery
                );


            const headingLevel =
                getHeadingLevelNumber(
                    doc.headingLevel
                );


            let score =
                Number(
                    hit.score ||
                    0
                );


            // ==========================================
            // أولوية العناوين
            // ==========================================

            if (
                doc.isHeading
            ) {

                score +=
                    30;

            }


            score +=
                headingRelevance;


            // ==========================================
            // أولوية المبحث/المطلب
            // ==========================================

            if (
                headingLevel ===
                2
            ) {

                score +=
                    12;

            }
            else if (
                headingLevel ===
                3
            ) {

                score +=
                    15;

            }
            else if (
                headingLevel ===
                4
            ) {

                score +=
                    8;

            }


            // ==========================================
            // تطابق السؤال كاملًا داخل الفقرة
            // ==========================================

            const normalizedText =
                normalizeSearchText(
                    doc.text ||
                    ""
                );


            if (
                normalizedText.includes(
                    normalizedQuery
                )
            ) {

                score +=
                    20;

            }


            // ==========================================
            // عدد كلمات السؤال الموجودة
            // ==========================================

            let matchedTokens =
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

                        matchedTokens +=
                            1;

                    }

                }
            );


            const queryCoverage =
                queryTokens.length >
                0
                    ? matchedTokens /
                      queryTokens.length
                    : 0;


            score +=
                queryCoverage *
                12;


            // ==========================================
            // محاولة تحديد نوع المطابقة
            // ==========================================

            let matchType =
                "word";


            if (
                normalizedText.includes(
                    normalizedQuery
                )
            ) {

                matchType =
                    "exact";

            }
            else if (
                doc.isHeading ||
                headingRelevance >=
                    25
            ) {

                matchType =
                    "heading";

            }


            scoredResults.push({

                id:
                    String(
                        doc.id ||
                        ""
                    ),

                paragraphIndex:
                    paragraphIndex,

                paragraphId:
                    String(
                        doc.id ||
                        (
                            "p-" +
                            paragraphIndex
                        )
                    ),

                text:
                    String(
                        doc.text ||
                        ""
                    ),

                heading:
                    heading,

                headingIndex:
                    Number(
                        doc.headingIndex ||
                        -1
                    ),

                headingLevel:
                    String(
                        doc.headingLevel ||
                        ""
                    ),

                isHeading:
                    Boolean(
                        doc.isHeading
                    ),

                score:
                    score,

                oramaScore:
                    Number(
                        hit.score ||
                        0
                    ),

                matchedTokens:
                    matchedTokens,

                queryCoverage:
                    queryCoverage,

                headingRelevance:
                    headingRelevance,

                matchType:
                    matchType

            });

        }
    );


    // ==========================================
    // ترتيب النتائج
    // ==========================================

    scoredResults.sort(
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
                    b.isHeading
                ) !==
                Boolean(
                    a.isHeading
                )
            ) {

                return (
                    b.isHeading
                        ? 1
                        : -1
                );

            }


            if (
                b.queryCoverage !==
                a.queryCoverage
            ) {

                return (
                    b.queryCoverage -
                    a.queryCoverage
                );

            }


            return (
                a.paragraphIndex -
                b.paragraphIndex
            );

        }
    );


    // ==========================================
    // النتيجة النهائية
    // ==========================================

    const finalResults =
        scoredResults.slice(
            0,
            maxResults
        );


    return {

        query:
            normalizedQuery,

        count:
            scoredResults.length,

        results:
            finalResults,

        matchedTerms:
            queryTokens,

        matchedFamilies:
            [],

        totalQueryTerms:
            queryTokens.length,

        profile:
            profile,

        indexedOccurrences:
            scoredResults.length,

        headings:
            headings

    };

}


// =====================================================
// Search Orama Headings
// البحث المباشر في العناوين
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
        sortHeadings(
            structureData.headings
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

                headingIndex:
                    "number",

                heading:
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

                    headingIndex:
                        Number(
                            heading.index
                        ),

                    heading:
                        String(
                            heading.text ||
                            ""
                        )
                            .trim(),

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


    await insertMultiple(
        db,
        records
    );


    const normalizedQuery =
        normalizeOramaQuery(
            query
        );


    if (
        !normalizedQuery
    ) {

        return [];

    }


    let result;


    try {

        result =
            await search(
                db,
                {

                    term:
                        normalizedQuery,

                    properties:
                        [
                            "heading"
                        ],

                    tolerance:
                        2,

                    limit:
                        20

                }
            );

    }
    catch (error) {

        console.error(
            "فشل البحث في عناوين Orama:",
            error
        );

        return [];

    }


    if (
        !result ||
        !Array.isArray(
            result.hits
        )
    ) {

        return [];

    }


    return result.hits.map(
        function (
            hit
        ) {

            return {

                score:
                    Number(
                        hit.score ||
                        0
                    ),

                headingIndex:
                    Number(
                        hit.document.headingIndex
                    ),

                heading:
                    String(
                        hit.document.heading ||
                        ""
                    ),

                style:
                    String(
                        hit.document.style ||
                        ""
                    ),

                level:
                    Number(
                        hit.document.level ||
                        9
                    )

            };

        }
    );

}


// =====================================================
// Get Heading Paragraph Range
// الحصول على نطاق فقرات المطلب
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
        sortHeadings(
            structureData.headings
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

            if (
                !paragraph
            ) {

                return false;

            }


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
// Build Orama Heading Expansion
// توسيع نتيجة العنوان إلى محتوى المطلب
// =====================================================

function expandHeadingResult(
    structureData,
    result,
    maxParagraphs
) {

    if (
        !result ||
        !result.isHeading
    ) {

        return result;

    }


    const headingIndex =
        Number(
            result.paragraphIndex
        );


    const range =
        getHeadingParagraphRange(
            structureData,
            headingIndex
        );


    const limit =
        typeof maxParagraphs ===
            "number"
                ? Math.max(
                    1,
                    maxParagraphs
                )
                : 3;


    const selected =
        range.slice(
            0,
            limit
        );


    const content =
        selected
            .map(
                function (
                    paragraph
                ) {

                    return String(
                        paragraph.text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim();

                }
            )
            .filter(
                function (
                    text
                ) {

                    return text !== "";

                }
            )
            .join(
                " "
            );


    return Object.assign(
        {},
        result,
        {

            headingMatch:
                true,

            sectionParagraphs:
                selected,

            sectionText:
                content

        }
    );

}


// =====================================================
// Add Neighbor Paragraphs
// إضافة الفقرة السابقة واللاحقة
// =====================================================

function addNeighborParagraphs(
    structureData,
    result
) {

    if (
        !structureData ||
        !result
    ) {

        return result;

    }


    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    const index =
        Number(
            result.paragraphIndex
        );


    const previous =
        paragraphs.find(
            function (
                paragraph
            ) {

                return (
                    Number(
                        paragraph.index
                    ) ===
                    index - 1
                );

            }
        );


    const next =
        paragraphs.find(
            function (
                paragraph
            ) {

                return (
                    Number(
                        paragraph.index
                    ) ===
                    index + 1
                );

            }
        );


    return Object.assign(
        {},
        result,
        {

            previousParagraphText:
                previous
                    ? String(
                        previous.text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim()
                    : "",

            nextParagraphText:
                next
                    ? String(
                        next.text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim()
                    : ""

        }
    );

}


// =====================================================
// Search + Expand + Neighbors
// الواجهة الموحدة للاسترجاع
// =====================================================

async function searchOramaForRetrieval(
    documentItem,
    query,
    options
) {

    const settings =
        options ||
        {};


    const maxResults =
        typeof settings.maxResults ===
            "number"
                ? settings.maxResults
                : 8;


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    const searchResult =
        await searchOramaDocument(
            documentItem,
            query,
            settings
        );


    const expandedResults =
        searchResult.results.map(
            function (
                result
            ) {

                let item =
                    expandHeadingResult(
                        structureData,
                        result,
                        settings.headingParagraphs || 3
                    );


                item =
                    addNeighborParagraphs(
                        structureData,
                        item
                    );


                return item;

            }
        );


    // ==========================================
    // منع تكرار الفقرات
    // ==========================================

    const unique =
        new Map();


    expandedResults.forEach(
        function (
            result
        ) {

            const key =
                String(
                    result.paragraphIndex
                );


            if (
                !unique.has(
                    key
                ) ||
                Number(
                    result.score ||
                    0
                ) >
                Number(
                    unique.get(
                        key
                    ).score ||
                    0
                )
            ) {

                unique.set(
                    key,
                    result
                );

            }

        }
    );


    let results =
        Array.from(
            unique.values()
        );


    // ==========================================
    // إعادة الترتيب بعد التوسعة
    // ==========================================

    results.sort(
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


    results =
        results.slice(
            0,
            maxResults
        );


    return {

        query:
            searchResult.query,

        count:
            searchResult.count,

        results:
            results,

        matchedTerms:
            searchResult.matchedTerms,

        matchedFamilies:
            searchResult.matchedFamilies,

        totalQueryTerms:
            searchResult.totalQueryTerms,

        profile:
            searchResult.profile,

        indexedOccurrences:
            searchResult.indexedOccurrences

    };

}


// =====================================================
// Compatibility Bridge
// الاسم الجديد بدل searchIndexedDocument
// =====================================================
//
// لا نعيد بناء المحرك القديم.
// هذا الاسم سيستخدمه الجزء الثالث للـ AI.
// =====================================================

async function searchIndexedDocument(
    documentId,
    query,
    options
) {

    const documentItem =
        documents.find(
            function (
                documentItem
            ) {

                return (
                    documentItem &&
                    String(
                        documentItem.id
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


    return await searchOramaForRetrieval(
        documentItem,
        query,
        options || {}
    );

}


// =====================================================
// Build Retrieval Context
// تحويل نتائج Orama إلى سياق AI
// =====================================================

function buildRetrievalContext(
    searchResult,
    options
) {

    const settings =
        options ||
        {};


    const maxResults =
        typeof settings.maxResults ===
            "number"
                ? settings.maxResults
                : 4;


    const maxChars =
        typeof settings.maxChars ===
            "number"
                ? settings.maxChars
                : 3500;


    const includeNeighbors =
        settings.includeNeighbors !==
        false;


    if (
        !searchResult ||
        !Array.isArray(
            searchResult.results
        )
    ) {

        return {

            query:
                "",

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


    const results =
        searchResult.results
            .filter(
                function (
                    result
                ) {

                    return (
                        result &&
                        typeof result.text ===
                            "string"
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


    const selected =
        [];


    const selectedParagraphIndexes =
        new Set();


    for (
        let i = 0;
        i < results.length;
        i++
    ) {

        const candidate =
            results[i];


        const paragraphIndex =
            Number(
                candidate.paragraphIndex
            );


        if (
            selectedParagraphIndexes.has(
                paragraphIndex
            )
        ) {

            continue;

        }


        const candidateText =
            String(
                candidate.sectionText ||
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


        selected.push(
            candidate
        );


        selectedParagraphIndexes.add(
            paragraphIndex
        );


        if (
            selected.length >=
            maxResults
        ) {

            break;

        }

    }


    const contexts =
        [];


    let totalChars =
        0;


    selected.forEach(
        function (
            result,
            index
        ) {

            const remainingChars =
                maxChars -
                totalChars;


            if (
                remainingChars <=
                0
            ) {

                return;

            }


            let mainContext =
                String(
                    result.sectionText ||
                    result.text ||
                    ""
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            if (
                !mainContext
            ) {

                return;

            }


            const heading =
                String(
                    result.heading ||
                    ""
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            let previousContext =
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


            let nextContext =
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
                selectedParagraphIndexes.has(
                    Number(
                        result.paragraphIndex
                    ) - 1
                )
            ) {

                previousContext =
                    "";

            }


            if (
                selectedParagraphIndexes.has(
                    Number(
                        result.paragraphIndex
                    ) + 1
                )
            ) {

                nextContext =
                    "";

            }


            const reserved =
                220;


            const available =
                Math.max(
                    300,
                    remainingChars -
                    reserved
                );


            if (
                mainContext.length >
                available
            ) {

                mainContext =
                    mainContext.substring(
                        0,
                        available
                    ) +
                    "…";

            }


            let context =
                mainContext;


            let remaining =
                available -
                context.length;


            if (
                includeNeighbors &&
                previousContext &&
                remaining >
                150
            ) {

                const allowed =
                    remaining -
                    1;


                const previousPart =
                    previousContext.length >
                    allowed
                        ? previousContext.substring(
                            Math.max(
                                0,
                                previousContext.length -
                                allowed
                            )
                        ) +
                        "…"
                        : previousContext;


                context =
                    previousPart +
                    " " +
                    context;

            }


            remaining =
                available -
                context.length;


            if (
                includeNeighbors &&
                nextContext &&
                remaining >
                150
            ) {

                const allowed =
                    remaining -
                    1;


                const nextPart =
                    nextContext.length >
                    allowed
                        ? nextContext.substring(
                            0,
                            allowed
                        ) +
                        "…"
                        : nextContext;


                context =
                    context +
                    " " +
                    nextPart;

            }


            const item = {

                rank:
                    index + 1,

                paragraphIndex:
                    result.paragraphIndex,

                heading:
                    heading,

                score:
                    Number(
                        result.score ||
                        0
                    ),

                matchType:
                    result.matchType ||
                    "word",

                headingMatch:
                    Boolean(
                        result.headingMatch
                    ),

                sectionText:
                    String(
                        result.sectionText ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim(),

                previousParagraph:
                    previousContext,

                mainParagraph:
                    String(
                        result.text ||
                        ""
                    )
                        .replace(
                            /\s+/g,
                            " "
                        )
                        .trim(),

                nextParagraph:
                    nextContext,

                context:
                    context

            };


            contexts.push(
                item
            );


            totalChars +=
                context.length;

        }
    );


    const textParts =
        [];


    contexts.forEach(
        function (
            item
        ) {

            let block =
                "[مقطع " +
                String(
                    item.rank
                ) +
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


            textParts.push(
                block
            );

        }
    );


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
            textParts.join(
                "\n\n---\n\n"
            )

    };

}


// =====================================================
// Debug Bridge
// اختبار محرك Orama الجديد
// =====================================================

window.testOramaSearch =
    async function (
        query
    ) {

        console.log(
            "بدء اختبار Orama الجديد:",
            query
        );


        if (
            !currentDocument
        ) {

            console.warn(
                "لا يوجد مستند نشط."
            );


            return [];

        }


        try {

            const result =
                await searchOramaDocument(
                    currentDocument,
                    query,
                    {

                        profile:
                            "general",

                        maxResults:
                            10

                    }
                );


            console.log(
                "===================================="
            );


            console.log(
                "Orama Search"
            );


            console.log(
                "السؤال:",
                query
            );


            console.log(
                "عدد النتائج:",
                result.count
            );


            console.log(
                "===================================="
            );


            result.results.forEach(
                function (
                    item,
                    index
                ) {

                    console.log(
                        "#" +
                        (
                            index + 1
                        ),
                        {

                            score:
                                item.score,

                            oramaScore:
                                item.oramaScore,

                            heading:
                                item.heading,

                            paragraphIndex:
                                item.paragraphIndex,

                            matchType:
                                item.matchType,

                            headingRelevance:
                                item.headingRelevance,

                            queryCoverage:
                                item.queryCoverage,

                            text:
                                String(
                                    item.text ||
                                    ""
                                ).substring(
                                    0,
                                    300
                                )

                        }
                    );

                }
            );


            return result.results;

        }
        catch (
            error
        ) {

            console.error(
                "فشل اختبار Orama:",
                error
            );


            return [];

        }

    };


// =====================================================
// Debug Heading Search
// اختبار البحث في العناوين
// =====================================================

window.testOramaHeadingSearch =
    async function (
        query
    ) {

        if (
            !currentDocument
        ) {

            console.warn(
                "لا يوجد مستند نشط."
            );


            return [];

        }


        try {

            const results =
                await searchOramaHeadings(
                    currentDocument,
                    query
                );


            console.log(
                "======================================"
            );


            console.log(
                "بحث Orama في العناوين"
            );


            console.log(
                "السؤال:",
                query
            );


            console.log(
                "عدد النتائج:",
                results.length
            );


            console.log(
                "======================================"
            );


            results.forEach(
                function (
                    item,
                    index
                ) {

                    console.log(
                        "#" +
                        (
                            index + 1
                        ),
                        item
                    );

                }
            );


            return results;

        }
        catch (
            error
        ) {

            console.error(
                "فشل اختبار عناوين Orama:",
                error
            );


            return [];

        }

    };


// =====================================================
// Standalone Orama Test
// اختبار مستقل للتأكد من أن Orama نفسه يعمل
// =====================================================

window.testOramaStandalone =
    async function (
        query
    ) {

        if (
            !currentDocument
        ) {

            console.warn(
                "لا يوجد مستند نشط."
            );


            return [];

        }


        try {

            const structureData =
                await ensureDocumentStructure(
                    currentDocument
                );


            if (
                !structureData
            ) {

                return [];

            }


            const paragraphs =
                Array.isArray(
                    structureData.paragraphs
                )
                    ? structureData.paragraphs
                    : [];


            const headings =
                sortHeadings(
                    structureData.headings
                );


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

                        paragraphIndex:
                            "number",

                        text:
                            "string",

                        heading:
                            "string"

                    },

                    language:
                        "arabic"

                });


            const records =
                paragraphs
                    .filter(
                        function (
                            paragraph
                        ) {

                            return (
                                paragraph &&
                                String(
                                    paragraph.text ||
                                    ""
                                ).trim()
                            );

                        }
                    )
                    .map(
                        function (
                            paragraph
                        ) {

                            const heading =
                                findNearestHeading(
                                    Number(
                                        paragraph.index
                                    ),
                                    headings
                                );


                            return {

                                id:
                                    "p-" +
                                    String(
                                        paragraph.index
                                    ),

                                paragraphIndex:
                                    Number(
                                        paragraph.index
                                    ),

                                text:
                                    String(
                                        paragraph.text ||
                                        ""
                                    )
                                        .trim(),

                                heading:
                                    heading
                                        ? String(
                                            heading.text ||
                                            ""
                                        ).trim()
                                        : ""

                            };

                        }
                    );


            console.log(
                "عدد السجلات التي ستدخل Orama:",
                records.length
            );


            if (
                records.length ===
                0
            ) {

                return [];

            }


            await insertMultiple(
                db,
                records
            );


            console.log(
                "تم إدخال السجلات إلى Orama."
            );


            const normalizedQuery =
                normalizeOramaQuery(
                    query
                );


            const result =
                await search(
                    db,
                    {

                        term:
                            normalizedQuery,

                        properties:
                            [
                                "text",
                                "heading"
                            ],

                        boost:
                            {
                                heading:
                                    5
                            },

                        tolerance:
                            2,

                        limit:
                            10

                    }
                );


            console.log(
                "===================================="
            );


            console.log(
                "Orama Standalone Test"
            );


            console.log(
                "السؤال:",
                query
            );


            console.log(
                "عدد النتائج:",
                result.count
            );


            console.log(
                "===================================="
            );


            const finalResults =
                Array.isArray(
                    result.hits
                )
                    ? result.hits.map(
                        function (
                            hit
                        ) {

                            return {

                                score:
                                    Number(
                                        hit.score ||
                                        0
                                    ),

                                heading:
                                    hit.document.heading,

                                paragraphIndex:
                                    hit.document.paragraphIndex,

                                paragraphId:
                                    hit.document.id,

                                text:
                                    String(
                                        hit.document.text ||
                                        ""
                                    )

                            };

                        }
                    )
                    : [];


            finalResults.forEach(
                function (
                    item,
                    index
                ) {

                    console.log(
                        "#" +
                        (
                            index + 1
                        ),
                        item
                    );

                }
            );


            return finalResults;

        }
        catch (
            error
        ) {

            console.error(
                "فشل اختبار Orama المستقل:",
                error
            );


            return [];

        }

    };


// =====================================================
// Test Current Orama Index
// اختبار حالة فهرس المستند الحالي
// =====================================================

window.testCurrentDocumentOrama =
    async function () {

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
                await ensureDocumentStructure(
                    currentDocument
                );


            const db =
                await ensureOramaRetrievalIndex(
                    currentDocument,
                    structureData
                );


            console.log(
                "======================================"
            );


            console.log(
                "حالة فهرس Orama"
            );


            console.log(
                "المستند:",
                currentDocument.name
            );


            console.log(
                "Document ID:",
                currentDocument.id
            );


            console.log(
                "الفقرات:",
                structureData.paragraphCount
            );


            console.log(
                "العناوين:",
                structureData.headingCount
            );


            console.log(
                "Orama DB:",
                db
            );


            console.log(
                "Cache Key:",
                oramaRetrievalCacheKey
            );


            console.log(
                "======================================"
            );


            return db;

        }
        catch (
            error
        ) {

            console.error(
                "فشل فحص فهرس Orama:",
                error
            );


            return null;

        }

    };


// =====================================================
// PART 2 END
// =====================================================
// ======================================
// Word AI Assistant
// PART 3 / 4
// الذكاء الاصطناعي + الإعدادات + النماذج
// متوافق مع Orama الجديد في PART 2
// ======================================


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
// JSON Reader
// =====================================================

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
                String(
                    result.error.message
                );


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
// Extract OpenAI / OpenRouter / Groq Answer
// =====================================================

function extractOpenAIStyleAnswer(
    result,
    providerName
) {

    if (
        result &&
        Array.isArray(
            result.choices
        ) &&
        result.choices.length > 0
    ) {

        const choice =
            result.choices[0];


        if (
            choice &&
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
// Extract Gemini Answer
// =====================================================

function extractGeminiAnswer(
    result
) {

    if (
        result &&
        Array.isArray(
            result.candidates
        ) &&
        result.candidates.length > 0
    ) {

        const candidate =
            result.candidates[0];


        if (
            candidate &&
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
// AI Settings Storage
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


// =====================================================
// Save AI Settings
// =====================================================

function saveAISettings(
    data
) {

    localStorage.setItem(
        "AI_SETTINGS",
        JSON.stringify(
            data ||
            {}
        )
    );

}


// =====================================================
// Provider Information
// =====================================================

function updateProviderInfo() {

    if (
        !providerInfo ||
        !provider
    ) {

        return;

    }


    const value =
        String(
            provider.value ||
            "openrouter"
        ).toLowerCase();


    if (
        value ===
        "openrouter"
    ) {

        providerInfo.innerHTML =
            "OpenRouter: سيتم جلب النماذج المجانية المتاحة.";

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
        "سيتم تحديد طريقة الاتصال بحسب المزود.";

}


// =====================================================
// Load Saved Settings
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
            savedModel
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


            if (!settingsWindow) {

                console.warn(
                    "عنصر settings-window غير موجود في الصفحة."
                );

                return;

            }


            loadSettings();


            settingsWindow.classList.add(
                "open"
            );

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
                        ? String(
                            provider.value ||
                            "openrouter"
                        )
                        : "openrouter",

                key:
                    apiKey
                        ? String(
                            apiKey.value ||
                            ""
                        ).trim()
                        : "",

                model:
                    modelSelect
                        ? String(
                            modelSelect.value ||
                            ""
                        ).trim()
                        : ""

            };


            if (
                !settings.key
            ) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ يرجى إدخال مفتاح API.";

                }

                return;

            }


            if (
                !settings.model
            ) {

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
// Load Models
// =====================================================

async function loadModels() {

    const selectedProvider =
        provider
            ? String(
                provider.value ||
                "openrouter"
            ).toLowerCase()
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


            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⏳ جاري تحديث النماذج...";

            }


            try {

                await loadModels();

            }
            catch (error) {

                if (settingsStatus) {

                    settingsStatus.innerHTML =
                        "⚠ " +
                        (
                            error &&
                            error.message
                                ? error.message
                                : "تعذر تحديث النماذج"
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
// Groq Models
// =====================================================

async function loadGroqModels() {

    const key =
        apiKey
            ? String(
                apiKey.value ||
                ""
            ).trim()
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
        !result ||
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
                        item.active !== false
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


// =====================================================
// OpenRouter Models
// =====================================================

async function loadOpenRouterModels() {

    const key =
        apiKey
            ? String(
                apiKey.value ||
                ""
            ).trim()
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
        !result ||
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
                    !item.id
                ) {

                    return false;

                }


                return String(
                    item.id
                ).endsWith(
                    ":free"
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
// OpenAI Models
// =====================================================

async function loadOpenAIModels() {

    const key =
        apiKey
            ? String(
                apiKey.value ||
                ""
            ).trim()
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
        !result ||
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
                        id.startsWith("gpt-") ||
                        id.startsWith("o1") ||
                        id.startsWith("o3") ||
                        id.startsWith("o4")
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


// =====================================================
// Gemini Models
// =====================================================

async function loadGeminiModels() {

    const key =
        apiKey
            ? String(
                apiKey.value ||
                ""
            ).trim()
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
        !result ||
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

                    if (
                        !item ||
                        !item.name ||
                        !Array.isArray(
                            item.supportedGenerationMethods
                        )
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
            )
            .map(
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


    if (
        savedModel
    ) {

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


        if (
            exists
        ) {

            modelSelect.value =
                savedModel;

        }

    }

}


// =====================================================
// AI Connection Test
// =====================================================

async function testAIConnection() {

    const data = {

        provider:
            provider
                ? String(
                    provider.value ||
                    "openrouter"
                ).toLowerCase()
                : "openrouter",

        key:
            apiKey
                ? String(
                    apiKey.value ||
                    ""
                ).trim()
                : "",

        model:
            modelSelect
                ? String(
                    modelSelect.value ||
                    ""
                ).trim()
                : ""

    };


    if (
        !data.key
    ) {

        throw new Error(
            "يرجى إدخال مفتاح API أولاً."
        );

    }


    if (
        !data.model
    ) {

        throw new Error(
            "يرجى تحديد نموذج الذكاء الاصطناعي أولاً."
        );

    }


    // ==================================
    // OpenRouter
    // ==================================

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


        return "✓ تم الاتصال بـ OpenRouter بنجاح";

    }


    // ==================================
    // OpenAI
    // ==================================

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


        return "✓ تم الاتصال بـ OpenAI بنجاح";

    }


    // ==================================
    // Gemini
    // ==================================

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

                                    role:
                                        "user",

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


        return "✓ تم الاتصال بـ Gemini بنجاح";

    }


    // ==================================
    // Groq
    // ==================================

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
                            error &&
                            error.message
                                ? error.message
                                : "تعذر الاتصال"
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
// Retrieval Profile
// تحديد طبيعة السؤال
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
    // تعريف
    // ==================================

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


    // ==================================
    // أثر
    // ==================================

    if (
        /اثر|أثر|تاثير|تأثير|نتائج|ينتج عن|يترتب على|انعكاس/
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
    // أسباب
    // ==================================

    if (
        /لماذا|سبب|اسباب|أسباب|علة|علل|لأن|لان|بسبب/
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
    // موقع / موضع
    // ==================================

    if (
        /اين|أين|موضع|موضعه|الفصل|المبحث|المطلب|الصفحة/
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
// Retrieval Limits
// تحديد حجم السياق
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
// Estimate Tokens
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
// يستخدم Orama عبر searchIndexedDocument
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


        // ==================================
        // Orama هو محرك البحث الأساسي
        // ==================================

        const searchResult =
            await searchIndexedDocument(
                currentDocument.id,
                query,
                {

                    profile:
                        retrievalProfile.type,

                    maxResults:
                        Math.max(
                            maxResults * 2,
                            10
                        ),

                    headingParagraphs:
                        3

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


        // ==================================
        // مصادر الإحالات
        // ==================================

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
                Array.isArray(
                    retrieval.contexts
                )
                    ? retrieval.contexts
                    : [],

            text:
                retrieval.text ||
                ""

        };

    }
    catch (error) {

        currentCitationSources =
            [];


        console.warn(
            "تعذر استرجاع سياق المستند عبر Orama:",
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
// AI Request
// الطلب غير المتدفق
// =====================================================

async function askAI(
    text
) {

    const data =
        getSavedSettings();


    const selectedProvider =
        String(
            data.provider ||
            "openrouter"
        ).toLowerCase();


    const key =
        String(
            data.key ||
            ""
        ).trim();


    const model =
        String(
            data.model ||
            ""
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


    // ==================================
    // أولًا نسترجع سياق المستند
    // قبل حساب تاريخ المحادثة
    // ==================================

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


    // ==================================
    // محتوى السؤال
    // ==================================

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
                            ? documentContext.matchedFamilies.join("، ")
                            : "لا توجد"
                    ),

                "",

                "=== المادة المستخرجة ===",
                documentContext.text,

                "",

                "=== قواعد الإجابة ===",
                "اعتمد على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",
                "استخرج الأفكار المرتبطة بالسؤال فقط.",
                "ادمج الأفكار المتشابهة.",
                "أجب عن جميع جوانب السؤال التي تدعمها المادة.",
                "لا تضف معلومة غير موجودة في المادة المستخرجة.",
                "إذا لم تكف المادة للإجابة عن جزء من السؤال، صرّح بذلك.",
                "لا تستخدم المعرفة العامة لسد النقص إلا إذا طلب المستخدم ذلك صراحة.",
                "حافظ على العربية والأسلوب الأكاديمي.",
                "ضع الإحالات بصيغة [مقطع X].",
                "لا تخترع أرقام المقاطع.",
                "قدّم إجابة تركيبية لا تلخيصًا منفصلًا لكل مقطع.",
                "استبعد المقاطع التي تحتوي كلمات السؤال دون أن تجيب عنه مباشرة."

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


    // ==================================
    // OpenRouter
    // ==================================

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
                            key,

                        "HTTP-Referer":
                            window.location.href,

                        "X-Title":
                            "Research Tools"

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


    // ==================================
    // OpenAI
    // ==================================

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


    // ==================================
    // Gemini
    // ==================================

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


    // ==================================
    // Groq
    // ==================================

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
// PART 3 END
// =====================================================
 // ======================================
// Word AI Assistant
// PART 4 / 4
// Streaming + Send + Initialization
// ======================================


// =====================================================
// Stream Groq AI
// =====================================================

async function streamGroqAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        String(
            data.key ||
            ""
        ).trim();


    const model =
        String(
            data.model ||
            ""
        ).trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح Groq من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج Groq."
        );

    }


    // ==================================
    // سياق المستند
    // ==================================

    const documentContext =
        await buildAIDocumentContext(
            text
        );


    // ==================================
    // تاريخ المحادثة
    // ==================================

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
    // محتوى السؤال
    // ==================================

    let userContent =
        text;


    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "أنت مساعد بحث أكاديمي يعمل على مستند Word.",
                "",

                "=== سؤال المستخدم ===",
                text,

                "",

                "=== اسم المستند ===",
                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",
                documentContext.text,

                "",

                "=== قواعد الإجابة ===",
                "اعتمد على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",
                "استخرج الأفكار المرتبطة بالسؤال فقط.",
                "ادمج الأفكار المتشابهة.",
                "أجب عن جميع جوانب السؤال التي تدعمها المادة.",
                "لا تضف معلومات غير موجودة في المادة المستخرجة.",
                "إذا لم تكف المادة للإجابة عن جزء من السؤال، صرّح بذلك.",
                "لا تستخدم المعرفة العامة لسد النقص إلا إذا طلب المستخدم ذلك صراحة.",
                "حافظ على العربية والأسلوب الأكاديمي.",
                "ضع الإحالات بصيغة [مقطع X].",
                "لا تخترع أرقام المقاطع.",
                "قدّم إجابة تركيبية مترابطة."

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


    // ==================================
    // إرسال الطلب
    // ==================================

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
                            2500,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (
        !response.ok
    ) {

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


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من Groq."
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


        const events =
            buffer.split(
                "\n\n"
            );


        buffer =
            events.pop() ||
            "";


        events.forEach(
            function (
                event
            ) {

                event
                    .split("\n")
                    .forEach(
                        function (
                            line
                        ) {

                            const cleanLine =
                                line.trim();


                            if (
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
                    );

            }
        );

    }


    if (
        buffer.trim()
    ) {

        // قد توجد بيانات نهائية غير مكتملة
        const remainingLines =
            buffer
                .split("\n")
                .map(
                    function (
                        line
                    ) {

                        return line.trim();

                    }
                )
                .filter(
                    function (
                        line
                    ) {

                        return line.startsWith(
                            "data:"
                        );

                    }
                );


        remainingLines.forEach(
            function (
                line
            ) {

                const dataText =
                    line
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


                try {

                    const parsed =
                        JSON.parse(
                            dataText
                        );


                    const delta =
                        parsed &&
                        parsed.choices &&
                        parsed.choices[0] &&
                        parsed.choices[0].delta
                            ? parsed.choices[0].delta.content
                            : "";


                    if (
                        typeof delta ===
                            "string" &&
                        delta
                    ) {

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

                }
                catch (
                    error
                ) {

                    // تجاهل الجزء غير القابل للتحليل

                }

            }
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من Groq عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// Stream Gemini AI
// =====================================================

async function streamGeminiAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        String(
            data.key ||
            ""
        ).trim();


    const model =
        String(
            data.model ||
            ""
        ).trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح Gemini من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج Gemini من الإعدادات."
        );

    }


    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const systemInstruction = [

        "أنت مساعد بحث أكاديمي يعمل على مستندات Word.",

        "اعتمد على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",

        "أجب عن السؤال مباشرة وبأسلوب أكاديمي واضح.",

        "رتب الإجابة وفق محاور السؤال.",

        "ادمج الأفكار المتشابهة.",

        "استبعد المعلومة الجانبية التي لا تجيب عن السؤال.",

        "لا تضف معلومة غير موجودة في المادة المستخرجة.",

        "إذا لم تكف المادة للإجابة عن جزء من السؤال، صرّح بذلك.",

        "لا تستخدم المعرفة العامة لسد النقص إلا إذا طلب المستخدم ذلك صراحة.",

        "ضع الإحالات بصيغة [مقطع X].",

        "لا تخترع أرقام المقاطع.",

        "قدّم خلاصة تركيبية مترابطة."

    ].join(
        "\n"
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

        const messagesWithoutCurrent =
            currentChat.messages.slice(
                0,
                -1
            );


        messagesWithoutCurrent
            .slice(
                -historyLimit
            )
            .forEach(
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
                                ? "model"
                                : "user",

                        parts: [

                            {
                                text:
                                    messageText
                            }

                        ]

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

                "=== سؤال المستخدم ===",
                text,

                "",

                "=== اسم المستند ===",
                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة ===",
                documentContext.text

            ].join(
                "\n"
            );

    }


    conversationMessages.push({

        role:
            "user",

        parts: [

            {
                text:
                    userContent
            }

        ]

    });


    const cleanModel =
        normalizeGeminiModel(
            model
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
                        key

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


    if (
        !response.ok
    ) {

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


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من Gemini."
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


    function processGeminiLine(
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


                if (
                    part.thought ===
                    true
                ) {

                    return;

                }


                const delta =
                    part.text;


                if (
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
        );

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
            )
            .replace(
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


            processGeminiLine(
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

        processGeminiLine(
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


// =====================================================
// Stream OpenRouter AI
// =====================================================

async function streamOpenRouterAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        String(
            data.key ||
            ""
        ).trim();


    const model =
        String(
            data.model ||
            ""
        ).trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenRouter من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج OpenRouter."
        );

    }


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

        currentChat.messages
            .slice(
                0,
                -1
            )
            .slice(
                -historyLimit
            )
            .forEach(
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

                "أنت مساعد بحث أكاديمي.",

                "",

                "=== سؤال المستخدم ===",
                text,

                "",

                "=== اسم المستند ===",
                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",
                documentContext.text,

                "",

                "=== قواعد الإجابة ===",
                "اعتمد على المادة المستخرجة بوصفها المصدر الأساسي.",
                "أجب عن السؤال مباشرة.",
                "ادمج الأفكار المتشابهة.",
                "لا تضف معلومات غير موجودة في المادة.",
                "ضع الإحالات بصيغة [مقطع X].",
                "لا تخترع أرقام المقاطع.",
                "قدّم إجابة مترابطة."

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
                        key,

                    "HTTP-Referer":
                        window.location.href,

                    "X-Title":
                        "Research Tools"

                },

                body:
                    JSON.stringify({

                        model:
                            model,

                        messages:
                            conversationMessages,

                        max_tokens:
                            2500,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (
        !response.ok
    ) {

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


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من OpenRouter."
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


    function processOpenRouterLine(
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
            )
            .replace(
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


            processOpenRouterLine(
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

        processOpenRouterLine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من OpenRouter عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// Stream OpenAI AI
// =====================================================

async function streamOpenAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        String(
            data.key ||
            ""
        ).trim();


    const model =
        String(
            data.model ||
            ""
        ).trim();


    if (!key) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenAI من الإعدادات."
        );

    }


    if (!model) {

        throw new Error(
            "لم يتم تحديد نموذج OpenAI."
        );

    }


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

        currentChat.messages
            .slice(
                0,
                -1
            )
            .slice(
                -historyLimit
            )
            .forEach(
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

                "أنت مساعد بحث أكاديمي.",

                "",

                "=== سؤال المستخدم ===",
                text,

                "",

                "=== اسم المستند ===",
                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",
                documentContext.text,

                "",

                "=== قواعد الإجابة ===",
                "اعتمد على المادة المستخرجة بوصفها المصدر الأساسي.",
                "أجب عن السؤال مباشرة.",
                "ادمج الأفكار المتشابهة.",
                "لا تضف معلومات غير موجودة في المادة.",
                "ضع الإحالات بصيغة [مقطع X].",
                "لا تخترع أرقام المقاطع.",
                "قدّم إجابة مترابطة."

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
                            2500,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (
        !response.ok
    ) {

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


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من OpenAI."
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


    function processOpenAILine(
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
            )
            .replace(
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


            processOpenAILine(
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

        processOpenAILine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من OpenAI عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// Send Message
// =====================================================

async function sendMessage() {

    if (!input) {

        return;

    }


    const text =
        input.value.trim();


    if (
        !text
    ) {

        return;

    }


    // ==================================
    // إنشاء محادثة عند الحاجة
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
    // تثبيت المحادثة
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


        if (
            !alreadyExists
        ) {

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
                    new Date().toISOString();


                saveProjects();

            }

        }


        saveChats();

    }


    // ==================================
    // رسالة المستخدم
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
    // رسالة التحميل
    // ==================================

    const loading =
        document.createElement(
            "div"
        );


    loading.className =
        "message ai-message";


    loading.innerHTML =
        "⏳ جاري التفكير...";


    if (
        chatArea
    ) {

        chatArea.appendChild(
            loading
        );


        chatArea.scrollTop =
            chatArea.scrollHeight;

    }


    // ==================================
    // المزود الحالي
    // ==================================

    const savedSettings =
        getSavedSettings();


    const selectedProvider =
        String(
            savedSettings.provider ||
            "openrouter"
        ).toLowerCase();


    let pendingRenderText =
        "";


    let renderTimer =
        null;


    function renderStreamingText() {

        if (
            !loading
        ) {

            return;

        }


        if (
            !pendingRenderText
        ) {

            loading.innerHTML =
                "⏳ جاري التفكير...";

        }
        else {

            loading.innerHTML =
                formatAIMessage(
                    pendingRenderText,
                    currentCitationSources
                );

        }


        if (
            chatArea
        ) {

            chatArea.scrollTop =
                chatArea.scrollHeight;

        }


        renderTimer =
            null;

    }


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

                    renderStreamingText();

                },
                60
            );

    }


    try {

        let answer =
            "";


        // ==================================
        // Groq
        // ==================================

        if (
            selectedProvider ===
            "groq"
        ) {

            answer =
                await streamGroqAI(
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

        }


        // ==================================
        // Gemini
        // ==================================

        else if (
            selectedProvider ===
            "gemini"
        ) {

            answer =
                await streamGeminiAI(
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

        }


        // ==================================
        // OpenRouter
        // ==================================

        else if (
            selectedProvider ===
            "openrouter"
        ) {

            answer =
                await streamOpenRouterAI(
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

        }


        // ==================================
        // OpenAI
        // ==================================

        else if (
            selectedProvider ===
            "openai"
        ) {

            answer =
                await streamOpenAI(
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

        }


        else {

            throw new Error(
                "مزود الذكاء الاصطناعي غير معروف: " +
                selectedProvider
            );

        }


        // ==================================
        // إلغاء المؤقت
        // ==================================

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


        pendingRenderText =
            String(
                answer ||
                ""
            );


        renderStreamingText();


        // ==================================
        // إزالة التحميل
        // ==================================

        if (
            loading &&
            loading.parentNode
        ) {

            loading.remove();

        }


        // ==================================
        // حفظ إجابة AI
        // ==================================

        currentChat.messages.push({

            role:
                "ai",

            text:
                String(
                    answer ||
                    ""
                ),

            citationSources:
                Array.isArray(
                    currentCitationSources
                )
                    ? currentCitationSources.map(
                        function (
                            source
                        ) {

                            return {

                                rank:
                                    source.rank,

                                paragraphIndex:
                                    source.paragraphIndex,

                                heading:
                                    source.heading ||
                                    "",

                                text:
                                    source.text ||
                                    ""

                            };

                        }
                    )
                    : []

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();

    }
    catch (error) {

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


        currentChat.messages.push({

            role:
                "ai",

            text:
                "خطأ: " +
                (
                    error &&
                    error.message
                        ? error.message
                        : "حدث خطأ غير معروف"
                )

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();


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
// Final Initialization
// =====================================================

initializeSidebarSections();

renderProjects();

renderExpandedProjects();

renderDocuments();

renderChatList();

renderSidebarChats();

renderRecentChats();

renderChat();

loadSettings();


// =====================================================
// End Office.onReady
// =====================================================

});


// =====================================================
// Sidebar Pin
// يبقى خارج Office.onReady
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


    if (
        sidebarPinned
    ) {

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
// PART 4 END
// =====================================================