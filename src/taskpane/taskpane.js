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
// ======================================

function setCurrentDocument(
    documentItem
) {

    if (!documentItem) {

        currentDocument =
            null;


        oramaDocumentDb =
            null;


        oramaDocumentId =
            null;


        oramaRetrievalDb =
            null;


        oramaRetrievalCacheKey =
            "";


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
    // إعادة استخدام البنية الموجودة
    // ==================================

    if (
        documentItem.readStatus ===
        "read"
    ) {

        ensureDocumentStructure(
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
                    "تعذر تحديث بنية المستند:",
                    error
                );

                renderDocuments();

            }
        );


        renderDocuments();


        return;

    }


    // ==================================
    // قراءة المستند الجديد
    // ==================================

    readCurrentWordDocument(
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
            // العائلة الصرفية
            // ==================================

            const matchedFamilies =
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
// Search Indexed Document
// Orama + Arabic Families + Heading Ranking
// + Query Profile + Comparison Relation
// =====================================================

async function searchIndexedDocument(
    documentId,
    query,
    options
) {

    const settings =
        options || {};


    // ==================================
    // نوع السؤال
    // ==================================

    const retrievalProfile =
        settings.profile ||
        getRetrievalProfile(
            query
        ).type;


    // ==================================
    // تطبيع السؤال
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
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            comparisonConceptFamilies:
                [],

            totalQueryTerms:
                0,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // المستند
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


    if (!documentItem) {

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


    if (!structureData) {

        throw new Error(
            "لا توجد بنية محفوظة لهذا المستند."
        );

    }


    // ==================================
    // فهرس Orama
    // ==================================

    const db =
        await ensureOramaRetrievalIndex(
            documentItem,
            structureData
        );


    if (!db) {

        throw new Error(
            "تعذر إنشاء فهرس Orama."
        );

    }


    // ==================================
    // كلمات البحث المهمة
    // ==================================

    const queryTokens =
        getSearchQueryTokens(
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
                [],

            matchedTerms:
                [],

            matchedFamilies:
                [],

            comparisonConceptFamilies:
                [],

            totalQueryTerms:
                0,

            indexedOccurrences:
                0

        };

    }


    // ==================================
    // عائلات كلمات السؤال
    // ==================================

    const queryFamilies =
        [];


    queryTokens.forEach(
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


    // ==================================
    // في المقارنة:
    // بعض الكلمات تصف نوع السؤال
    // وليست مفاهيم المقارنة نفسها
    // ==================================

    const comparisonMarkerFamilies =
        new Set([

            "فرق",
            "فروق",
            "مقارن",
            "مقارنه",
            "تمييز",
            "خلاف"

        ]);


    const comparisonConceptFamilies =
        retrievalProfile ===
        "comparison"

            ? queryFamilies.filter(
                function (
                    family
                ) {

                    return !comparisonMarkerFamilies.has(
                        family
                    );

                }
            )

            : [];


    // ==================================
    // كل نتائج Orama
    // ==================================

    const searchResults =
        [];


    // ==================================
    // Orama Search
    // ==================================

    function runOramaSearch(
        term,
        exact
    ) {

        if (
            !term
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


            return window.Orama.search(
                db,
                {

                    term:
                        term,

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
                        100

                }
            );

        }
        catch (error) {

            console.warn(
                "فشل بحث Orama:",
                term,
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
    // البحث في السؤال كاملًا
    // ==================================

    const fullQueryResult =
        runOramaSearch(
            searchTerm,
            false
        );


    if (
        fullQueryResult &&
        Array.isArray(
            fullQueryResult.hits
        )
    ) {

        searchResults.push(
            ...fullQueryResult.hits
        );

    }


    // ==================================
    // البحث في الكلمات المهمة
    // يساعد الأسئلة المركبة والطويلة
    // ==================================

    queryTokens.forEach(
        function (
            token
        ) {

            const result =
                runOramaSearch(
                    token,
                    false
                );


            if (
                result &&
                Array.isArray(
                    result.hits
                )
            ) {

                searchResults.push(
                    ...result.hits
                );

            }

        }
    );


    // ==================================
    // الكلمات المطابقة فعليًا
    // ==================================

    const matchedTerms =
        queryTokens.filter(
            function (
                token
            ) {

                const normalizedToken =
                    normalizeSearchText(
                        token
                    );


                return searchResults.some(
                    function (
                        hit
                    ) {

                        if (
                            !hit ||
                            !hit.document
                        ) {

                            return false;

                        }


                        const text =
                            normalizeSearchText(
                                hit.document.text ||
                                ""
                            );


                        const heading =
                            normalizeSearchText(
                                hit.document.heading ||
                                ""
                            );


                        return (
                            text.includes(
                                normalizedToken
                            ) ||
                            heading.includes(
                                normalizedToken
                            )
                        );

                    }
                );

            }
        );


    // ==================================
    // العائلات الموجودة فعليًا
    // ==================================

    const matchedFamilies =
        queryFamilies.filter(
            function (
                family
            ) {

                return searchResults.some(
                    function (
                        hit
                    ) {

                        if (
                            !hit ||
                            !hit.document
                        ) {

                            return false;

                        }


                        const text =
                            String(
                                hit.document.text ||
                                ""
                            );


                        const tokens =
                            tokenizeDocumentText(
                                text
                            );


                        return tokens.some(
                            function (
                                token
                            ) {

                                return (
                                    getConservativeFamilyKey(
                                        token,
                                        null
                                    ) ===
                                    family
                                );

                            }
                        );

                    }
                );

            }
        );


    // ==================================
    // إزالة التكرار حسب الفقرة
    // ==================================

    const merged =
        new Map();


    searchResults.forEach(
        function (
            hit
        ) {

            if (
                !hit ||
                !hit.document
            ) {

                return;

            }


            const paragraphIndex =
                Number(
                    hit.document.paragraphIndex
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
                    hit.score ||
                    0
                );


            const existing =
                merged.get(
                    paragraphIndex
                );


            if (!existing) {

                merged.set(
                    paragraphIndex,
                    {

                        document:
                            hit.document,

                        oramaScore:
                            oramaScore

                    }
                );

            }
            else {

                existing.oramaScore =
                    Math.max(
                        existing.oramaScore,
                        oramaScore
                    );

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
            : [];


    // =================================================
    // Profile Score
    // ترجيح النتيجة بحسب نوع السؤال
    // =================================================

    function getProfileScore(
        resultData
    ) {

        if (
            !resultData
        ) {

            return 0;

        }


        const text =
            normalizeSearchText(
                resultData.text ||
                ""
            );


        const heading =
            normalizeSearchText(
                resultData.heading ||
                ""
            );


        let profileScore =
            0;


        // ==================================
        // Definition
        // ==================================

        if (
            retrievalProfile ===
            "definition"
        ) {

            if (
                heading.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    25;

            }


            if (
                text.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    15;

            }


            if (
                /تعريف|المقصود|المراد|يعني|يقصد|هو ان|هي ان/
                    .test(
                        text
                    )
            ) {

                profileScore +=
                    12;

            }

        }


        // ==================================
        // Effect
        // ==================================

        else if (
            retrievalProfile ===
            "effect"
        ) {

            if (
                /اثر|تاثير|نتيجة|ينتج|يترتب|انعكاس/
                    .test(
                        text
                    )
            ) {

                profileScore +=
                    12;

            }


            if (
                heading.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    15;

            }

        }


        // ==================================
        // Comparison
        // ==================================

        else if (
            retrievalProfile ===
            "comparison"
        ) {

            if (
                /الفرق|الفروق|مقارنة|يقارن|التمييز|يفترق|خلاف/
                    .test(
                        text
                    )
            ) {

                profileScore +=
                    12;

            }


            if (
                heading.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    15;

            }

        }


        // ==================================
        // Causes
        // ==================================

        else if (
            retrievalProfile ===
            "causes"
        ) {

            if (
                /سبب|اسباب|علة|علل|لان|بسبب|من اسباب/
                    .test(
                        text
                    )
            ) {

                profileScore +=
                    12;

            }


            if (
                heading.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    15;

            }

        }


        // ==================================
        // Location
        // ==================================

        else if (
            retrievalProfile ===
            "location"
        ) {

            if (
                /الفصل|المبحث|المطلب|الباب|الصفحة|موضع|موضعه/
                    .test(
                        text
                    )
            ) {

                profileScore +=
                    10;

            }


            if (
                heading.includes(
                    searchTerm
                )
            ) {

                profileScore +=
                    20;

            }

        }


        return profileScore;

    }


    // ==================================
    // النتائج النهائية
    // ==================================

    const results =
        [];


    merged.forEach(
        function (
            item
        ) {

            const paragraph =
                item.document;


            const originalText =
                String(
                    paragraph.text ||
                    ""
                );


            const normalizedText =
                normalizeSearchText(
                    originalText
                );


            // ==================================
            // كلمات السؤال الموجودة في الفقرة
            // ==================================

            const matchedTokenCount =
                queryTokens.filter(
                    function (
                        token
                    ) {

                        return normalizedText.includes(
                            normalizeSearchText(
                                token
                            )
                        );

                    }
                ).length;


            const queryCoverage =
                queryTokens.length >
                0

                    ? matchedTokenCount /
                      queryTokens.length

                    : 0;


            // ==================================
            // عائلات الفقرة
            // ==================================

            const paragraphFamilies =
                [];


            const paragraphTokens =
                tokenizeDocumentText(
                    originalText
                );


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


            // ==================================
            // عدد عائلات السؤال الموجودة
            // ==================================

            const matchedFamilyCount =
                matchedFamilies.filter(
                    function (
                        family
                    ) {

                        return paragraphFamilies.includes(
                            family
                        );

                    }
                ).length;


            const familyCoverage =
                matchedFamilies.length >
                0

                    ? matchedFamilyCount /
                      matchedFamilies.length

                    : 0;


            // ==================================
            // تغطية مفاهيم المقارنة
            // ==================================

            let comparisonCoverage =
                0;


            if (
                retrievalProfile ===
                "comparison"
            ) {

                const matchedComparisonFamilies =
                    comparisonConceptFamilies.filter(
                        function (
                            family
                        ) {

                            return paragraphFamilies.includes(
                                family
                            );

                        }
                    );


                comparisonCoverage =
                    comparisonConceptFamilies.length >
                    0

                        ? matchedComparisonFamilies.length /
                          comparisonConceptFamilies.length

                        : 0;

            }


            // ==================================
            // الفلترة الأولية
            // ==================================

            const minimumCoverage =
                queryTokens.length >= 3
                    ? 0.50
                    : 0.34;


            if (
                retrievalProfile ===
                    "comparison"
            ) {

                // لا نحذف المقاطع التي تحمل مفهومًا
                // واحدًا فقط، لأن بعضها قد يكون مهمًا
                // في شرح أحد طرفي المقارنة.
                //
                // لكن يجب أن تحمل على الأقل
                // مفهومًا من المفاهيم الأساسية.

                if (
                    comparisonCoverage ===
                        0 &&
                    familyCoverage <
                        minimumCoverage
                ) {

                    return;

                }

            }
            else {

                if (
                    queryCoverage <
                        minimumCoverage &&
                    familyCoverage <
                        minimumCoverage
                ) {

                    return;

                }

            }


            // ==================================
            // أقرب عنوان
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
                    headings[i];


                if (
                    Number(
                        heading.index
                    ) <=
                    Number(
                        paragraph.paragraphIndex
                    )
                ) {

                    nearestHeading =
                        heading;

                    break;

                }

            }


            // ==================================
            // درجة العنوان
            // ==================================

            let headingScore =
                0;


            let headingCoverage =
                0;


            if (
                nearestHeading
            ) {

                const normalizedHeading =
                    normalizeSearchText(
                        nearestHeading.text
                    );


                queryTokens.forEach(
                    function (
                        token
                    ) {

                        const normalizedToken =
                            normalizeSearchText(
                                token
                            );


                        if (
                            normalizedHeading.includes(
                                normalizedToken
                            )
                        ) {

                            headingScore +=
                                6;

                        }

                    }
                );


                const matchedHeadingTokens =
                    queryTokens.filter(
                        function (
                            token
                        ) {

                            return normalizedHeading.includes(
                                normalizeSearchText(
                                    token
                                )
                            );

                        }
                    ).length;


                headingCoverage =
                    queryTokens.length >
                    0

                        ? matchedHeadingTokens /
                          queryTokens.length

                        : 0;


                if (
                    normalizedHeading ===
                    searchTerm
                ) {

                    headingScore +=
                        40;

                }
                else if (
                    normalizedHeading.includes(
                        searchTerm
                    )
                ) {

                    headingScore +=
                        25;

                }


                // ==================================
                // عنوان يجمع مفاهيم المقارنة
                // ==================================

                if (
                    retrievalProfile ===
                    "comparison" &&
                    comparisonConceptFamilies.length >
                        0
                ) {

                    let headingConceptCount =
                        0;


                    comparisonConceptFamilies.forEach(
                        function (
                            family
                        ) {

                            const headingTokens =
                                tokenizeDocumentText(
                                    normalizedHeading
                                );


                            const found =
                                headingTokens.some(
                                    function (
                                        token
                                    ) {

                                        return (
                                            getConservativeFamilyKey(
                                                token,
                                                null
                                            ) ===
                                            family
                                        );

                                    }
                                );


                            if (
                                found
                            ) {

                                headingConceptCount +=
                                    1;

                            }

                        }
                    );


                    if (
                        headingConceptCount ===
                        comparisonConceptFamilies.length
                    ) {

                        headingScore +=
                            25;

                    }

                    else if (
                        headingConceptCount >
                        0
                    ) {

                        headingScore +=
                            8;

                    }

                }

            }


            // ==================================
            // قرب الكلمات
            // ==================================

            let proximityScore =
                0;


            // ==================================
            // في المقارنة:
            // نحسب قرب مفاهيم المقارنة الأساسية
            // بدل الاعتماد على كلمات مثل "الفرق"
            // ==================================

            if (
                retrievalProfile ===
                "comparison" &&
                comparisonConceptFamilies.length >
                    1
            ) {

                const familyPositions =
                    [];


                paragraphTokens.forEach(
                    function (
                        token,
                        tokenIndex
                    ) {

                        const family =
                            getConservativeFamilyKey(
                                token,
                                null
                            );


                        if (
                            comparisonConceptFamilies.includes(
                                family
                            )
                        ) {

                            familyPositions.push({

                                family:
                                    family,

                                position:
                                    tokenIndex

                            });

                        }

                    }
                );


                if (
                    familyPositions.length >
                    1
                ) {

                    familyPositions.sort(
                        function (
                            a,
                            b
                        ) {

                            return (
                                a.position -
                                b.position
                            );

                        }
                    );


                    let bestSpan =
                        null;


                    for (
                        let left = 0;

                        left <
                            familyPositions.length;

                        left++
                    ) {

                        const firstFamily =
                            familyPositions[
                                left
                            ].family;


                        for (
                            let right =
                                left + 1;

                            right <
                                familyPositions.length;

                            right++
                        ) {

                            const secondFamily =
                                familyPositions[
                                    right
                                ].family;


                            if (
                                firstFamily ===
                                    secondFamily
                            ) {

                                continue;

                            }


                            const span =
                                familyPositions[
                                    right
                                ].position -
                                familyPositions[
                                    left
                                ].position +
                                1;


                            if (
                                bestSpan ===
                                    null ||
                                span <
                                    bestSpan
                            ) {

                                bestSpan =
                                    span;

                            }

                        }

                    }


                    if (
                        bestSpan !==
                            null
                    ) {

                        proximityScore =
                            20 /
                            Math.max(
                                bestSpan,
                                1
                            );

                    }

                }

            }

            else {

                // ==================================
                // الأسئلة غير المقارنة
                // ==================================

                const tokenPositions =
                    [];


                queryTokens.forEach(
                    function (
                        token
                    ) {

                        const normalizedToken =
                            normalizeSearchText(
                                token
                            );


                        const position =
                            normalizedText.indexOf(
                                normalizedToken
                            );


                        if (
                            position >=
                                0
                        ) {

                            tokenPositions.push(
                                position
                            );

                        }

                    }
                );


                if (
                    tokenPositions.length >
                    1
                ) {

                    tokenPositions.sort(
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
                        tokenPositions[
                            tokenPositions.length - 1
                        ] -
                        tokenPositions[0] +
                        1;


                    proximityScore =
                        20 /
                        Math.max(
                            span,
                            1
                        );

                }

            }


            // ==================================
            // التطابق الكامل
            // ==================================

            const exactPhrase =
                normalizedText.includes(
                    searchTerm
                );


            // ==================================
            // Relation Score
            // قياس وجود علاقة صريحة
            // بين مفاهيم المقارنة
            // ==================================

            let relationScore =
                0;


            if (
                retrievalProfile ===
                "comparison"
            ) {

                const normalizedOriginalText =
                    normalizedText;


                const relationPatterns = [

                    /علاق[ةه]\s+بين/,

                    /الفرق\s+بين/,

                    /الفروق\s+بين/,

                    /مقارن[ةه]\s+بين/,

                    /يقارن\s+بين/,

                    /التمييز\s+بين/,

                    /يفترق/,

                    /خلاف\s+بين/

                ];


                const hasRelationPattern =
                    relationPatterns.some(
                        function (
                            pattern
                        ) {

                            return pattern.test(
                                normalizedOriginalText
                            );

                        }
                    );


                if (
                    hasRelationPattern
                ) {

                    relationScore +=
                        30;

                }


                // اجتماع جميع المفاهيم
                if (
                    comparisonCoverage >=
                    1
                ) {

                    relationScore +=
                        20;

                }


                // قرب المفاهيم
                if (
                    proximityScore >
                    0.5
                ) {

                    relationScore +=
                        15;

                }


                // تعزيز إضافي عندما يجتمع
                // المفهومان في جملة قصيرة
                if (
                    comparisonCoverage >=
                        1 &&
                    proximityScore >=
                        1
                ) {

                    relationScore +=
                        15;

                }

            }


            // ==================================
            // الدرجة الأساسية
            // Orama هو الأساس
            // ==================================

            let score =
                Number(
                    item.oramaScore ||
                    0
                );


            // ==================================
            // تغطية السؤال
            // ==================================

            score +=
                queryCoverage *
                20;


            // ==================================
            // تغطية العائلات
            // ==================================

            score +=
                familyCoverage *
                15;


            // ==================================
            // أولوية العنوان
            // ==================================

            score +=
                headingScore;


            // ==================================
            // تغطية العنوان
            // ==================================

            score +=
                headingCoverage *
                15;


            // ==================================
            // قرب الكلمات
            // ==================================

            score +=
                proximityScore;


            // ==================================
            // التطابق الكامل
            // ==================================

            if (
                exactPhrase
            ) {

                score +=
                    15;

            }


            // ==================================
            // المطابقة متعددة المفاهيم
            // ==================================

            if (
                matchedFamilyCount >=
                2
            ) {

                score +=
                    20;

            }


            if (
                matchedTokenCount >=
                2
            ) {

                score +=
                    15;

            }


            // ==================================
            // درجة نوع السؤال
            // ==================================

            const profileScore =
                getProfileScore({

                    text:
                        originalText,

                    heading:
                        nearestHeading
                            ? nearestHeading.text
                            : ""

                });


            score +=
                profileScore;


            // ==================================
            // اجتماع مفاهيم المقارنة
            // ==================================

            if (
                retrievalProfile ===
                "comparison"
            ) {

                score +=
                    comparisonCoverage *
                    25;

            }


            // ==================================
            // علاقة المفاهيم
            // ==================================

            score +=
                relationScore;


            // ==================================
            // أولوية المقارنة بين المفاهيم
            //
            // لا نحذف النتيجة،
            // وإنما نعطي أفضلية للمقاطع
            // التي تجمع جميع المفاهيم الأساسية.
            // ==================================

            if (
                retrievalProfile ===
                    "comparison" &&
                comparisonCoverage <
                    1
            ) {

                score *=
                    0.65;

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
                retrievalProfile ===
                    "comparison" &&
                comparisonCoverage >=
                    1 &&
                relationScore >
                    0
            ) {

                matchType =
                    "comparison";

            }
            else if (
                headingScore >
                0
            ) {

                matchType =
                    "heading";

            }
            else if (
                matchedFamilyCount >
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


            const firstToken =
                queryTokens.find(
                    function (
                        token
                    ) {

                        return normalizedText.includes(
                            normalizeSearchText(
                                token
                            )
                        );

                    }
                );


            if (
                firstToken
            ) {

                const normalizedFirstToken =
                    normalizeSearchText(
                        firstToken
                    );


                const position =
                    normalizedText.indexOf(
                        normalizedFirstToken
                    );


                const start =
                    Math.max(
                        0,
                        position -
                        120
                    );


                const end =
                    Math.min(
                        normalizedText.length,
                        position +
                        normalizedFirstToken.length +
                        320
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

            results.push({

                paragraphIndex:
                    paragraph.paragraphIndex,

                paragraphId:
                    paragraph.paragraphId,

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

                score:
                    score,

                oramaScore:
                    item.oramaScore,

                profile:
                    retrievalProfile,

                profileScore:
                    profileScore,

                relationScore:
                    relationScore,

                comparisonCoverage:
                    comparisonCoverage,

                matchType:
                    matchType,

                matchedTerms:
                    matchedTerms,

                matchedFamilies:
                    matchedFamilies,

                matchedFamilyCount:
                    matchedFamilyCount,

                queryCoverage:
                    queryCoverage,

                familyCoverage:
                    familyCoverage,

                headingScore:
                    headingScore,

                headingCoverage:
                    headingCoverage,

                proximityScore:
                    proximityScore,

                exactPhrase:
                    exactPhrase

            });

        }
    );


    // ==================================
    // الترتيب النهائي
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
                retrievalProfile ===
                "comparison"
            ) {

                if (
                    b.comparisonCoverage !==
                    a.comparisonCoverage
                ) {

                    return (
                        b.comparisonCoverage -
                        a.comparisonCoverage
                    );

                }


                if (
                    b.relationScore !==
                    a.relationScore
                ) {

                    return (
                        b.relationScore -
                        a.relationScore
                    );

                }

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


            if (
                b.familyCoverage !==
                a.familyCoverage
            ) {

                return (
                    b.familyCoverage -
                    a.familyCoverage
                );

            }


            return (
                a.paragraphIndex -
                b.paragraphIndex
            );

        }
    );


    // ==================================
    // تحديد عدد المرشحين
    // قبل بناء السياق
    // ==================================

    const candidateLimit =
        retrievalProfile ===
            "comparison"

            ? 30

            : retrievalProfile ===
                "definition"

                ? 20

                : 25;


    if (
        results.length >
        candidateLimit
    ) {

        results.splice(
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


    const allDocumentTokens =
        tokenizeDocumentText(
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
                )
        );


    matchedFamilies.forEach(
        function (
            family
        ) {

            allDocumentTokens.forEach(
                function (
                    token
                ) {

                    if (
                        getConservativeFamilyKey(
                            token,
                            null
                        ) ===
                        family
                    ) {

                        indexedOccurrences +=
                            1;

                    }

                }
            );

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
            matchedTerms,

        matchedFamilies:
            matchedFamilies,

        comparisonConceptFamilies:
            comparisonConceptFamilies,

        totalQueryTerms:
            queryTokens.length,

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
// تحويل نتائج Orama إلى سياق ذكي
// =====================================================

async function buildRetrievalContext(
    searchResult,
    options
) {

    const settings =
        options || {};


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

            let structureData = null;

    if (currentDocument) {

        try {

            structureData =
                await ensureDocumentStructure(
                    currentDocument
                );

        }
        catch (error) {

            console.warn(
                "تعذر تحميل بنية المستند للسياق:",
                error
            );

        }

    }

    const paragraphMap =
        new Map();

    if (
        structureData &&
        Array.isArray(
            structureData.paragraphs
        )
    ) {

        structureData.paragraphs.forEach(
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

    }
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
    // خريطة فقرات المستند
    // ==================================

    const paragraphMap =
        new Map();


    if (
        currentDocument
    ) {

        try {

            const structureData =
                await ensureDocumentStructure(
                    currentDocument
                );


            if (
                structureData &&
                Array.isArray(
                    structureData.paragraphs
                )
            ) {

                structureData.paragraphs.forEach(
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

            }

        }
        catch (error) {

            console.warn(
                "تعذر تحميل بنية الفقرات للسياق:",
                error
            );

        }

    }

    // ==================================
    // ترتيب النتائج
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

                const scoreA =
                    Number(
                        a.score ||
                        0
                    );

                const scoreB =
                    Number(
                        b.score ||
                        0
                    );


                if (
                    scoreB !==
                    scoreA
                ) {

                    return (
                        scoreB -
                        scoreA
                    );

                }


                // ==================================
                // عند التعادل:
                // نفضل نتيجة تجمع مفاهيم المقارنة
                // ==================================

                const relationA =
                    Number(
                        a.relationScore ||
                        0
                    );

                const relationB =
                    Number(
                        b.relationScore ||
                        0
                    );


                if (
                    relationB !==
                    relationA
                ) {

                    return (
                        relationB -
                        relationA
                    );

                }


                const coverageA =
                    Number(
                        a.queryCoverage ||
                        0
                    );

                const coverageB =
                    Number(
                        b.queryCoverage ||
                        0
                    );


                if (
                    coverageB !==
                    coverageA
                ) {

                    return (
                        coverageB -
                        coverageA
                    );

                }


                return (
                    Number(
                        a.paragraphIndex ||
                        0
                    ) -
                    Number(
                        b.paragraphIndex ||
                        0
                    )
                );

            }
        );


    // ==================================
    // اختيار النتائج
    // ==================================

    const selected =
        [];


    const selectedParagraphIndexes =
        new Set();


    const MAX_TEXT_OVERLAP =
        0.75;


    for (
        let i = 0;
        i < results.length;
        i++
    ) {

        const candidate =
            results[i];


        if (
            selectedParagraphIndexes.has(
                Number(
                    candidate.paragraphIndex
                )
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


        if (!candidateText) {

            continue;

        }


        let tooSimilar =
            false;


        for (
            let j = 0;
            j < selected.length;
            j++
        ) {

            const selectedText =
                String(
                    selected[j].context ||
                    selected[j].text ||
                    ""
                )
                    .replace(
                        /\s+/g,
                        " "
                    )
                    .trim();


            if (!selectedText) {

                continue;

            }


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


        if (tooSimilar) {

            continue;

        }


        selected.push(
            candidate
        );


        selectedParagraphIndexes.add(
            Number(
                candidate.paragraphIndex
            )
        );


        if (
            selected.length >=
            maxResults
        ) {

            break;

        }

    }



    // =================================================
    // بناء السياقات
    // =================================================

    const contexts =
        [];


    let totalChars =
        0;


    // -------------------------------------------------
    // جلب بنية المستند بطريقة متزامنة مع async wrapper
    // -------------------------------------------------
    // buildRetrievalContext نفسها بقيت synchronous
    // لذلك نستخدم البيانات الموجودة داخل النتائج أولًا.
    // ويتم إكمال previous/next من نتائج Orama إن وجدت.
    // -------------------------------------------------

    selected.forEach(
        function (
            result,
            index
        ) {

            let mainContext =
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


            if (!mainContext) {

                return;

            }


            let previousContext =
                "";

            let nextContext =
                "";


            // ==================================
            // الفقرة السابقة من المستند الحقيقي
            // ==================================

            if (
                includeNeighbors &&
                paragraphMap.size > 0
            ) {

                const previousParagraph =
                    paragraphMap.get(
                        Number(
                            result.paragraphIndex
                        ) - 1
                    );


                if (
                    previousParagraph &&
                    previousParagraph.text
                ) {

                    previousContext =
                        String(
                            previousParagraph.text
                        )
                            .replace(
                                /\s+/g,
                                " "
                            )
                            .trim();

                }

            }


            // ==================================
            // الفقرة التالية من المستند الحقيقي
            // ==================================

            if (
                includeNeighbors &&
                paragraphMap.size > 0
            ) {

                const nextParagraph =
                    paragraphMap.get(
                        Number(
                            result.paragraphIndex
                        ) + 1
                    );


                if (
                    nextParagraph &&
                    nextParagraph.text
                ) {

                    nextContext =
                        String(
                            nextParagraph.text
                        )
                            .replace(
                                /\s+/g,
                                " "
                            )
                            .trim();

                }

            }


            // ==================================
            // منع التكرار
            // ==================================

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


            const heading =
                String(
                    result.heading ||
                    ""
                )
                    .trim();


            const remainingChars =
                maxChars -
                totalChars;


            if (
                remainingChars <=
                0
            ) {

                return;

            }


            const reservedForMetadata =
                250;


            const availableChars =
                Math.max(
                    300,
                    remainingChars -
                    reservedForMetadata
                );


            let context =
                mainContext;


            let remainingForNeighbors =
                availableChars -
                context.length;


            // ==================================
            // السابق
            // ==================================

            if (
                includeNeighbors &&
                previousContext &&
                remainingForNeighbors >
                    150
            ) {

                const allowedPreviousLength =
                    Math.max(
                        0,
                        remainingForNeighbors -
                        1
                    );


                if (
                    allowedPreviousLength >
                    100
                ) {

                    const previousPart =
                        previousContext.length >
                            allowedPreviousLength
                            ? previousContext.substring(
                                Math.max(
                                    0,
                                    previousContext.length -
                                    allowedPreviousLength
                                )
                            ) +
                            "…"
                            : previousContext;


                    context =
                        previousPart +
                        " " +
                        context;

                }

            }


            remainingForNeighbors =
                availableChars -
                context.length;


            // ==================================
            // التالي
            // ==================================

            if (
                includeNeighbors &&
                nextContext &&
                remainingForNeighbors >
                    150
            ) {

                const allowedNextLength =
                    Math.max(
                        0,
                        remainingForNeighbors -
                        1
                    );


                if (
                    allowedNextLength >
                    100
                ) {

                    const nextPart =
                        nextContext.length >
                            allowedNextLength
                            ? nextContext.substring(
                                0,
                                allowedNextLength
                            ) +
                            "…"
                            : nextContext;


                    context =
                        context +
                        " " +
                        nextPart;

                }

            }


            // ==================================
            // العنصر النهائي
            // ==================================

            const item = {

                rank:
                    index + 1,

                paragraphIndex:
                    Number(
                        result.paragraphIndex
                    ),

                paragraphId:
                    String(
                        result.paragraphId ||
                        result.paragraphIndex
                    ),

                heading:
                    heading,

                score:
                    Number(
                        result.score ||
                        0
                    ),

                matchType:
                    result.matchType ||
                    "family",

                matchedFamilies:
                    Array.isArray(
                        result.matchedFamilies
                    )
                        ? result.matchedFamilies
                        : [],

                matchedTerms:
                    Array.isArray(
                        result.matchedTerms
                    )
                        ? result.matchedTerms
                        : [],

                familyOccurrences:
                    Number(
                        result.familyOccurrencesInParagraph ||
                        0
                    ),

                exactWordMatches:
                    Number(
                        result.exactWordMatches ||
                        0
                    ),

                previousParagraph:
                    previousContext,

                mainParagraph:
                    mainContext,

                nextParagraph:
                    nextContext,

                context:
                    context,

                source:
                    result

            };


            contexts.push(
                item
            );


            totalChars +=
                context.length;

        }
    );


    // =================================================
    // بناء النص المرسل إلى الذكاء الاصطناعي
    // =================================================

    const textParts =
        [];


    contexts.forEach(
        function (
            item
        ) {

            let block =
                "[مقطع " +
                item.rank +
                "]" +
                "\n";


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


    const finalText =
        textParts.join(
            "\n\n---\n\n"
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
            finalText

    };

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


        // ==================================
        // البحث عبر Orama
        // ==================================

        const searchResult =
            await searchIndexedDocument(
                currentDocument.id,
                query,
                {

                    profile:
                        retrievalProfile.type,

                    maxResults:
                        retrievalProfile.maxResults

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


        // ==================================
        // إعدادات المزود
        // ==================================

        const settings =
            getSavedSettings();


        const retrievalLimits =
            getRetrievalLimits(
                settings.provider,
                settings.model
            );


        let maxResults =
            Math.min(
                retrievalLimits.maxResults,
                retrievalProfile.maxResults
            );


        let maxChars =
            Math.min(
                retrievalLimits.maxChars,
                retrievalProfile.maxChars
            );


        // ==================================
        // المقارنة تحتاج مساحة أكبر
        // ==================================

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
        // بناء السياق
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
        // حفظ مصادر الإحالات
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

                            paragraphId:
                                item.paragraphId,

                            heading:
                                item.heading ||
                                "",

                            mainParagraph:
                                item.mainParagraph ||
                                "",

                            text:
                                item.context ||
                                item.mainParagraph ||
                                "",

                            score:
                                Number(
                                    item.score ||
                                    0
                                ),

                            matchType:
                                item.matchType ||
                                "family"

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

            text:
                "",

            error:
                error &&
                error.message
                    ? error.message
                    : String(error)

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
// Format AI Message
// تحويل Markdown + إحالات المقاطع
// =====================================================

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


        // ==================================
        // تحويل الإحالات
        // [مقطع 1]
        // [مقطع 1، 2]
        // [مقطع 1, 2]
        // ==================================

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
                        ranks
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


// =====================================================
// Citation Click Handler
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
// الانتقال إلى المقطع الأصلي في Word
// =====================================================

async function openCitationInWord(
    rank
) {

    const source =
        getCitationSource(
            rank
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


                // ==================================
                // النص الأساسي للمقطع
                // ==================================

                let searchText =
                    prepareCitationSearchText(
                        source.mainParagraph ||
                        source.text ||
                        "",
                        180
                    );


                if (!searchText) {

                    throw new Error(
                        "لا يوجد نص صالح للمقطع."
                    );

                }


                // ==================================
                // المحاولة الأولى
                // ==================================

                let results =
                    await searchCitationInWord(
                        body,
                        searchText
                    );


                if (
                    results.length >
                    0
                ) {

                    results[0].select(
                        "Select"
                    );


                    await context.sync();

                    return;

                }


                // ==================================
                // المحاولة الثانية
                // نص أقصر
                // ==================================

                const fallbackText =
                    prepareCitationSearchText(
                        source.mainParagraph ||
                        source.text ||
                        "",
                        80
                    );


                if (
                    fallbackText &&
                    fallbackText !==
                        searchText
                ) {

                    results =
                        await searchCitationInWord(
                            body,
                            fallbackText
                        );


                    if (
                        results.length >
                        0
                    ) {

                        results[0].select(
                            "Select"
                        );


                        await context.sync();

                        return;

                    }

                }


                // ==================================
                // المحاولة الثالثة
                // استخدام جزء من السياق
                // ==================================

                const contextText =
                    prepareCitationSearchText(
                        source.text ||
                        source.mainParagraph ||
                        "",
                        50
                    );


                if (
                    contextText &&
                    contextText !==
                        searchText &&
                    contextText !==
                        fallbackText
                ) {

                    results =
                        await searchCitationInWord(
                            body,
                            contextText
                        );


                    if (
                        results.length >
                        0
                    ) {

                        results[0].select(
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
    catch (error) {

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
// Ensure Citation Sources Are Stable
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

            return {

                rank:
                    Number(
                        source.rank ||
                        0
                    ),

                paragraphIndex:
                    Number(
                        source.paragraphIndex ||
                        0
                    ),

                paragraphId:
                    String(
                        source.paragraphId ||
                        ""
                    ),

                heading:
                    String(
                        source.heading ||
                        ""
                    ),

                mainParagraph:
                    String(
                        source.mainParagraph ||
                        ""
                    ),

                text:
                    String(
                        source.text ||
                        ""
                    ),

                score:
                    Number(
                        source.score ||
                        0
                    ),

                matchType:
                    String(
                        source.matchType ||
                        "family"
                    )

            };

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
// Build Streaming User Content
// =====================================================

async function buildStreamingContext(
    text
) {

    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const history =
        buildStreamingHistory(
            documentContext
        );


    let userContent =
        String(
            text ||
            ""
        );


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

                "ادمج المقاطع التي تحمل الفكرة نفسها.",

                "لا تحول كل مقطع مستخرج إلى فقرة مستقلة.",

                "ضع الإحالة [مقطع رقم] بعد الفكرة التي يدعمها المقطع.",

                "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع الإحالات.",

                "لا تكرر الإحالة نفسها دون فائدة.",

                "لا تخترع أرقام مقاطع.",

                "لا تنسب فكرة إلى مقطع لا يدعمها.",

                "وجود كلمات السؤال داخل المقطع لا يعني أن المقطع صالح للإجابة.",

                "قدّم خلاصة تركيبية للمادة المستخرجة."

            ].join(
                "\n"
            );

    }


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
    // ==================================

    const systemInstruction = [

        "أنت مساعد بحث أكاديمي يعمل على مستندات Word.",

        "اعتمد على المادة المستخرجة من المستند بوصفها المصدر الأساسي للإجابة.",

        "أجب عن السؤال مباشرة وبأسلوب أكاديمي واضح.",

        "رتب الإجابة وفق محاور السؤال.",

        "إذا كان السؤال يتضمن أكثر من جانب، فافصل بينها بوضوح.",

        "ادمج الأفكار المتشابهة في صياغة واحدة.",

        "لا تحول كل مقطع مستخرج إلى فقرة مستقلة.",

        "استبعد المعلومة الجانبية التي لا تجيب مباشرة عن السؤال.",

        "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع إحالاتها.",

        "لا تكرر الفكرة نفسها لمجرد ورودها في أكثر من مقطع.",

        "لا تضف معلومة أو حكمًا أو نسبة قول إلى المستند غير موجودة في المادة المستخرجة.",

        "إذا لم تكف المادة المستخرجة للإجابة عن جزء من السؤال، صرّح بذلك.",

        "لا تستخدم المعرفة العامة لسد نقص المستند إلا إذا طلب المستخدم ذلك صراحة.",

        "لا تذكر مشكلة الدراسة أو أهدافها أو منهجها أو أسئلتها إلا إذا طلب المستخدم ذلك صراحة.",

        "ضع الإحالات بعد الأفكار التي يدعمها المستند بصيغة [مقطع X].",

        "إذا كانت الفكرة مدعومة بأكثر من مقطع، استخدم إحالات متعددة.",

        "لا تخترع أرقام المقاطع.",

        "حافظ على لغة السؤال ولغة المستند.",

        "لا تبدأ باعتذار أو تمهيد غير ضروري.",

        "لا تعيد صياغة سؤال المستخدم.",

        "قدّم إجابة تركيبية مترابطة."

    ].join(
        "\n"
    );


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

        loadingElement.innerHTML =
            formatAIMessage(
                value,
                currentCitationSources
            );

    }


    if (chatArea) {

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
// Render Documents
// =====================================================

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
                    (
                        documentItem.indexTokenCount ||
                        0
                    ) +
                    " كلمة · " +
                    (
                        documentItem.indexUniqueTerms ||
                        0
                    ) +
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
            // Rename Document
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
            // Delete Document
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

                                        setCurrentDocument(
                                            null
                                        );

                                    }


                                    const remaining =
                                        currentProject
                                            ? getProjectDocuments(
                                                currentProject.id
                                            )
                                            : [];


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
// Render Projects
// =====================================================

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
                                async function () {

                                    const projectDocumentIds =
                                        Array.isArray(
                                            project.documents
                                        )
                                            ? project.documents
                                            : [];


                                    for (
                                        let i = 0;
                                        i <
                                        projectDocumentIds.length;
                                        i++
                                    ) {

                                        const id =
                                            projectDocumentIds[i];


                                        const doc =
                                            documents.find(
                                                function (
                                                    d
                                                ) {

                                                    return (
                                                        d &&
                                                        d.id ===
                                                            id
                                                    );

                                                }
                                            );


                                        if (!doc) {

                                            continue;

                                        }


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


                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return !projectDocumentIds.includes(
                                                    doc.id
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


                                        setCurrentDocument(
                                            null
                                        );

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
// Expanded Sidebar Toggle
// =====================================================

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
                            sidebarTogglePlaceholder
                                .nextSibling
                        );

                }

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

renderDocuments();

renderChatList();

renderSidebarChats();

renderRecentChats();

renderChat();

loadSettings();


// =====================================================
// Initial State
// =====================================================

if (currentProject) {

    renderDocuments();

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