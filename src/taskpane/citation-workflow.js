"use strict";

function normalizeCitationIdentity(value) {
    return String(value || "")
        .normalize("NFKC")
        .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
        .replace(/[\u200C-\u200F\u202A-\u202E]/g, "")
        .replace(/[أإآٱ]/g, "ا")
        .replace(/ى/g, "ي")
        .replace(/\s+/g, " ")
        .replace(/[،,:؛.()[\]{}«»"'ـ]/g, "")
        .trim()
        .toLowerCase();
}

function parseCitationJSON(answer) {
    const raw = String(answer || "")
        .trim()
        .replace(/^\x60{3}(?:json)?\s*/i, "")
        .replace(/\s*\x60{3}$/i, "")
        .trim();

    try {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.citations)) return parsed.citations;
    } catch (_) {
        // Try the first complete JSON object below.
    }

    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start >= 0 && end > start) {
        const parsed = JSON.parse(raw.slice(start, end + 1));
        if (parsed && Array.isArray(parsed.citations)) return parsed.citations;
    }

    throw new Error("تعذر قراءة نتيجة تحليل التوثيقات بصيغة منظمة.");
}

function createCitationSystemPrompt(options) {
    const settings = options || {};
    const authorDiscoveryInstruction = settings.addMissingAuthors === true
        ? "خيار استقصاء المؤلف مفعّل: إذا غاب اسم المؤلف عن الحاشية، فحدّد مؤلف الكتاب من معرفتك الببليوغرافية الموثوقة، وأعد الاسم العلمي المختصر. لا تخمّن النسبة؛ اترك المؤلف فارغًا إذا لم تكن متأكدًا."
        : "خيار استقصاء المؤلف غير مفعّل: لا تستكمل اسم مؤلف غائب عن الحاشية أو مواد السياق؛ اترك حقل author فارغًا إذا لم يرد الاسم.";

    return [
        "أنت محلل توثيقات أكاديمية متخصص في الحواشي العربية.",
        authorDiscoveryInstruction,
        "مهمتك استخراج الإحالات الببليوغرافية الموجودة في الحواشي السفلية والختامية فقط.",
        "لا تنشئ قائمة مراجع، ولا تنفذ توحيدًا ببليوغرافيًا عامًا، ولا تكتب شرحًا.",
        "أخرج كل إحالة مستقلة بوصفها occurrence منفصلًا، مع الاحتفاظ بمعرّف الحاشية.",
        "sourceText يجب أن يكون مقطعًا حرفيًا متصلًا من نص الحاشية المستهدفة، ويشمل هوية المؤلف والكتاب والبيانات الملحقة وموضع الاستشهاد متى كانت ضمن الإحالة.",
        "لا تضم إلى sourceText كلام الباحث التفسيري أو الاقتباس أو الإحالة المستقلة المجاورة.",
        "لا تختصر sourceText ولا تعِد صياغته ولا تصحح النص الأصلي.",
        "أعد author بصيغة الاسم العلمي المختصر التي تظهر أو تتضح من الحواشي المرسلة، وauthorFull بالصيغة الكاملة إذا وردت صراحة فقط. لا تخترع اسمًا.",
        "أعد title كما يدل عليه نص الإحالة. لا تدمج كتابًا بشرحه أو حاشيته أو مختصره، ولا تساوِ بين عنوانين لمجرد التشابه.",
        "استخرج volume وpage وpageRange كلًّا على حدة عندما تكون واضحة، ولا تستنتج قيمة غير موجودة.",
        "لا تعتبر تخريج الحديث أو الإحالة القرآنية أو الإحالة الداخلية إلى صفحات الدراسة مرجعًا كتابيًا.",
        "يمكن استخدام مواد السياق لفهم المصدر نفسه أو المرجع نفسه، لكن لا تستخرج منها إحالات جديدة. إذا تعذر تعيين المرجع السابق بثقة فاجعل needsReview=true.",
        "إذا تكرر المرجع في أكثر من حاشية، استخدم الصيغة نفسها للمؤلف والعنوان في النتائج متى كانت هويته واحدة.",
        "أعد JSON فقط بهذه البنية: {\"citations\":[{\"materialId\":\"\",\"sourceText\":\"\",\"author\":\"\",\"authorFull\":\"\",\"title\":\"\",\"type\":\"book\",\"volume\":\"\",\"page\":\"\",\"pageRange\":\"\",\"confidence\":0.0,\"needsReview\":false}]}",
        "لا تُسقط إحالة واضحة بسبب نقص بياناتها؛ اترك الحقول غير المعروفة فارغة وارفع needsReview عند الالتباس."
    ].join("\n");
}

function createCitationBatches(notes, maxChars, maxNotes) {
    const batches = [];
    let current = [];
    let currentChars = 0;

    (Array.isArray(notes) ? notes : []).forEach(function (note) {
        const text = String(note && (note.originalText || note.text) || "");
        if (!String(note && note.id || "").trim() || !text.trim()) return;

        const itemChars = JSON.stringify({
            materialId: String(note.id),
            source: String(note.source || ""),
            noteNumber: note.noteNumber == null ? null : note.noteNumber,
            text: text
        }).length;

        if (current.length && (currentChars + itemChars > maxChars || current.length >= maxNotes)) {
            batches.push(current);
            current = [];
            currentChars = 0;
        }

        current.push({
            id: String(note.id),
            source: String(note.source || ""),
            noteNumber: note.noteNumber == null ? null : note.noteNumber,
            originalText: text
        });
        currentChars += itemChars;
    });

    if (current.length) batches.push(current);
    return batches;
}

function createCitationUserPrompt(batch, contextNotes) {
    return JSON.stringify({
        targetNotes: (Array.isArray(batch) ? batch : []).map(function (note) {
            return {
                materialId: note.id,
                source: note.source,
                noteNumber: note.noteNumber == null ? null : note.noteNumber,
                text: note.originalText
            };
        }),
        contextOnly: (Array.isArray(contextNotes) ? contextNotes : []).map(function (note) {
            return {
                materialId: note.id,
                source: note.source,
                noteNumber: note.noteNumber == null ? null : note.noteNumber,
                text: String(note.originalText || "").slice(0, 700)
            };
        }),
        instruction: "استخرج إحالات targetNotes فقط. لا تنشئ أي نتيجة من contextOnly، وحافظ على materialId وsourceText حرفيًا."
    });
}

async function fetchCitationJSON(fetchImpl, url, requestOptions, readJSON, provider, timeoutMs) {
    const controller = typeof AbortController !== "undefined"
        ? new AbortController()
        : null;
    const providerNames = {
        duckai: "Duck.ai",
        gemini: "Gemini",
        groq: "Groq",
        openai: "OpenAI",
        openrouter: "OpenRouter",
        pollinations: "Pollinations"
    };
    const providerLabel = providerNames[provider] || provider;
    const timeoutError = new Error(
        "تجاوز الاتصال بمزوّد " + providerLabel +
        " مهلة الانتظار (" + Math.ceil(timeoutMs / 1000) +
        " ثانية). لم يكتمل التحليل؛ تحقق من اتصال الإنترنت وإعدادات المزوّد ثم أعد المحاولة."
    );
    timeoutError.code = "CITATION_REQUEST_TIMEOUT";
    timeoutError.retryable = true;

    let timedOut = false;
    let timeoutId = null;
    const timeoutPromise = new Promise(function (_, reject) {
        timeoutId = setTimeout(function () {
            timedOut = true;
            if (controller) controller.abort();
            reject(timeoutError);
        }, timeoutMs);
    });

    const requestPromise = (async function () {
        const options = Object.assign({}, requestOptions || {});
        if (controller) options.signal = controller.signal;

        const response = await fetchImpl(url, options);
        const result = await readJSON(response);
        return { response: response, result: result };
    })();

    try {
        return await Promise.race([requestPromise, timeoutPromise]);
    } catch (error) {
        if (timedOut || (error && error.name === "AbortError")) {
            throw timeoutError;
        }
        throw error;
    } finally {
        if (timeoutId !== null) clearTimeout(timeoutId);
    }
}

async function requestCitationCompletion(systemPrompt, userPrompt, settings, dependencies) {
    const provider = String(settings && settings.provider || "openrouter").trim().toLowerCase();
    const key = String(settings && settings.key || "").trim();
    const model = String(settings && settings.model || "").trim();
    const fetchImpl = dependencies.fetch || fetch;
    const requestTimeoutMs = Number(dependencies.requestTimeoutMs) > 0
        ? Math.floor(Number(dependencies.requestTimeoutMs))
        : 60000;

    if (!model) throw new Error("لم يتم تحديد نموذج الذكاء الاصطناعي.");
    if (provider !== "duckai" && !key) throw new Error("لم يتم إدخال مفتاح الذكاء الاصطناعي.");

    const readJSON = dependencies.readJSON;
    const getAPIError = dependencies.getAPIError;
    if (typeof readJSON !== "function" || typeof getAPIError !== "function") {
        throw new Error("خدمة الاتصال بالذكاء الاصطناعي غير مكتملة.");
    }

    if (provider === "gemini") {
        const normalizeGeminiModel = dependencies.normalizeGeminiModel;
        const extractGeminiAnswer = dependencies.extractGeminiAnswer;
        if (typeof normalizeGeminiModel !== "function" || typeof extractGeminiAnswer !== "function") {
            throw new Error("أدوات الاتصال بـ Gemini غير متاحة.");
        }

        const request = await fetchCitationJSON(
            fetchImpl,
            "https://generativelanguage.googleapis.com/v1beta/models/" +
                encodeURIComponent(normalizeGeminiModel(model)) +
                ":generateContent?key=" + encodeURIComponent(key),
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    systemInstruction: { parts: [{ text: systemPrompt }] },
                    contents: [{ role: "user", parts: [{ text: userPrompt }] }],
                    generationConfig: {
                        temperature: 0.1,
                        responseMimeType: "application/json",
                        maxOutputTokens: 6000
                    }
                })
            },
            readJSON,
            provider,
            requestTimeoutMs
        );
        const response = request.response;
        const result = request.result;
        if (!response.ok) throw new Error(getAPIError(result, "فشل تحليل التوثيقات عبر Gemini."));
        if (String(result && result.candidates && result.candidates[0] && result.candidates[0].finishReason || "").toUpperCase() === "MAX_TOKENS") {
            throw new Error("توقف تحليل التوثيقات بسبب حد الإخراج؛ لم تُعتمد النتيجة الجزئية.");
        }

        const answer = extractGeminiAnswer(result);
        if (!String(answer || "").trim()) throw new Error("أعاد Gemini استجابة فارغة أثناء تحليل التوثيقات.");
        return answer;
    }

    const endpoint = provider === "duckai"
        ? "http://127.0.0.1:8080/v1/chat/completions"
        : provider === "openai"
            ? "https://api.openai.com/v1/chat/completions"
            : provider === "groq"
                ? "https://api.groq.com/openai/v1/chat/completions"
                : provider === "pollinations"
                    ? "https://gen.pollinations.ai/v1/chat/completions"
                    : "https://openrouter.ai/api/v1/chat/completions";

    const headers = { "Content-Type": "application/json" };
    if (provider !== "duckai") headers.Authorization = "Bearer " + key;
    if (provider === "openrouter") {
        headers["HTTP-Referer"] = typeof window !== "undefined" ? window.location.href : "";
        headers["X-Title"] = "Research Tools";
    }

    const request = await fetchCitationJSON(
        fetchImpl,
        endpoint,
        {
            method: "POST",
            headers: headers,
            body: JSON.stringify({
                model: model,
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "user", content: userPrompt }
                ],
                temperature: 0.1,
                response_format: { type: "json_object" },
                max_tokens: provider === "groq" ? 4096 : 7000
            })
        },
        readJSON,
        provider,
        requestTimeoutMs
    );
    const response = request.response;
    const result = request.result;
    if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بمزود الذكاء الاصطناعي أثناء تحليل التوثيقات."));

    const finishReason = String(result && result.choices && result.choices[0] && result.choices[0].finish_reason || "").toLowerCase();
    if (["length", "max_tokens", "max_output_tokens"].includes(finishReason)) {
        throw new Error("توقف تحليل التوثيقات بسبب حد الإخراج؛ لم تُعتمد النتيجة الجزئية.");
    }

    const extractOpenAIStyleAnswer = dependencies.extractOpenAIStyleAnswer;
    if (typeof extractOpenAIStyleAnswer !== "function") {
        throw new Error("أداة قراءة استجابة نموذج الذكاء الاصطناعي غير متاحة.");
    }

    const answer = extractOpenAIStyleAnswer(result, provider);
    if (!String(answer || "").trim()) throw new Error("أعاد مزود الذكاء الاصطناعي استجابة فارغة أثناء تحليل التوثيقات.");
    return answer;
}

async function analyzeCitations(notes, settings, dependencies, options) {
    const provider = String(settings && settings.provider || "openrouter").trim().toLowerCase();
    const maxChars = provider === "groq" ? 4200 : provider === "duckai" ? 6500 : 8500;
    const batches = createCitationBatches(notes, maxChars, 12);
    if (!batches.length) return { citations: [], batchCount: 0, noteCount: 0 };

    const systemPrompt = createCitationSystemPrompt(options);
    const allCitations = [];
    const allNotes = Array.isArray(notes) ? notes : [];
    const reportProgress = typeof dependencies.onProgress === "function"
        ? function (details) {
            try {
                dependencies.onProgress(Object.assign({}, details));
            } catch (error) {
                // Progress UI must never interrupt citation extraction.
                if (typeof console !== "undefined" && console.warn) {
                    console.warn("تعذر تحديث مؤشر تقدم التوثيقات:", error);
                }
            }
        }
        : function () {};
    const totalNotes = batches.reduce(function (total, batch) {
        return total + batch.length;
    }, 0);
    let completedNotes = 0;

    reportProgress({
        stage: "started",
        totalBatches: batches.length,
        completedBatches: 0,
        noteCount: totalNotes,
        completedNotes: 0,
        percent: 10
    });

    for (let index = 0; index < batches.length; index++) {
        const batch = batches[index];
        reportProgress({
            stage: "batch-start",
            currentBatch: index + 1,
            totalBatches: batches.length,
            completedBatches: index,
            batchNoteCount: batch.length,
            noteCount: totalNotes,
            completedNotes: completedNotes,
            percent: 10 + Math.round((completedNotes / totalNotes) * 62)
        });
        const firstIndex = allNotes.findIndex(function (note) {
            return String(note && note.id || "") === batch[0].id;
        });
        const contextNotes = firstIndex > 0
            ? allNotes.slice(Math.max(0, firstIndex - 2), firstIndex)
            : [];
        const userPrompt = createCitationUserPrompt(batch, contextNotes);

        let answer = "";
        let lastError = null;
        const attempts = provider === "duckai" ? 2 : 1;

        for (let attempt = 0; attempt < attempts; attempt++) {
            try {
                answer = await requestCitationCompletion(systemPrompt, userPrompt, settings, dependencies);
                break;
            } catch (error) {
                lastError = error;
                if (attempt + 1 >= attempts) throw error;
                await new Promise(function (resolve) {
                    setTimeout(resolve, 450 * (attempt + 1));
                });
            }
        }

        if (!answer && lastError) throw lastError;
        const batchCitations = parseCitationJSON(answer);
        const allowedIds = new Set(batch.map(function (note) { return note.id; }));

        batchCitations.forEach(function (citation) {
            const materialId = String(citation && citation.materialId || "").trim();
            if (!allowedIds.has(materialId)) return;
            allCitations.push(Object.assign({}, citation, { materialId: materialId }));
        });

        completedNotes += batch.length;
        reportProgress({
            stage: "batch-complete",
            currentBatch: index + 1,
            totalBatches: batches.length,
            completedBatches: index + 1,
            batchNoteCount: batch.length,
            noteCount: totalNotes,
            completedNotes: completedNotes,
            batchCitationCount: batchCitations.length,
            citationCount: allCitations.length,
            percent: 10 + Math.round((completedNotes / totalNotes) * 62)
        });
    }

    const order = new Map(allNotes.map(function (note, index) {
        return [String(note && note.id || ""), index];
    }));

    allCitations.sort(function (a, b) {
        return (order.get(a.materialId) == null ? Number.MAX_SAFE_INTEGER : order.get(a.materialId)) -
            (order.get(b.materialId) == null ? Number.MAX_SAFE_INTEGER : order.get(b.materialId));
    });

    reportProgress({
        stage: "complete",
        totalBatches: batches.length,
        completedBatches: batches.length,
        noteCount: totalNotes,
        completedNotes: totalNotes,
        citationCount: allCitations.length,
        percent: 74
    });

    return {
        citations: allCitations,
        batchCount: batches.length,
        noteCount: totalNotes
    };
}

function buildUnifiedCitationReferences(citations, materials, validateSourceSpan) {
    const notesById = new Map();
    (Array.isArray(materials) ? materials : []).forEach(function (note) {
        notesById.set(String(note && note.id || ""), note);
    });

    const grouped = new Map();
    let skippedCount = 0;

    (Array.isArray(citations) ? citations : []).forEach(function (citation) {
        const materialId = String(citation && citation.materialId || "").trim();
        const note = notesById.get(materialId);
        const title = String(citation && citation.title || "").trim();
        const author = String(citation && citation.author || "").trim();
        const sourceTextCandidate = String(citation && citation.sourceText || "").trim();

        if (!note || !title || !sourceTextCandidate) {
            skippedCount++;
            return;
        }

        const originalText = String(note.originalText || note.text || "");
        const sourceText = typeof validateSourceSpan === "function"
            ? String(validateSourceSpan(originalText, sourceTextCandidate) || "")
            : (originalText.includes(sourceTextCandidate) ? sourceTextCandidate : "");

        if (!sourceText) {
            skippedCount++;
            return;
        }

        const type = String(citation && citation.type || "book").trim().toLowerCase();
        const identityKey = [
            type,
            normalizeCitationIdentity(author),
            normalizeCitationIdentity(title)
        ].join("|");

        let reference = grouped.get(identityKey);
        if (!reference) {
            reference = {
                id: "citation-" + (grouped.size + 1),
                type: type,
                author: author,
                authorFull: String(citation && citation.authorFull || "").trim(),
                title: title,
                locations: [],
                variants: [],
                occurrences: [],
                confidence: Number(citation && citation.confidence || 0),
                needsReview: citation && citation.needsReview === true
            };
            grouped.set(identityKey, reference);
        } else {
            if (!reference.authorFull && citation && citation.authorFull) {
                reference.authorFull = String(citation.authorFull).trim();
            }
            reference.confidence = Math.min(
                Number(reference.confidence || 0),
                Number(citation && citation.confidence || 0)
            );
            reference.needsReview = reference.needsReview || (citation && citation.needsReview === true);
        }

        const occurrence = {
            materialId: materialId,
            source: String(note.source || ""),
            noteNumber: note.noteNumber == null ? null : note.noteNumber,
            sourceText: sourceText,
            sourceTextConfidence: Number(citation && citation.confidence || 0),
            volume: String(citation && citation.volume || "").trim(),
            page: String(citation && citation.page || "").trim(),
            pageRange: String(citation && citation.pageRange || "").trim()
        };

        reference.occurrences.push(occurrence);
        reference.variants.push({
            text: sourceText,
            materialId: materialId,
            source: String(note.source || ""),
            noteNumber: note.noteNumber == null ? null : note.noteNumber,
            confidence: Number(citation && citation.confidence || 0)
        });

        const locationKey = [occurrence.volume, occurrence.page, occurrence.pageRange].join("|");
        if (
            (occurrence.volume || occurrence.page || occurrence.pageRange) &&
            !reference.locations.some(function (location) {
                return [
                    String(location && location.volume || ""),
                    String(location && location.page || ""),
                    String(location && location.pageRange || "")
                ].join("|") === locationKey;
            })
        ) {
            reference.locations.push({
                volume: occurrence.volume,
                page: occurrence.page,
                pageRange: occurrence.pageRange
            });
        }
    });

    return {
        references: Array.from(grouped.values()),
        skippedCount: skippedCount
    };
}

module.exports = {
    normalizeCitationIdentity: normalizeCitationIdentity,
    parseCitationJSON: parseCitationJSON,
    createCitationSystemPrompt: createCitationSystemPrompt,
    createCitationBatches: createCitationBatches,
    createCitationUserPrompt: createCitationUserPrompt,
    analyzeCitations: analyzeCitations,
    buildUnifiedCitationReferences: buildUnifiedCitationReferences
};
