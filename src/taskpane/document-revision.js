(function (global) {
    "use strict";

    const MAX_DIFF_CELLS = 2000000;
    let pendingTarget = null;

    function isRequest(value) {
        const text = String(value || "").trim();
        if (!text) return false;

        return /(?:صحح|تصحيح|دقق|تدقيق|نقح|تنقيح|أعد\s*صياغ|إعادة\s*صياغ|صياغة|حرر\s+(?:النص|الفقرة)|راجع\s+(?:لغوي|النص|الفقرة|المستند)|حسّن\s+(?:الأسلوب|الصياغة|العبارة|ترابط)|حسن\s+(?:الأسلوب|الصياغة|العبارة|ترابط)|تحسين\s+(?:الصياغة|الأسلوب|الترابط)|اجعل\s+(?:النص|العبارة|الفقرة)\s+(?:أوضح|أكثر\s+(?:وضوحا|وضوحًا|رصانة|سلاسة|أكاديمية)))/i.test(text);
    }

    function buildRequest(userInstruction, originalText) {
        return {
            systemPrompt: [
                "أنت محرر لغوي وأكاديمي متخصص في تحرير النصوص العربية والبحوث العلمية.",
                "افهم طلب الباحث من عبارته وحدد بنفسك نوع العمل: تصحيح لغوي، أو إعادة صياغة، أو الجمع بينهما. لا تطلب منه اختيار نمط من قائمة.",
                "في التصحيح اللغوي: أصلح الأخطاء الإملائية والنحوية والصرفية والتركيبية وعلامات الترقيم، ولا تغير الصياغة الصحيحة لمجرد تحسينها.",
                "في إعادة الصياغة: حسّن بناء الجمل وترابطها ووضوحها ورصانتها، مع المحافظة الدقيقة على المعنى والأفكار والمعلومات والاستدلالات.",
                "في الجمع بينهما: نفذ التصحيح والتحسين بالقدر الذي يقتضيه طلب الباحث، دون توسع غير مطلوب.",
                "في جميع الأحوال: لا تختلق معلومة أو شاهدًا أو إحالة أو مصدرًا، ولا تحذف مضمونًا علميًا، ولا تستبدل المصطلحات المتخصصة بمرادفات قد تغير دلالتها. حافظ على ترتيب الفقرات وفواصلها ولا تضف تنسيق Markdown.",
                "حافظ على الآيات القرآنية والأحاديث والنقول الحرفية والمراجع والتوثيقات كما وردت، ولا تعد صياغتها.",
                "إذا كان طلب الباحث غامضًا إلى حد يمنع معرفة التعديل المقصود، فأعد سؤال استيضاح واحدًا ولا تعدل النص.",
                "أعد JSON صالحًا فقط، دون Markdown أو سياج شيفرة أو نص خارجه.",
                "عند إمكان تنفيذ الطلب استخدم: {\"status\":\"ready\",\"mode\":\"correction أو rewrite أو combined\",\"text\":\"النص كاملًا بعد المعالجة\"}.",
                "عند الحاجة إلى الاستيضاح استخدم: {\"status\":\"clarify\",\"question\":\"سؤال واحد محدد\"}.",
                "قيمة mode يجب أن تكون واحدة من هذه القيم حرفيًا: correction أو rewrite أو combined.",
                "لا تضف مقدمة أو خاتمة إلى النص الناتج."
            ].join("\n"),
            userPrompt: [
                "طلب الباحث:",
                String(userInstruction || "").trim(),
                "",
                "النص المطلوب معالجته:",
                "<<<BEGIN_TEXT>>>",
                String(originalText || ""),
                "<<<END_TEXT>>>",
                "",
                "نفذ الطلب على النص المحصور بين العلامتين فقط، وأعد النص كاملًا."
            ].join("\n")
        };
    }

    function parseResponse(value) {
        let raw = String(value || "").trim();
        raw = raw.replace(/^\uFEFF/, "");
        raw = raw.replace(/^\`\`\`(?:json)?\s*/i, "").replace(/\s*\`\`\`$/i, "");

        const first = raw.indexOf("{");
        const last = raw.lastIndexOf("}");
        if (first >= 0 && last > first) raw = raw.slice(first, last + 1);

        let parsed;
        try {
            parsed = JSON.parse(raw);
        } catch (_) {
            throw new Error("تعذر التحقق من نتيجة المعالجة؛ لم تُطبق أي تغييرات في المستند.");
        }

        if (parsed && parsed.status === "clarify" && String(parsed.question || "").trim()) {
            return { status: "clarify", question: String(parsed.question).trim() };
        }

        const mode = String(parsed && parsed.mode || "");
        const text = String(parsed && parsed.text || "");
        if (!parsed || parsed.status !== "ready" || !text.trim() ||
            !["correction", "rewrite", "combined"].includes(mode)) {
            throw new Error("نتيجة المعالجة غير مكتملة أو غير صالحة؛ لم يُمس المستند.");
        }

        return { status: "ready", mode: mode, text: text };
    }

    function normalizeLineBreaks(originalValue, revisedValue) {
        const original = String(originalValue || "");
        const revised = String(revisedValue || "");
        const originalBreaks = original.match(/\r\n|\r|\n/g) || [];
        const revisedBreaks = revised.match(/\r\n|\r|\n/g) || [];
        if (originalBreaks.length !== revisedBreaks.length) return revised;

        let index = 0;
        return revised.replace(/\r\n|\r|\n/g, function () {
            return originalBreaks[index++];
        });
    }

    function tokenize(value) {
        const text = String(value || "");
        const regex = /\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu;
        const tokens = [];
        let match;
        while ((match = regex.exec(text)) !== null) {
            tokens.push({ value: match[0], start: match.index, end: match.index + match[0].length });
        }
        return tokens;
    }

    function splitText(value, maxChars) {
        const text = String(value || "");
        const limit = Math.max(1000, Number(maxChars) || 4200);
        if (text.length <= limit) return [text];

        const chunks = [];
        let start = 0;
        while (start < text.length) {
            let end = Math.min(start + limit, text.length);
            if (end < text.length) {
                const minimum = start + Math.floor(limit * 0.55);
                const paragraphBreak = Math.max(
                    text.lastIndexOf("\n", end),
                    text.lastIndexOf("\r", end)
                );
                const sentenceBreak = Math.max(
                    text.lastIndexOf(". ", end),
                    text.lastIndexOf("؟ ", end),
                    text.lastIndexOf("! ", end),
                    text.lastIndexOf("؛ ", end),
                    text.lastIndexOf("۔ ", end)
                );
                const spaceBreak = text.lastIndexOf(" ", end);

                if (paragraphBreak >= minimum) {
                    end = paragraphBreak + 1;
                    if (text[paragraphBreak] === "\r" && text[paragraphBreak + 1] === "\n") end++;
                } else if (sentenceBreak >= minimum) {
                    end = sentenceBreak + 1;
                } else if (spaceBreak >= minimum) {
                    end = spaceBreak + 1;
                }
            }

            if (end <= start) end = Math.min(start + limit, text.length);
            chunks.push(text.slice(start, end));
            start = end;
        }

        return chunks;
    }

    function fallbackDiff(oldText, newText) {
        let start = 0;
        const maxPrefix = Math.min(oldText.length, newText.length);
        while (start < maxPrefix && oldText[start] === newText[start]) start++;

        let oldEnd = oldText.length;
        let newEnd = newText.length;
        while (oldEnd > start && newEnd > start && oldText[oldEnd - 1] === newText[newEnd - 1]) {
            oldEnd--;
            newEnd--;
        }

        if (oldEnd === start && newEnd === start) return [];
        return [{ start: start, length: oldEnd - start, replacement: newText.slice(start, newEnd) }];
    }

    function diffText(oldValue, newValue) {
        const oldText = String(oldValue || "");
        const newText = String(newValue || "");
        if (oldText === newText) return [];

        const a = tokenize(oldText);
        const b = tokenize(newText);
        const columns = b.length + 1;
        const cells = (a.length + 1) * columns;

        if (cells > MAX_DIFF_CELLS) return fallbackDiff(oldText, newText);

        const matrix = new Uint32Array(cells);
        for (let i = a.length - 1; i >= 0; i--) {
            const row = i * columns;
            const nextRow = (i + 1) * columns;
            for (let j = b.length - 1; j >= 0; j--) {
                matrix[row + j] = a[i].value === b[j].value
                    ? matrix[nextRow + j + 1] + 1
                    : Math.max(matrix[nextRow + j], matrix[row + j + 1]);
            }
        }

        const operations = [];
        let i = 0;
        let j = 0;
        let oldOffset = 0;
        let pending = null;

        function begin() {
            if (!pending) pending = { start: oldOffset, end: oldOffset, replacement: "" };
        }
        function flush() {
            if (!pending) return;
            operations.push({
                start: pending.start,
                length: pending.end - pending.start,
                replacement: pending.replacement
            });
            pending = null;
        }

        while (i < a.length || j < b.length) {
            if (i < a.length && j < b.length && a[i].value === b[j].value) {
                flush();
                oldOffset = a[i].end;
                i++;
                j++;
                continue;
            }

            const deleteScore = i < a.length ? matrix[(i + 1) * columns + j] : -1;
            const insertScore = j < b.length ? matrix[i * columns + j + 1] : -1;

            if (i < a.length && (j >= b.length || deleteScore >= insertScore)) {
                begin();
                oldOffset = a[i].end;
                pending.end = oldOffset;
                i++;
            } else if (j < b.length) {
                begin();
                pending.replacement += b[j].value;
                j++;
            }
        }

        flush();
        return operations;
    }

    async function releasePendingTarget() {
        if (!pendingTarget || !pendingTarget.range) return;
        const previous = pendingTarget;
        pendingTarget = null;
        try {
            await Word.run(async function (context) {
                previous.range.untrack();
                await context.sync();
            });
        } catch (_) {}
    }

    async function capture() {
        if (typeof Word === "undefined" || !Word.run) {
            throw new Error("اتصال Word غير متاح؛ افتح مستند Word ثم أعد المحاولة.");
        }

        await releasePendingTarget();
        let result = null;
        await Word.run(async function (context) {
            const selection = context.document.getSelection();
            selection.load("text");
            const bodyRange = context.document.body.getRange();
            bodyRange.load("text");
            await context.sync();

            const selectedText = String(selection.text || "");
            const targetType = selectedText.trim() ? "selection" : "document";
            const range = targetType === "selection" ? selection : bodyRange;
            const originalText = String(range.text || "");

            if (!originalText.trim()) {
                throw new Error("لا يوجد نص قابل للمعالجة في التحديد أو المستند المفتوح.");
            }

            range.track();
            await context.sync();

            result = {
                id: "revision-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
                targetType: targetType,
                originalText: originalText
            };
            pendingTarget = { id: result.id, range: range };
        });

        return result;
    }

    async function resolveTarget(context, revision) {
        const body = context.document.body;
        if (revision.targetType === "document") {
            return body.getRange();
        }

        if (pendingTarget && pendingTarget.id === revision.id && pendingTarget.range) {
            context.trackedObjects.add(pendingTarget.range);
            return pendingTarget.range;
        }

        const matches = body.search(String(revision.originalText || ""), {
            matchCase: true,
            matchWholeWord: false
        });
        matches.load("items");
        await context.sync();
        matches.items.forEach(function (match) { match.load("text,start,end"); });
        await context.sync();

        const exactMatches = matches.items.filter(function (match) {
            return String(match.text || "") === String(revision.originalText || "");
        });
        if (exactMatches.length !== 1) {
            throw new Error("تعذر تحديد موضع النص الأصلي بصورة فريدة. أعد تحديد النص في Word ثم أرسل طلب المعالجة من جديد.");
        }

        return exactMatches[0];
    }

    async function apply(revision) {
        if (!revision || !revision.originalText || !revision.text) {
            throw new Error("بيانات التعديل غير مكتملة.");
        }
        if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
            throw new Error("إصدار Word الحالي لا يدعم تحديد مواضع التغييرات بدقة؛ لم يُعدّل المستند.");
        }

        const originalText = String(revision.originalText);
        const operations = diffText(originalText, revision.text);
        if (!operations.length) return { applied: false, unchanged: true, changedParts: 0 };

        let appliedCount = 0;

        await Word.run(async function (context) {
            const range = await resolveTarget(context, revision);
            range.load("text,start");
            await context.sync();

            if (String(range.text || "") !== originalText) {
                throw new Error("تغير النص الأصلي في Word بعد إرساله للمعالجة؛ أُلغي التطبيق حمايةً للمستند. أعد المعالجة على النص الحالي.");
            }

            const prepared = [];
            for (const operation of operations) {
                if (operation.length > 0) {
                    const oldText = originalText.slice(operation.start, operation.start + operation.length);
                    if (/[\r\n]/.test(oldText)) {
                        throw new Error("تتضمن النتيجة تغييرًا في فواصل الفقرات؛ أُلغي التطبيق حتى لا تتغير بنية المستند.");
                    }
                    const search = range.search(oldText, { matchCase: true, matchWholeWord: false });
                    search.load("items");
                    prepared.push({ operation: operation, search: search, kind: "replace" });
                } else {
                    const tokens = tokenize(originalText);
                    const nextToken = tokens.find(function (token) { return token.start === operation.start; });
                    const previousToken = tokens.slice().reverse().find(function (token) { return token.end === operation.start; });
                    const anchor = nextToken && !/^\s+$/.test(nextToken.value)
                        ? { token: nextToken, location: "Before" }
                        : previousToken && !/^\s+$/.test(previousToken.value)
                            ? { token: previousToken, location: "After" }
                            : null;

                    if (operation.start === 0) {
                        prepared.push({ operation: operation, kind: "boundary-start" });
                    } else if (operation.start === originalText.length) {
                        prepared.push({ operation: operation, kind: "boundary-end" });
                    } else if (anchor) {
                        const search = range.search(anchor.token.value, { matchCase: true, matchWholeWord: false });
                        search.load("items");
                        prepared.push({ operation: operation, search: search, kind: "insert", anchor: anchor });
                    } else {
                        throw new Error("تعذر تحديد موضع إضافة إحدى العبارات بأمان؛ لم يُعدّل المستند.");
                    }
                }
            }

            await context.sync();

            const resolved = [];
            for (const item of prepared) {
                if (item.kind === "boundary-start") {
                    resolved.push({ ...item, target: range.getRange("Start") });
                    continue;
                }
                if (item.kind === "boundary-end") {
                    resolved.push({ ...item, target: range.getRange("End") });
                    continue;
                }

                item.search.items.forEach(function (match) { match.load("text,start,end"); });
                await context.sync();
                const expectedStart = Number(range.start) + Number(
                    item.kind === "insert" && item.anchor.location === "After"
                        ? item.anchor.token.start
                        : item.operation.start
                );
                const matches = item.search.items.filter(function (match) {
                    return Number(match.start) === expectedStart &&
                        (item.kind !== "replace" || String(match.text || "") === originalText.slice(
                            item.operation.start,
                            item.operation.start + item.operation.length
                        ));
                });
                if (matches.length !== 1) {
                    throw new Error("تعذر مطابقة أحد مواضع التغيير بصورة دقيقة؛ أُلغي التطبيق دون تعديل المستند.");
                }
                resolved.push({ ...item, target: matches[0] });
            }

            resolved.sort(function (a, b) { return b.operation.start - a.operation.start; });
            const insertedRanges = [];
            resolved.forEach(function (item) {
                const location = item.kind === "insert" ? item.anchor.location : "Replace";
                const inserted = item.target.insertText(item.operation.replacement, location);
                if (item.operation.replacement) {
                    inserted.font.color = "#FF0000";
                    insertedRanges.push(inserted);
                }
            });

            insertedRanges.forEach(function (item) { item.load("text"); });
            await context.sync();
            appliedCount = insertedRanges.length;

            if (pendingTarget && pendingTarget.id === revision.id && pendingTarget.range) {
                try {
                    pendingTarget.range.untrack();
                    await context.sync();
                } catch (_) {}
                pendingTarget = null;
            }
        });

        return { applied: true, unchanged: false, changedParts: appliedCount };
    }

    const api = {
        isRequest: isRequest,
        buildRequest: buildRequest,
        parseResponse: parseResponse,
        tokenize: tokenize,
        diffText: diffText,
        splitText: splitText,
        normalizeLineBreaks: normalizeLineBreaks,
        capture: capture,
        apply: apply
    };

    global.DocumentRevision = api;
    if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
