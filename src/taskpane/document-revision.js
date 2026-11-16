(function (global) {
    "use strict";

    const DocumentRevisionPreflight = require("./document-revision-preflight");

    const MAX_DIFF_CELLS = 2000000;
    let pendingTarget = null;

    function isRequest(value) {
        const text = String(value || "").trim();
        if (!text) return false;

        return /(?:صحح|تصحيح|دقق|تدقيق|نقح|تنقيح|أعد\s*صياغ|إعادة\s*صياغ|صياغة|حرر\s+(?:النص|الفقرة)|راجع\s+(?:لغوي|النص|الفقرة|المستند)|حسّن\s+(?:الأسلوب|الصياغة|العبارة|ترابط)|حسن\s+(?:الأسلوب|الصياغة|العبارة|ترابط)|تحسين\s+(?:الصياغة|الأسلوب|الترابط)|(?:النص\s+)?المحدد\s+لغوي(?:ا|ًا)|(?:صحح|راجع|دقق|نقح)\s+(?:هذا\s+)?النص|اجعل\s+(?:النص|العبارة|الفقرة)\s+(?:أوضح|أكثر\s+(?:وضوحا|وضوحًا|رصانة|سلاسة|أكاديمية)))/i.test(text);
    }

    function buildRequest(userInstruction, originalText) {
        const sourceText = String(originalText || "");
        let requestText = sourceText;
        let maskState = null;

        if (pendingTarget && Array.isArray(pendingTarget.protectedRanges)) {
            const chunkOffset = Number(pendingTarget.requestCursor) || 0;
            maskState = DocumentRevisionPreflight.maskChunk(
                sourceText,
                pendingTarget.protectedRanges,
                chunkOffset,
                pendingTarget.id
            );
            requestText = maskState.text;
            pendingTarget.lastMaskedChunk = maskState;
            pendingTarget.requestCursor = chunkOffset + sourceText.length;
        }

        return {
            systemPrompt: [
                "أنت محرر لغوي وأكاديمي متخصص في تحرير النصوص العربية والبحوث العلمية.",
                "افهم طلب الباحث من عبارته وحدد بنفسك نوع العمل: تصحيح لغوي، أو إعادة صياغة، أو الجمع بينهما. لا تطلب منه اختيار نمط من قائمة.",
                "في التصحيح اللغوي: لا تعِد كتابة النص ولا الفقرة. اكتشف الأخطاء الحقيقية فقط، وأخرج لكل خطأ تصحيحًا موضعيًا مستقلًا.",
                "استجابة التصحيح اللغوي يجب أن تكون JSON بهذا الشكل: {\"status\":\"ready\",\"mode\":\"correction\",\"edits\":[{\"original\":\"المقطع الخاطئ حرفيًا\",\"replacement\":\"الصواب\",\"contextBefore\":\"السياق السابق الحرفي\",\"contextAfter\":\"السياق اللاحق الحرفي\",\"reason\":\"سبب موجز\"}]} .",
                "يجب أن يكون original مطابقًا حرفيًا لمقطع متصل من النص، وألا يتضمن فاصل فقرة. أدرج السياق السابق واللاحق كما وردا حرفيًا، واجعل كل تصحيح مستقلًا وغير متداخل. إذا لم توجد أخطاء فأعد edits مصفوفة فارغة.",
                "لا تقترح تحسينات أسلوبية في وضع التصحيح اللغوي، ولا تغير صياغة صحيحة. لا تدمج عدة مواضع في تصحيح واحد إلا إذا كانت متجاورة.",
                "في إعادة الصياغة أو الجمع بين التصحيح والتحسين، استمر في إخراج النص الكامل بعد المعالجة باستخدام الحقل text.",
                "لا تختلق معلومة أو شاهدًا أو إحالة أو مصدرًا، ولا تحذف مضمونًا علميًا، ولا تستبدل المصطلحات المتخصصة بمرادفات قد تغير دلالتها. حافظ على ترتيب الفقرات وفواصلها، ولا تضف تنسيق Markdown.",
                "حافظ على الآيات القرآنية والأحاديث والنقول الحرفية والمراجع والتوثيقات كما وردت، ولا تعد صياغتها.",
                "قد تظهر في النص علامات حماية خاصة على صورة [[WB_KEEP_...]] بدل فقرات تحتوي على تغييرات متعقبة سابقة. هذه العلامات تمثل نصًا موجودًا في المستند ويجب إبقاؤها حرفيًا، مرة واحدة، وفي موضعها وترتيبها؛ لا تترجمها ولا تعدلها ولا تحذفها ولا تنقلها.",
                "إذا كان طلب الباحث غامضًا إلى حد يمنع معرفة التعديل المقصود، فأعد سؤال استيضاح واحدًا ولا تعدل النص.",
                "أعد JSON صالحًا فقط، دون Markdown أو سياج شيفرة أو نص خارجه.",
                "في وضع rewrite أو combined استخدم: {\"status\":\"ready\",\"mode\":\"rewrite أو combined\",\"text\":\"النص كاملًا بعد المعالجة\"}.",
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
                requestText,
                "<<<END_TEXT>>>",
                "",
                "نفذ طلب الباحث على النص المحصور بين العلامتين. في التصحيح اللغوي أخرج قائمة edits فقط، وفي إعادة الصياغة أو الجمع أخرج النص الكامل. حافظ على علامات الحماية كما هي تمامًا."
            ].join("\n")
        };
    }

    function parseResponse(value, originalValue) {
        const sourceText = String(originalValue || "");
        const maskState = pendingTarget && pendingTarget.lastMaskedChunk &&
            pendingTarget.lastMaskedChunk.source === sourceText
            ? pendingTarget.lastMaskedChunk
            : null;
        let raw = String(value || "").trim();
        raw = raw.replace(new RegExp("^" + String.fromCharCode(65279)), "");
        const tick = String.fromCharCode(96);
        raw = raw.replace(new RegExp("^" + tick + "{3}(?:json)?\\s*", "i"), "")
            .replace(new RegExp("\\s*" + tick + "{3}$"), "");

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
            if (pendingTarget) pendingTarget.lastMaskedChunk = null;
            return { status: "clarify", question: String(parsed.question).trim() };
        }

        const mode = String(parsed && parsed.mode || "");
        if (!parsed || parsed.status !== "ready" ||
            !["correction", "rewrite", "combined"].includes(mode)) {
            throw new Error("نتيجة المعالجة غير مكتملة أو غير صالحة؛ لم يُمس المستند.");
        }

        if (mode === "correction") {
            if (!Array.isArray(parsed.edits)) {
                throw new Error("لم تُرجع المراجعة قائمة تصحيحات موضعية صالحة.");
            }
            const source = sourceText;
            const edits = [];
            parsed.edits.forEach(function (item) {
                const original = String(item && item.original || "");
                const replacement = String(item && item.replacement || "");
                const contextBefore = String(item && item.contextBefore || "");
                const contextAfter = String(item && item.contextAfter || "");
                const reason = String(item && item.reason || "").trim();
                if (!original || !reason || /[\r\n]/.test(original + replacement)) {
                    throw new Error("تتضمن قائمة التصحيحات عنصرًا ناقصًا أو يمتد عبر فاصل فقرة.");
                }

                const candidates = [];
                const exactMatches = [];
                let cursor = 0;
                while ((cursor = source.indexOf(original, cursor)) !== -1) {
                    exactMatches.push(cursor);
                    const before = source.slice(Math.max(0, cursor - contextBefore.length), cursor);
                    const after = source.slice(cursor + original.length, cursor + original.length + contextAfter.length);
                    if ((!contextBefore || before === contextBefore) && (!contextAfter || after === contextAfter)) {
                        candidates.push(cursor);
                    }
                    cursor += original.length;
                }

                // Context supplied by a language model can be slightly imprecise.
                // If the exact phrase occurs only once, that occurrence is still unambiguous.
                // When it occurs more than once, require the context to identify exactly one.
                const resolvedStart = candidates.length === 1
                    ? candidates[0]
                    : (exactMatches.length === 1 ? exactMatches[0] : -1);
                if (resolvedStart < 0) {
                    throw new Error("تعذر تحديد موضع التصحيح بصورة فريدة: " + original);
                }
                edits.push({
                    start: resolvedStart,
                    original: original,
                    replacement: replacement,
                    reason: reason,
                    contextBefore: contextBefore,
                    contextAfter: contextAfter
                });
            });

            edits.sort(function (a, b) { return a.start - b.start; });
            for (let index = 1; index < edits.length; index++) {
                if (edits[index - 1].start + edits[index - 1].original.length > edits[index].start) {
                    throw new Error("تتداخل مواضع التصحيحات؛ أُلغي إعدادها لتجنب تغيير النص مرتين.");
                }
            }

            DocumentRevisionPreflight.assertNoProtectedEdits(edits, maskState);

            let revised = source;
            edits.slice().reverse().forEach(function (edit) {
                revised = revised.slice(0, edit.start) + edit.replacement +
                    revised.slice(edit.start + edit.original.length);
            });
            if (pendingTarget) pendingTarget.lastMaskedChunk = null;
            return { status: "ready", mode: mode, text: revised, edits: edits };
        }

        let text = String(parsed && parsed.text || "");
        if (!text.trim()) {
            throw new Error("نتيجة إعادة الصياغة غير مكتملة؛ لم يُمس المستند.");
        }
        if (maskState) {
            text = DocumentRevisionPreflight.restoreChunk(text, maskState);
        }
        if (pendingTarget) pendingTarget.lastMaskedChunk = null;
        return { status: "ready", mode: mode, text: text };
    }

    function hasSameLineBreakCount(originalValue, revisedValue) {
        const count = function (value) {
            return (String(value || "").match(/\r\n|\r|\n/g) || []).length;
        };
        return count(originalValue) === count(revisedValue);
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

    // Normalize a single paragraph's model output and restore its original edge separators.
    // Each model request handles one paragraph. Restore source separators and
    // flatten any line breaks introduced inside the revised paragraph.
    // Batched requests may contain several original paragraphs.
    // Unlike normalizeChunkLineBreaks (single-paragraph mode), preserve every
    // internal line break and restore the exact original separator sequence.
    function normalizeBatchLineBreaks(originalValue, revisedValue) {
        const original = String(originalValue || "");
        const revised = String(revisedValue || "");

        if (!hasSameLineBreakCount(original, revised)) {
            throw new Error(
                "لم تحافظ معالجة إحدى الدفعات على فواصل الفقرات الأصلية؛ " +
                "أُلغي تجهيز التعديلات حمايةً لبنية المستند. أعد المحاولة مع طلب المحافظة على الفقرات."
            );
        }

        return normalizeLineBreaks(original, revised);
    }

    function normalizeChunkLineBreaks(originalValue, revisedValue) {
        const original = String(originalValue || "");
        const revised = String(revisedValue || "");
        const breakPattern = /\r\n|\r|\n/g;

        const leadingMatch = original.match(/^(?:(?:\r\n|\r|\n))+/);
        const leading = leadingMatch ? leadingMatch[0] : "";
        const afterLeading = original.slice(leading.length);
        const trailingMatch = afterLeading.match(/(?:(?:\r\n|\r|\n))+$/);
        const trailing = trailingMatch ? trailingMatch[0] : "";
        const revisedLeadingMatch = revised.match(/^(?:(?:\r\n|\r|\n))+/);
        const revisedLeading = revisedLeadingMatch ? revisedLeadingMatch[0] : "";
        const revisedAfterLeading = revised.slice(revisedLeading.length);
        const revisedTrailingMatch = revisedAfterLeading.match(/(?:(?:\r\n|\r|\n))+$/);
        const revisedTrailing = revisedTrailingMatch ? revisedTrailingMatch[0] : "";
        let revisedCore = revisedAfterLeading.slice(0, revisedAfterLeading.length - revisedTrailing.length);

        // A paragraph must remain a paragraph; model-generated line breaks are flattened.
        revisedCore = revisedCore.replace(breakPattern, " ");
        return leading + revisedCore + trailing;
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

    // Resolve Word search results by their document order, without Range.start/end.
    // Those position properties require WordApiDesktop 1.4, while this add-in
    // intentionally supports the broader WordApi 1.3 baseline.
    function resolveSearchMatchByOccurrence(searchItems, sourceValue, expectedText, expectedStart) {
        const source = String(sourceValue || "");
        const phrase = String(expectedText || "");
        if (!phrase) {
            throw new Error("تعذر تحديد عبارة البحث المطلوبة داخل النص الأصلي.");
        }

        const positions = [];
        let cursor = 0;
        while ((cursor = source.indexOf(phrase, cursor)) !== -1) {
            positions.push(cursor);
            cursor += phrase.length;
        }

        const occurrenceIndex = positions.indexOf(Number(expectedStart));
        if (occurrenceIndex < 0) {
            throw new Error("موضع التعديل المحسوب لا يطابق أي ورود حرفي في النص الأصلي.");
        }

        const exactMatches = Array.from(searchItems || []).filter(function (match) {
            return String(match && match.text || "") === phrase;
        });

        if (exactMatches.length !== positions.length) {
            throw new Error(
                "تعذر مطابقة نتائج Word مع النص الأصلي بأمان (مواضع النص: " +
                positions.length + "، نتائج Word المطابقة: " + exactMatches.length + ")."
            );
        }

        const target = exactMatches[occurrenceIndex];
        if (!target || String(target.text || "") !== phrase) {
            throw new Error("تعذر تعيين نتيجة البحث المقابلة لموضع التعديل بأمان.");
        }
        return target;
    }

    function splitLongParagraph(value, limit) {
        const text = String(value || "");
        if (text.length <= limit) return [text];
        const chunks = [];
        let start = 0;
        while (start < text.length) {
            let end = Math.min(start + limit, text.length);
            if (end < text.length) {
                const minimum = start + Math.floor(limit * 0.55);
                const sentenceBreak = Math.max(
                    text.lastIndexOf(". ", end - 1),
                    text.lastIndexOf("؟ ", end - 1),
                    text.lastIndexOf("! ", end - 1),
                    text.lastIndexOf("؛ ", end - 1),
                    text.lastIndexOf("۔ ", end - 1)
                );
                const spaceBreak = text.lastIndexOf(" ", end - 1);
                if (sentenceBreak >= minimum) end = sentenceBreak + 1;
                else if (spaceBreak >= minimum) end = spaceBreak + 1;
            }
            if (end <= start) end = Math.min(start + limit, text.length);
            chunks.push(text.slice(start, end));
            start = end;
        }
        return chunks;
    }

    function getRecommendedChunkSize(userInstruction) {
        const instruction = String(userInstruction || "").trim();
        const asksRewrite = /(?:أعد\s*صياغ|إعادة\s*صياغ|حسّن|حسن\s+(?:الأسلوب|الصياغة)|تحسين\s+(?:الصياغة|الأسلوب)|اجعل\s+(?:النص|الفقرة|العبارة).{0,45}(?:أوضح|أكثر|أكاديمية|رصانة|سلاسة))/i.test(instruction);
        const asksCorrection = /(?:صحح|تصحيح|دقق|تدقيق|نقح|تنقيح|التصحيح\s+اللغوي|المراجعة\s+اللغوية|راجع\s+(?:لغوي|النص|الفقرة|المستند))/i.test(instruction);

        // Correction responses contain only localized edits, so they can safely
        // inspect larger batches. Rewriting returns the full text and needs a
        // smaller batch to stay within provider output limits.
        if (asksRewrite) return 8000;
        if (asksCorrection) return 16000;
        return 12000;
    }

    function splitText(value, maxChars) {
        const text = String(value || "");
        const limit = Math.max(1000, Number(maxChars) || 12000);
        if (!text) return [text];

        // Preserve every original paragraph separator, but pack adjacent
        // paragraphs into bounded batches instead of issuing one request per
        // paragraph. Only exceptionally long paragraphs are split internally.
        const chunks = [];
        let current = "";
        let cursor = 0;

        function flush() {
            if (!current) return;
            chunks.push(current);
            current = "";
        }

        function appendSegment(segment) {
            if (!segment) return;
            if (current && current.length + segment.length > limit) flush();
            current += segment;
            if (current.length >= limit) flush();
        }

        while (cursor < text.length) {
            let breakIndex = cursor;
            while (breakIndex < text.length &&
                text[breakIndex] !== "\r" &&
                text[breakIndex] !== "\n") {
                breakIndex++;
            }

            const paragraph = text.slice(cursor, breakIndex);
            let separator = "";
            if (breakIndex < text.length) {
                separator = text[breakIndex] === "\r" && text[breakIndex + 1] === "\n"
                    ? "\r\n"
                    : text[breakIndex];
            }

            if (paragraph.trim()) {
                const pieces = splitLongParagraph(paragraph, limit);
                pieces.forEach(function (piece, index) {
                    appendSegment(piece + (index === pieces.length - 1 ? separator : ""));
                });
            } else {
                // Keep blank lines and whitespace-only paragraphs verbatim.
                appendSegment(paragraph + separator);
            }

            cursor = breakIndex + separator.length;
        }

        flush();
        return chunks.length ? chunks : [text];
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

    function diffTextPreservingParagraphs(oldValue, newValue) {
        const oldText = String(oldValue || "");
        const newText = String(newValue || "");
        if (!hasSameLineBreakCount(oldText, newText)) {
            throw new Error("غيّرت المعالجة عدد فواصل الفقرات؛ أُلغي التطبيق للمحافظة على بنية المستند. أعد المحاولة مع طلب المحافظة على الفقرات.");
        }

        const oldParts = oldText.split(/(\r\n|\r|\n)/);
        const newParts = newText.split(/(\r\n|\r|\n)/);
        const operations = [];
        let offset = 0;

        for (let index = 0; index < oldParts.length; index++) {
            const oldPart = oldParts[index];
            const newPart = newParts[index] === undefined ? "" : newParts[index];
            if (index % 2 === 0) {
                diffText(oldPart, newPart).forEach(function (operation) {
                    operations.push({
                        start: offset + operation.start,
                        length: operation.length,
                        replacement: operation.replacement
                    });
                });
            } else if (oldPart !== newPart && oldPart.length !== newPart.length) {
                throw new Error("تعذر مطابقة فواصل الفقرات بأمان؛ أُلغي التطبيق للمحافظة على بنية المستند.");
            }
            offset += oldPart.length;
        }

        return operations;
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

            let protectedRanges = [];
            if (targetType === "selection") {
                if (typeof range.getOoxml !== "function") {
                    throw new Error("يتعذر فحص التغييرات المتعقبة في التحديد بهذا الإصدار من Word؛ لم يُرسل النص إلى الذكاء الاصطناعي.");
                }
                const selectionOoxml = range.getOoxml();
                await context.sync();
                if (DocumentRevisionPreflight.hasTrackedChanges(selectionOoxml.value)) {
                    throw new Error("يتضمن التحديد تغييرات متعقبة سابقة. اقبلها أو ارفضها في Word قبل تدقيق هذا التحديد؛ لم يُرسل النص إلى الذكاء الاصطناعي.");
                }
            } else {
                const paragraphs = context.document.body.paragraphs;
                paragraphs.load("items/text");
                await context.sync();

                if (!paragraphs.items.every(function (paragraph) {
                    return typeof paragraph.getOoxml === "function";
                })) {
                    throw new Error("يتعذر فحص التغييرات المتعقبة في فقرات المستند بهذا الإصدار من Word؛ أُوقف التدقيق حمايةً للمراجعات السابقة.");
                }

                const paragraphResults = paragraphs.items.map(function (paragraph) {
                    return paragraph.getOoxml();
                });
                await context.sync();

                const paragraphRecords = paragraphs.items.map(function (paragraph, index) {
                    return {
                        index: index,
                        text: paragraph.text,
                        ooxml: paragraphResults[index].value
                    };
                });
                protectedRanges = DocumentRevisionPreflight.buildProtectedParagraphRanges(
                    originalText,
                    paragraphRecords
                );
            }

            range.track();
            await context.sync();

            result = {
                id: "revision-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
                targetType: targetType,
                originalText: originalText,
                protectedParagraphCount: protectedRanges.length
            };
            pendingTarget = {
                id: result.id,
                range: range,
                protectedRanges: protectedRanges,
                requestCursor: 0,
                lastMaskedChunk: null
            };
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
        matches.items.forEach(function (match) { match.load("text"); });
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
        if (!hasSameLineBreakCount(originalText, revision.text)) {
            throw new Error("غيّرت المعالجة عدد فواصل الفقرات؛ أُلغي التطبيق للمحافظة على بنية المستند. أعد المحاولة مع طلب المحافظة على الفقرات.");
        }
        const safeRevisedText = normalizeLineBreaks(originalText, revision.text);
        const operations = diffTextPreservingParagraphs(originalText, safeRevisedText);
        if (!operations.length) return { applied: false, unchanged: true, changedParts: 0 };

        let appliedCount = 0;

        await Word.run(async function (context) {
            const range = await resolveTarget(context, revision);
            range.load("text");
            await context.sync();

            if (String(range.text || "") !== originalText) {
                throw new Error("تغير النص الأصلي في Word بعد إرساله للمعالجة؛ أُلغي التطبيق حمايةً للمستند. أعد المعالجة على النص الحالي.");
            }

            let paragraphTargets = [];
            if (revision.targetType === "document" && range.paragraphs) {
                const paragraphCollection = range.paragraphs;
                paragraphCollection.load("items/text");
                await context.sync();

                try {
                    const mappings = DocumentRevisionPreflight.mapParagraphRanges(
                        originalText,
                        paragraphCollection.items.map(function (paragraph) {
                            return paragraph.text;
                        })
                    );
                    paragraphTargets = mappings.map(function (mapping) {
                        return {
                            paragraphIndex: mapping.paragraphIndex,
                            start: mapping.start,
                            end: mapping.end,
                            text: mapping.text,
                            paragraph: paragraphCollection.items[mapping.paragraphIndex]
                        };
                    }).filter(function (mapping) {
                        return !!mapping.paragraph;
                    });
                } catch (_) {
                    // Keep the established exact-occurrence resolver as a safe fallback
                    // when a Word host exposes paragraph boundaries differently.
                    paragraphTargets = [];
                }
            }

            function findParagraphTarget(start, length) {
                const mapping = DocumentRevisionPreflight.findParagraphRange(
                    paragraphTargets,
                    start,
                    length
                );
                return mapping || null;
            }

            const prepared = [];
            for (const operation of operations) {
                if (operation.length > 0) {
                    const oldText = originalText.slice(operation.start, operation.start + operation.length);
                    if (/[\r\n]/.test(oldText)) {
                        throw new Error("تتضمن النتيجة تغييرًا في فواصل الفقرات؛ أُلغي التطبيق حتى لا تتغير بنية المستند.");
                    }
                    const paragraphTarget = findParagraphTarget(operation.start, operation.length);
                    const searchRange = paragraphTarget ? paragraphTarget.paragraph : range;
                    const search = searchRange.search(oldText, { matchCase: true, matchWholeWord: false });
                    search.load("items");
                    prepared.push({
                        operation: operation,
                        search: search,
                        kind: "replace",
                        searchSourceText: paragraphTarget ? paragraphTarget.text : originalText,
                        searchExpectedStart: paragraphTarget ? operation.start - paragraphTarget.start : operation.start
                    });
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
                        const paragraphTarget = findParagraphTarget(
                            anchor.token.start,
                            anchor.token.value.length
                        );
                        const searchRange = paragraphTarget ? paragraphTarget.paragraph : range;
                        const search = searchRange.search(anchor.token.value, { matchCase: true, matchWholeWord: false });
                        search.load("items");
                        prepared.push({
                            operation: operation,
                            search: search,
                            kind: "insert",
                            anchor: anchor,
                            searchSourceText: paragraphTarget ? paragraphTarget.text : originalText,
                            searchExpectedStart: paragraphTarget ? anchor.token.start - paragraphTarget.start : anchor.token.start
                        });
                    } else {
                        throw new Error("تعذر تحديد موضع إضافة إحدى العبارات بأمان؛ لم يُعدّل المستند.");
                    }
                }
            }

            await context.sync();

            // Resolve queued searches in one grouped load, then map each result
            // against its paragraph-local text and expected occurrence.
            const searchable = prepared.filter(function (item) { return !!item.search; });
            searchable.forEach(function (item) {
                item.search.items.forEach(function (match) { match.load("text"); });
            });
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

                const expectedText = item.kind === "replace"
                    ? originalText.slice(item.operation.start, item.operation.start + item.operation.length)
                    : item.anchor.token.value;
                const expectedStart = item.searchExpectedStart !== undefined
                    ? item.searchExpectedStart
                    : (item.kind === "insert" && item.anchor.location === "After"
                        ? item.anchor.token.start
                        : item.operation.start);
                const target = resolveSearchMatchByOccurrence(
                    item.search.items,
                    item.searchSourceText || originalText,
                    expectedText,
                    expectedStart
                );
                resolved.push({ ...item, target: target });
            }

            resolved.sort(function (a, b) { return b.operation.start - a.operation.start; });
            const insertedRanges = [];
            resolved.forEach(function (item) {
                const location = item.kind === "insert" ? item.anchor.location : "Replace";
                const inserted = item.target.insertText(item.operation.replacement, location);
                if (item.operation.replacement) {
                    inserted.font.color = "#FF0000";
                    insertedRanges.push({ range: inserted, expectedText: item.operation.replacement });
                }
            });

            insertedRanges.forEach(function (item) { item.range.load("text"); });
            await context.sync();

            const unverifiable = insertedRanges.find(function (item) {
                return String(item.range.text || "") !== item.expectedText;
            });
            if (unverifiable) {
                throw new Error("لم يجتز أحد التعديلات التحقق بعد إدراجه في Word؛ راجع المستند قبل متابعة العمل.");
            }
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
        resolveSearchMatchByOccurrence: resolveSearchMatchByOccurrence,
        diffText: diffText,
        diffTextPreservingParagraphs: diffTextPreservingParagraphs,
        getRecommendedChunkSize: getRecommendedChunkSize,
        splitText: splitText,
        normalizeLineBreaks: normalizeLineBreaks,
        normalizeBatchLineBreaks: normalizeBatchLineBreaks,
        normalizeChunkLineBreaks: normalizeChunkLineBreaks,
        hasSameLineBreakCount: hasSameLineBreakCount,
        capture: capture,
        apply: apply
    };

    global.DocumentRevision = api;
    if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
