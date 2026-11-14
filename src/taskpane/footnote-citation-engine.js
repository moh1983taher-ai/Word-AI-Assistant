(function (global) {
    "use strict";

    const DIGIT = "[0-9٠-٩۰-۹]";
    const NUMBER = DIGIT + "{1,6}(?:\\s*\\/\\s*" + DIGIT + "{1,6})?(?:\\s*[-–—]\\s*" + DIGIT + "{1,6})?";
    const LOCATION_LABEL = "(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\\s*)?";
    const LOCATION_SEPARATOR = "(?:\\s*[،,؛;]\\s*|\\s+)";
    const LOCATION_SUFFIX = new RegExp(
        LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER +
        "(?:" + LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER + ")*" +
        "(?:\\s*(?:[،,؛;]\\s*)?(?:وما بعدها|وما بعده|وما يليها|وما يليه|وما بعد ذلك))?\\s*$",
        "u"
    );

    function normalizeIdentityText(value) {
        return String(value ?? "")
            .normalize("NFC")
            .toLowerCase()
            .replace(/[\u064B-\u065F\u0670\u0640]/gu, "")
            .replace(/[\u200B-\u200F\u202A-\u202E\uFEFF]/gu, "")
            .replace(/[أإآٱ]/gu, "ا")
            .replace(/ى/gu, "ي")
            .replace(/[ؤ]/gu, "و")
            .replace(/[ئ]/gu, "ي")
            .replace(/ة/gu, "ه")
            .replace(/[\s\u00A0\p{P}\p{S}]+/gu, "");
    }

    function referenceKey(reference) {
        if (!reference || typeof reference !== "object") return "";
        const author = normalizeIdentityText(reference.author);
        const title = normalizeBibliographicTitleKey(reference.title);
        if (!title) return "";
        return author + "::" + title;
    }

    function uniqueReferenceEntries(entries) {
        const seen = new Set();
        const result = [];
        for (const item of Array.isArray(entries) ? entries : []) {
            const key = referenceKey(item && item.reference);
            if (key && seen.has(key)) continue;
            if (key) seen.add(key);
            result.push(item);
        }
        return result;
    }

    function getIdentityText(value) {
        let text = String(value ?? "").trim();
        if (!text) return "";
        // Editorial cues are outside the bibliographic identity and must survive replacement.
        text = text.replace(/^(?:(?:انظر|ينظر|راجع|يراجع|قارن|نقلاً عن|نقلًا عن)\s*[:：]?\s*)+/u, "").trim();
        const match = LOCATION_SUFFIX.exec(text);
        if (!match || match.index <= 0) return text;
        return text.slice(0, match.index)
            .trim()
            .replace(/[،,؛;:\s]+$/gu, "")
            .trim();
    }

    function isHadithTakhrijReference(reference, occurrence = null) {
        if (!reference || typeof reference !== "object") return false;
        const type = String(reference.type || reference.referenceType || reference.kind || "").trim().toLowerCase();
        if (["hadith", "حديث", "hadith-takhrij", "takhrij"].includes(type)) return true;

        const variantTexts = (Array.isArray(reference.variants) ? reference.variants : []).map(item => {
            if (typeof item === "string") return item;
            if (!item || typeof item !== "object") return "";
            return String(item.text || item.originalText || item.cleanedText || item.citation || item.value || "");
        });

        const evidence = [
            reference.author, reference.title, reference.originalText,
            reference.cleanedText, reference.notes,
            occurrence && occurrence.sourceText,
            occurrence && occurrence.originalText,
            ...variantTexts
        ].filter(Boolean).join(" ");

        return /(?:^|[\s،,؛:])(?:أخرجه|أخرج|رواه|روى|خرَّجه|خرّجه|خرجه|تخريج(?:\s+الحديث)?|عزاه)(?=$|[\s،,؛:])|حديث\s*(?:رقم|برقم)|رقم\s+الحديث|متفق\s+عليه/u.test(evidence);
    }

    function hasUnparsedLocationSuffix(value) {
        const text = String(value ?? "").trim();
        if (!text) return false;
        return /(?:^|[،,؛;\s])(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*[0-9٠-٩۰-۹]/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6}(?:\s*[،,؛;]\s*[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6})?\s*$/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*[-–—]\s*[0-9٠-٩۰-۹]{1,6}(?:\s*[،,؛;]?\s*وما بعدها)?\s*$/u.test(text);
    }

    function findExactOccurrences(text, needle) {
        const source = String(text ?? "");
        const target = String(needle ?? "");
        if (!target) return [];
        const positions = [];
        let from = 0;
        while (from <= source.length - target.length) {
            const index = source.indexOf(target, from);
            if (index < 0) break;
            positions.push(index);
            from = index + Math.max(1, target.length);
        }
        return positions;
    }

        function findValidatedSourceSpan(noteText, candidateText) {
        const source = String(noteText ?? "");
        const candidate = String(candidateText ?? "").trim();
        if (!source || !candidate) return "";
        const target = normalizeIdentityText(candidate);
        if (target.length < 4) return "";

        const normalized = [];
        const rawMap = [];
        for (let index = 0; index < source.length; index++) {
            const part = normalizeIdentityText(source[index]);
            if (!part) continue;
            normalized.push(part);
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        const haystack = normalized.join("");
        let from = 0;
        while (from <= haystack.length - target.length) {
            const position = haystack.indexOf(target, from);
            if (position < 0) break;
            const rawStart = rawMap[position];
            const rawEnd = rawMap[position + target.length - 1] + 1;
            from = position + Math.max(1, target.length);
            if (!Number.isInteger(rawStart) || !Number.isInteger(rawEnd) || rawEnd <= rawStart) continue;

            const previous = rawStart > 0 ? source[rawStart - 1] : "";
            const next = source.slice(rawEnd).match(/^\s*/u)?.[0].length || 0;
            const following = source[rawEnd + next] || "";
            if (previous && /[\p{L}\p{N}]/u.test(previous)) {
                const beforePrevious = rawStart > 1 ? source[rawStart - 2] : "";
                const isAttachedKaf = previous === "ك" &&
                    /^ال/u.test(source.slice(rawStart, rawStart + 2)) &&
                    (!beforePrevious || !/[\p{L}\p{N}]/u.test(beforePrevious));
                if (!isAttachedKaf) continue;
            }
            if (following && /[\p{L}\p{N}]/u.test(following)) continue;
            return source.slice(rawStart, rawEnd);
        }
        return "";
    }

function getNoteIndex(suggestion, source) {
        const materialId = String(suggestion && suggestion.materialId || "");
        const pattern = source === "footnote" ? /^footnote-(\d+)$/ : /^endnote-(\d+)$/;
        const match = materialId.match(pattern);
        if (match) {
            const value = Number(match[1]);
            return Number.isInteger(value) && value > 0 ? value - 1 : -1;
        }
        const noteNumber = Number(suggestion && suggestion.noteNumber);
        return Number.isInteger(noteNumber) && noteNumber > 0 ? noteNumber - 1 : -1;
    }

    function normalizeBibliographicTitle(value, author, locationText = "") {
        let title = String(value ?? "").trim();
        const location = String(locationText ?? "").trim();
        if (title && location) {
            const suffixCandidates = [location];
            if (/^ص/u.test(location)) {
                const pageValue = location.replace(/^ص\s*/u, "").trim();
                if (pageValue) {
                    suffixCandidates.push("ص" + pageValue);
                    suffixCandidates.push("صفحة " + pageValue);
                }
            }
            for (const suffix of new Set(suffixCandidates)) {
                const position = title.lastIndexOf(suffix);
                if (position < 0) continue;
                const tail = title.slice(position + suffix.length);
                const hasBoundary = position === 0 || /[\s،,؛:]/u.test(title[position - 1]);
                if (tail.trim() === "" && hasBoundary) {
                    title = title.slice(0, position).replace(/[،,؛:\s]+$/u, "").trim();
                    break;
                }
            }
        }
        const authorKey = normalizeIdentityText(author);
        if (!title) return "";

        if (authorKey) {
            const parts = title.split(/[،,؛;]+/u).map(part => part.trim()).filter(Boolean);
            const authorFiltered = parts.filter(part => normalizeIdentityText(part) !== authorKey);
            if (authorFiltered.length && authorFiltered.length !== parts.length) {
                title = authorFiltered.join("، ");
            }
        }

        const tokens = title.split(/\s+/u).filter(Boolean);
        const cleaned = [];

        function tokenInfo(token) {
            const raw = String(token || "");
            const core = raw.replace(/^[،,؛:()[\]{}"'\u0060]+/u, "").replace(/[،,؛:()[\]{}"'\u0060]+$/u, "");
            const hasWaw = /^و(?=.{2,})/u.test(core);
            return {
                key: normalizeIdentityText(hasWaw ? core.slice(1) : core),
                hasWaw: hasWaw,
                suffix: (raw.match(/[،,؛:.)\]}]+$/u) || [""])[0]
            };
        }

        for (const token of tokens) {
            const current = tokenInfo(token);
            const previous = cleaned.length ? tokenInfo(cleaned[cleaned.length - 1]) : null;
            const duplicate = previous && current.key &&
                current.key === previous.key &&
                (current.hasWaw || previous.hasWaw);

            if (duplicate) {
                if (current.suffix && !/[،,؛:.)\]}]+$/u.test(cleaned[cleaned.length - 1])) {
                    cleaned[cleaned.length - 1] += current.suffix;
                }
                continue;
            }
            cleaned.push(token);
        }

        return cleaned.join(" ").replace(/\s+([،,؛:.)])/gu, "$1").trim();
    }


    function findNormalizedSpans(sourceText,candidateText){
        const source=String(sourceText??""),target=normalizeIdentityText(candidateText);if(!source||!target)return[];
        const normalized=[],rawMap=[];for(let i=0;i<source.length;i++){const p=normalizeIdentityText(source[i]);if(!p)continue;normalized.push(p);for(let j=0;j<p.length;j++)rawMap.push(i);}
        const hay=normalized.join(""),out=[];let from=0;while(from<=hay.length-target.length){const p=hay.indexOf(target,from);if(p<0)break;from=p+Math.max(1,target.length);const start=rawMap[p],end=rawMap[p+target.length-1]+1;if(Number.isInteger(start)&&Number.isInteger(end)&&end>start)out.push({start,end,text:source.slice(start,end)});}return out;
    }
    function findBibliographicCitationSpans(noteText,sourceText,reference){
        const source=String(noteText??"");if(!source||!reference||typeof reference!=="object")return[];
        const title=getIdentityText(reference.title||""),author=String(reference.author||"").trim();if(!title)return[];
        const candidates=[];const add=v=>{v=String(v||"").trim();const k=normalizeIdentityText(v);if(k.length>=5&&!candidates.some(q=>normalizeIdentityText(q)===k))candidates.push(v);};
        add(title);const short=title.replace(/^كتاب\s+/u,"").trim();if(short!==title)add(short);const first=short.split(/\s+/u)[0]||"";if(normalizeIdentityText(first).length>=5)add(first);
        const out=[];for(const candidate of candidates){for(const span of findNormalizedSpans(source,candidate)){
            const prev=source.slice(0,span.start).slice(-1);if(prev&&/[\p{L}\p{N}]/u.test(prev))continue;
            const tail=source.slice(span.end).replace(/^[\s،,؛:]+/u,"");
            const location=/^(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*)?[0-9٠-٩۰-۹]/u.test(tail);
            const terminal=!tail||/^[.؛،,:؟!\)\]}]/u.test(tail);
            let start=span.start,end=span.end;
            if(author){
                const as=findNormalizedSpans(source,author);
                const before=as.filter(a=>a.end<=span.start&&/^[\s،,؛:]+$/u.test(source.slice(a.end,span.start)));
                const after=as.filter(a=>a.start>=span.end&&/^[\s،,؛:]+$/u.test(source.slice(span.end,a.start)));
                if(before.length===1)start=before[0].start;
                else if(after.length===1){
                    const afterAuthor=source.slice(after[0].end).replace(/^[\s،,؛:]+/u,"");
                    const authorThenLocation=/^(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*)?[0-9٠-٩۰-۹]/u.test(afterAuthor);
                    const authorThenTerminal=!afterAuthor||/^[.؛،,:؟!\)\]}]/u.test(afterAuthor);
                    if(authorThenLocation||authorThenTerminal)end=after[0].end;
                }
            }
            const finalTail=source.slice(end).replace(/^[\s،,؛:]+/u,"");
            const hasFinalLocation=/^(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*)?[0-9٠-٩۰-۹]/u.test(finalTail);
            const hasFinalTerminal=!finalTail||/^[.؛،,:؟!\)\]}]/u.test(finalTail);
            if(!hasFinalLocation&&!hasFinalTerminal&&!location&&!terminal)continue;
            if(!out.some(q=>q.start===start&&q.end===end))out.push({start,end,text:source.slice(start,end).trim()});
        }if(out.length)break;}return out.sort((a,b)=>a.start-b.start);
    }
    function normalizeBibliographicTitleKey(value){
        const key=normalizeIdentityText(getIdentityText(value));
        // Explicitly verified title alias: al-Farahidi's "Kitab al-Ayn" is
        // also cited as "al-Ayn". Do not strip "Kitab" generically because
        // it can be a meaningful part of other work titles.
        const aliases = new Map([
            [normalizeIdentityText("كتاب العين"), normalizeIdentityText("العين")],
            [normalizeIdentityText("العين"), normalizeIdentityText("العين")]
        ]);
        return aliases.get(key) || key;
    }

    function buildPairs(noteText,suggestion){
        const refs=Array.isArray(suggestion&&suggestion.references)?suggestion.references:[],news=Array.isArray(suggestion&&suggestion.suggestedTexts)?suggestion.suggestedTexts:[],olds=Array.isArray(suggestion&&suggestion.sourceTexts)?suggestion.sourceTexts:[];
        const count=Math.min(refs.length,news.length,olds.length),pairs=[];pairs.collapsedCount=0;pairs.unchangedCount=0;const counts=new Map();
        for(let i=0;i<count;i++){const ref=refs[i],sourceText=String(olds[i]??"").trim(),replacementText=getIdentityText(news[i]);if(!sourceText||!replacementText)continue;
            const spans=findBibliographicCitationSpans(noteText,sourceText,ref);if(!spans.length)continue;
            const key=referenceKey(ref)+"::"+normalizeIdentityText(getIdentityText(sourceText));if(!key)continue;
            const ordinal=counts.get(key)||0;if(spans.length<=ordinal)continue;const span=spans[ordinal];counts.set(key,ordinal+1);
            if(normalizeIdentityText(span.text)===normalizeIdentityText(replacementText)){pairs.unchangedCount++;continue;}
            pairs.push({sourceText,identityText:span.text,replacementText,start:span.start,end:span.end,expectedText:span.text,sourceOccurrenceIndex:ordinal,referenceKey:referenceKey(ref)||key});
        }
        pairs.sort((a,b)=>a.start-b.start||b.end-a.end);const safe=[];for(const pair of pairs){const prev=safe[safe.length-1];if(prev&&pair.start<prev.end)continue;safe.push(pair);}safe.collapsedCount=pairs.collapsedCount;safe.unchangedCount=pairs.unchangedCount;return safe;
    }
    function selectOccurrence(matches, occurrenceIndex) {
        if (!Array.isArray(matches) || !Number.isInteger(occurrenceIndex) || occurrenceIndex < 0) {
            return null;
        }
        return matches[occurrenceIndex] || null;
    }

    function findUniqueBibliographicAnchor(noteText, reference) {
        const text = String(noteText ?? "");
        if (!text.trim() || !reference || typeof reference !== "object") return "";

        const author = String(reference.author ?? "").trim();
        const title = normalizeBibliographicTitle(getIdentityText(reference.title ?? ""), author);
        if (!title) return "";

        const candidates = [];
        const addCandidate = value => {
            const candidate = String(value ?? "").trim();
            const key = normalizeIdentityText(candidate);
            if (key.length < 8 || candidates.some(item => normalizeIdentityText(item) === key)) return;
            candidates.push(candidate);
        };

        // A fallback must prove the complete author/title identity.
        // Matching the title alone is unsafe: common titles can appear in prose
        // or in a different citation and must never authorize replacement.
        if (!author || !title) return "";
        addCandidate(author + "، " + title);
        addCandidate(title + "، " + author);
        addCandidate(author + " " + title);
        addCandidate(title + " " + author);

        let normalized = "";
        const rawMap = [];
        for (let index = 0; index < text.length; index++) {
            const part = normalizeIdentityText(text[index]);
            normalized += part;
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        candidates.sort((a, b) => normalizeIdentityText(b).length - normalizeIdentityText(a).length);

        for (const candidate of candidates) {
            const target = normalizeIdentityText(candidate);
            const positions = [];
            let from = 0;
            while (from <= normalized.length - target.length) {
                const position = normalized.indexOf(target, from);
                if (position < 0) break;
                positions.push(position);
                from = position + Math.max(1, target.length);
            }

            // A fallback is safe only when the bibliographic identity occurs once.
            // Repeated occurrences must be anchored by an explicit source variant.
            if (positions.length !== 1) continue;

            const position = positions[0];
            const rawStart = rawMap[position];
            const rawEnd = rawMap[position + target.length - 1] + 1;
            if (Number.isInteger(rawStart) && Number.isInteger(rawEnd) && rawEnd > rawStart) {
                const validated = findValidatedSourceSpan(text, text.slice(rawStart, rawEnd));
                if (validated) return validated;
            }
        }

        return "";
    }

    function searchOptions() {
        return {
            matchCase: false,
            matchWholeWord: false,
            matchWildcards: false,
            ignorePunct: true,
            ignoreSpace: true
        };
    }

    async function apply(suggestions) {
        const items = Array.isArray(suggestions) ? suggestions : [];
        if (!items.length) return { applied: 0, skipped: 0, unchanged: 0 };

        const summaries = { applied: 0, skipped: 0, unchanged: 0 };

        await Word.run(async function (context) {
            const body = context.document.body;
            const footnotes = body.footnotes;
            const endnotes = body.endnotes;
            footnotes.load("items/body");
            endnotes.load("items/body");
            await context.sync();

            const noteEntries = [];
            function addNotes(collection, source) {
                collection.items.forEach(function (note, index) {
                    const range = note.body.getRange();
                    range.load("text");
                    noteEntries.push({
                        source: source,
                        index: index,
                        range: range,
                        text: ""
                    });
                });
            }
            addNotes(footnotes, "footnote");
            addNotes(endnotes, "endnote");
            if (noteEntries.length) await context.sync();

            noteEntries.forEach(function (entry) {
                entry.text = String(entry.range.text ?? "").trim();
            });

            const noteByKey = new Map(noteEntries.map(entry => [
                entry.source + ":" + entry.index,
                entry
            ]));

            const operations = [];
            const options = searchOptions();

            for (const suggestion of items) {
                const source = suggestion && suggestion.source;
                if (source !== "footnote" && source !== "endnote") {
                    summaries.skipped++;
                    continue;
                }

                const noteIndex = getNoteIndex(suggestion, source);
                const entry = noteByKey.get(source + ":" + noteIndex);
                if (!entry || normalizeIdentityText(entry.text) !==
                    normalizeIdentityText(suggestion.originalText)) {
                    summaries.skipped += Math.max(
                        1,
                        Math.min(
                            Array.isArray(suggestion && suggestion.references) ? suggestion.references.length : 0,
                            Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts.length : 0,
                            Array.isArray(suggestion && suggestion.suggestedTexts) ? suggestion.suggestedTexts.length : 0
                        )
                    );
                    continue;
                }

                const pairs = buildPairs(entry.text, suggestion);
                const requestedCount = Math.min(
                    Array.isArray(suggestion && suggestion.references) ? suggestion.references.length : 0,
                    Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts.length : 0,
                    Array.isArray(suggestion && suggestion.suggestedTexts) ? suggestion.suggestedTexts.length : 0
                );
                summaries.skipped += Math.max(0, requestedCount - pairs.length -
                    (pairs.collapsedCount || 0) - (pairs.unchangedCount || 0));
                summaries.unchanged += pairs.unchangedCount || 0;

                for (const pair of pairs) {
                    operations.push({ entry: entry, pair: pair });
                }
            }

            if (!operations.length) return;

            // Bound the number of queued Word searches. Each batch is fully
            // resolved before the next one, while all mutations remain deferred.
            const SEARCH_BATCH_SIZE = 60;
            const resolved = [];

            for (let offset = 0; offset < operations.length; offset += SEARCH_BATCH_SIZE) {
                const batch = operations.slice(offset, offset + SEARCH_BATCH_SIZE);
                const fullSearches = batch.map(function (operation) {
                    return {
                        entry: operation.entry,
                        pair: operation.pair,
                        fullSearch: operation.entry.range.search(
                            operation.pair.sourceText,
                            options
                        ).load("items")
                    };
                });

                await context.sync();

                const fullCandidates = [];
                for (const operation of fullSearches) {
                    for (const match of operation.fullSearch.items) {
                        match.load("text");
                        fullCandidates.push(match);
                    }
                }
                if (fullCandidates.length) await context.sync();

                const identitySearches = [];
                for (const operation of fullSearches) {
                    const expectedFull = normalizeIdentityText(operation.pair.sourceText);
                    const matches = operation.fullSearch.items.filter(match =>
                        normalizeIdentityText(match.text) === expectedFull
                    );
                    const fullMatch = selectOccurrence(matches, operation.pair.sourceOccurrenceIndex);
                    if (!fullMatch) {
                        summaries.skipped++;
                        continue;
                    }

                    const identitySearch = fullMatch.search(
                        operation.pair.identityText,
                        options
                    ).load("items");

                    identitySearches.push({
                        entry: operation.entry,
                        pair: operation.pair,
                        identitySearch: identitySearch
                    });
                }

                if (identitySearches.length) await context.sync();

                const identityCandidates = [];
                for (const operation of identitySearches) {
                    for (const match of operation.identitySearch.items) {
                        match.load("text");
                        identityCandidates.push(match);
                    }
                }
                if (identityCandidates.length) await context.sync();

                for (const operation of identitySearches) {
                    const expectedIdentity = normalizeIdentityText(operation.pair.identityText);
                    const matches = operation.identitySearch.items.filter(match =>
                        normalizeIdentityText(match.text) === expectedIdentity
                    );
                    if (matches.length !== 1) {
                        summaries.skipped++;
                        continue;
                    }

                    const range = matches[0];
                    if (String(range.text ?? "") !== operation.pair.expectedText) {
                        summaries.skipped++;
                        continue;
                    }

                    resolved.push({
                        entry: operation.entry,
                        pair: operation.pair,
                        range: range,
                        insertedRange: null
                    });
                }
            }

            // Re-read every affected note immediately before queuing mutations.
            const affectedEntries = Array.from(new Set(resolved.map(item => item.entry)));
            for (const entry of affectedEntries) entry.range.load("text");
            if (affectedEntries.length) await context.sync();

            for (const entry of affectedEntries) {
                if (String(entry.range.text ?? "").trim() !== entry.text) {
                    throw new Error("تغير نص إحدى الحواشي بعد إعداد المعاينة؛ أُلغي الاستبدال حمايةً للمستند. أعد التحليل ثم حاول مجددًا.");
                }
            }

            // Sort from the end of each note to the beginning; distinct ranges remain stable.
            resolved.sort((a, b) =>
                a.entry.source.localeCompare(b.entry.source) ||
                a.entry.index - b.entry.index ||
                b.pair.start - a.pair.start
            );

            for (const item of resolved) {
                item.insertedRange = item.range.insertText(
                    item.pair.replacementText,
                    "Replace"
                );
                item.insertedRange.load("text");
            }

            if (resolved.length) await context.sync();

            for (const item of resolved) {
                if (normalizeIdentityText(item.insertedRange.text) ===
                    normalizeIdentityText(item.pair.replacementText)) {
                    summaries.applied++;
                } else {
                    summaries.skipped++;
                }
            }
        });

        return {
            applied: summaries.applied,
            skipped: summaries.skipped,
            unchanged: summaries.unchanged
        };
    }

    global.FootnoteCitationEngine = {
        apply: apply,
        buildPairs: buildPairs,
        findExactOccurrences: findExactOccurrences,
        findValidatedSourceSpan: findValidatedSourceSpan,
        getIdentityText: getIdentityText,
        hasUnparsedLocationSuffix: hasUnparsedLocationSuffix,
        normalizeIdentityText: normalizeIdentityText,
        referenceKey: referenceKey,
        normalizeBibliographicTitleKey: normalizeBibliographicTitleKey,
        findBibliographicCitationSpans: findBibliographicCitationSpans,
        uniqueReferenceEntries: uniqueReferenceEntries,
        normalizeBibliographicTitle: normalizeBibliographicTitle,
        findUniqueBibliographicAnchor: findUniqueBibliographicAnchor,
        isHadithTakhrijReference: isHadithTakhrijReference,
        selectOccurrence: selectOccurrence
    };

    if (typeof module !== "undefined" && module.exports) {
        module.exports = global.FootnoteCitationEngine;
    }
})(typeof window !== "undefined" ? window : globalThis);
