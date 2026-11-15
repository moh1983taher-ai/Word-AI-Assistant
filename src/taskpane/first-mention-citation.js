(function (global) {
    "use strict";

    function normalize(value) {
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

    function findSpans(source, value) {
        const target = normalize(value);
        if (target.length < 3) return [];

        const targets = [target];
        if (target.startsWith("ال") && target.length > 3) {
            targets.push(target.slice(1));
        }

        let normalized = "";
        const rawMap = [];
        for (let index = 0; index < source.length; index++) {
            const part = normalize(source[index]);
            if (!part) continue;
            normalized += part;
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        const spans = [];
        const seen = new Set();
        const isWord = character =>
            Boolean(character) && /[\p{L}\p{N}]/u.test(character);

        for (let targetIndex = 0; targetIndex < targets.length; targetIndex++) {
            const candidate = targets[targetIndex];
            let from = 0;

            while (from <= normalized.length - candidate.length) {
                const position = normalized.indexOf(candidate, from);
                if (position < 0) break;
                from = position + Math.max(1, candidate.length);

                let start = rawMap[position];
                const end = rawMap[position + candidate.length - 1] + 1;
                if (!Number.isInteger(start) || !Number.isInteger(end) || end <= start) continue;

                const previous = start > 0 ? source[start - 1] : "";
                const beforePrevious = start > 1 ? source[start - 2] : "";

                // Support the written double-lam form "للأزهري" without
                // accepting a name embedded inside an unrelated word.
                const attachedDoubleLam = targetIndex === 1 &&
                    previous === "ل" &&
                    (!beforePrevious || !isWord(beforePrevious));
                if (isWord(previous) && !attachedDoubleLam) continue;
                if (isWord(source[end] || "")) continue;
                if (attachedDoubleLam) start--;

                const key = start + ":" + end;
                if (!seen.has(key)) {
                    seen.add(key);
                    spans.push({ start, end, text: source.slice(start, end) });
                }
            }
        }

        return spans.sort((left, right) => left.start - right.start || left.end - right.end);
    }

    function build(sourceText, reference, fullAuthor, allowMissingAuthor = false) {
        const source = String(sourceText ?? "");
        const authorToUse = String(fullAuthor ?? "").trim();
        if (!source || !reference || typeof reference !== "object") return source;

        const title = String(reference.title ?? "").trim();
        if (!title) return source;

        const titleSpans = findSpans(source, title);
        if (!titleSpans.length) return source;

        if (!authorToUse) return source;

        const authorCandidates = Array.from(new Set([
            authorToUse,
            reference.author
        ].map(value => String(value ?? "").trim()).filter(Boolean)));

        const punctuationOnly = /^[\s\u00A0،,؛;:()[\]{}«»“”"'־\-–—]*$/u;
        let anyAuthorFound = false;

        // Prefer a full author only when it is actually adjacent to the title.
        // If it appears later in editor/publication metadata, try the canonical
        // short name before deciding that the citation is ambiguous.
        for (const candidate of authorCandidates) {
            const authorSpans = findSpans(source, candidate).filter(authorSpan =>
                !titleSpans.some(titleSpan =>
                    authorSpan.start < titleSpan.end && titleSpan.start < authorSpan.end
                )
            );
            if (authorSpans.length) anyAuthorFound = true;

            const pairs = [];
            for (const titleSpan of titleSpans) {
                for (const authorSpan of authorSpans) {
                    let separator = "";
                    if (titleSpan.end <= authorSpan.start) {
                        separator = source.slice(titleSpan.end, authorSpan.start);
                    } else if (authorSpan.end <= titleSpan.start) {
                        separator = source.slice(authorSpan.end, titleSpan.start);
                    } else {
                        continue;
                    }

                    if (!punctuationOnly.test(separator)) continue;
                    pairs.push({ titleSpan, authorSpan });
                }
            }

            if (pairs.length > 1) return source;
            if (pairs.length === 1) {
                const selected = pairs[0].authorSpan;
                const matchedTitle = pairs[0].titleSpan;

                // This option is explicitly title-author on the first mention.
                // Reorder only the adjacent identity when the source is author-title;
                // otherwise replace only the author and retain every separator.
                if (selected.end <= matchedTitle.start) {
                    const identityStart = selected.start;
                    const identityEnd = matchedTitle.end;
                    return source.slice(0, identityStart) +
                        matchedTitle.text + "، " + authorToUse +
                        source.slice(identityEnd);
                }

                if (normalize(selected.text) === normalize(authorToUse)) return source;
                return source.slice(0, selected.start) +
                    authorToUse +
                    source.slice(selected.end);
            }
        }

        // Never duplicate an author that is already present in a non-adjacent
        // part of the same citation, such as the editor or biographical data.
        if (anyAuthorFound || !allowMissingAuthor || titleSpans.length !== 1) {
            return source;
        }

        const titleSpan = titleSpans[0];
        return source.slice(0, titleSpan.end) +
            "، " + authorToUse +
            source.slice(titleSpan.end);
    }

    const api = { build: build };
    global.FirstMentionCitation = api;

    if (typeof module !== "undefined" && module.exports) {
        module.exports = api;
    }
})(typeof window !== "undefined" ? window : globalThis);
