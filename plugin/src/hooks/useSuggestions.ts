import { useRef, useState } from "react";
import { closeApplySession, sendApply, type ApplyEdit } from "../api/apply";
import {
  applyDocxToWord,
  getDocumentBytes,
  getTrackChanges,
} from "../onlyoffice/document";
import type {
  ProposedEdit,
  ProposedEditResult,
  Suggestion,
  SuggestionReady,
  SuggestionSet,
} from "../types";
import { base64ToArrayBuffer } from "../utils/base64";

/** Failure codes for a suggestion whose text changed after it was proposed. */
export const STALE_ANCHOR_CODES = new Set([
  "anchor_not_found",
  "multiple_matches",
]);

type Deferred<T> = {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (reason: unknown) => void;
};

/**
 * One server session that every accept joins while any apply is in flight:
 * its children share a batch, so each returned revision holds all of them.
 */
type ApplySession = {
  sessionId: Deferred<string>;
  nextIndex: number;
  inFlight: number;
  /** Compare runs one revision at a time, and never an older one after a newer one. */
  revisions: { chain: Promise<void>; latest: number };
};

function createDeferred<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((onResolve, onReject) => {
    resolve = onResolve;
    reject = onReject;
  });
  // Joiners await it; the opener's own failure is reported on its cards.
  promise.catch(() => undefined);
  return { promise, resolve, reject };
}

function mapSuggestions(
  sets: Record<string, SuggestionSet>,
  ids: ReadonlySet<string>,
  update: (suggestion: Suggestion) => Suggestion,
): Record<string, SuggestionSet> {
  return Object.fromEntries(
    Object.entries(sets).map(([setId, set]) => [
      setId,
      {
        ...set,
        suggestions: set.suggestions.map((suggestion) =>
          ids.has(suggestion.id) ? update(suggestion) : suggestion,
        ),
      },
    ]),
  );
}

function toSuggestion(setId: string, edit: ProposedEdit): Suggestion {
  return {
    id: `${setId}:${edit.index}`,
    index: edit.index,
    old: edit.old,
    proposedNew: edit.new,
    new: edit.new,
    status: "pending",
  };
}

function toSuggestionSet(id: string, result: ProposedEditResult): SuggestionSet {
  return {
    id,
    css: result.css,
    suggestions: result.suggestions.map((edit) => toSuggestion(id, edit)),
  };
}

/** Adds one streamed suggestion; its CSS is appended and deduplicated on render. */
function addStreamedSuggestion(
  sets: Record<string, SuggestionSet>,
  { tool_call_id: setId, suggestion, css }: SuggestionReady,
): Record<string, SuggestionSet> {
  const set = sets[setId] ?? { id: setId, css: "", suggestions: [] };
  const others = set.suggestions.filter(
    (existing) => existing.index !== suggestion.index,
  );
  return {
    ...sets,
    [setId]: {
      ...set,
      css: `${set.css}\n${css}`,
      suggestions: [...others, toSuggestion(setId, suggestion)].sort(
        (a, b) => a.index - b.index,
      ),
    },
  };
}

function describeSuggestion(suggestion: Suggestion): string {
  switch (suggestion.status) {
    case "applied":
      return suggestion.new === suggestion.proposedNew
        ? "applied"
        : "applied after the user edited it";
    case "rejected":
      return "rejected";
    case "failed":
      return STALE_ANCHOR_CODES.has(suggestion.failure?.code ?? "")
        ? "failed because the text had changed"
        : `failed (${suggestion.failure?.reason ?? "unknown error"})`;
    case "applying":
      return "still being applied";
    case "pending":
      return "not reviewed yet";
  }
}

/** The history note that tells the agent what became of its suggestions. */
function describeSuggestionReview(set: SuggestionSet | undefined) {
  if (!set?.suggestions.length) return undefined;
  const outcomes = set.suggestions.map(
    (suggestion) => `#${suggestion.index} ${describeSuggestion(suggestion)}`,
  );
  return `Suggestion review: ${outcomes.join(", ")}.`;
}

async function getDocumentForApply(): Promise<Uint8Array> {
  // Reads the editor's tracking mode, which applyDocxToWord relies on.
  await getTrackChanges();
  return getDocumentBytes();
}

export function useSuggestions(author: string) {
  const [sets, setSets] = useState<Record<string, SuggestionSet>>({});
  const [applying, setApplying] = useState(false);
  const session = useRef<ApplySession | null>(null);
  // Ids already sent, so a quick second accept can't send a card twice.
  const claimed = useRef(new Set<string>());

  function updateSuggestions(ids: string[], patch: Partial<Suggestion>) {
    setSets((previous) =>
      mapSuggestions(previous, new Set(ids), (suggestion) => ({
        ...suggestion,
        ...patch,
      })),
    );
  }

  function failStillApplying(ids: string[], reason: string) {
    setSets((previous) =>
      mapSuggestions(previous, new Set(ids), (suggestion) =>
        suggestion.status === "applying"
          ? { ...suggestion, status: "failed", failure: { code: null, reason } }
          : suggestion,
      ),
    );
  }

  function applyRevision(
    current: ApplySession,
    revision: number,
    docxB64: string,
  ): Promise<void> {
    const revisions = current.revisions;
    revisions.chain = revisions.chain
      .catch(() => undefined)
      .then(async () => {
        if (revision <= revisions.latest) return;
        await applyDocxToWord(base64ToArrayBuffer(docxB64));
        revisions.latest = revision;
      });
    return revisions.chain;
  }

  async function apply(ids: string[]) {
    const wanted = new Set(ids);
    const edits: ApplyEdit[] = Object.values(sets)
      .flatMap((set) => set.suggestions)
      .filter(
        (suggestion) =>
          wanted.has(suggestion.id) &&
          suggestion.status === "pending" &&
          !claimed.current.has(suggestion.id),
      )
      .map(({ id, old, new: replacement }) => ({ id, old, new: replacement }));
    if (edits.length === 0) return;
    const editIds = edits.map((edit) => edit.id);
    for (const id of editIds) claimed.current.add(id);
    updateSuggestions(editIds, { status: "applying" });

    const opening = session.current === null;
    session.current ??= {
      sessionId: createDeferred<string>(),
      nextIndex: 0,
      inFlight: 0,
      revisions: { chain: Promise.resolve(), latest: -1 },
    };
    const current = session.current;
    const startIndex = current.nextIndex;
    current.nextIndex += edits.length;
    current.inFlight += 1;
    setApplying(true);
    try {
      const target = opening
        ? { file: await getDocumentForApply() }
        : { sessionId: await current.sessionId.promise, startIndex };
      await sendApply({ ...target, author, edits }, async (event) => {
        if (event.type === "session") {
          current.sessionId.resolve(event.sessionId);
        } else if (event.type === "edit_applied") {
          await applyRevision(current, event.revision, event.docx_b64);
        } else if (event.type === "suggestion_applied") {
          updateSuggestions([event.id], { status: "applied" });
        } else if (event.type === "suggestion_failed") {
          updateSuggestions([event.id], {
            status: "failed",
            failure: { code: event.code, reason: event.reason },
          });
        }
      });
      failStillApplying(editIds, "The apply ended before this edit finished.");
    } catch (error) {
      if (opening) current.sessionId.reject(error);
      failStillApplying(
        editIds,
        error instanceof Error ? error.message : String(error),
      );
    } finally {
      current.inFlight -= 1;
      if (current.inFlight === 0) {
        session.current = null;
        await current.revisions.chain.catch(() => undefined);
        const sessionId = await current.sessionId.promise.catch(() => null);
        if (sessionId) await closeApplySession(sessionId).catch(console.error);
        setApplying(false);
      }
    }
  }

  return {
    sets,
    applying,
    apply,
    addSuggestion: (event: SuggestionReady) =>
      setSets((previous) => addStreamedSuggestion(previous, event)),
    // The call's final result is authoritative: it replaces the streamed cards,
    // dropping any its batch check rejected. Cards can't be acted on while the
    // agent is still running, so nothing the user did is lost.
    addSet: (id: string, result: ProposedEditResult) =>
      setSets((previous) => ({ ...previous, [id]: toSuggestionSet(id, result) })),
    reject: (id: string) => updateSuggestions([id], { status: "rejected" }),
    edit: (id: string, html: string) => updateSuggestions([id], { new: html }),
    getReview: (setId: string) => describeSuggestionReview(sets[setId]),
  };
}
