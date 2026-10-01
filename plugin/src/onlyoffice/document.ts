import { fetchOnlyOfficeFile, uploadRevision } from "../api/document";
import { sleep } from "../utils/time";
import {
  callCommand,
  executeMethod,
  withCompareFunctionPatcher,
} from "./commands";
import type { CompareFunctionPatcherClass } from "./patches/types";

const COMPARE_TIMEOUT_MS = 60_000;
const COMPARE_POLL_MS = 150;

// Whether the editor records edits as tracked changes, read when a turn starts.
let tracked = false;

// Runs inside the editor, so it may only use editor globals.
function readEditorState(Patcher: CompareFunctionPatcherClass) {
  // asc_coAuthoringGetUsers reports the participants synchronously.
  let editors = 0;
  function countEditors(users: Record<string, { asc_getView(): boolean }>) {
    for (const id of Object.keys(users)) {
      if (!users[id].asc_getView()) editors += 1;
    }
  }
  editor.asc_registerCallback("asc_onAuthParticipantsChanged", countEditors);
  editor.asc_coAuthoringGetUsers();
  editor.asc_unregisterCallback("asc_onAuthParticipantsChanged", countEditors);

  let canPatch = false;
  try {
    canPatch = new Patcher().isSupported();
  } catch {}

  return {
    editors,
    tracked: Api.GetDocument().IsTrackRevisions(),
    canPatch,
  };
}

// Revisions reach the editor through Compare, which refuses to run while
// someone else edits the document and needs its functions patched to keep the
// document's existing tracked changes.
async function checkEditor() {
  const state = await callCommand(withCompareFunctionPatcher(readEditorState));
  if (!state.canPatch) {
    throw new Error("This ONLYOFFICE version isn't supported");
  }
  if (state.editors > 1) {
    throw new Error("Close the document for other editors to edit it");
  }
  return state;
}

// Returns whether Vespper should record its edits as tracked changes. It never
// does: Compare records them itself when the editor tracks changes.
export async function getTrackChanges(): Promise<boolean> {
  ({ tracked } = await checkEditor());
  return false;
}

export async function getDocumentBytes(): Promise<Uint8Array<ArrayBuffer>> {
  const url = await executeMethod<string>("GetFileToDownload", ["docx"]);
  if (!url) throw new Error("ONLYOFFICE did not return a DOCX download URL");
  return fetchOnlyOfficeFile(url);
}

export async function applyDocxToWord(
  data: ArrayBuffer | Uint8Array
): Promise<void> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  await checkEditor();

  // Untracked edits are credited to a throwaway author so that exactly those
  // changes can be accepted afterwards, leaving existing ones pending.
  const author = tracked ? null : `vespper-${crypto.randomUUID()}`;
  Asc.scope.documentUpdate = { url: await uploadRevision(bytes), author };
  await callCommand(
    withCompareFunctionPatcher((Patcher) => {
      // sdkjs's Asc.c_oAscAsyncActionType and Asc.c_oAscAsyncAction, which a
      // command can't reach: callCommand shadows Asc with its own object.
      enum AsyncActionType {
        BlockInteraction = 1,
      }
      enum AsyncAction {
        SlowOperation = 11,
      }

      const { url, author } = Asc.scope.documentUpdate as {
        url: string;
        author: string | null;
      };
      const patcher = new Patcher();

      const state: DocumentUpdateState = {
        status: "pending",
        error: 0,
        cancel: () => finish("error"),
      };
      function finish(status: "done" | "error") {
        if (state.status !== "pending") return;
        state.status = status;
        patcher.unpatch();
        editor.asc_unregisterCallback("asc_onEndAction", onEndAction);
        editor.asc_unregisterCallback("asc_onError", onError);
      }
      // Compare blocks interaction with a slow operation until the diff has
      // been applied to the document.
      function onEndAction(type: AsyncActionType, id: AsyncAction) {
        if (
          type === AsyncActionType.BlockInteraction &&
          id === AsyncAction.SlowOperation
        ) {
          finish("done");
        }
      }
      function onError(id: number) {
        state.error = id;
        finish("error");
      }

      editor.vespperDocumentUpdate = state;
      editor.asc_registerCallback("asc_onEndAction", onEndAction);
      editor.asc_registerCallback("asc_onError", onError);
      // A throw would make callCommand retry, patching a second time.
      try {
        patcher.patch(author);
        editor.asc_CompareDocumentUrl(
          url,
          new AscCommonWord.ComparisonOptions()
        );
      } catch {
        finish("error");
      }
      return true;
    })
  );

  await waitForCompare();
  if (author) await acceptChangesBy(author);
}

async function waitForCompare(): Promise<void> {
  const deadline = Date.now() + COMPARE_TIMEOUT_MS;
  while (Date.now() < deadline) {
    await sleep(COMPARE_POLL_MS);
    const state = await callCommand(() => {
      const current = editor.vespperDocumentUpdate;
      return current ? { status: current.status, error: current.error } : null;
    });
    if (state?.status === "pending") continue;
    if (state?.status === "done") return;
    throw new Error(
      `ONLYOFFICE could not compare the document (error ${state?.error ?? "unknown"})`
    );
  }
  // Compare never finished, so its patches must not outlive this revision.
  await callCommand(() => {
    editor.vespperDocumentUpdate?.cancel?.();
    return true;
  });
  throw new Error("Timed out comparing the document");
}

async function acceptChangesBy(author: string): Promise<void> {
  Asc.scope.author = author;
  const unaccepted = await callCommand(() => {
    const author = Asc.scope.author as string;
    const changesBy = () =>
      editor.asc_GetTrackRevisionsReportByAuthors()[author] ?? [];
    // Accepting a change reshapes the content around it, so the rest of a
    // report can point at positions that no longer exist. Accepting those
    // throws and the editor then rolls back the whole command, so the report
    // is read again after every change.
    let changes = changesBy();
    while (changes.length) {
      editor.asc_AcceptChanges(changes[0]);
      const remaining = changesBy();
      if (remaining.length >= changes.length) break;
      changes = remaining;
    }
    return changes.length;
  });
  if (unaccepted) {
    throw new Error(`${unaccepted} compared changes could not be accepted`);
  }
}
