import { randomUUID } from "node:crypto";
import { PassThrough, Readable } from "node:stream";
import { finished } from "node:stream/promises";
import { Agent } from "@mastra/core/agent";
import { createTool } from "@mastra/core/tools";
import { MCPClient } from "@mastra/mcp";
import Vespper from "vespper";
import {
  DEFAULT_MODEL,
  MAX_ROUNDS,
  META_BATCH_ID,
  META_EDIT_INDEX,
  REASONING_EFFORT,
  REASONING_SUMMARY,
} from "./config";
import { createEditInputParser } from "./edit-input";
import {
  CommittedDocumentSchema,
  type EditInputParser,
  type EditPair,
  type RunAgentTurnOptions,
} from "./types";
import { setLastModifiedBy } from "./utils";

const SYSTEM_PROMPT = `You are a DOCX editing assistant. You will receive an existing document and a natural-language instruction.
Edit the existing document through edit_document. Do not regenerate it. Preserve formatting, styles, tables, numbering, headers, footers, images, and unrelated content. Make the smallest edits that satisfy the instruction.`;

export async function* runAgent(options: RunAgentTurnOptions) {
  const client = new Vespper({
    apiKey: options.apiKey,
    mcpUrl: options.mcpUrl,
  });
  const sessionId = await client.openSession(options.docBytes);
  const mcp = new MCPClient({
    id: randomUUID(),
    servers: {
      vespperDocx: {
        url: new URL(client.mcpUrl),
        requestInit: {
          headers: { Authorization: client.authorizationHeader },
        },
      },
    },
  });
  let editCount = 0;
  let latest: { revision: number; document: Buffer } | undefined;
  const output = new PassThrough({ objectMode: true });
  const childCalls: Promise<unknown>[] = [];
  let activeInput:
    | { batchId: string; parser: EditInputParser }
    | undefined;

  try {
    await client.patchMCPTools({
      mcp,
      sessionId,
      author: options.author,
      trackChanges: options.trackChanges,
      async onDocumentUpdated(update) {
        const document = await setLastModifiedBy(
          update.document,
          options.author
        );
        if (!latest || update.revision > latest.revision) {
          latest = { revision: update.revision, document };
        }
        output.write({
          type: "edit_applied",
          docx_b64: document.toString("base64"),
          edit_count: update.editCount,
          revision: update.revision,
        });
      },
    });
    const tools: any = await mcp.listTools();
    const readDocument = tools.vespperDocx_read_document;
    const searchDocument = tools.vespperDocx_search_document;
    const mcpEdit = tools.vespperDocx_edit_document;
    if (!readDocument || !searchDocument || !mcpEdit) {
      throw new Error("Vespper did not advertise the required DOCX tools");
    }

    function sendEdit(
      toolCallId: string,
      abortSignal: AbortSignal | undefined,
      index: number,
      edit: EditPair
    ) {
      childCalls.push(
        mcpEdit
          .execute(
            { edits: [edit] },
            {
              _meta: {
                [META_BATCH_ID]: toolCallId,
                [META_EDIT_INDEX]: index,
              },
              abortSignal,
            }
          )
          .catch(() => undefined)
      );
    }

    const editDocument = createTool({
      id: "edit_document",
      description: mcpEdit.description,
      inputSchema: mcpEdit.inputSchema,
      onInputStart: ({ toolCallId, abortSignal }) => {
        if (activeInput && !activeInput.parser.ended) {
          throw new Error("Parallel edit_document calls are unsupported");
        }
        activeInput = {
          batchId: toolCallId,
          parser: createEditInputParser({
            onEdit: sendEdit.bind(null, toolCallId, abortSignal),
            onError: () => {
              if (activeInput?.batchId === toolCallId) activeInput = undefined;
            },
          }),
        };
      },
      onInputDelta: ({ toolCallId, inputTextDelta }) => {
        if (activeInput?.batchId !== toolCallId) return;
        activeInput.parser.write(inputTextDelta);
      },
      onInputAvailable: ({ toolCallId }) => {
        if (activeInput?.batchId !== toolCallId) return;
        activeInput.parser.finish();
        activeInput = undefined;
      },
      execute: async (input, context) => {
        const batchId = context.agent?.toolCallId;
        if (!batchId) throw new Error("edit_document has no tool call ID");
        await Promise.allSettled([...childCalls]);
        const raw = await mcpEdit.execute(input, {
          _meta: { [META_BATCH_ID]: batchId },
          abortSignal: context.abortSignal,
        });
        const committed = CommittedDocumentSchema.safeParse(raw);
        if (committed.success) editCount += committed.data.count;
        if (!raw || typeof raw !== "object") return raw;
        const { base64: _base64, ...modelResult } = raw as Record<
          string,
          unknown
        >;
        return modelResult;
      },
    });

    const agent = new Agent({
      id: "onlyoffice-editor",
      name: "ONLYOFFICE Editor",
      instructions: SYSTEM_PROMPT,
      model: options.model || DEFAULT_MODEL,
      tools: {
        read_document: readDocument,
        search_document: searchDocument,
        edit_document: editDocument,
      },
    });
    const stream = await agent.stream(options.messages, {
      maxSteps: MAX_ROUNDS,
      abortSignal: options.signal,
      providerOptions: {
        openai: {
          reasoningSummary: REASONING_SUMMARY,
          reasoningEffort: REASONING_EFFORT,
        },
      },
    });
    const modelOutput = Readable.from(stream.fullStream, {
      objectMode: true,
    });
    modelOutput.pipe(output, { end: false });

    const completion = (async () => {
      try {
        await finished(modelOutput);
        await Promise.allSettled(childCalls);
        const document =
          latest?.document ??
          Buffer.from(client.getSessionDocument(sessionId));
        output.end({
          type: "done",
          docx_b64: document.toString("base64"),
          edit_count: editCount,
        });
      } catch (error) {
        output.destroy(
          error instanceof Error ? error : new Error(String(error))
        );
      }
    })();

    for await (const event of output) yield event;
    await completion;
  } finally {
    try {
      await mcp.disconnect();
    } finally {
      await client.closeSession(sessionId).catch(console.error);
    }
  }
}
