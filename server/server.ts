import fs from "node:fs/promises";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import express, { type Request, type Response } from "express";
import multer from "multer";
import { Document, Packer, Paragraph, TextRun } from "docx";
import {
  APP_DOCKER_URL,
  APP_PUBLIC_URL,
  DATA_DIR,
  DEFAULT_MODEL,
  DOCUMENT_PATH,
  DOCX_AUTHOR,
  DOCX_TYPE,
  getEditorConfig,
  getModelApiKeyNames,
  hasModelApiKey,
  HEALTH,
  MAX_REVISIONS,
  ONLYOFFICE_HOSTNAMES,
  ONLYOFFICE_PUBLIC_URL,
  PORT,
  ROOT_DIR,
  VESPPER_API_KEY,
  VESPPER_MCP_URL,
} from "./config";
import { runAgent } from "./agent";
import { DocumentVersionGuard } from "./document-version";
import { parseChatMessages } from "./types";

const documentVersionGuard = new DocumentVersionGuard();
const revisions = new Map<string, Buffer>();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024, fieldSize: 128 * 1024 * 1024 },
});

async function ensureDocument() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    await fs.access(DOCUMENT_PATH);
  } catch {
    const document = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "Welcome to Vespper for ONLYOFFICE",
                  bold: true,
                  size: 32,
                }),
              ],
            }),
            new Paragraph(
              "Open the Vespper panel on the right and describe the edits you want.",
            ),
          ],
        },
      ],
    });
    await fs.writeFile(DOCUMENT_PATH, await Packer.toBuffer(document));
  }
}

let keyGeneration = 0;

async function documentKey() {
  const file = await fs.readFile(DOCUMENT_PATH);
  return createHash("sha256")
    .update(file)
    .update(String(keyGeneration))
    .digest("hex")
    .slice(0, 32);
}

// Only fetch Document Server links, and always through ONLYOFFICE_PUBLIC_URL,
// which is reachable from this process.
function documentServerUrl(raw: string) {
  const url = new URL(raw, ONLYOFFICE_PUBLIC_URL);
  if (!ONLYOFFICE_HOSTNAMES.has(url.hostname)) {
    throw new Error("Refusing to fetch a non-ONLYOFFICE URL");
  }
  const { protocol, host } = new URL(ONLYOFFICE_PUBLIC_URL);
  url.protocol = protocol;
  url.host = host;
  return url;
}

async function fetchOnlyOfficeFile(raw: string) {
  const response = await fetch(documentServerUrl(raw));
  if (!response.ok) {
    throw new Error(`ONLYOFFICE download failed (${response.status})`);
  }
  return Buffer.from(await response.arrayBuffer());
}

const app = express();
app.use((_req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (_req.method === "OPTIONS") return res.sendStatus(204);
  next();
});
app.use(express.json({ limit: "2mb" }));

app.get("/health", (_req, res) => {
  res.json(HEALTH);
});

app.get("/api/editor-config", async (_req, res) => {
  const key = await documentKey();
  const config = getEditorConfig(key);

  res.json(config);
});

app.get("/document.docx", async (_req, res) => {
  res.type(DOCX_TYPE).sendFile(DOCUMENT_PATH);
});

app.put(
  "/api/document",
  express.raw({ type: "*/*", limit: "50mb" }),
  async (req: Request, res: Response) => {
    const body = Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body);
    if (body.length < 512) {
      return res.status(400).send("Refusing to save an invalid DOCX");
    }
    documentVersionGuard.markReplaced(await documentKey());
    await fs.writeFile(DOCUMENT_PATH, body);
    // Identical bytes would otherwise reuse the old key, and with it the
    // Document Server's cached session (or cached open failure).
    keyGeneration += 1;
    return res.sendStatus(204);
  },
);

app.post(
  "/api/revisions",
  express.raw({ type: "*/*", limit: "50mb" }),
  (req: Request, res: Response) => {
    const body = Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body);
    if (body.length < 512) {
      return res.status(400).send("Refusing to store an invalid DOCX");
    }
    const id = randomUUID();
    revisions.set(id, body);
    if (revisions.size > MAX_REVISIONS) {
      const oldest = revisions.keys().next().value;
      if (oldest) revisions.delete(oldest);
    }
    return res.json({ url: `${APP_DOCKER_URL}/api/revisions/${id}.docx` });
  },
);

app.get("/api/revisions/:file", (req, res) => {
  const body = revisions.get(req.params.file.replace(/\.docx$/, ""));
  if (!body) return res.sendStatus(404);
  return res.type(DOCX_TYPE).send(body);
});

app.post("/api/onlyoffice/fetch", async (req, res) => {
  try {
    const body = await fetchOnlyOfficeFile(String(req.body?.url ?? ""));
    res.type(DOCX_TYPE).send(body);
  } catch (error) {
    res
      .status(502)
      .send(error instanceof Error ? error.message : String(error));
  }
});

app.post("/api/onlyoffice/callback", async (req, res) => {
  try {
    if (
      [2, 6].includes(Number(req.body?.status)) &&
      req.body?.url &&
      documentVersionGuard.shouldSaveCallback(req.body?.key)
    ) {
      await fs.writeFile(
        DOCUMENT_PATH,
        await fetchOnlyOfficeFile(String(req.body.url)),
      );
    }
    res.json({ error: 0 });
  } catch (error) {
    console.error("ONLYOFFICE callback failed:", error);
    res.json({ error: 1 });
  }
});

app.post(
  "/api/word/process",
  upload.single("file"),
  async (req: Request, res: Response) => {
    if (!VESPPER_API_KEY?.startsWith("sk_live_")) {
      return res.status(500).json({
        error: "Set a valid VESPPER_API_KEY in onlyoffice-add-in/.env.",
      });
    }
    if (!req.file?.buffer.length) {
      return res.status(400).json({ error: "Missing DOCX upload" });
    }

    let messages;
    try {
      messages = parseChatMessages(req.body.messages);
    } catch (error) {
      return res.status(400).json({
        error: error instanceof Error ? error.message : "Invalid messages",
      });
    }
    const model = String(req.body.model || DEFAULT_MODEL);
    if (!hasModelApiKey(model)) {
      return res.status(500).json({
        error: `Set ${getModelApiKeyNames(model).join(" or ")} for ${model}.`,
      });
    }

    res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    const controller = new AbortController();
    req.once("aborted", () => controller.abort());

    const stream = runAgent({
      docBytes: req.file.buffer,
      messages,
      author: String(req.body.author || DOCX_AUTHOR),
      model,
      trackChanges: String(req.body.trackChanges ?? "true") === "true",
      mcpUrl: VESPPER_MCP_URL,
      apiKey: VESPPER_API_KEY,
      signal: controller.signal,
    });
    try {
      for await (const event of stream) {
        if (res.destroyed) break;
        res.write(
          `${JSON.stringify(event, (key, value) => {
            if (key === "html") return undefined;
            if (value instanceof Error) return value.message;
            return value;
          })}\n`,
        );
      }
    } catch (error) {
      if (!controller.signal.aborted && !res.destroyed) {
        res.write(
          `${JSON.stringify({
            type: "error",
            detail: error instanceof Error ? error.message : String(error),
          })}\n`,
        );
      }
    } finally {
      if (!res.destroyed) res.end();
    }
    return;
  },
);

app.use(
  "/plugin",
  express.static(path.join(ROOT_DIR, "plugin"), {
    etag: false,
    lastModified: false,
    setHeaders: (res) => res.setHeader("Cache-Control", "no-store"),
  }),
);
app.use(express.static(path.join(ROOT_DIR, "public")));

await ensureDocument();
app.listen(PORT, () => {
  console.log(`Vespper ONLYOFFICE: ${APP_PUBLIC_URL}`);
  console.log(`ONLYOFFICE Docs: ${ONLYOFFICE_PUBLIC_URL}`);
});
