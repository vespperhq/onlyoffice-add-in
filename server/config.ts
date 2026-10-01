import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

export const ROOT_DIR = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  ".."
);
dotenv.config({ path: path.join(ROOT_DIR, ".env") });

export const PORT = Number(process.env.PORT ?? 3101);
export const DOCX_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
export const DATA_DIR = path.join(ROOT_DIR, "data");
export const DOCUMENT_PATH = path.join(DATA_DIR, "document.docx");
export const MAX_REVISIONS = 20;
// public/index.html and plugin/index.html load the Document Server's scripts
// from this address too.
export const ONLYOFFICE_PUBLIC_URL = "http://localhost";
// Hostnames the Document Server may put in the download links it hands to
// the plugin: its public name, or a name it only has inside Docker.
export const ONLYOFFICE_HOSTNAMES = new Set([
  new URL(ONLYOFFICE_PUBLIC_URL).hostname,
  "localhost",
  "127.0.0.1",
  "onlyoffice",
  "host.docker.internal",
]);
export const APP_PUBLIC_URL =
  process.env.APP_PUBLIC_URL ?? `http://localhost:${PORT}`;
// This server's address from inside the Document Server container, which
// downloads the document and revisions from it and posts save callbacks to it.
export const APP_DOCKER_URL =
  process.env.APP_DOCKER_URL ?? `http://host.docker.internal:${PORT}`;
export const VESPPER_API_KEY = process.env.VESPPER_API_KEY ?? "";
export const VESPPER_MCP_URL = (
  process.env.VESPPER_MCP_URL ?? "https://mcp.vespper.com/mcp"
).replace(/\/+$/, "");
export const DOCX_AUTHOR = process.env.DOCX_AUTHOR ?? "Vespper Agent";
// The agent proposes edits as suggestion cards; false applies them as it writes.
export const USE_SUGGESTIONS = process.env.USE_SUGGESTIONS !== "false";
export const DEFAULT_MODEL =
  process.env.ONLYOFFICE_AGENT_MODEL ??
  process.env.WORD_AGENT_MODEL ??
  "openai/gpt-5.6-sol";
export const REASONING_EFFORT =
  process.env.AGENT_REASONING_EFFORT ?? "medium";
export const REASONING_SUMMARY =
  process.env.AGENT_REASONING_SUMMARY ?? "detailed";
export const MAX_ROUNDS = 24;
export const PLUGIN_GUID = "asc.{A79D9B7A-2BFA-4EB1-A4C5-7EAD5B6F80D1}";
export const AVAILABLE_MODELS = [
  "openai/gpt-5.6-sol",
  "openai/gpt-5.5",
  "anthropic/claude-sonnet-4.5",
  "google/gemini-2.5-flash",
  "google/gemini-2.5-pro",
];

export const META_BATCH_ID = "com.vespper/batch-id";
export const META_EDIT_INDEX = "com.vespper/edit-index";

export function getModelApiKeyNames(model: string): string[] {
  const provider = model.includes("/") ? model.split("/", 1)[0] : "";
  if (provider === "openai") return ["OPENAI_API_KEY"];
  if (provider === "anthropic") return ["ANTHROPIC_API_KEY"];
  if (provider === "google") {
    return ["GOOGLE_API_KEY", "GOOGLE_GENERATIVE_AI_API_KEY"];
  }
  return [];
}

export function hasModelApiKey(model: string): boolean {
  return getModelApiKeyNames(model).some((name) => Boolean(process.env[name]));
}

export const HEALTH = {
  ok: true,
  vespperConfigured: Boolean(VESPPER_API_KEY),
  agentConfigured: hasModelApiKey(DEFAULT_MODEL),
  agentModel: DEFAULT_MODEL,
  defaultModel: DEFAULT_MODEL,
  availableModels: AVAILABLE_MODELS,
  mcpUrl: VESPPER_MCP_URL,
  editor: "onlyoffice",
  suggestions: USE_SUGGESTIONS,
};

// The Document Server editor config for the document with this key.
export function getEditorConfig(key: string) {
  return {
    document: {
      fileType: "docx",
      key,
      title: "Vespper document.docx",
      url: `${APP_DOCKER_URL}/document.docx?v=${key}`,
      permissions: {
        edit: true,
        download: true,
        review: true,
      },
    },
    documentType: "word",
    editorConfig: {
      callbackUrl: `${APP_DOCKER_URL}/api/onlyoffice/callback`,
      mode: "edit",
      user: { id: "local-user", name: "Local user" },
      customization: {
        autosave: true,
        forcesave: true,
      },
      plugins: {
        autostart: [PLUGIN_GUID],
        pluginsData: [`${APP_PUBLIC_URL}/plugin/config.json`],
      },
    },
    height: "100%",
    width: "100%",
    type: "desktop",
  };
}
