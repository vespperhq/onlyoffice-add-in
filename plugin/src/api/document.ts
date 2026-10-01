const DOCX_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// Downloads a file the Document Server exported, through this app's server.
export async function fetchOnlyOfficeFile(
  url: string
): Promise<Uint8Array<ArrayBuffer>> {
  const response = await fetch("/api/onlyoffice/fetch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ url }),
  });
  if (!response.ok) throw new Error(await response.text());
  return new Uint8Array(await response.arrayBuffer());
}

// Stores a revision for the Document Server to download; returns its URL.
export async function uploadRevision(bytes: Uint8Array): Promise<string> {
  const response = await fetch("/api/revisions", {
    method: "POST",
    headers: { "Content-Type": DOCX_TYPE },
    body: Uint8Array.from(bytes).buffer,
  });
  if (!response.ok) throw new Error(await response.text());
  const { url } = (await response.json()) as { url: string };
  return url;
}
