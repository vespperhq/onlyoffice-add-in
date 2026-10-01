import JSZip from "jszip";

const CORE_PATH = "docProps/core.xml";
const LAST_MODIFIED_BY =
  /<cp:lastModifiedBy\s*\/>|<cp:lastModifiedBy>[\s\S]*?<\/cp:lastModifiedBy>/;

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ONLYOFFICE's Compare credits the tracked changes it creates to the revised
// document's last author and has no option for choosing another one.
export async function setLastModifiedBy(
  docx: Uint8Array,
  author: string,
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx);
  const core = await zip.file(CORE_PATH)?.async("string");
  if (core === undefined) return Buffer.from(docx);

  const tag = `<cp:lastModifiedBy>${escapeXml(author)}</cp:lastModifiedBy>`;
  zip.file(
    CORE_PATH,
    core
      .replace(LAST_MODIFIED_BY, "")
      .replace("</cp:coreProperties>", `${tag}</cp:coreProperties>`),
  );
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
