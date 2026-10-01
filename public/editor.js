(function startEditor() {
  const DOCX_TYPE =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const status = document.getElementById("status");
  let editor;

  function showStatus(text) {
    status.textContent = text;
    status.hidden = !text;
  }

  async function loadConfig() {
    const response = await fetch("/api/editor-config");
    if (!response.ok) throw new Error(await response.text());
    return response.json();
  }

  async function mountEditor() {
    editor?.destroyEditor();
    editor = undefined;
    showStatus("Starting ONLYOFFICE…");
    try {
      const config = await loadConfig();
      config.events = {
        onDocumentReady() {
          showStatus("");
        },
        onError(event) {
          showStatus(`ONLYOFFICE error: ${event.data?.errorCode ?? "unknown"}`);
        },
        onRequestRefreshFile() {
          void refreshEditor();
        },
      };
      editor = new DocsAPI.DocEditor("editor", config);
      window.docEditor = editor;
    } catch (error) {
      showStatus(
        error instanceof Error ? error.message : "Failed to start ONLYOFFICE"
      );
    }
  }

  // The Document Server asks for this when the file's version changed under an
  // open editor. refreshFile swaps it in place, keeping the editor frame.
  async function refreshEditor() {
    if (!editor) return mountEditor();
    try {
      editor.refreshFile(await loadConfig());
    } catch (error) {
      console.error("refreshFile failed; remounting the editor", error);
      await mountEditor();
    }
  }

  async function openFile(file) {
    const response = await fetch("/api/document", {
      method: "PUT",
      headers: { "Content-Type": DOCX_TYPE },
      body: file,
    });
    if (!response.ok) {
      showStatus(await response.text());
      return;
    }
    await mountEditor();
  }

  document.getElementById("open-file").addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (file) void openFile(file);
  });

  void mountEditor();
})();
