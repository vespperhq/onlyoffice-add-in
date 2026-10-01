# Vespper ONLYOFFICE add-in

A local ONLYOFFICE Docs editor with the same Vespper chat experience as the
Microsoft Word example. It exports the active DOCX, sends it to the Vespper
agent, streams agent activity and edited DOCX revisions back into the editor,
supports selected-text context and pasted images, and preserves Vespper tracked
changes.

This example is also available through [Vespper Examples](https://github.com/vespperhq/examples).

## Requirements

- Node.js 22.13 or newer
- Docker Desktop
- Port 80 free: the Document Server must be reachable at the same address from
  the browser and from inside its container
- A Vespper API key
- An API key for the selected model

## Run locally

Clone this repository, then install and start it:

```bash
git clone https://github.com/vespperhq/onlyoffice-add-in.git
cd onlyoffice-add-in
cp .env.example .env
# Add VESPPER_API_KEY and OPENAI_API_KEY to .env
npm install
npm start
```

Open <http://localhost:3101>. The first Docker startup can take a few minutes.
The Vespper plugin starts automatically; click the **V** icon on the right if
its panel is collapsed.

Stop the editor with:

```bash
npm run stop
```

Use **Open .docx** above the editor to load a document. It replaces the current
document, which is stored at `data/document.docx`, and always opens a fresh
editing session. Delete that file while the app is stopped to recreate the
welcome document.

## How document updates work

The plugin calls ONLYOFFICE `GetFileToDownload("docx")` to export the active
file. The local server proxies that signed download URL and submits the DOCX to
Vespper.

Each committed revision is applied in place with ONLYOFFICE's Compare, without
reloading, so undo history, the cursor, and the chat panel survive. Vespper
edits the document directly, and Compare records the differences:

- With tracked changes on, they stay as tracked changes. Compare credits them
  to the revision's last author, which the server sets to the tracked-change
  author.
- With tracked changes off, they are accepted.

The plugin hosts each revision on the local server (`/api/revisions`) so the
Document Server can fetch it.

Compare accepts a document's existing tracked changes before it diffs, so the
plugin patches the editor functions involved for the duration of each
comparison (`plugin/src/onlyoffice/patches`). Most ONLYOFFICE builds minify
them under names that change every release, so the plugin finds them at
runtime; this was verified on ONLYOFFICE Docs 8.3.3 through 9.4.0. On a build
where it can't, or while someone else edits the document (Compare refuses to
run then), a turn fails with an error instead of applying the edits.

This example is intentionally local and disables ONLYOFFICE JWT validation. Do
not expose it directly to a network. A production deployment should enable JWT,
authenticate the app routes, restrict the download proxy, and define a
multi-user document versioning strategy.
