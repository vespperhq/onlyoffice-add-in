# Vespper ONLYOFFICE add-in

A local ONLYOFFICE Docs editor with a Vespper chat plugin, connecting a Mastra
agent to [Vespper](https://vespper.com) through the Vespper TypeScript SDK. It
offers the same chat experience as the Microsoft Word example.

The example includes the following features:

- Streamed edits
- Suggestion cards to review edits before applying them
- Tracked changes
- Selected-text context
- Pasted image context
- Multi-model selection

This example is also available through [Vespper Examples](https://github.com/vespperhq/examples).

## Prerequisites

- Node.js 22.13 or newer
- [Docker Desktop](https://www.docker.com/products/docker-desktop/), running.
  The ONLYOFFICE Document Server runs in a container.
- Port 80 free on your machine. The Document Server must be reachable at the
  same address from the browser and from inside its container.
- A [Vespper account](https://app.vespper.com)
- A model-provider API key for OpenAI, Anthropic, or Google

## 1. Get a Vespper API key

1. [Sign up for Vespper](https://app.vespper.com).
2. Open the [API keys page](https://app.vespper.com/keys).
3. Create a key and copy the `sk_live_...` secret immediately. It is shown only
   once.

## 2. Install the example

Clone this repository, then install the dependencies:

```bash
git clone https://github.com/vespperhq/onlyoffice-add-in.git
cd onlyoffice-add-in
npm install
cp .env.example .env
```

Edit `.env` and add your Vespper key plus the key for the model provider you
want to use:

```bash
VESPPER_API_KEY=sk_live_your_key_here
OPENAI_API_KEY=sk-your_openai_key_here
```

Only one model-provider key is required:

| Model family     | Environment variable                               |
| ---------------- | -------------------------------------------------- |
| OpenAI GPT       | `OPENAI_API_KEY`                                   |
| Anthropic Claude | `ANTHROPIC_API_KEY`                                |
| Google Gemini    | `GOOGLE_API_KEY` or `GOOGLE_GENERATIVE_AI_API_KEY` |

OpenAI GPT 5.6 Sol is the default. To start with another model, set
`ONLYOFFICE_AGENT_MODEL` in `.env`, for example:

```bash
ONLYOFFICE_AGENT_MODEL=anthropic/claude-sonnet-4.5
```

By default, the agent proposes edits as suggestion cards that you accept or
reject. To have it apply edits to the document as it writes them, set:

```bash
USE_SUGGESTIONS=false
```

## 3. Run it

```bash
npm start
```

This starts the ONLYOFFICE Document Server in Docker, builds the plugin, and
starts the local server at <http://localhost:3101>. The first run pulls the
Document Server image, so it can take a few minutes before the editor loads.

Open <http://localhost:3101>. The Vespper plugin starts automatically; click
the **V** icon on the left if its panel is collapsed. Drag the panel's edge to
widen it (up to 600 px).

To verify configuration, open <http://localhost:3101/health>. Both
`vespperConfigured` and `agentConfigured` should be `true`.

Stop the Document Server with:

```bash
npm stop
```

## Try an edit

1. Open <http://localhost:3101>. A welcome document is loaded the first time.
   To edit your own file, use **Open .docx** above the editor.
2. Open the Vespper panel.
3. Ask for an edit.
4. Optionally select document text or paste images into the prompt.

That's it. You should see the document being changed.

The plugin follows the editor's **Track Changes** setting. When it is on, new
edits are recorded as tracked changes. When it is off, they are applied
directly.

The settings button lets you change the tracked-change author used in
tracked mode from the default `Vespper Agent`.

The open document is stored at `data/document.docx`. Opening a file replaces
it and starts a fresh editing session. To get the welcome document back, stop
the app and delete that file.

## Development commands

| Command             | Purpose                                                 |
| ------------------- | ------------------------------------------------------- |
| `npm start`         | Start the Document Server, then build and serve the app |
| `npm stop`          | Stop the Document Server container                      |
| `npm run dev`       | Build and serve the app with a watcher, without Docker  |
| `npm run build`     | Build the plugin assets once                            |
| `npm test`          | Run the server and plugin tests                         |
| `npm run typecheck` | Type-check the server and plugin                        |

## Project layout

```text
onlyoffice-add-in/
├── server/              Local server and Mastra agent
├── plugin/
│   ├── config.json      ONLYOFFICE plugin manifest
│   └── src/             React chat panel and ONLYOFFICE integration
├── public/              Page that hosts the ONLYOFFICE editor
├── shared/              Message contracts shared by client and server
├── data/                The open document (created on first run)
├── docker-compose.yml   ONLYOFFICE Document Server
├── .env.example
└── package.json
```

## Security

`.env` is ignored by Git. Vespper and model-provider keys are read only by the
local Node.js server and are never bundled into the plugin.

This example is intentionally local and disables ONLYOFFICE JWT validation. Do
not expose it directly to a network. A production deployment should enable JWT,
authenticate the app routes, restrict the download proxy, and define a
multi-user document versioning strategy.

## Updating the examples collection

Push changes to this repository normally. The workflow in
[Vespper Examples](https://github.com/vespperhq/examples) checks all of its
submodules every five minutes and updates their pointers automatically.
No workflow or secret is required in this repository. GitHub may delay scheduled
runs, so an update can take longer than five minutes.
