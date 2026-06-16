# SuperDoc Editor

Open-source Vue 3 and Vite app that hosts the SuperDoc document editor for iframe-based integrations.

This project is intended to be deployed as a standalone editor surface. A parent application embeds it in an iframe and communicates with it through `window.postMessage`.

## Features

- Embeddable SuperDoc document editor.
- `.docx` loading and exporting.
- Optional toolbar controls for opening files, first-page header/footer handling, and page breaks.
- Parent-controlled initialization, document replacement, export, empty-state checks, and redline diff generation.
- Parent-origin allow-listing for deployed environments.

## License

SuperDoc Editor is open source and licensed under the MIT License. See `LICENSE` for the full license text.

## Getting Started

Install dependencies:

```sh
npm install
```

Run the local development server:

```sh
npm run dev
```

The dev server defaults to `http://localhost:3013`.

Copy `.env.example` to `.env` when you need local environment overrides. Set `VITE_ALLOWED_PARENT_ORIGINS` to a comma-separated list of parent app origins that are allowed to send editor messages.

## Scripts

- `npm run dev`: start the Vite dev server.
- `npm run type-check`: run TypeScript and Vue type checks.
- `npm run build`: type-check and build the production app.
- `npm run preview`: preview the production build locally.

## Parent Integration

Embed the deployed editor URL in an iframe:

```html
<iframe
  src="https://your-editor-host.example.com"
  title="SuperDoc Editor"
></iframe>
```

All messages must include the `superdoc-editor` namespace and a `type`.

```ts
iframe.contentWindow?.postMessage(
  {
    namespace: 'superdoc-editor',
    type: 'init',
    requestId: 'init-1',
    payload: {
      document: null,
      user: { name: 'Editor User', email: 'editor@example.com' },
      documentMode: 'editing',
      role: 'editor',
      hideToolbar: false,
      showOpenDocx: true,
      showDifferentFirstPage: true,
      showPageBreak: true,
      trackChangesVisible: false,
    },
  },
  'https://your-editor-host.example.com',
)
```

Parent-to-iframe messages:

- `init`: initialize editor configuration.
- `setDocument`: replace the currently loaded document.
- `exportDocx`: export the current editor content as a `.docx` file.
- `isEmpty`: check whether the editor content is empty.
- `runDiff`: generate a redlined `.docx` from original and revised documents.

Iframe-to-parent messages:

- `ready`: editor initialized.
- `update`: editor content changed.
- `docxSelected`: user selected a local `.docx` from the toolbar.
- `exportDocxResult`: result for an export request.
- `isEmptyResult`: result for an empty-content request.
- `runDiffResult`: result for a redline diff request.
- `error`: structured error for failed requests.

## Deployment

Build the app before deploying:

```sh
npm run build
```

Deploy the generated `dist/` directory to your static host, then configure the parent app with the deployed editor URL. In production, set `VITE_ALLOWED_PARENT_ORIGINS` so only trusted parent origins can initialize and control the editor.
