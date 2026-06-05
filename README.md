# SuperDoc Editor

Standalone Vue/Vite app that hosts the SuperDoc document editor for embedding in another app through an iframe.

## License

This project is licensed under the GNU Affero General Public License v3.0 only (`AGPL-3.0-only`). See `LICENSE` for the full license text.

If you deploy a modified version of this app for users over a network, AGPLv3 requires you to make the corresponding source code available to those users.

## Local Development

```sh
npm install
npm run dev
```

The dev server defaults to `http://localhost:3013`.

Copy `.env.example` to `.env` when you need local environment overrides.

## Scripts

```sh
npm run dev
npm run type-check
npm run build
npm run preview
```

## Parent Integration

The parent app embeds this app in an iframe and communicates with it using `window.postMessage`.

Parent-to-iframe messages:

- `init`: initial editor configuration.
- `setDocument`: replace the currently loaded document.
- `exportDocx`: export the current editor content as a `.docx` file.
- `isEmpty`: check whether the editor content is empty.

Iframe-to-parent messages:

- `ready`: editor initialized.
- `update`: editor content changed.
- `docxSelected`: user selected a local `.docx` from the toolbar.
- `exportDocxResult`: result for an export request.
- `isEmptyResult`: result for an empty-content request.
- `error`: structured error for failed requests.

Set `VITE_ALLOWED_PARENT_ORIGINS` to a comma-separated list of allowed parent origins in deployed environments.

## Publishing to GitHub

Before pushing this project as its own repository:

1. Confirm `LICENSE` and `package.json` still reflect `AGPL-3.0-only`.
2. Set the GitHub repository visibility to public.
3. Add the deployed editor URL to the parent app as `VITE_SUPERDOC_EDITOR_URL`.
4. Configure `VITE_ALLOWED_PARENT_ORIGINS` in deployed environments.
