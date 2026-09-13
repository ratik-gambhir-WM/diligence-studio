# diligence-studio

The repository is an npm workspace with two explicit runtime boundaries:

```text
diligence-studio/
├── web/       React 19 and Vite browser application
└── server/    Express API, SQLite persistence, and Node PowerPoint tooling
```

The web package never imports the server or Node-only PowerPoint importer. The server owns its
runtime code, tests, database, assets, and command-line PowerPoint tools.

From the repository root:

```sh
npm install
npx playwright install chromium
npm run dev
npm run server:dev
npm run typecheck
npm test
npm run build
```

`npm run dev` starts the web app at Vite's local URL and proxies `/api/v1/*` requests to the API at
`http://127.0.0.1:43127`. `npm run server:dev` starts that API. A production deployment must provide
the same `/api/v1` reverse-proxy boundary, or set the public `VITE_API_BASE_URL` (including the API
version path) at build time. The browser sends `VITE_APP_ID` with API calls and defaults to
`DiligenceStudio_WestMonroe`; set a distinct value for each consuming app.
The API uses headless Chromium to render imported-template previews from the same SVG model as the
canvas, so the web app must be reachable at `TEMPLATE_PREVIEW_RENDER_URL` while imports run.
Package-specific commands
can also be run with `npm run <command> --workspace @diligence-studio/web` or
`npm run <command> --workspace @diligence-studio/server`.
