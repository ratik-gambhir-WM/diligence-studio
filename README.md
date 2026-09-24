# diligence-studio

The repository is an npm workspace with two explicit runtime boundaries:

```text
diligence-studio/
├── web/       React 19 and Vite browser application
└── server/    Express API, SQLite persistence, and Node PowerPoint tooling
```

The web package never imports the server or Node-only PowerPoint importer. The server owns its
runtime code, tests, database, assets, and command-line PowerPoint tools.

## Microsoft sign-in

The app uses Microsoft Entra ID for the application session. Register a confidential web application
and add this redirect URI for local development:

```text
http://localhost:43127/api/auth/callback
```

Grant delegated Microsoft Graph permissions for `User.Read`, `Files.Read`, and `Sites.Read.All`,
along with the standard OpenID scopes. Copy `.env.example` to `.env` and set the Microsoft values
before starting the API service. Set `MICROSOFT_SUPPORT=true` in both the server `.env` and
`web/.env` to enable the signed-in experience and server-side API protection. Sessions are held in
memory by the API service, so restarting it signs users out.

See [docs/microsoft-sign-in-flow-design.md](docs/microsoft-sign-in-flow-design.md) for the complete request sequence, Entra registration steps,
configuration reference, troubleshooting, and security review notes.

## Run locally

From the repository root:

```sh
npm install
npx playwright install chromium
npm run dev
npm run api
```

`npm run dev` starts the web app and proxies `/api/*` requests to the API at
`http://127.0.0.1:43127`. `npm run api` starts the API with `.env` loaded. You can also use
`npm run server:dev` when environment variables are already available in the shell.

The browser uses `/api/v1` for template, import, and export operations. A production deployment must
provide the same reverse-proxy boundary, or set the public `VITE_API_BASE_URL` at build time. The
browser sends `VITE_APP_ID` with API calls and defaults to `DiligenceStudio_WestMonroe`; set a
distinct value for each consuming app.

The API uses headless Chromium to render imported-template previews from the same SVG model as the
canvas, so the web app must be reachable at `TEMPLATE_PREVIEW_RENDER_URL` while imports run.

Run checks from the repository root:

```sh
npm run typecheck
npm test
npm run build
```

Package-specific commands can also be run with `npm run <command> --workspace
@diligence-studio/web` or `npm run <command> --workspace @diligence-studio/server`.
