# Publish the Office Add-in

This project keeps two manifest files:

- `manifest.xml`: production manifest. It points to `https://633qaq.github.io/officezhushou/office.html` and does not include `localhost`.
- `manifest.dev.xml`: local development manifest. It points to `http://localhost:3456/office.html` for sideload testing.

Static publishing is handled by `.github/workflows/pages.yml`. The workflow copies only the files required by the Office add-in into a `dist` artifact.

## Pre-publish checklist

1. Publish the repository with GitHub Pages and confirm these files are reachable over HTTPS:
   - `office.html`
   - `js/app-config.js`
   - `js/ppt-engine.js`
   - `assets/icon-16.png`
   - `assets/icon-32.png`
   - `assets/icon-80.png`
   - `assets/icon-128.png`
   - `docs/legal/privacy.md`
   - `docs/legal/terms.md`

2. Open the task pane URL and confirm it loads:
   - `https://633qaq.github.io/officezhushou/office.html`

3. Validate the production manifest:

```powershell
npx office-addin-manifest validate manifest.xml
```

4. Choose a distribution path:
   - Personal or small-scope use: sideload `manifest.xml` in Office.
   - Internal organization use: upload `manifest.xml` through Microsoft 365 centralized deployment.
   - Public marketplace use: submit through Microsoft AppSource and provide required privacy, terms, support, and branding information.

User-facing installation and first-run instructions are in `USER_GUIDE.md`.

## Backend deployment

GitHub Pages can host the task pane UI, but it cannot run the Node.js API in `server/src`.

If all users need server mode, deploy `server` to a public HTTPS host such as Azure App Service, Render, Railway, Fly.io, or your own server. After deployment:

1. Change the frontend default server URL to your HTTPS API origin.
2. Add that HTTPS API origin to `AppDomains` in `manifest.xml`.
3. Configure `JWT_SECRET` and server-side AI provider API keys in the backend `.env`.

The frontend default server URL lives in `js/app-config.js`:

```js
window.OFFICE_ASSISTANT_CONFIG = {
  defaultServerUrl: 'https://office-assistant-api.example.com',
};
```

The backend deployment template lives in `server/.env.example`.

Without a deployed backend, users can still use direct mode. In that mode, API keys are stored in the user's local browser storage, and some model providers may be blocked by browser CORS rules.
