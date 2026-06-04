Status: ready-for-agent

# PRD: Publishable Office Add-in

## Problem Statement

The user wants the Office AI Assistant to be publishable as a real Office add-in that can be added from Office, not just run as a local project. The current code can run locally, but publication requires a valid production manifest, HTTPS-hosted task pane assets, clear backend deployment choices, privacy and terms documents, and repeatable validation.

## Solution

Prepare the project as a publishable Office add-in with separate production and development manifests, GitHub Pages static hosting, a documented backend deployment path, publication-ready legal drafts, and validation tasks that can be run before upload or AppSource submission.

## User Stories

1. As the publisher, I want a production manifest that points to HTTPS assets, so that Office can load the add-in outside my local machine.
2. As the publisher, I want a development manifest that points to localhost, so that I can keep testing changes quickly.
3. As the publisher, I want static assets deployable to GitHub Pages, so that the Office task pane has a stable public URL.
4. As the publisher, I want icons hosted with the app, so that manifest validation and Office ribbon display do not depend on placeholder image URLs.
5. As the publisher, I want server mode to use a configurable public backend URL, so that I can deploy the backend separately from the static frontend.
6. As a user, I want direct mode to remain available, so that I can use my own AI provider credentials without relying on a shared backend.
7. As a user, I want clear errors when server mode is not configured, so that I understand why login or document saving is unavailable.
8. As the publisher, I want an environment template for backend deployment, so that a host can be configured without reverse-engineering required variables.
9. As the publisher, I want privacy and terms drafts, so that AppSource or organization deployment prerequisites are visible early.
10. As the publisher, I want a publish checklist, so that I can validate the manifest and hosted URLs before distribution.
11. As a maintainer, I want release tasks tracked as issues, so that future agents can complete work in independent slices.
12. As a maintainer, I want tests for backend smoke paths, so that publication work does not break authentication, settings, or document storage.

## Implementation Decisions

- Keep `manifest.xml` as the production manifest and keep `manifest.dev.xml` for local sideload testing.
- Host the static task pane and icon assets through GitHub Pages.
- Keep the Node backend separate from GitHub Pages because GitHub Pages cannot run the backend API.
- Add runtime frontend configuration for the deployed backend origin.
- Preserve direct mode as a supported mode, but document the API-key and CORS tradeoffs.
- Treat server mode as optional until a public HTTPS backend is deployed.
- Keep issue tracking in local markdown under `.scratch/`.

## Testing Decisions

- Test manifest validity through `office-addin-manifest validate`.
- Test backend publication smoke paths through the existing Node integration test runner.
- Test XML parseability for both production and development manifests.
- Test the published GitHub Pages URL manually after deployment.
- Test Office sideloading separately for the production and development manifests.

## Out of Scope

- AppSource submission itself.
- Legal review of privacy and terms text.
- Production backend hosting account setup.
- Payment, billing, or multi-tenant administration.
- Full cross-platform Office QA across all Office clients.

## Further Notes

The production add-in can be distributed before server mode exists, but users should understand that only direct mode is guaranteed until the backend is hosted over HTTPS.

