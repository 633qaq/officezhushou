Status: ready-for-agent

# Validate Office sideloading

## Parent

`.scratch/publish-office-addin/PRD.md`

## What to build

Create and execute the validation path for both local development sideloading and production manifest sideloading.

## Acceptance criteria

- [ ] `manifest.dev.xml` opens the local task pane when the local server is running.
- [ ] `manifest.xml` opens the HTTPS-hosted task pane after GitHub Pages deployment.
- [ ] Word and PowerPoint ribbon commands both show the add-in button.
- [ ] Read selection, write back, copy, direct mode, and server-mode error states are checked.

## Blocked by

- `.scratch/publish-office-addin/issues/01-publish-static-taskpane.md`

