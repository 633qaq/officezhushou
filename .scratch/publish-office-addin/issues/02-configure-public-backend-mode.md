Status: ready-for-agent

# Configure public backend mode

## Parent

`.scratch/publish-office-addin/PRD.md`

## What to build

Make server mode point at a configurable public HTTPS backend and fail clearly when no backend URL is configured.

## Acceptance criteria

- [ ] The frontend has a runtime configuration file for the backend origin.
- [ ] Published pages do not silently default to `localhost`.
- [ ] Local development still defaults to `http://localhost:3456`.
- [ ] Server mode actions show a clear error when no backend URL is configured.

## Blocked by

None - can start immediately

