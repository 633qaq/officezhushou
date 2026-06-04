Status: ready-for-agent

# Prepare backend deployment

## Parent

`.scratch/publish-office-addin/PRD.md`

## What to build

Prepare the Node backend for deployment to a public HTTPS host with documented environment variables and safe production defaults.

## Acceptance criteria

- [ ] `.env.example` is readable and complete.
- [ ] Production deployment instructions name all required variables.
- [ ] `JWT_SECRET`, `CORS_ORIGIN`, AI provider keys, database path, and Ollama URL are documented.
- [ ] Backend smoke tests pass after configuration changes.

## Blocked by

None - can start immediately

