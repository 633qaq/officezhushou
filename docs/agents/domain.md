# Domain Docs

How the engineering skills should consume this repo's domain documentation when exploring the codebase.

## Before exploring, read these

- `CONTEXT.md` at the repo root, if it exists
- `docs/adr/`, reading ADRs relevant to the area being changed

If these files do not exist yet, proceed silently. They can be created later as the repo evolves.

## File structure

This repo uses a single-context layout:

```text
/
|- CONTEXT.md
|- docs/adr/
`- src/
```

## Use the glossary's vocabulary

When naming domain concepts in issues, plans, tests, or proposals, prefer the terms defined in `CONTEXT.md`.

## Flag ADR conflicts

If a proposed change contradicts an existing ADR, call that out explicitly instead of silently overriding it.
