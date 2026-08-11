# Repository Instructions

## Public Safety

- This is a public repository. Use only generated synthetic data and generic
  names.
- Never commit real workbooks, exports, screenshots, prompts, transcripts,
  endpoints, credentials, user paths, employer details, internal product names,
  colleague names, network locations, or copied report formulas.
- Do not copy source or history from private repositories. Reimplement generic
  behavior from public requirements and synthetic tests.
- Refer to the output style as a "dense management report".

## Delivery

- Work from `plan.md`, one numbered task at a time.
- Keep `docs/agent_handoff.md` current before stopping.
- Run relevant tests and public-safety checks before delivery.
- Completed scoped work must be committed, pushed, and present on `main`.
- Preserve unrelated work and generated local artifacts.

## Product Boundaries

- The model may propose only validated report-specification operations.
- No arbitrary formulas, VBA, COM calls, shell commands, filesystem tools,
  workbook saves, publishing, or deletion may be exposed to the model.
- The add-in may modify only explicitly owned managed objects.
- Never truncate data silently or infer a missing reporting year.
- Continuous, specific progress is required for every long-running operation.
